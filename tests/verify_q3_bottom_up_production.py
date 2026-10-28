"""Independently audit saved Q3 Excel/PPT and unchanged original monthly files.

Run: uv run python tests/verify_q3_bottom_up_production.py
Historical pre-fix evidence in 2026-Q3-indicator-hierarchy-audit.json is preserved.
"""

from collections import Counter
from decimal import Decimal
from fractions import Fraction
import json
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation

from audit_q3_indicator_hierarchy import (
    MONTHS, ROOT, decimal, inspect_hierarchy, mean, normalized, read_detail, rounded,
)
from hangbo.pipeline.cumulative import calculate_period
from hangbo.summary.table import SUMMARY_COLUMNS, SUMMARY_ROW_DEFINITIONS
from hangbo.survey.customer_category_rules import DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES
from verify_summary_algorithm_regression import BASELINE, LABELS, digest


def main():
    detail_dir = ROOT / "data/satisfaction_detail/2026/Q3"
    reports = [read_detail(path) for path in sorted(detail_dir.glob("*.xlsx"))]
    hierarchy = [inspect_hierarchy(report) for report in reports]
    counts = {"detail_workbooks": len(reports), "satisfaction_rows": sum(len(r["records"]) for r in reports)}
    for report in hierarchy:
        assert report["total"] == report["section_mean"], report
        assert all(section["matches_children"] for section in report["sections"]), report
    counts["customer_section_means"] = len(hierarchy)
    counts["section_child_means"] = sum(len(report["sections"]) for report in hierarchy)
    for report in reports:
        for record in report["records"]:
            value = decimal(record["value"])
            if value is not None:
                assert value == rounded(value) == Decimal(f"{float(value):.2f}"), record

    # Independently reproduce leaves from raw monthly calculations and the
    # documented period/customer weights, without calling cumulative aggregation.
    periods = [calculate_period(ROOT / "data/raw/2026" / label, label,
                                sheet_name="问卷数据", calculation_mode="template") for label in MONTHS]
    leaves = []
    for report in hierarchy:
        role = Path(report["file"]).stem
        for section in report["sections"]:
            for child in section["children"]:
                terms = []
                for period in periods:
                    stats = period.reports.get(role)
                    valid_count = sum(row.display_name is not None and row.values["总分"] is not None
                                      for row in period.rows)
                    if stats is None or stats.satisfaction is None:
                        continue
                    source_section = next((s for s in stats.sections if s.name == section["name"]), None)
                    metric = next((m for m in source_section.metrics if m.name == child["name"]), None) if source_section else None
                    if metric and metric.satisfaction is not None:
                        terms.append((Fraction(str(metric.satisfaction)), Fraction(1, len(periods) * valid_count)))
                average = sum(value * weight for value, weight in terms) / sum(weight for _, weight in terms) if terms else None
                expected = rounded(Decimal(average.numerator) / Decimal(average.denominator)) if average is not None else None
                assert decimal(child["value"]) == expected, (role, child, expected)
                leaves.append({"role": role, "cell": child["cell"], "value": child["value"], "expected": expected})
    counts["raw_monthly_leaf_reproductions"] = len(leaves)

    workbook = load_workbook(ROOT / "data/satisfaction_summary/2026/Q3/Q3客户类型满意度汇总表.xlsx", data_only=True)
    cross_table = []
    bottom = []
    try:
        sheet = workbook["汇总表"]
        for index, definition in enumerate(SUMMARY_ROW_DEFINITIONS, 3):
            aliases = {normalized(alias) for alias in definition.source_aliases}
            matched = [r for r in reports if normalized(Path(r["file"]).stem) in aliases
                       or normalized(r["records"][0]["name"]) in aliases]
            for col, indicator in enumerate(SUMMARY_COLUMNS, 3):
                sources = []
                for report in matched:
                    for selector in definition.selectors.get(indicator, ()):
                        candidates = [r for r in report["records"] if r["kind"] == selector.source_type
                                      and (selector.source_type == "overall" or r["name"] == selector.source_name)
                                      and decimal(r["value"]) is not None]
                        if candidates:
                            sources.append(candidates[0]["value"])
                            break
                actual = decimal(sheet.cell(index, col).value)
                assert actual == rounded(mean(sources)), (definition.display_name, indicator, actual, sources)
                cross_table.append({"customer": definition.display_name, "indicator": indicator,
                                    "cell": sheet.cell(index, col).coordinate, "value": actual})
        for col, indicator in enumerate(SUMMARY_COLUMNS, 3):
            expected = rounded(mean(row["value"] for row in cross_table if row["indicator"] == indicator and row["customer"]))
            assert decimal(sheet.cell(22, col).value) == expected, indicator
            bottom.append({"indicator": indicator, "value": expected})
        assert sheet["C22"].value == 9.73
        assert sheet["F8"].value == 9.86
    finally:
        workbook.close()
    counts["summary_detail_positions"] = len(cross_table)
    counts["summary_bottom_means"] = len(bottom)

    # Compare every numeric satisfaction table row in the actual saved PPT.
    presentation = Presentation(ROOT / "data/ppt/2026/Q3/Q3满意度报告.pptx")
    ppt_checks = []
    covered = set()
    for index, slide in enumerate(presentation.slides, 1):
        table_rows = [tuple(cell.text for cell in row.cells) for shape in slide.shapes if shape.has_table for row in shape.table.rows]
        if not table_rows:
            continue
        report, = [r for r in reports if any(row[0] == r["records"][0]["name"] for row in table_rows)]
        expected = Counter((r["name"], decimal(r["value"])) for r in report["records"] if decimal(r["value"]) is not None)
        actual = Counter()
        for row in table_rows:
            try:
                score = Decimal(row[1])
            except Exception:
                continue
            actual[(row[0], score)] += 1
        assert actual == expected, (index, report["file"], actual - expected, expected - actual)
        covered.add(report["file"])
        ppt_checks.append({"slide": index, "detail": report["file"], "numeric_rows": sum(actual.values())})
    assert len(covered) == len(reports), (covered, len(reports))
    counts["ppt_customer_tables"] = len(ppt_checks)
    counts["ppt_numeric_satisfaction_rows"] = sum(item["numeric_rows"] for item in ppt_checks)

    baseline = json.loads(BASELINE.read_text())
    protected = {name: expected for name, expected in baseline["protected_excel_hashes"].items()
                 if Path(name).parts[3] in LABELS or Path(name).parts[:4] == ("data", "raw", "2026", "Q3")}
    assert all(digest(ROOT / name) == expected for name, expected in protected.items()), "Original monthly/raw Excel changed"
    counts["original_monthly_and_q3_raw_files_unchanged"] = len(protected)
    payload = {"counts": counts, "mismatches": 0, "overall_satisfaction": 9.73,
               "cross_table": cross_table, "bottom": bottom, "leaf_reproduction": leaves, "ppt": ppt_checks}
    output = ROOT / "logs/verification/q3-bottom-up-production.json"
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str) + "\n")
    print(json.dumps({"counts": counts, "mismatches": 0, "overall_satisfaction": 9.73}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
