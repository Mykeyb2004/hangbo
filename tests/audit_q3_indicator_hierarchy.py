"""Historical pre-fix audit of Q3 hierarchy and presentation rounding.

Run: uv run python tests/audit_q3_indicator_hierarchy.py
Writes evidence under logs/verification; never saves an Excel or PowerPoint file.
Requires pre-fix reports: its monthly-parent reproduction assertion intentionally
describes the old algorithm. For current acceptance use verify_q3_bottom_up_production.py.
"""

from decimal import Decimal, ROUND_HALF_UP
import hashlib
import json
from pathlib import Path
import subprocess
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook

from hangbo.pipeline.cumulative import build_cumulative_result, calculate_period
from hangbo.summary.table import SUMMARY_COLUMNS, SUMMARY_ROW_DEFINITIONS, row_kind
from hangbo.survey.customer_category_rules import DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES


ROOT = Path(__file__).resolve().parents[1]
MONTHS = ("7月", "8月", "9月")
OUTPUT = ROOT / "logs/verification/2026-Q3-indicator-hierarchy-audit.json"


def decimal(value):
    return Decimal(str(value)) if isinstance(value, (int, float, Decimal)) else None


def rounded(value):
    value = decimal(value)
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if value is not None else None


def mean(values):
    defined = [decimal(value) for value in values if decimal(value) is not None]
    return sum(defined) / len(defined) if defined else None


def normalized(value):
    return str(value).strip().replace(" ", "").replace("（", "(").replace("）", ")").replace("／", "/")


def read_detail(path):
    workbook = load_workbook(path, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        records = [
            {"cell": f"B{index}", "name": sheet.cell(index, 1).value,
             "kind": row_kind(sheet, index, 1), "value": sheet.cell(index, 2).value,
             "number_format": sheet.cell(index, 2).number_format}
            for index in range(2, sheet.max_row + 1)
            if sheet.cell(index, 1).value is not None
        ]
        return {"file": str(path.relative_to(ROOT)), "sheet": sheet.title, "records": records}
    finally:
        workbook.close()


def inspect_hierarchy(report):
    records = report["records"]
    sections = []
    for record in records:
        if record["kind"] == "section":
            sections.append({**record, "children": []})
        elif record["kind"] == "metric" and sections:
            sections[-1]["children"].append(record)
    for section in sections:
        section["parent_rounded"] = rounded(section["value"])
        section["mean_of_raw_children"] = rounded(mean(c["value"] for c in section["children"]))
        section["mean_of_rounded_children"] = rounded(mean(rounded(c["value"]) for c in section["children"]))
        section["difference"] = (
            section["parent_rounded"] - section["mean_of_rounded_children"]
            if section["parent_rounded"] is not None and section["mean_of_rounded_children"] is not None else None
        )
        section["matches_children"] = section["parent_rounded"] == section["mean_of_rounded_children"]
    overall = next(record for record in records if record["kind"] == "overall")
    return {"file": report["file"], "total": rounded(overall["value"]),
            "section_mean": rounded(mean(s["parent_rounded"] for s in sections)), "sections": sections}


def main():
    protected = {
        path: hashlib.sha256(path.read_bytes()).hexdigest()
        for label in (*MONTHS, "Q3")
        for category in ("raw", "satisfaction_detail", "satisfaction_summary", "ppt")
        for path in (ROOT / "data" / category / "2026" / label).glob("*")
        if path.suffix in {".xlsx", ".pptx"}
    }
    directory = ROOT / "data/satisfaction_detail/2026/Q3"
    reports = [read_detail(path) for path in sorted(directory.glob("*.xlsx"))]
    hierarchy = [inspect_hierarchy(report) for report in reports]
    cross_table = []
    summary_path = ROOT / "data/satisfaction_summary/2026/Q3/Q3客户类型满意度汇总表.xlsx"
    workbook = load_workbook(summary_path, data_only=True)
    try:
        sheet = workbook["汇总表"]
        for row_index, definition in enumerate(SUMMARY_ROW_DEFINITIONS, 3):
            aliases = {normalized(alias) for alias in definition.source_aliases}
            matched = [report for report in reports
                       if normalized(Path(report["file"]).stem) in aliases
                       or normalized(report["records"][0]["name"]) in aliases]
            for col_index, indicator in enumerate(SUMMARY_COLUMNS, 3):
                selectors = definition.selectors.get(indicator, ())
                sources = []
                for report in matched:
                    for selector in selectors:
                        candidates = [record for record in report["records"]
                                      if record["kind"] == selector.source_type
                                      and (selector.source_type == "overall" or record["name"] == selector.source_name)
                                      and decimal(record["value"]) is not None]
                        if candidates:
                            sources.append({"file": report["file"], "sheet": report["sheet"], **candidates[0]})
                            break
                expected = rounded(mean(rounded(source["value"]) for source in sources))
                cell = sheet.cell(row_index, col_index)
                actual = decimal(cell.value)
                cross_table.append({"customer": definition.display_name, "indicator": indicator,
                                    "summary_cell": cell.coordinate, "summary_value": actual,
                                    "expected_from_detail": expected, "sources": sources,
                                    "matches": actual == expected})
        bottom = []
        for col_index, indicator in enumerate(SUMMARY_COLUMNS, 3):
            expected = rounded(mean(item["summary_value"] for item in cross_table
                                    if item["indicator"] == indicator and item["customer"]))
            actual = decimal(sheet.cell(22, col_index).value)
            bottom.append({"indicator": indicator, "cell": sheet.cell(22, col_index).coordinate,
                           "actual": actual, "customer_mean": expected, "matches": actual == expected})
    finally:
        workbook.close()

    rounding_differences = []
    for report in reports:
        for record in report["records"]:
            if decimal(record["value"]) is not None and rounded(record["value"]) != Decimal(f"{record['value']:.2f}"):
                rounding_differences.append({"file": report["file"], **record,
                                             "decimal_half_up": rounded(record["value"]),
                                             "python_float_format": f"{record['value']:.2f}"})

    # Reproduce the stored quarter sections/children from independently recomputed months.
    periods = tuple(calculate_period(ROOT / "data/raw/2026" / label, label,
                                     sheet_name="问卷数据", calculation_mode="template") for label in MONTHS)
    result = build_cumulative_result(periods)
    display_names = {rule.name: rule.customer_category for rule in DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES}
    reproduction = []
    for report in hierarchy:
        role = Path(report["file"]).stem
        terms = [(period, period.reports[role],
                  result.period_row_weights[(period.label, display_names[role], "总分")])
                 for period in periods if role in period.reports and period.reports[role].satisfaction is not None]
        for section in report["sections"]:
            source = []
            for period, stats, weight in terms:
                monthly = next((item for item in stats.sections if item.name == section["name"]), None)
                if monthly is not None:
                    source.append({"period": period.label, "weight": weight, "parent": monthly.satisfaction,
                                   "children": {metric.name: metric.satisfaction for metric in monthly.metrics}})
            section["monthly_sources"] = source
            support_sets = [tuple(item["period"] for item in source if item["children"].get(child["name"]) is not None)
                            for child in section["children"] if child["value"] is not None]
            section["different_month_support_for_children"] = len(set(support_sets)) > 1
            section["monthly_parent_already_differs_from_children"] = [
                item["period"] for item in source
                if rounded(item["parent"]) != rounded(mean(item["children"].values()))
            ]
            for record in [section, *section["children"]]:
                defined = [(item["parent"] if record is section else item["children"].get(record["name"]), item["weight"])
                           for item in source]
                defined = [(decimal(value), decimal(weight)) for value, weight in defined if value is not None]
                expected = sum(value * weight for value, weight in defined) / sum(weight for _, weight in defined) if defined else None
                actual = decimal(record["value"])
                matches = actual is None and expected is None or (
                    actual is not None and expected is not None and abs(actual - expected) < Decimal("0.000000000001"))
                reproduction.append({"file": report["file"], "cell": record["cell"], "name": record["name"],
                                     "actual": actual, "monthly_weighted_reproduction": expected, "matches": matches})

    # Read the actual saved PowerPoint tables, rather than infer their formatting.
    ppt_path = ROOT / "data/ppt/2026/Q3/Q3满意度报告.pptx"
    ppt_evidence = []
    if ppt_path.exists():
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        with ZipFile(ppt_path) as archive:
            for name in archive.namelist():
                if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                    slide = ElementTree.fromstring(archive.read(name))
                    titles = [element.text for element in slide.findall(".//a:t", ns)]
                    if any("参会客户" in (title or "") or "住宿团队" in (title or "") for title in titles):
                        rows = [["".join(cell.itertext()) for cell in row.findall("a:tc", ns)]
                                for row in slide.findall(".//a:tr", ns)]
                        rows = [row for row in rows if row and row[0] in {"配套服务", "餐饮服务", "标识标牌清晰"}]
                        if rows:
                            ppt_evidence.append({"slide": name, "title": titles[0], "rows": rows})

    monthly_hierarchy = {}
    for label in MONTHS:
        checks = [inspect_hierarchy(read_detail(path))
                  for path in sorted((ROOT / "data/satisfaction_detail/2026" / label).glob("*.xlsx"))]
        monthly_hierarchy[label] = [{"file": report["file"], "cell": section["cell"], "name": section["name"],
                                     "actual": section["parent_rounded"], "children_mean": section["mean_of_rounded_children"]}
                                    for report in checks for section in report["sections"] if not section["matches_children"]]

    mismatches = [{"file": report["file"], **section} for report in hierarchy
                  for section in report["sections"] if not section["matches_children"]]
    protected_changes = [str(path.relative_to(ROOT)) for path, old_hash in protected.items()
                         if hashlib.sha256(path.read_bytes()).hexdigest() != old_hash]
    counts = {
        "detail_workbooks": len(reports), "satisfaction_rows": sum(len(report["records"]) for report in reports),
        "cross_table_positions": len(cross_table), "cross_table_numeric_positions": sum(item["summary_value"] is not None for item in cross_table),
        "cross_table_mismatches": sum(not item["matches"] for item in cross_table),
        "total_to_section_checks": len(hierarchy), "total_to_section_mismatches": sum(item["total"] != item["section_mean"] for item in hierarchy),
        "bottom_mean_checks": len(bottom), "bottom_mean_mismatches": sum(not item["matches"] for item in bottom),
        "first_level_sections": sum(len(item["sections"]) for item in hierarchy),
        "valid_first_level_sections": sum(section["value"] is not None for item in hierarchy for section in item["sections"]),
        "section_to_children_mismatches": len(mismatches), "affected_detail_workbooks": len({item["file"] for item in mismatches}),
        "difference_exactly_001": sum(abs(item["difference"]) == Decimal("0.01") for item in mismatches),
        "difference_greater_than_001": sum(abs(item["difference"]) > Decimal("0.01") for item in mismatches),
        "formatting_differences": len(rounding_differences),
        "monthly_reproduction_checks": len(reproduction), "monthly_reproduction_mismatches": sum(not item["matches"] for item in reproduction),
        "protected_files": len(protected), "protected_files_changed": len(protected_changes),
    }
    payload = {"commit": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT).decode().strip(),
               "scope": "satisfaction column; blank values excluded; Decimal half-up to two decimal places",
               "counts": counts, "cross_table": cross_table, "bottom_means": bottom, "hierarchy": hierarchy,
               "section_child_mismatches": mismatches, "rounding_differences": rounding_differences,
               "monthly_reproduction": reproduction, "monthly_existing_section_mismatches": monthly_hierarchy,
               "ppt_evidence": ppt_evidence, "protected_hashes": {str(p.relative_to(ROOT)): h for p, h in protected.items()},
               "protected_changes": protected_changes}
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str) + "\n")
    print(json.dumps(counts, ensure_ascii=False, indent=2))
    print("Mismatches (first-level score / mean of rounded children):")
    for item in mismatches:
        print(Path(item["file"]).name, item["cell"], item["name"], item["parent_rounded"], item["mean_of_rounded_children"],
              "different_month_support=", item["different_month_support_for_children"],
              "monthly_parent_mismatches=", item["monthly_parent_already_differs_from_children"])
    assert not protected_changes, "Audit changed production files"
    assert all(item["matches"] for item in reproduction), "Monthly sources do not explain saved quarter values"


if __name__ == "__main__":
    main()
