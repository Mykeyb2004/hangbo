"""Regenerate reports in isolation and compare before/after workbook cell values.

Capture before the change, then verify after it:
    uv run python tests/verify_summary_algorithm_regression.py capture
    uv run python tests/verify_summary_algorithm_regression.py verify

Production Excel files are only read. Evidence is written under logs/verification.
"""

import argparse
from contextlib import redirect_stdout
import hashlib
import io
import json
from pathlib import Path
import shutil
from tempfile import TemporaryDirectory
import tomllib

from openpyxl import load_workbook

from hangbo.pipeline.cumulative import run_cumulative_batch
from hangbo.summary.table import SUMMARY_COLUMNS, build_summary_rows, load_report_snapshots
from hangbo.survey.stats import run_directory_batch
from hangbo.summary.table import generate_summary_report


ROOT = Path(__file__).resolve().parents[1]
LABELS = ("1-2月", *(f"{month}月" for month in range(3, 10)))
BASELINE = ROOT / "logs/verification/summary-algorithm-before.json"
REPORT = ROOT / "logs/verification/summary-algorithm-regression.json"
CONFIG = tomllib.loads((ROOT / "pipeline.defaults.toml").read_text())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def snapshot(path):
    workbook = load_workbook(path, data_only=False)
    try:
        return {sheet.title: {cell.coordinate: cell.value for row in sheet for cell in row if cell.value is not None}
                for sheet in workbook}
    finally:
        workbook.close()


def detail_snapshots(directory):
    return {path.name: snapshot(path) for path in sorted(directory.glob("*.xlsx"))}


def saved_month(label):
    return {"details": detail_snapshots(ROOT / "data/satisfaction_detail/2026" / label),
            "summary": snapshot(ROOT / "data/satisfaction_summary/2026" / label / f"{label}客户类型满意度汇总表.xlsx")}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("capture", "verify"))
    mode = parser.parse_args().mode
    if mode == "capture" and BASELINE.exists():
        raise SystemExit(f"Baseline already exists; preserve it: {BASELINE}")
    baseline = json.loads(BASELINE.read_text()) if mode == "verify" else None
    protected = {str(path.relative_to(ROOT)): digest(path)
                 for label in (*LABELS, "Q3", "1-9月")
                 for category in ("raw", "satisfaction_detail", "satisfaction_summary")
                 for path in (ROOT / "data" / category / "2026" / label).glob("*.xlsx")}
    if baseline:
        assert protected == baseline["protected_excel_hashes"], "Production Excel changed since capture"
    results = {}
    cross_table_counts = {}
    with TemporaryDirectory(prefix="hangbo-summary-regression-") as directory:
        work = Path(directory)
        for label in LABELS:
            raw = work / "raw" / label
            shutil.copytree(ROOT / "data/raw/2026" / label, raw)
            detail = work / "detail" / label
            with redirect_stdout(io.StringIO()):
                run_directory_batch(input_dir=raw, output_dir=detail,
                                    sheet_name=CONFIG["sheet_name"], calculation_mode=CONFIG["calculation_mode"])
            summary = generate_summary_report(input_dir=detail, output_dir=work / "summary" / label,
                                             output_name=f"{label}客户类型满意度汇总表.xlsx")
            results[label] = {"details": detail_snapshots(detail), "summary": snapshot(summary)}
            # Existing single-month files, as well as the before-change engine, are controls.
            assert results[label] == saved_month(label), (label, "regenerated monthly values differ from existing files")
            if baseline:
                assert results[label] == baseline["monthly"][label], (label, "monthly values changed")
            print(f"{label}: {len(results[label]['details'])} detail workbooks and summary unchanged", flush=True)

        cumulative = {}
        for label in ("Q3", "1-9月"):
            detail = work / "cumulative" / label / "detail"
            summary = work / "cumulative" / label / "summary.xlsx"
            with redirect_stdout(io.StringIO()):
                result = run_cumulative_batch(
                    year="2026", batch=label, data_root=ROOT / "data",
                    detail_output_dir=detail, summary_output_path=summary,
                    sheet_name=CONFIG["sheet_name"], calculation_mode=CONFIG["calculation_mode"],
                    january_february_mode=CONFIG["january_february_mode"],
                    log_path=work / "cumulative" / label / "cumulative.json",
                )
            current = {"details": detail_snapshots(detail), "total_values": result.total_values}
            cumulative[label] = current
            if baseline:
                assert current == baseline["cumulative"][label], (label, "cumulative details or period totals changed")
            rebuilt = build_summary_rows(load_report_snapshots(detail))
            saved = snapshot(summary)["汇总表"]
            mismatches = []
            for row_index, row in enumerate(rebuilt, 3):
                for col_index, indicator in enumerate(SUMMARY_COLUMNS, 3):
                    cell = f"{chr(64 + col_index)}{row_index}"
                    actual = saved.get(cell)
                    actual = actual if isinstance(actual, (int, float)) else None
                    if actual != row.values[indicator]:
                        mismatches.append({"cell": cell, "indicator": indicator, "actual": actual,
                                           "from_detail": row.values[indicator]})
            cross_table_counts[label] = {"compared_cells": len(rebuilt) * len(SUMMARY_COLUMNS),
                                         "mismatches": mismatches}
            if baseline:
                assert not mismatches, (label, mismatches)
            print(f"{label}: {len(current['details'])} detail workbooks; {len(mismatches)} cross-table differences", flush=True)
    assert all(digest(ROOT / path) == expected for path, expected in protected.items()), "Production Excel was modified"
    if mode == "capture":
        BASELINE.parent.mkdir(parents=True, exist_ok=True)
        BASELINE.write_text(json.dumps({"monthly": results, "cumulative": cumulative,
                                        "cross_table": cross_table_counts, "protected_excel_hashes": protected},
                                       ensure_ascii=False, indent=2) + "\n")
        print(f"Baseline saved: {BASELINE}")
    else:
        report = {"monthly_labels": list(LABELS), "monthly_values_unchanged": True,
                  "monthly_detail_workbooks": sum(len(r["details"]) for r in results.values()),
                  "cumulative_details_unchanged": True, "cumulative_period_totals_unchanged": True,
                  "cross_table": cross_table_counts, "protected_excel_files_unchanged": len(protected)}
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
        print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
