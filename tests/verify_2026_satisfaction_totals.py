"""Read-only reconciliation of period and cumulative satisfaction scores.

Run: uv run python tests/verify_2026_satisfaction_totals.py
Only JSON evidence under logs/ is written. Production functions are not called.
The immutable category rules/templates are reused, while arithmetic is independent.
"""

from __future__ import annotations

import hashlib
import json
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils import column_index_from_string

from hangbo.survey.stats import CUSTOMER_CATEGORY_RULE_BY_NAME, TEMPLATE_DEFINITIONS


ROOT = Path(__file__).resolve().parents[1]
PERIODS = ("1-2月", *(f"{month}月" for month in range(3, 10)))
CUMULATIVE = "1-9月"
RAW = ROOT / "data/raw/2026"
RULES = sorted(
    (rule for rule in CUSTOMER_CATEGORY_RULE_BY_NAME.values() if rule.directory_enabled),
    key=lambda rule: rule.sequence_number,
)


def average(values, rounded=True):
    defined = [Decimal(str(value)) for value in values if value is not None]
    if not defined:
        return None
    result = sum(defined) / len(defined)
    if rounded:
        result = result.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return float(result)


def canonical(value):
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    try:
        return str(Decimal(text).normalize())
    except Exception:
        return text


def text_column(frame, letter):
    return frame.iloc[:, column_index_from_string(letter) - 1].fillna("").astype(str).str.strip()


def number_column(frame, letter):
    return pd.to_numeric(frame.iloc[:, column_index_from_string(letter) - 1], errors="coerce")


def calculate(frame, rule):
    if frame is None:
        return {"total": None, "unrounded_total": None, "count": 0, "sections": {}}
    if rule.aggregate_rule_names:
        components = {
            name: calculate(frame, CUSTOMER_CATEGORY_RULE_BY_NAME[name])
            for name in rule.aggregate_rule_names
        }
        sections = {}
        names = dict.fromkeys(name for item in components.values() for name in item["sections"])
        for name in names:
            matching = [item["sections"][name] for item in components.values() if name in item["sections"]]
            sections[name] = {
                "mean": average([item["mean"] for item in matching]),
                "unrounded_mean": average([item["unrounded_mean"] for item in matching], False),
                "metrics": {},
            }
        return {
            "total": average([item["total"] for item in components.values()]),
            "unrounded_total": average([item["unrounded_total"] for item in components.values()], False),
            "count": sum(item["count"] for item in components.values()),
            "sections": sections,
            "components": components,
        }
    mask = text_column(frame, rule.data_column).isin(rule.data_values)
    if rule.auxiliary_column:
        mask &= text_column(frame, rule.auxiliary_column).isin(rule.auxiliary_values)
    sections = {}
    for section in TEMPLATE_DEFINITIONS[rule.template_name].sections:
        metrics = {}
        for metric in section.metrics:
            series = number_column(frame, metric.satisfaction_column)
            lower = number_column(frame, metric.satisfaction_gt_zero_column or metric.satisfaction_column)
            upper = number_column(frame, metric.satisfaction_lt_eleven_column or metric.satisfaction_column)
            values = series[mask & lower.gt(0) & upper.lt(11)].dropna().tolist()
            metrics[metric.name] = {
                "column": metric.satisfaction_column,
                "count": len(values),
                "sum": float(sum(Decimal(str(value)) for value in values)),
                "mean": average(values),
                "unrounded_mean": average(values, False),
            }
        sections[section.name] = {
            "mean": average([item["mean"] for item in metrics.values()]),
            "unrounded_mean": average([item["unrounded_mean"] for item in metrics.values()], False),
            "metrics": metrics,
        }
    return {
        "total": average([item["mean"] for item in sections.values()]),
        "unrounded_total": average([item["unrounded_mean"] for item in sections.values()], False),
        "count": int(mask.sum()),
        "sections": sections,
    }


def read_summary(period):
    path = ROOT / "data/satisfaction_summary/2026" / period / f"{period}客户类型满意度汇总表.xlsx"
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(workbook.active.values)
    workbook.close()
    return {
        "path": str(path),
        "total": rows[-1][2],
        "type_scores": {RULES[index].name: rows[index + 2][2] for index in range(13)}
        | {RULES[index].name: rows[index + 3][2] for index in range(13, len(RULES))},
    }


def scoring_columns(filename):
    columns = {0}
    for rule in CUSTOMER_CATEGORY_RULE_BY_NAME.values():
        if rule.source_file_name != filename:
            continue
        for letter in (rule.data_column, rule.auxiliary_column):
            if letter:
                columns.add(column_index_from_string(letter) - 1)
        if rule.template_name:
            for section in TEMPLATE_DEFINITIONS[rule.template_name].sections:
                for metric in section.metrics:
                    for letter in (
                        metric.satisfaction_column,
                        metric.satisfaction_gt_zero_column,
                        metric.satisfaction_lt_eleven_column,
                    ):
                        if letter:
                            columns.add(column_index_from_string(letter) - 1)
    return sorted(columns)


def row_counter(frame, filename):
    selected = frame.iloc[:, scoring_columns(filename)].copy()
    for name in ("年份", "月份"):
        if name in frame:
            selected[name] = frame[name]
    return Counter(tuple(canonical(value) for value in row) for row in selected.itertuples(index=False, name=None))


def run():
    source_files = sorted({rule.source_file_name for rule in RULES})
    paths = sorted(
        set(RAW.glob("*/*.xlsx"))
        | set((ROOT / "data/satisfaction_summary/2026").glob("*/*.xlsx"))
        | set((ROOT / "data/satisfaction_detail/2026").glob("*/*.xlsx"))
        | set((ROOT / "src").rglob("*.py"))
        | set(ROOT.glob("*.py"))
    )
    before = {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
    frames = {}
    raw_checks = {}
    for filename in source_files:
        for period in (*PERIODS, CUMULATIVE):
            path = RAW / period / filename
            frames[period, filename] = pd.read_excel(path, sheet_name="问卷数据") if path.exists() else None
        cumulative = frames[CUMULATIVE, filename]
        combined = Counter()
        for period in PERIODS:
            frame = frames[period, filename]
            if frame is not None:
                combined.update(row_counter(frame, filename))
        actual = row_counter(cumulative, filename)
        raw_checks[filename] = {
            "cumulative_rows": len(cumulative),
            "period_rows": sum(len(frames[period, filename]) for period in PERIODS if frames[period, filename] is not None),
            "scoring_records_equal": actual == combined,
            "extra_in_cumulative": sum((actual - combined).values()),
            "missing_in_cumulative": sum((combined - actual).values()),
        }
    result = {"raw_checks": raw_checks, "periods": {}}
    for period in (*PERIODS, CUMULATIVE):
        summary = read_summary(period)
        recalculated = {
            rule.name: calculate(frames[period, rule.source_file_name], rule)
            for rule in RULES
        }
        total = average([item["total"] for item in recalculated.values()])
        summary.update({
            "valid_types": sum(isinstance(value, (int, float)) for value in summary["type_scores"].values()),
            "raw_recalculated_total": total,
            "unrounded_raw_total": average([item["unrounded_total"] for item in recalculated.values()], False),
            "raw_recalculated": recalculated,
            "type_mismatches": {
                name: {"stored": summary["type_scores"][name], "recalculated": item["total"]}
                for name, item in recalculated.items() if summary["type_scores"][name] != item["total"]
            },
        })
        result["periods"][period] = summary
    result["period_total_mean"] = average([result["periods"][period]["total"] for period in PERIODS], False)
    result["recalculated_period_total_mean"] = average([result["periods"][period]["raw_recalculated_total"] for period in PERIODS], False)
    result["mean_by_type_then_total"] = average([
        average([result["periods"][period]["type_scores"][rule.name] for period in PERIODS], False)
        for rule in RULES
    ], False)
    changed = [str(path) for path in paths if before[str(path)] != hashlib.sha256(path.read_bytes()).hexdigest()]
    result["input_hash_check"] = {"files_checked": len(paths), "changed": changed}
    assert not changed, "Input/source files changed while verifying"
    assert result["periods"][CUMULATIVE]["raw_recalculated_total"] == 9.67
    assert not result["periods"][CUMULATIVE]["type_mismatches"]
    output = ROOT / "logs/verify_2026_satisfaction_totals.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "period_total_mean": result["period_total_mean"],
        "recalculated_period_total_mean": result["recalculated_period_total_mean"],
        "mean_by_type_then_total": result["mean_by_type_then_total"],
        "raw_checks": raw_checks,
        "periods": {period: {key: item[key] for key in ("total", "valid_types", "raw_recalculated_total", "unrounded_raw_total", "type_mismatches")} for period, item in result["periods"].items()},
        "input_hash_check": result["input_hash_check"],
        "evidence_path": str(output),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    run()
