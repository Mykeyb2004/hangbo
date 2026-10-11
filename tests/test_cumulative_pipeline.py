from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from hangbo.pipeline.cumulative import (
    PeriodResult,
    build_cumulative_result,
    is_cumulative_batch,
    resolve_period_labels,
)
from hangbo.summary.table import SUMMARY_COLUMNS, SummaryRowResult


def make_rows(values: dict[str, float | None]) -> tuple[SummaryRowResult, ...]:
    rows = []
    for name in ("甲", "乙"):
        row_values = {column: None for column in SUMMARY_COLUMNS}
        row_values.update(values.get(name, {}))
        rows.append(
            SummaryRowResult(
                category_label="一、会展客户",
                display_name=name,
                values=row_values,
                applicable_columns=frozenset(row_values),
            )
        )
    return tuple(rows)


class CumulativePipelineTest(unittest.TestCase):
    def test_batch_detection_keeps_1_2_month_as_single_period(self) -> None:
        self.assertFalse(is_cumulative_batch("1-2月"))
        self.assertTrue(is_cumulative_batch("1-9月"))
        self.assertTrue(is_cumulative_batch("1-12月"))
        self.assertTrue(is_cumulative_batch("Q1"))

    def test_period_resolution_supports_grouped_and_split_january_february(self) -> None:
        with TemporaryDirectory() as directory:
            root = Path(directory)
            for label in ("1-2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月"):
                (root / label).mkdir()
            self.assertEqual(
                resolve_period_labels("1-9月", root),
                ("1-2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月"),
            )

            (root / "1月").mkdir()
            (root / "2月").mkdir()
            (root / "1-2月").rmdir()
            self.assertEqual(
                resolve_period_labels("1-9月", root),
                ("1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月"),
            )

    def test_full_year_period_count_changes_when_january_february_is_split(self) -> None:
        with TemporaryDirectory() as directory:
            root = Path(directory)
            self.assertEqual(
                len(resolve_period_labels("1-12月", root, january_february_mode="grouped")),
                11,
            )
            self.assertEqual(
                len(resolve_period_labels("1-12月", root, january_february_mode="split")),
                12,
            )

    def test_cumulative_total_is_average_of_period_totals_and_rows_reconcile(self) -> None:
        periods = (
            PeriodResult("1-2月", make_rows({"甲": {"总分": 10.0, "产品服务": 10.0}, "乙": {"总分": 8.0, "产品服务": 8.0}})),
            PeriodResult("3月", make_rows({"甲": {"总分": 6.0, "产品服务": 6.0}, "乙": {"总分": None, "产品服务": None}})),
        )
        result = build_cumulative_result(periods)
        self.assertEqual(result.total_values["总分"], 7.5)
        self.assertEqual(result.total_values["产品服务"], 7.5)
        self.assertAlmostEqual(
            sum(
                result.row_weights[(row.display_name, "总分")] * (row.values["总分"] or 0)
                for row in result.rows
                if row.display_name
            ),
            7.5,
        )
        self.assertTrue(all(abs(value) < 1e-12 for value in result.checks.values()))


if __name__ == "__main__":
    unittest.main()
