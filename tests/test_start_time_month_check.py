from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from check_start_time_month import (
    DEFAULT_FIELD_NAME,
    DEFAULT_SHEET_NAME,
    analyze_directory,
    format_directory_summary,
    summarize_workbook,
)


def write_workbook(
    output_path: Path,
    start_times: list[str | None] | None = None,
    *,
    include_target_sheet: bool = True,
    include_field: bool = True,
    other_sheet_with_field: bool = False,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active

    if include_target_sheet:
        worksheet.title = DEFAULT_SHEET_NAME
        if include_field:
            worksheet.append([DEFAULT_FIELD_NAME, "其他字段"])
            for value in start_times or []:
                worksheet.append([value, "x"])
        else:
            worksheet.append(["别的字段", "其他字段"])
            for value in start_times or []:
                worksheet.append([value, "x"])
    else:
        worksheet.title = "其他sheet"
        worksheet.append([DEFAULT_FIELD_NAME, "其他字段"])
        for value in start_times or []:
            worksheet.append([value, "x"])

    if other_sheet_with_field:
        extra_sheet = workbook.create_sheet("附加sheet")
        extra_sheet.append([DEFAULT_FIELD_NAME, "其他字段"])
        extra_sheet.append(["2026-02-20 12:00:00", "y"])

    workbook.save(output_path)


class StartTimeMonthCheckTest(unittest.TestCase):
    def test_analyze_directory_reports_shared_month_and_ignores_hidden_excel_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir)
            write_workbook(
                input_dir / "展览-2.xlsx",
                ["2026-02-11 17:36:50", "2026-02-15 09:12:10"],
            )
            write_workbook(
                input_dir / "会议-2.xlsx",
                ["2026-02-28 16:54:26", None, "2026-02-01 08:00:00"],
            )
            (input_dir / "._会议-2.xlsx").write_text("not an excel file", encoding="utf-8")

            summary = analyze_directory(input_dir)

            self.assertTrue(summary.all_valid_values_in_one_month)
            self.assertTrue(summary.all_files_in_one_month)
            self.assertEqual(summary.detected_months, ("2026-02",))
            self.assertEqual([item.path.name for item in summary.file_summaries], ["会议-2.xlsx", "展览-2.xlsx"])

    def test_analyze_directory_marks_cross_month_missing_field_and_empty_values(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir)
            write_workbook(
                input_dir / "餐饮-4.xlsx",
                ["2026-01-31 23:59:59", "2026-02-01 00:00:00"],
            )
            write_workbook(
                input_dir / "酒店住宿.xlsx",
                ["2026-02-10 10:00:00"],
                include_field=False,
            )
            write_workbook(
                input_dir / "空数据.xlsx",
                [None, None],
            )

            summary = analyze_directory(input_dir)
            report = format_directory_summary(summary)

            self.assertFalse(summary.all_valid_values_in_one_month)
            self.assertFalse(summary.all_files_in_one_month)
            self.assertEqual(summary.detected_months, ("2026-01", "2026-02"))

            by_name = {item.path.name: item for item in summary.file_summaries}
            self.assertEqual(by_name["酒店住宿.xlsx"].status, "missing_field")
            self.assertEqual(by_name["空数据.xlsx"].status, "no_valid_values")
            self.assertEqual(by_name["餐饮-4.xlsx"].months, ("2026-01", "2026-02"))

            self.assertIn("2026-01, 2026-02", report)
            self.assertIn("跨月", report)
            self.assertIn("字段缺失", report)
            self.assertIn("无有效时间值", report)

    def test_summarize_workbook_reports_missing_target_sheet_even_if_other_sheet_has_field(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "8月酒店过程分析.xlsx"
            write_workbook(
                input_path,
                ["2025-08-24 16:37:29"],
                include_target_sheet=False,
                other_sheet_with_field=True,
            )

            summary = summarize_workbook(input_path)

            self.assertEqual(summary.status, "missing_sheet")
            self.assertEqual(summary.months, ())
            self.assertEqual(summary.valid_count, 0)


if __name__ == "__main__":
    unittest.main()
