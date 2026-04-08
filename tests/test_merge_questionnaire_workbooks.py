from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from merge_questionnaire_workbooks import (
    DEFAULT_SHEET_NAME,
    merge_workbooks_by_filename,
    format_merge_summary,
)


def write_workbook(
    output_path: Path,
    headers: list[str],
    rows: list[list[object]],
    *,
    sheet_name: str = DEFAULT_SHEET_NAME,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(output_path)


class MergeQuestionnaireWorkbooksTest(unittest.TestCase):
    def test_merge_workbooks_by_filename_merges_questionnaire_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            january_dir = root / "1月"
            february_dir = root / "2月"
            output_dir = root / "合并结果"
            january_dir.mkdir()
            february_dir.mkdir()

            write_workbook(
                january_dir / "展览.xlsx",
                headers=["姓名", "得分"],
                rows=[["张三", 90], ["李四", 88]],
            )
            write_workbook(
                february_dir / "展览.xlsx",
                headers=["姓名", "得分"],
                rows=[["王五", 95]],
            )

            summary = merge_workbooks_by_filename(
                [january_dir, february_dir],
                output_dir=output_dir,
            )

            self.assertEqual(len(summary.results), 1)
            result = summary.results[0]
            self.assertEqual(result.status, "merged")
            self.assertEqual(result.merged_rows, 3)
            self.assertEqual(result.output_path, output_dir / "展览.xlsx")
            self.assertTrue(result.output_path.exists())

            worksheet = load_workbook(result.output_path)[DEFAULT_SHEET_NAME]
            self.assertEqual(worksheet.max_row, 4)
            self.assertEqual(worksheet.cell(row=1, column=1).value, "姓名")
            self.assertEqual(worksheet.cell(row=1, column=2).value, "得分")
            self.assertEqual(worksheet.cell(row=2, column=1).value, "张三")
            self.assertEqual(worksheet.cell(row=3, column=1).value, "李四")
            self.assertEqual(worksheet.cell(row=4, column=1).value, "王五")

    def test_merge_workbooks_by_filename_skips_when_columns_differ(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            first_dir = root / "目录A"
            second_dir = root / "目录B"
            output_dir = root / "合并结果"
            first_dir.mkdir()
            second_dir.mkdir()

            write_workbook(
                first_dir / "会议.xlsx",
                headers=["姓名", "得分", "城市"],
                rows=[["张三", 90, "上海"]],
            )
            write_workbook(
                second_dir / "会议.xlsx",
                headers=["姓名", "得分", "电话"],
                rows=[["李四", 88, "123456"]],
            )

            summary = merge_workbooks_by_filename(
                [first_dir, second_dir],
                output_dir=output_dir,
            )
            report = format_merge_summary(summary)

            self.assertEqual(len(summary.results), 1)
            result = summary.results[0]
            self.assertEqual(result.status, "column_mismatch")
            self.assertFalse((output_dir / "会议.xlsx").exists())
            self.assertIn("列名不一致", report)
            self.assertIn("仅 /", report)
            self.assertIn("城市", report)
            self.assertIn("电话", report)

    def test_merge_workbooks_by_filename_skips_when_sheet_missing(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            first_dir = root / "目录A"
            second_dir = root / "目录B"
            output_dir = root / "合并结果"
            first_dir.mkdir()
            second_dir.mkdir()

            write_workbook(
                first_dir / "参展商.xlsx",
                headers=["姓名", "得分"],
                rows=[["张三", 90]],
            )
            write_workbook(
                second_dir / "参展商.xlsx",
                headers=["姓名", "得分"],
                rows=[["李四", 88]],
                sheet_name="其他sheet",
            )

            summary = merge_workbooks_by_filename(
                [first_dir, second_dir],
                output_dir=output_dir,
            )
            report = format_merge_summary(summary)

            self.assertEqual(len(summary.results), 1)
            result = summary.results[0]
            self.assertEqual(result.status, "missing_sheet")
            self.assertFalse((output_dir / "参展商.xlsx").exists())
            self.assertIn("缺少 问卷数据 sheet", report)
            self.assertIn("目录B", report)


if __name__ == "__main__":
    unittest.main()
