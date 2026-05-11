from __future__ import annotations

import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path

from openpyxl import Workbook

from fill_year_month_columns import DirectoryUpdateSummary, FileUpdateResult
from merge_sample_summary import (
    BatchNameError,
    MixedSourceYearMonthError,
    SourcePreparationError,
    build_merge_sample_paths,
    discover_source_directories,
    iter_source_excel_paths,
    parse_number_selection,
    prepare_source_directories,
    validate_batch_name,
)


def write_questionnaire_workbook(
    output_path: Path,
    headers: list[str],
    rows: list[list[str]],
    sheet_name: str = "问卷数据",
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(output_path)


class MergeSampleSummaryHelpersTest(unittest.TestCase):
    def test_discover_source_directories_lists_only_direct_folders_sorted_by_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            raw_year_dir = Path(temp_dir) / "data" / "raw" / "2026"
            raw_year_dir.mkdir(parents=True)
            (raw_year_dir / "3月").mkdir()
            (raw_year_dir / "1-2月").mkdir()
            (raw_year_dir / "说明.txt").write_text("ignore", encoding="utf-8")
            (raw_year_dir / "Q1").mkdir()
            (raw_year_dir / "Q1" / "nested").mkdir()

            result = discover_source_directories(raw_year_dir)

        self.assertEqual([item.name for item in result], ["1-2月", "3月", "Q1"])

    def test_discover_source_directories_rejects_missing_year_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaisesRegex(FileNotFoundError, "年份原始数据目录不存在"):
                discover_source_directories(Path(temp_dir) / "data" / "raw" / "2026")

    def test_discover_source_directories_rejects_file_year_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            raw_year_dir = Path(temp_dir) / "data" / "raw" / "2026"
            raw_year_dir.parent.mkdir(parents=True)
            raw_year_dir.write_text("not a directory", encoding="utf-8")

            with self.assertRaisesRegex(FileNotFoundError, "年份原始数据目录不存在"):
                discover_source_directories(raw_year_dir)

    def test_discover_source_directories_rejects_empty_year_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            raw_year_dir = Path(temp_dir) / "data" / "raw" / "2026"
            raw_year_dir.mkdir(parents=True)

            with self.assertRaisesRegex(ValueError, "没有可选择的来源目录"):
                discover_source_directories(raw_year_dir)

    def test_parse_number_selection_supports_commas_and_ranges(self) -> None:
        result = parse_number_selection("1, 3-5, 2", item_count=5)

        self.assertEqual(result, (0, 2, 3, 4, 1))

    def test_parse_number_selection_deduplicates_in_first_seen_order(self) -> None:
        result = parse_number_selection("1-3,2", item_count=3)

        self.assertEqual(result, (0, 1, 2))

    def test_parse_number_selection_rejects_out_of_range_and_empty_values(self) -> None:
        with self.assertRaisesRegex(ValueError, "至少选择一个来源目录"):
            parse_number_selection("", item_count=3)
        with self.assertRaisesRegex(ValueError, "超出范围"):
            parse_number_selection("4", item_count=3)
        with self.assertRaisesRegex(ValueError, "范围起点不能大于终点"):
            parse_number_selection("3-1", item_count=3)

    def test_parse_number_selection_rejects_empty_tokens(self) -> None:
        for raw_value in ("1,,2", "1,"):
            with self.assertRaisesRegex(ValueError, "选择编号不能为空"):
                parse_number_selection(raw_value, item_count=3)

    def test_validate_batch_name_rejects_empty_separator_and_source_conflict(self) -> None:
        selected_dirs = (Path("data/raw/2026/1-2月"), Path("data/raw/2026/3月"))

        self.assertEqual(validate_batch_name(" Q1 ", selected_dirs), "Q1")
        for raw_name in ("", "  ", "../Q1", "Q1/backup", "1-2月"):
            with self.assertRaises(BatchNameError):
                validate_batch_name(raw_name, selected_dirs)

    def test_validate_batch_name_rejects_dot_only_names(self) -> None:
        for raw_name in (".", "..", "..."):
            with self.assertRaises(BatchNameError):
                validate_batch_name(raw_name, ())

    def test_build_merge_sample_paths_uses_existing_directory_contract(self) -> None:
        paths = build_merge_sample_paths(
            year="2026",
            batch_name="Q1",
            data_root=Path("data"),
        )

        self.assertEqual(paths.raw_year_dir, Path("data/raw/2026"))
        self.assertEqual(paths.merged_raw_dir, Path("data/raw/2026/Q1"))
        self.assertEqual(paths.sample_summary_dir, Path("data/sample_summary/2026/Q1"))
        self.assertEqual(
            paths.sample_summary_path,
            Path("data/sample_summary/2026/Q1/Q1客户类型样本统计表.xlsx"),
        )

    def test_build_merge_sample_paths_strips_year_and_batch_name(self) -> None:
        paths = build_merge_sample_paths(
            year=" 2026 ",
            batch_name=" Q1 ",
            data_root=Path("data"),
        )

        self.assertEqual(paths.year, "2026")
        self.assertEqual(paths.batch_name, "Q1")
        self.assertEqual(paths.raw_year_dir, Path("data/raw/2026"))
        self.assertEqual(paths.merged_raw_dir, Path("data/raw/2026/Q1"))

    def test_build_merge_sample_paths_accepts_string_data_root(self) -> None:
        paths = build_merge_sample_paths(
            year="2026",
            batch_name="Q1",
            data_root="data",
        )

        self.assertEqual(paths.data_root, Path("data"))
        self.assertEqual(paths.raw_year_dir, Path("data/raw/2026"))


class MergeSampleSummaryPreparationTest(unittest.TestCase):
    def test_iter_source_excel_paths_lists_only_direct_normal_xlsx_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source_dir = Path(temp_dir) / "1-2月"
            source_dir.mkdir()
            normal_workbook = source_dir / "展览.xlsx"
            write_questionnaire_workbook(normal_workbook, ["姓名", "年份", "月份"], [["张三", "2026", "1"]])
            write_questionnaire_workbook(source_dir / "~$展览.xlsx", ["姓名"], [["张三"]])
            write_questionnaire_workbook(source_dir / "._展览.xlsx", ["姓名"], [["张三"]])
            nested_dir = source_dir / "nested"
            nested_dir.mkdir()
            write_questionnaire_workbook(nested_dir / "会议.xlsx", ["姓名"], [["张三"]])

            result = iter_source_excel_paths(source_dir)

        self.assertEqual(result, (normal_workbook,))

    def test_prepare_source_directories_autofills_only_single_month_dirs(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            raw_year_dir = Path(temp_dir)
            month_dir = raw_year_dir / "3月"
            mixed_dir = raw_year_dir / "1-2月"
            month_dir.mkdir()
            mixed_dir.mkdir()
            write_questionnaire_workbook(
                mixed_dir / "展览.xlsx",
                ["姓名", "年份", "月份"],
                [["张三", "2026", "1"]],
            )

            with patch("merge_sample_summary.apply_year_month_to_directory") as apply_year_month:
                apply_year_month.return_value = DirectoryUpdateSummary(input_dir=month_dir, file_results=())
                prepare_source_directories(
                    (month_dir, mixed_dir),
                    year="2026",
                    sheet_name="问卷数据",
                )

        apply_year_month.assert_called_once_with(
            month_dir,
            year="2026",
            month="3",
            sheet_name="问卷数据",
        )

    def test_prepare_source_directories_blocks_mixed_dir_missing_year_month_headers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            mixed_dir = Path(temp_dir) / "1-2月"
            mixed_dir.mkdir()
            write_questionnaire_workbook(
                mixed_dir / "展览.xlsx",
                ["姓名", "月份"],
                [["张三", "1"]],
            )

            with self.assertRaisesRegex(MixedSourceYearMonthError, "缺少“年份”/“月份”列"):
                prepare_source_directories(
                    (mixed_dir,),
                    year="2026",
                    sheet_name="问卷数据",
                )

    def test_prepare_source_directories_blocks_single_month_autofill_failures(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            month_dir = Path(temp_dir) / "3月"
            month_dir.mkdir()
            skipped_path = month_dir / "展览.xlsx"
            summary = DirectoryUpdateSummary(
                input_dir=month_dir,
                file_results=(
                    FileUpdateResult(
                        path=skipped_path,
                        status="missing_sheet",
                        updated_rows=0,
                    ),
                ),
            )

            with patch("merge_sample_summary.apply_year_month_to_directory") as apply_year_month:
                apply_year_month.return_value = summary
                with self.assertRaises(SourcePreparationError) as error_context:
                    prepare_source_directories(
                        (month_dir,),
                        year="2026",
                        sheet_name="问卷数据",
                    )

        error_message = str(error_context.exception)
        self.assertIn(str(month_dir), error_message)
        self.assertIn("展览.xlsx", error_message)
        self.assertIn("missing_sheet", error_message)

    def test_prepare_source_directories_blocks_mixed_dir_without_usable_excel_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            mixed_dir = Path(temp_dir) / "1-2月"
            mixed_dir.mkdir()
            write_questionnaire_workbook(
                mixed_dir / "~$展览.xlsx",
                ["姓名", "年份", "月份"],
                [["张三", "2026", "1"]],
            )

            with self.assertRaisesRegex(MixedSourceYearMonthError, "没有可用的 Excel 文件"):
                prepare_source_directories(
                    (mixed_dir,),
                    year="2026",
                    sheet_name="问卷数据",
                )


if __name__ == "__main__":
    unittest.main()
