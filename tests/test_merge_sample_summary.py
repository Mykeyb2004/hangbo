from __future__ import annotations

import tempfile
import unittest
import curses
from unittest.mock import patch
from pathlib import Path

from openpyxl import Workbook

from fill_year_month_columns import DirectoryUpdateSummary, FileUpdateResult
from merge_questionnaire_workbooks import MergeResult, MergeSummary
from merge_sample_summary import (
    BatchNameError,
    MergeSampleRunConfig,
    MixedSourceYearMonthError,
    SourcePreparationError,
    build_merge_sample_paths,
    clear_generated_outputs,
    confirm_overwrite_if_needed,
    discover_source_directories,
    iter_source_excel_paths,
    parse_args,
    parse_number_selection,
    prompt_batch_name,
    prepare_source_directories,
    run_merge_sample_summary,
    select_directories,
    select_directories_by_number_prompt,
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
                apply_year_month.return_value = DirectoryUpdateSummary(
                    input_dir=month_dir,
                    file_results=(
                        FileUpdateResult(
                            path=month_dir / "展览.xlsx",
                            status="updated",
                            updated_rows=2,
                        ),
                    ),
                )
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

    def test_prepare_source_directories_blocks_empty_single_month_dir(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            month_dir = Path(temp_dir) / "3月"
            month_dir.mkdir()
            summary = DirectoryUpdateSummary(input_dir=month_dir, file_results=())

            with patch("merge_sample_summary.apply_year_month_to_directory") as apply_year_month:
                apply_year_month.return_value = summary
                with self.assertRaisesRegex(SourcePreparationError, "没有可用的 Excel 文件"):
                    prepare_source_directories(
                        (month_dir,),
                        year="2026",
                        sheet_name="问卷数据",
                    )

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


class MergeSampleSummaryRunTest(unittest.TestCase):
    def test_run_merge_sample_summary_merges_and_generates_only_sample_summary(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir) / "data"
            source_dirs = (
                data_root / "raw" / "2026" / "1-2月",
                data_root / "raw" / "2026" / "3月",
            )
            sample_config_path = Path(temp_dir) / "sample_table_config.json"
            expected_paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=data_root,
            )
            merge_output_dirs: list[Path] = []
            sample_output_dirs: list[Path] = []

            def merge_side_effect(
                input_dirs: tuple[Path, ...],
                *,
                output_dir: Path,
                sheet_name: str,
            ) -> MergeSummary:
                merge_output_dirs.append(output_dir)
                output_dir.mkdir(parents=True, exist_ok=True)
                temp_output_path = output_dir / "展览.xlsx"
                temp_output_path.write_text("new merged", encoding="utf-8")
                return MergeSummary(
                    input_dirs=input_dirs,
                    output_dir=output_dir,
                    results=(
                        MergeResult(
                            file_name="展览.xlsx",
                            source_paths=source_dirs,
                            status="merged",
                            merged_rows=3,
                            output_path=temp_output_path,
                        ),
                    ),
                )

            def generate_sample_side_effect(
                *,
                input_dir: Path,
                output_dir: Path,
                output_name: str,
                config_path: Path,
                source_sheet_name: str,
                default_year: str,
            ) -> Path:
                self.assertEqual(input_dir, expected_paths.merged_raw_dir)
                self.assertEqual(config_path, sample_config_path)
                self.assertEqual(source_sheet_name, "问卷数据")
                self.assertEqual(default_year, "2026")
                sample_output_dirs.append(output_dir)
                output_dir.mkdir(parents=True, exist_ok=True)
                temp_sample_path = output_dir / output_name
                temp_sample_path.write_text("new sample", encoding="utf-8")
                return temp_sample_path

            with (
                patch("merge_sample_summary.prepare_source_directories") as prepare_sources,
                patch("merge_sample_summary.merge_workbooks_by_filename") as merge_workbooks,
                patch("merge_sample_summary.generate_sample_table_report") as generate_sample_table,
            ):
                merge_workbooks.side_effect = merge_side_effect
                generate_sample_table.side_effect = generate_sample_side_effect

                result = run_merge_sample_summary(
                    MergeSampleRunConfig(
                        year="2026",
                        batch_name="Q1",
                        selected_dirs=source_dirs,
                        data_root=data_root,
                        sheet_name="问卷数据",
                        sample_config_path=sample_config_path,
                    )
                )

            prepare_sources.assert_called_once_with(
                source_dirs,
                year="2026",
                sheet_name="问卷数据",
            )
            self.assertEqual(len(merge_output_dirs), 1)
            self.assertNotEqual(merge_output_dirs[0], expected_paths.merged_raw_dir)
            self.assertEqual(len(sample_output_dirs), 1)
            self.assertNotEqual(sample_output_dirs[0], expected_paths.sample_summary_dir)
            self.assertEqual(
                generate_sample_table.call_args.kwargs["output_name"],
                expected_paths.sample_summary_path.name,
            )
            self.assertEqual(result.paths, expected_paths)
            self.assertEqual(result.merge_summary.output_dir, expected_paths.merged_raw_dir)
            self.assertEqual(
                result.merge_summary.results[0].output_path,
                expected_paths.merged_raw_dir / "展览.xlsx",
            )
            self.assertEqual(result.sample_summary_path, expected_paths.sample_summary_path)
            self.assertEqual(
                (expected_paths.merged_raw_dir / "展览.xlsx").read_text(encoding="utf-8"),
                "new merged",
            )
            self.assertEqual(
                expected_paths.sample_summary_path.read_text(encoding="utf-8"),
                "new sample",
            )

    def test_run_merge_sample_summary_stops_when_any_merge_result_failed(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir) / "data"
            source_dirs = (data_root / "raw" / "2026" / "1-2月",)
            config_path = Path(temp_dir) / "sample_table_config.json"
            expected_paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=data_root,
            )
            merge_output_dirs: list[Path] = []

            def merge_side_effect(
                input_dirs: tuple[Path, ...],
                *,
                output_dir: Path,
                sheet_name: str,
            ) -> MergeSummary:
                merge_output_dirs.append(output_dir)
                return MergeSummary(
                    input_dirs=input_dirs,
                    output_dir=output_dir,
                    results=(
                        MergeResult(
                            file_name="展览.xlsx",
                            source_paths=source_dirs,
                            status="missing_sheet",
                            missing_sheet_paths=source_dirs,
                        ),
                    ),
                )

            with (
                patch("merge_sample_summary.prepare_source_directories") as prepare_sources,
                patch("merge_sample_summary.merge_workbooks_by_filename") as merge_workbooks,
                patch("merge_sample_summary.generate_sample_table_report") as generate_sample_table,
            ):
                merge_workbooks.side_effect = merge_side_effect

                with self.assertRaisesRegex(RuntimeError, "合并阶段存在失败项") as error_context:
                    run_merge_sample_summary(
                        MergeSampleRunConfig(
                            year="2026",
                            batch_name="Q1",
                            selected_dirs=source_dirs,
                            data_root=data_root,
                            sheet_name="问卷数据",
                            sample_config_path=config_path,
                        )
                    )

        prepare_sources.assert_called_once_with(
            source_dirs,
            year="2026",
            sheet_name="问卷数据",
        )
        self.assertEqual(len(merge_output_dirs), 1)
        self.assertNotEqual(merge_output_dirs[0], expected_paths.merged_raw_dir)
        generate_sample_table.assert_not_called()
        self.assertIn("跳过/失败: 1", str(error_context.exception))
        self.assertIn(str(expected_paths.merged_raw_dir), str(error_context.exception))
        self.assertNotIn(str(merge_output_dirs[0]), str(error_context.exception))

    def test_run_merge_sample_summary_stops_when_merge_finds_no_results(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir) / "data"
            source_dirs = (data_root / "raw" / "2026" / "1-2月",)
            config_path = Path(temp_dir) / "sample_table_config.json"
            expected_paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=data_root,
            )
            merge_output_dirs: list[Path] = []

            def merge_side_effect(
                input_dirs: tuple[Path, ...],
                *,
                output_dir: Path,
                sheet_name: str,
            ) -> MergeSummary:
                merge_output_dirs.append(output_dir)
                return MergeSummary(
                    input_dirs=input_dirs,
                    output_dir=output_dir,
                    results=(),
                )

            with (
                patch("merge_sample_summary.prepare_source_directories") as prepare_sources,
                patch("merge_sample_summary.merge_workbooks_by_filename") as merge_workbooks,
                patch("merge_sample_summary.generate_sample_table_report") as generate_sample_table,
            ):
                merge_workbooks.side_effect = merge_side_effect

                with self.assertRaisesRegex(RuntimeError, "合并阶段存在失败项") as error_context:
                    run_merge_sample_summary(
                        MergeSampleRunConfig(
                            year="2026",
                            batch_name="Q1",
                            selected_dirs=source_dirs,
                            data_root=data_root,
                            sheet_name="问卷数据",
                            sample_config_path=config_path,
                        )
                    )

        prepare_sources.assert_called_once_with(
            source_dirs,
            year="2026",
            sheet_name="问卷数据",
        )
        self.assertEqual(len(merge_output_dirs), 1)
        self.assertNotEqual(merge_output_dirs[0], expected_paths.merged_raw_dir)
        generate_sample_table.assert_not_called()
        self.assertIn("未找到可处理的 xlsx 文件", str(error_context.exception))
        self.assertIn(str(expected_paths.merged_raw_dir), str(error_context.exception))
        self.assertNotIn(str(merge_output_dirs[0]), str(error_context.exception))

    def test_run_merge_sample_summary_keeps_existing_outputs_when_source_prep_fails(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir) / "data"
            source_dirs = (data_root / "raw" / "2026" / "1-2月",)
            config_path = Path(temp_dir) / "sample_table_config.json"
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=data_root,
            )
            paths.merged_raw_dir.mkdir(parents=True)
            existing_workbook = paths.merged_raw_dir / "展览.xlsx"
            existing_workbook.write_text("old merged", encoding="utf-8")
            paths.sample_summary_path.parent.mkdir(parents=True)
            paths.sample_summary_path.write_text("old sample", encoding="utf-8")

            with (
                patch(
                    "merge_sample_summary.prepare_source_directories",
                    side_effect=SourcePreparationError("prep failed"),
                ) as prepare_sources,
                patch("merge_sample_summary.merge_workbooks_by_filename") as merge_workbooks,
                patch("merge_sample_summary.generate_sample_table_report") as generate_sample_table,
            ):
                with self.assertRaisesRegex(SourcePreparationError, "prep failed"):
                    run_merge_sample_summary(
                        MergeSampleRunConfig(
                            year="2026",
                            batch_name="Q1",
                            selected_dirs=source_dirs,
                            data_root=data_root,
                            sheet_name="问卷数据",
                            sample_config_path=config_path,
                            overwrite=True,
                        )
                    )

            prepare_sources.assert_called_once_with(
                source_dirs,
                year="2026",
                sheet_name="问卷数据",
            )
            merge_workbooks.assert_not_called()
            generate_sample_table.assert_not_called()
            self.assertTrue(existing_workbook.exists())
            self.assertEqual(existing_workbook.read_text(encoding="utf-8"), "old merged")
            self.assertTrue(paths.sample_summary_path.exists())
            self.assertEqual(paths.sample_summary_path.read_text(encoding="utf-8"), "old sample")

    def test_run_merge_sample_summary_keeps_existing_sample_summary_when_merge_fails(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir) / "data"
            source_dirs = (data_root / "raw" / "2026" / "1-2月",)
            config_path = Path(temp_dir) / "sample_table_config.json"
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=data_root,
            )
            paths.merged_raw_dir.mkdir(parents=True)
            existing_workbook = paths.merged_raw_dir / "展览.xlsx"
            existing_workbook.write_text("old merged", encoding="utf-8")
            paths.sample_summary_path.parent.mkdir(parents=True)
            paths.sample_summary_path.write_text("old sample", encoding="utf-8")
            merge_output_dirs: list[Path] = []

            def merge_side_effect(
                input_dirs: tuple[Path, ...],
                *,
                output_dir: Path,
                sheet_name: str,
            ) -> MergeSummary:
                merge_output_dirs.append(output_dir)
                return MergeSummary(
                    input_dirs=input_dirs,
                    output_dir=output_dir,
                    results=(
                        MergeResult(
                            file_name="展览.xlsx",
                            source_paths=source_dirs,
                            status="missing_sheet",
                            missing_sheet_paths=source_dirs,
                        ),
                    ),
                )

            with (
                patch("merge_sample_summary.prepare_source_directories"),
                patch("merge_sample_summary.merge_workbooks_by_filename") as merge_workbooks,
                patch("merge_sample_summary.generate_sample_table_report") as generate_sample_table,
            ):
                merge_workbooks.side_effect = merge_side_effect

                with self.assertRaisesRegex(RuntimeError, "合并阶段存在失败项"):
                    run_merge_sample_summary(
                        MergeSampleRunConfig(
                            year="2026",
                            batch_name="Q1",
                            selected_dirs=source_dirs,
                            data_root=data_root,
                            sheet_name="问卷数据",
                            sample_config_path=config_path,
                            overwrite=True,
                        )
                    )

            self.assertEqual(len(merge_output_dirs), 1)
            self.assertNotEqual(merge_output_dirs[0], paths.merged_raw_dir)
            generate_sample_table.assert_not_called()
            self.assertTrue(paths.sample_summary_path.exists())
            self.assertEqual(paths.sample_summary_path.read_text(encoding="utf-8"), "old sample")

    def test_run_merge_sample_summary_keeps_existing_raw_workbooks_when_merge_fails_after_partial_temp_success(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir) / "data"
            source_dirs = (data_root / "raw" / "2026" / "1-2月",)
            config_path = Path(temp_dir) / "sample_table_config.json"
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=data_root,
            )
            paths.merged_raw_dir.mkdir(parents=True)
            existing_merged = paths.merged_raw_dir / "展览.xlsx"
            existing_other = paths.merged_raw_dir / "会议.xlsx"
            existing_merged.write_text("old merged", encoding="utf-8")
            existing_other.write_text("old other", encoding="utf-8")
            merge_output_dirs: list[Path] = []

            def merge_side_effect(
                input_dirs: tuple[Path, ...],
                *,
                output_dir: Path,
                sheet_name: str,
            ) -> MergeSummary:
                merge_output_dirs.append(output_dir)
                output_dir.mkdir(parents=True, exist_ok=True)
                temp_output_path = output_dir / "展览.xlsx"
                temp_output_path.write_text("temp merged", encoding="utf-8")
                return MergeSummary(
                    input_dirs=input_dirs,
                    output_dir=output_dir,
                    results=(
                        MergeResult(
                            file_name="展览.xlsx",
                            source_paths=source_dirs,
                            status="merged",
                            merged_rows=3,
                            output_path=temp_output_path,
                        ),
                        MergeResult(
                            file_name="会议.xlsx",
                            source_paths=source_dirs,
                            status="missing_sheet",
                            missing_sheet_paths=source_dirs,
                        ),
                    ),
                )

            with (
                patch("merge_sample_summary.prepare_source_directories"),
                patch("merge_sample_summary.merge_workbooks_by_filename") as merge_workbooks,
                patch("merge_sample_summary.generate_sample_table_report") as generate_sample_table,
            ):
                merge_workbooks.side_effect = merge_side_effect

                with self.assertRaisesRegex(RuntimeError, "合并阶段存在失败项"):
                    run_merge_sample_summary(
                        MergeSampleRunConfig(
                            year="2026",
                            batch_name="Q1",
                            selected_dirs=source_dirs,
                            data_root=data_root,
                            sheet_name="问卷数据",
                            sample_config_path=config_path,
                            overwrite=True,
                        )
                    )

            self.assertEqual(len(merge_output_dirs), 1)
            self.assertNotEqual(merge_output_dirs[0], paths.merged_raw_dir)
            generate_sample_table.assert_not_called()
            self.assertTrue(existing_merged.exists())
            self.assertEqual(existing_merged.read_text(encoding="utf-8"), "old merged")
            self.assertTrue(existing_other.exists())
            self.assertEqual(existing_other.read_text(encoding="utf-8"), "old other")

    def test_run_merge_sample_summary_keeps_existing_sample_summary_when_sample_generation_fails(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir) / "data"
            source_dirs = (data_root / "raw" / "2026" / "1-2月",)
            config_path = Path(temp_dir) / "sample_table_config.json"
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=data_root,
            )
            paths.sample_summary_path.parent.mkdir(parents=True)
            paths.sample_summary_path.write_text("old sample", encoding="utf-8")
            sample_output_dirs: list[Path] = []

            def merge_side_effect(
                input_dirs: tuple[Path, ...],
                *,
                output_dir: Path,
                sheet_name: str,
            ) -> MergeSummary:
                output_dir.mkdir(parents=True, exist_ok=True)
                temp_output_path = output_dir / "展览.xlsx"
                temp_output_path.write_text("new merged", encoding="utf-8")
                return MergeSummary(
                    input_dirs=input_dirs,
                    output_dir=output_dir,
                    results=(
                        MergeResult(
                            file_name="展览.xlsx",
                            source_paths=source_dirs,
                            status="merged",
                            merged_rows=3,
                            output_path=temp_output_path,
                        ),
                    ),
                )

            def generate_sample_side_effect(
                *,
                input_dir: Path,
                output_dir: Path,
                output_name: str,
                config_path: Path,
                source_sheet_name: str,
                default_year: str,
            ) -> Path:
                sample_output_dirs.append(output_dir)
                raise RuntimeError("sample failed")

            with (
                patch("merge_sample_summary.prepare_source_directories"),
                patch("merge_sample_summary.merge_workbooks_by_filename") as merge_workbooks,
                patch("merge_sample_summary.generate_sample_table_report") as generate_sample_table,
            ):
                merge_workbooks.side_effect = merge_side_effect
                generate_sample_table.side_effect = generate_sample_side_effect

                with self.assertRaisesRegex(RuntimeError, "sample failed"):
                    run_merge_sample_summary(
                        MergeSampleRunConfig(
                            year="2026",
                            batch_name="Q1",
                            selected_dirs=source_dirs,
                            data_root=data_root,
                            sheet_name="问卷数据",
                            sample_config_path=config_path,
                            overwrite=True,
                        )
                    )

            self.assertEqual(len(sample_output_dirs), 1)
            self.assertNotEqual(sample_output_dirs[0], paths.sample_summary_dir)
            self.assertTrue(paths.sample_summary_path.exists())
            self.assertEqual(paths.sample_summary_path.read_text(encoding="utf-8"), "old sample")

    def test_clear_generated_outputs_removes_only_generated_workbooks_and_sample_summary(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=Path(temp_dir) / "data",
            )
            paths.merged_raw_dir.mkdir(parents=True)
            generated_workbook = paths.merged_raw_dir / "展览.xlsx"
            keep_text = paths.merged_raw_dir / "keep.txt"
            generated_workbook.write_text("generated", encoding="utf-8")
            keep_text.write_text("keep", encoding="utf-8")
            paths.sample_summary_path.parent.mkdir(parents=True)
            paths.sample_summary_path.write_text("summary", encoding="utf-8")

            clear_generated_outputs(paths)

            self.assertFalse(generated_workbook.exists())
            self.assertTrue(keep_text.exists())
            self.assertFalse(paths.sample_summary_path.exists())


class MergeSampleSummaryInteractionTest(unittest.TestCase):
    def test_select_directories_by_number_prompt_reprompts_until_valid_selection(self) -> None:
        source_dirs = (Path("1-2月"), Path("3月"), Path("4月"))
        inputs = iter(("bad", "1,3"))
        outputs: list[str] = []

        result = select_directories_by_number_prompt(
            source_dirs,
            input_func=lambda _: next(inputs),
            output_func=outputs.append,
        )

        self.assertEqual(result, (Path("1-2月"), Path("4月")))
        self.assertTrue(any("选择无效" in line for line in outputs))

    def test_confirm_overwrite_if_needed_returns_true_when_outputs_do_not_exist(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=Path(temp_dir) / "data",
            )

            result = confirm_overwrite_if_needed(
                paths,
                input_func=lambda _: self.fail("should not prompt without existing outputs"),
                output_func=lambda _: self.fail("should not print without existing outputs"),
            )

        self.assertTrue(result)

    def test_confirm_overwrite_if_needed_respects_user_answer_when_outputs_exist(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=Path(temp_dir) / "data",
            )
            paths.merged_raw_dir.mkdir(parents=True)
            paths.sample_summary_path.parent.mkdir(parents=True)
            paths.sample_summary_path.write_text("existing", encoding="utf-8")
            outputs: list[str] = []

            denied = confirm_overwrite_if_needed(
                paths,
                input_func=lambda _: "n",
                output_func=outputs.append,
            )
            accepted = confirm_overwrite_if_needed(
                paths,
                input_func=lambda _: "yes",
                output_func=outputs.append,
            )

        self.assertFalse(denied)
        self.assertTrue(accepted)
        self.assertTrue(any("已存在" in line for line in outputs))

    def test_prompt_batch_name_reprompts_after_invalid_input(self) -> None:
        selected_dirs = (Path("1-2月"), Path("3月"))
        inputs = iter(("1-2月", " Q1 "))
        outputs: list[str] = []

        result = prompt_batch_name(
            selected_dirs,
            input_func=lambda _: next(inputs),
            output_func=outputs.append,
        )

        self.assertEqual(result, "Q1")
        self.assertTrue(any("批次名称无效" in line for line in outputs))

    def test_select_directories_falls_back_to_numbered_prompt_on_curses_error(self) -> None:
        source_dirs = (Path("1-2月"), Path("3月"))
        outputs: list[str] = []

        with (
            patch(
                "merge_sample_summary.select_directories_with_curses",
                side_effect=curses.error("no terminal"),
            ),
            patch(
                "merge_sample_summary.select_directories_by_number_prompt",
                return_value=(Path("3月"),),
            ) as numbered_prompt,
        ):
            result = select_directories(source_dirs, output_func=outputs.append)

        self.assertEqual(result, (Path("3月"),))
        numbered_prompt.assert_called_once_with(source_dirs, output_func=outputs.append)
        self.assertTrue(any("降级为编号选择" in line for line in outputs))

    def test_parse_args_accepts_required_year_and_default_config(self) -> None:
        args = parse_args(["--year", "2026"])

        self.assertEqual(args.year, "2026")
        self.assertEqual(args.config, Path("pipeline.defaults.toml"))


if __name__ == "__main__":
    unittest.main()
