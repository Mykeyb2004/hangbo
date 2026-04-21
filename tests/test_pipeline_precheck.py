from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook

from pipeline_paths import STANDARD_SOURCE_FILE_NAMES, build_pipeline_paths
from pipeline_precheck import run_precheck


def write_workbook(
    path: Path,
    headers: list[str],
    sheet_name: str = "问卷数据",
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(headers)
    workbook.save(path)


def write_standard_workbooks(
    raw_dir: Path,
    headers: list[str],
    sheet_name: str = "问卷数据",
) -> None:
    for file_name in STANDARD_SOURCE_FILE_NAMES:
        write_workbook(raw_dir / file_name, headers, sheet_name=sheet_name)


class PipelinePrecheckTest(unittest.TestCase):
    def test_missing_raw_dir_returns_blocking_issue(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )

            result = run_precheck(paths, sheet_name="问卷数据", single_month=3)

            self.assertEqual(len(result.blocking_issues), 1)
            self.assertEqual(result.blocking_issues[0].code, "missing_raw_dir")
            self.assertIn("原始批次目录不存在", result.blocking_issues[0].message)
            self.assertFalse(result.warning_issues)
            self.assertFalse(result.should_autofill_year_month)

    def test_missing_standard_sources_returns_aggregate_and_per_file_issues(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            paths.raw_dir.mkdir(parents=True)

            result = run_precheck(paths, sheet_name="问卷数据", single_month=3)

            issue_codes = [issue.code for issue in result.blocking_issues]
            self.assertEqual(issue_codes.count("missing_standard_sources"), 1)
            self.assertEqual(
                issue_codes.count("missing_source_file"),
                len(STANDARD_SOURCE_FILE_NAMES),
            )
            self.assertFalse(result.warning_issues)
            self.assertFalse(result.should_autofill_year_month)

    def test_single_month_missing_year_month_columns_warns_and_enables_autofill(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            write_standard_workbooks(paths.raw_dir, ["客户", "满意度"])

            with patch(
                "pipeline_precheck.run_unmapped_audit",
                return_value=(0, paths.unmapped_log_path),
            ):
                result = run_precheck(paths, sheet_name="问卷数据", single_month=3)

            self.assertFalse(result.blocking_issues)
            self.assertEqual(len(result.warning_issues), 1)
            self.assertEqual(result.warning_issues[0].code, "autofill_year_month")
            self.assertTrue(result.should_autofill_year_month)

    def test_combined_batch_missing_year_month_columns_blocks_without_autofill(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "Q1",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            write_standard_workbooks(paths.raw_dir, ["客户", "满意度"])

            result = run_precheck(paths, sheet_name="问卷数据", single_month=None)

            self.assertEqual(len(result.blocking_issues), 1)
            self.assertEqual(result.blocking_issues[0].code, "missing_year_month_columns")
            self.assertIn("缺少“年份”/“月份”列", result.blocking_issues[0].message)
            self.assertFalse(result.warning_issues)
            self.assertFalse(result.should_autofill_year_month)

    def test_combined_batch_missing_year_month_skips_unmapped_audit(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "Q1",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            write_standard_workbooks(paths.raw_dir, ["客户", "满意度"])

            with patch("pipeline_precheck.run_unmapped_audit") as audit_mock:
                result = run_precheck(paths, sheet_name="问卷数据", single_month=None)

            self.assertTrue(result.blocking_issues)
            audit_mock.assert_not_called()

    def test_missing_sheet_returns_blocking_issue(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            write_standard_workbooks(
                paths.raw_dir,
                ["年份", "月份", "客户"],
                sheet_name="其他",
            )

            result = run_precheck(paths, sheet_name="问卷数据", single_month=3)

            self.assertEqual(len(result.blocking_issues), 1)
            self.assertEqual(result.blocking_issues[0].code, "missing_sheet")
            self.assertIn("缺少 sheet", result.blocking_issues[0].message)
            self.assertFalse(result.should_autofill_year_month)

    def test_missing_sheet_skips_unmapped_audit(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            write_standard_workbooks(
                paths.raw_dir,
                ["年份", "月份", "客户"],
                sheet_name="其他",
            )

            with patch("pipeline_precheck.run_unmapped_audit") as audit_mock:
                result = run_precheck(paths, sheet_name="问卷数据", single_month=3)

            self.assertTrue(result.blocking_issues)
            audit_mock.assert_not_called()

    def test_corrupt_workbook_returns_precheck_error_instead_of_crashing(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            write_standard_workbooks(paths.raw_dir, ["年份", "月份", "客户"])
            corrupt_path = paths.raw_dir / STANDARD_SOURCE_FILE_NAMES[0]
            corrupt_path.write_text("not an xlsx file", encoding="utf-8")

            with patch("pipeline_precheck.run_unmapped_audit") as audit_mock:
                result = run_precheck(paths, sheet_name="问卷数据", single_month=3)

            self.assertEqual(len(result.blocking_issues), 1)
            self.assertEqual(result.blocking_issues[0].code, "precheck_error")
            self.assertIn("预查错过程失败", result.blocking_issues[0].message)
            self.assertIn(corrupt_path.name, result.blocking_issues[0].message)
            audit_mock.assert_not_called()

    def test_unmapped_records_audit_returns_blocking_issue(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            write_standard_workbooks(paths.raw_dir, ["年份", "月份", "客户"])

            with patch(
                "pipeline_precheck.run_unmapped_audit",
                return_value=(2, paths.unmapped_log_path),
            ):
                result = run_precheck(paths, sheet_name="问卷数据", single_month=3)

            self.assertEqual(len(result.blocking_issues), 1)
            self.assertEqual(result.blocking_issues[0].code, "unmapped_customer_records")
            self.assertIn("未映射标签", result.blocking_issues[0].message)
            self.assertFalse(result.warning_issues)
            self.assertFalse(result.should_autofill_year_month)

    def test_unmapped_audit_exception_returns_precheck_error(self) -> None:
        with TemporaryDirectory() as temp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(temp_dir) / "data",
                logs_root=Path(temp_dir) / "logs",
            )
            write_standard_workbooks(paths.raw_dir, ["年份", "月份", "客户"])

            with patch(
                "pipeline_precheck.run_unmapped_audit",
                side_effect=RuntimeError("boom"),
            ):
                result = run_precheck(paths, sheet_name="问卷数据", single_month=3)

            self.assertEqual(len(result.blocking_issues), 1)
            self.assertEqual(result.blocking_issues[0].code, "precheck_error")
            self.assertIn("预查错过程失败", result.blocking_issues[0].message)
            self.assertFalse(result.warning_issues)
            self.assertFalse(result.should_autofill_year_month)


if __name__ == "__main__":
    unittest.main()
