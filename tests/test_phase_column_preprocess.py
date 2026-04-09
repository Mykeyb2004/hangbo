from __future__ import annotations

import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

import pandas as pd

from phase_column_preprocess import (
    STATUS_ALREADY_PROCESSED,
    STATUS_FILE_MISSING,
    STATUS_INSUFFICIENT_COLUMNS,
    STATUS_MISSING_SHEET,
    STATUS_NO_PHASE_MARKER,
    STATUS_SAVE_ERROR,
    STATUS_UPDATED,
    PhaseColumnPreprocessResult,
    format_result_message,
    format_summary_message,
    is_phase_marker_value,
    main,
    preprocess_phase_column_if_needed,
    process_phase_column_workbook,
    sheet_has_phase_marker_in_third_column,
)
from survey_stats import EXHIBITOR_ROLE_NAME, excel_column_to_index


def build_mock_dataframe(role_name: str, role_column: str = "E") -> pd.DataFrame:
    column_count = excel_column_to_index("CF") + 1
    columns = [f"col_{index + 1}" for index in range(column_count)]
    rows = [[None for _ in range(column_count)] for _ in range(2)]

    rows[0][excel_column_to_index(role_column)] = role_name
    rows[1][excel_column_to_index(role_column)] = "其他身份"

    for column_name in (
        "AA",
        "AB",
        "AC",
        "AD",
        "AE",
        "AF",
        "AG",
        "AH",
        "AI",
        "AJ",
        "AK",
        "AL",
        "AM",
        "AN",
        "AO",
        "AP",
        "AQ",
        "AR",
        "AS",
        "AT",
        "AU",
        "AV",
        "AW",
        "AY",
        "AZ",
        "BA",
        "BB",
        "BC",
        "BD",
        "BE",
        "BF",
        "BG",
        "BH",
        "BI",
        "BJ",
        "BK",
        "BL",
        "BM",
        "BN",
        "BO",
        "BP",
        "BR",
        "BS",
        "BT",
        "BU",
        "BV",
        "BW",
        "BX",
        "BY",
        "BZ",
        "CA",
        "CB",
        "CC",
        "CE",
        "CF",
        "K",
        "L",
        "G",
        "H",
        "I",
        "J",
        "M",
        "N",
        "O",
        "P",
        "Q",
        "R",
        "S",
        "T",
        "U",
        "V",
        "W",
        "X",
        "Y",
        "Z",
    ):
        rows[0][excel_column_to_index(column_name)] = 9
        rows[1][excel_column_to_index(column_name)] = 1

    return pd.DataFrame(rows, columns=columns)


def build_shifted_dataframe_with_phase_column(
    role_name: str,
    role_column: str = "E",
    phase_values: tuple[str, str] = ("一期", "二期"),
) -> pd.DataFrame:
    df = build_mock_dataframe(role_name, role_column=role_column)
    df.insert(2, "phase_marker", list(phase_values))
    return df


def build_two_column_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ["row-1", "ok"],
            ["row-2", "ok"],
        ],
        columns=["col_1", "col_2"],
    )


def build_already_processed_dataframe(
    role_name: str,
    role_column: str = "E",
    phase_values: tuple[str, str] = ("一期", "二期"),
) -> pd.DataFrame:
    df = build_mock_dataframe(role_name, role_column=role_column)
    df["phase_marker"] = list(phase_values)
    return df


class PhaseColumnPreprocessTest(unittest.TestCase):
    def test_is_phase_marker_value_recognizes_supported_patterns(self) -> None:
        self.assertTrue(is_phase_marker_value("一期"))
        self.assertTrue(is_phase_marker_value(" 第三期 "))
        self.assertTrue(is_phase_marker_value("2期"))
        self.assertFalse(is_phase_marker_value("第二轮"))
        self.assertFalse(is_phase_marker_value("展览"))
        self.assertFalse(is_phase_marker_value(None))

    def test_sheet_has_phase_marker_in_third_column_detects_shifted_workbook(self) -> None:
        df = build_shifted_dataframe_with_phase_column(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "shifted.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="问卷数据", index=False)

            self.assertTrue(sheet_has_phase_marker_in_third_column(input_file, "问卷数据"))

    def test_sheet_has_phase_marker_in_third_column_returns_false_when_not_shifted(self) -> None:
        df = build_mock_dataframe(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "normal.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="问卷数据", index=False)

            self.assertFalse(sheet_has_phase_marker_in_third_column(input_file, "问卷数据"))

    def test_preprocess_phase_column_if_needed_moves_third_column_to_end_and_saves_workbook(self) -> None:
        df = build_shifted_dataframe_with_phase_column(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "shifted.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="问卷数据", index=False)

            notice = preprocess_phase_column_if_needed(input_file, "问卷数据")

            self.assertIsNotNone(notice)
            reloaded_df = pd.read_excel(input_file, sheet_name="问卷数据")
            self.assertEqual(reloaded_df.iloc[0, excel_column_to_index("E")], EXHIBITOR_ROLE_NAME)
            self.assertEqual(reloaded_df.iloc[0, reloaded_df.shape[1] - 1], "一期")
            self.assertEqual(reloaded_df.iloc[1, reloaded_df.shape[1] - 1], "二期")

    def test_preprocess_phase_column_if_needed_stays_silent_when_marker_is_only_found_in_other_column(self) -> None:
        df = build_already_processed_dataframe(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "already.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="问卷数据", index=False)

            notice = preprocess_phase_column_if_needed(input_file, "问卷数据")

            self.assertIsNone(notice)

    def test_process_phase_column_workbook_reports_missing_sheet(self) -> None:
        df = build_shifted_dataframe_with_phase_column(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "shifted.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="其他sheet", index=False)

            result = process_phase_column_workbook(input_file, "问卷数据")

            self.assertEqual(result.status, STATUS_MISSING_SHEET)
            self.assertIn("未找到 sheet", format_result_message(result))

    def test_process_phase_column_workbook_reports_insufficient_columns(self) -> None:
        df = build_two_column_dataframe()

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "short.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="问卷数据", index=False)

            result = process_phase_column_workbook(input_file, "问卷数据")

            self.assertEqual(result.status, STATUS_INSUFFICIENT_COLUMNS)
            self.assertIn("文件列数不足", format_result_message(result))

    def test_process_phase_column_workbook_reports_no_phase_marker(self) -> None:
        df = build_mock_dataframe(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "normal.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="问卷数据", index=False)

            result = process_phase_column_workbook(input_file, "问卷数据")

            self.assertEqual(result.status, STATUS_NO_PHASE_MARKER)
            self.assertIn("未发现期次特征列", format_result_message(result))

    def test_process_phase_column_workbook_reports_already_processed_when_marker_found_in_other_column(self) -> None:
        df = build_already_processed_dataframe(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "already.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="问卷数据", index=False)

            result = process_phase_column_workbook(input_file, "问卷数据")

            self.assertEqual(result.status, STATUS_ALREADY_PROCESSED)
            self.assertTrue(result.matched_markers)
            self.assertIn("可能已经处理过", format_result_message(result))
            self.assertIn("phase_marker", format_result_message(result))

    def test_format_result_message_covers_save_error(self) -> None:
        result = PhaseColumnPreprocessResult(
            path=Path("demo.xlsx"),
            sheet_name="问卷数据",
            status=STATUS_SAVE_ERROR,
            error_message="permission denied",
        )

        message = format_result_message(result)
        self.assertIn("保存失败", message)
        self.assertIn("permission denied", message)

    def test_format_summary_message_lists_result_groups(self) -> None:
        summary_message = format_summary_message(
            type(
                "Summary",
                (),
                {
                    "updated_count": 1,
                    "updated_results": (
                        PhaseColumnPreprocessResult(
                            path=Path("updated.xlsx"),
                            sheet_name="问卷数据",
                            status=STATUS_UPDATED,
                        ),
                    ),
                    "already_processed_count": 1,
                    "already_processed_results": (
                        PhaseColumnPreprocessResult(
                            path=Path("already.xlsx"),
                            sheet_name="问卷数据",
                            status=STATUS_ALREADY_PROCESSED,
                            matched_markers=("一期", "二期"),
                        ),
                    ),
                    "no_phase_marker_count": 1,
                    "no_phase_marker_results": (
                        PhaseColumnPreprocessResult(
                            path=Path("normal.xlsx"),
                            sheet_name="问卷数据",
                            status=STATUS_NO_PHASE_MARKER,
                        ),
                    ),
                    "insufficient_columns_count": 1,
                    "insufficient_columns_results": (
                        PhaseColumnPreprocessResult(
                            path=Path("short.xlsx"),
                            sheet_name="问卷数据",
                            status=STATUS_INSUFFICIENT_COLUMNS,
                        ),
                    ),
                    "failed_count": 1,
                    "failed_results": (
                        PhaseColumnPreprocessResult(
                            path=Path("missing.xlsx"),
                            sheet_name="问卷数据",
                            status=STATUS_FILE_MISSING,
                        ),
                    ),
                },
            )()
        )

        self.assertIn("成功处理（1）", summary_message)
        self.assertIn("未处理，疑似已处理过（1）", summary_message)
        self.assertIn("未处理，不含期次特征列（1）", summary_message)
        self.assertIn("未处理，列数不足（1）", summary_message)
        self.assertIn("失败（1）", summary_message)
        self.assertIn("updated.xlsx / 问卷数据", summary_message)
        self.assertIn("already.xlsx / 问卷数据", summary_message)
        self.assertIn("normal.xlsx / 问卷数据", summary_message)
        self.assertIn("short.xlsx / 问卷数据", summary_message)
        self.assertIn("missing.xlsx / 问卷数据", summary_message)
        self.assertIn("总结：共检查 5 个文件；成功处理 1 个", summary_message)

    def test_main_prints_terminal_messages_for_mixed_inputs(self) -> None:
        shifted_df = build_shifted_dataframe_with_phase_column(EXHIBITOR_ROLE_NAME)
        already_df = build_already_processed_dataframe(EXHIBITOR_ROLE_NAME)
        normal_df = build_mock_dataframe(EXHIBITOR_ROLE_NAME)
        short_df = build_two_column_dataframe()

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            shifted_file = temp_path / "shifted.xlsx"
            already_file = temp_path / "already.xlsx"
            normal_file = temp_path / "normal.xlsx"
            short_file = temp_path / "short.xlsx"
            missing_file = temp_path / "missing.xlsx"

            with pd.ExcelWriter(shifted_file, engine="openpyxl") as writer:
                shifted_df.to_excel(writer, sheet_name="问卷数据", index=False)
            with pd.ExcelWriter(already_file, engine="openpyxl") as writer:
                already_df.to_excel(writer, sheet_name="问卷数据", index=False)
            with pd.ExcelWriter(normal_file, engine="openpyxl") as writer:
                normal_df.to_excel(writer, sheet_name="问卷数据", index=False)
            with pd.ExcelWriter(short_file, engine="openpyxl") as writer:
                short_df.to_excel(writer, sheet_name="问卷数据", index=False)

            buffer = io.StringIO()
            with redirect_stdout(buffer):
                exit_code = main(
                    [
                        str(shifted_file),
                        str(already_file),
                        str(normal_file),
                        str(short_file),
                        str(missing_file),
                    ]
                )

            output = buffer.getvalue()
            self.assertEqual(exit_code, 1)
            self.assertIn("[INFO] 开始检查文件", output)
            self.assertIn("[OK] 已完成预处理", output)
            self.assertIn("可能已经处理过", output)
            self.assertIn("[INFO] 未发现期次特征列", output)
            self.assertIn("[WARN] 文件列数不足", output)
            self.assertIn("[ERROR] 文件不存在", output)
            self.assertIn("[INFO] 处理结束汇总：", output)
            self.assertTrue(output.rstrip().endswith("失败 1 个。"))
            self.assertIn("成功处理（1）", output)
            self.assertIn("未处理，疑似已处理过（1）", output)
            self.assertIn("未处理，不含期次特征列（1）", output)
            self.assertIn("未处理，列数不足（1）", output)
            self.assertIn("失败（1）", output)
            reloaded_df = pd.read_excel(shifted_file, sheet_name="问卷数据")
            self.assertEqual(reloaded_df.iloc[0, reloaded_df.shape[1] - 1], "一期")

    def test_format_result_message_covers_file_missing(self) -> None:
        result = PhaseColumnPreprocessResult(
            path=Path("missing.xlsx"),
            sheet_name="问卷数据",
            status=STATUS_FILE_MISSING,
        )

        self.assertIn("文件不存在", format_result_message(result))

    def test_process_phase_column_workbook_reports_updated_status(self) -> None:
        df = build_shifted_dataframe_with_phase_column(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "shifted.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="问卷数据", index=False)

            result = process_phase_column_workbook(input_file, "问卷数据")

            self.assertEqual(result.status, STATUS_UPDATED)
            self.assertIn("一期", result.matched_markers)


if __name__ == "__main__":
    unittest.main()
