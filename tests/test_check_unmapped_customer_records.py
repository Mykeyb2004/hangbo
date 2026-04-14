from __future__ import annotations

import argparse
import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

import pandas as pd

from check_unmapped_customer_records import (
    DEFAULT_LOG_DIR,
    DEFAULT_SHEET_NAME,
    audit_source_file,
    format_directory_audit_report,
    run_audit_command,
    run_directory_audit,
    write_audit_log,
)


def build_meeting_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "A": ["row1", "row2", "row3"],
            "B": ["x", "y", "z"],
            "C": ["会议", "", "会议"],
            "D": ["unused", "unused", "unused"],
            "E": ["会议主承办", "参会人员", "酒店参会客户"],
        }
    )


def build_hotel_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "A": ["row1", "row2", "row3"],
            "B": ["x", "y", "z"],
            "C": ["散客", "未知客群", "住宿团队"],
        }
    )


class CheckUnmappedCustomerRecordsTest(unittest.TestCase):
    def test_audit_source_file_reports_unmapped_combination_and_blank_auxiliary(self) -> None:
        audit = audit_source_file(
            source_file_name="会议.xlsx",
            workbook_path=Path("/tmp/会议.xlsx"),
            sheet_name=DEFAULT_SHEET_NAME,
            df=build_meeting_dataframe(),
        )

        self.assertEqual(audit.data_column, "E")
        self.assertEqual(audit.auxiliary_column, "C")
        self.assertEqual(len(audit.unmapped_records), 2)

        first_record = audit.unmapped_records[0]
        self.assertEqual(first_record.excel_row_number, 3)
        self.assertEqual(first_record.auxiliary_value, "")
        self.assertEqual(first_record.data_value, "参会人员")
        self.assertIn("辅助标签为空", first_record.reason)

        second_record = audit.unmapped_records[1]
        self.assertEqual(second_record.excel_row_number, 4)
        self.assertEqual(second_record.auxiliary_value, "会议")
        self.assertEqual(second_record.data_value, "酒店参会客户")
        self.assertIn("未映射问题", second_record.reason)

    def test_audit_source_file_reports_unmapped_value_when_no_auxiliary_column(self) -> None:
        audit = audit_source_file(
            source_file_name="酒店.xlsx",
            workbook_path=Path("/tmp/酒店.xlsx"),
            sheet_name=DEFAULT_SHEET_NAME,
            df=build_hotel_dataframe(),
        )

        self.assertEqual(audit.data_column, "C")
        self.assertIsNone(audit.auxiliary_column)
        self.assertEqual(len(audit.unmapped_records), 1)

        record = audit.unmapped_records[0]
        self.assertEqual(record.excel_row_number, 3)
        self.assertEqual(record.data_value, "未知客群")
        self.assertIn("未映射问题", record.reason)

    def test_run_audit_command_prints_report_and_writes_log_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            data_dir = temp_path / "datas"
            log_dir = temp_path / DEFAULT_LOG_DIR
            data_dir.mkdir()

            meeting_path = data_dir / "会议.xlsx"
            hotel_path = data_dir / "酒店.xlsx"
            with pd.ExcelWriter(meeting_path, engine="openpyxl") as writer:
                build_meeting_dataframe().to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)
            with pd.ExcelWriter(hotel_path, engine="openpyxl") as writer:
                build_hotel_dataframe().to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            report = run_directory_audit(data_dir, sheet_name=DEFAULT_SHEET_NAME)
            log_path = log_dir / "audit.log"
            report_text = format_directory_audit_report(report, log_path=log_path)
            write_audit_log(report_text, log_path)

            self.assertTrue(log_path.exists())
            logged_text = log_path.read_text(encoding="utf-8")
            self.assertIn("未映射记录数: 3", logged_text)
            self.assertIn("会议.xlsx", logged_text)
            self.assertIn("酒店.xlsx", logged_text)
            self.assertIn("缺少来源文件", logged_text)
            self.assertIn("会展服务商.xlsx", logged_text)

            buffer = io.StringIO()
            with redirect_stdout(buffer):
                run_audit_command(
                    argparse.Namespace(
                        input_dir=data_dir,
                        sheet_name=DEFAULT_SHEET_NAME,
                        log_dir=log_dir,
                        log_file=log_path,
                    )
                )

            output = buffer.getvalue()
            self.assertIn("客户映射核查结果", output)
            self.assertIn("日志文件", output)
            self.assertIn("行 3", output)
            self.assertIn("未知客群", output)


if __name__ == "__main__":
    unittest.main()
