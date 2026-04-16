from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from sample_table import (
    DEFAULT_SAMPLE_TABLE_TITLE,
    SAMPLE_TABLE_SHEET_NAME,
    build_sample_table_rows,
    generate_sample_table_report,
    load_sample_table_config,
    parse_sample_group_specs,
)


def build_source_dataframe(rows: list[dict[str, object]]) -> pd.DataFrame:
    columns = ["A", "B", "C", "D", "E", "F", "年份", "月份"]
    normalized_rows: list[list[object]] = []
    for row in rows:
        normalized_rows.append([row.get(column_name, "") for column_name in columns])
    return pd.DataFrame(normalized_rows, columns=columns)


def write_source_workbook(
    output_path: Path,
    rows: list[dict[str, object]],
) -> None:
    dataframe = build_source_dataframe(rows)
    dataframe.to_excel(output_path, sheet_name="问卷数据", index=False)


class SampleTableTest(unittest.TestCase):
    def test_load_sample_table_config_reads_default_targets(self) -> None:
        config = load_sample_table_config()

        self.assertEqual(config.title, DEFAULT_SAMPLE_TABLE_TITLE)
        self.assertEqual(config.sheet_name, SAMPLE_TABLE_SHEET_NAME)
        self.assertEqual(config.output_name, "客户类型样本统计表.xlsx")
        self.assertEqual(config.rows[0].display_name, "展览活动主（承）办")
        self.assertEqual(config.rows[0].target_sample_size, 38)
        self.assertEqual(config.rows[7].category_label, "二、酒店暗访（次）")
        self.assertEqual(config.rows[7].display_name, "")
        self.assertEqual(config.rows[7].target_sample_size, 4)
        self.assertEqual(config.rows[16].display_name, "酒店散客")
        self.assertEqual(config.rows[17].display_name, "酒店住宿团队")
        self.assertEqual(config.rows[18].display_name, "酒店参会客户")
        self.assertEqual(config.rows[19].display_name, "酒店会议活动主（承）办")
        self.assertEqual(config.rows[-1].display_name, "酒店餐饮客户")
        self.assertEqual(config.rows[-1].target_sample_size, 266)

    def test_build_sample_table_rows_uses_mapping_rules_and_special_overrides(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir)
            self._write_raw_sources(input_dir)

            result = build_sample_table_rows(input_dir=input_dir)
            row_by_label = {
                (row.category_label, row.display_name): row
                for row in result.rows
            }

            self.assertEqual(result.group_labels, ("1-2月", "3月"))

            organizer_row = row_by_label[("一、会展客户", "展览活动主（承）办")]
            self.assertEqual(organizer_row.target_sample_size, 38)
            self.assertEqual(organizer_row.actual_count, 1)
            self.assertEqual(organizer_row.group_counts, {"1-2月": 1, "3月": 0})

            lost_row = row_by_label[("一、会展客户", "会展流失主办客户")]
            self.assertEqual(lost_row.actual_count, 0)
            self.assertEqual(lost_row.group_counts, {"1-2月": 0, "3月": 0})

            audit_row = row_by_label[("二、酒店暗访（次）", "")]
            self.assertEqual(audit_row.target_sample_size, 4)
            self.assertEqual(audit_row.actual_count, 1)
            self.assertEqual(audit_row.group_counts, {"1-2月": 0, "3月": 0})

            guest_row = row_by_label[("六、酒店客户", "酒店散客")]
            self.assertEqual(guest_row.actual_count, 2)
            self.assertEqual(guest_row.group_counts, {"1-2月": 1, "3月": 1})

            meeting_attendee_row = row_by_label[("六、酒店客户", "酒店参会客户")]
            self.assertEqual(meeting_attendee_row.actual_count, 1)
            self.assertEqual(meeting_attendee_row.group_counts, {"1-2月": 0, "3月": 1})

            meeting_organizer_row = row_by_label[("六、酒店客户", "酒店会议活动主（承）办")]
            self.assertEqual(meeting_organizer_row.actual_count, 1)
            self.assertEqual(meeting_organizer_row.group_counts, {"1-2月": 0, "3月": 1})

            hotel_catering_row = row_by_label[("六、酒店客户", "酒店餐饮客户")]
            self.assertEqual(hotel_catering_row.actual_count, 4)
            self.assertEqual(hotel_catering_row.group_counts, {"1-2月": 2, "3月": 2})

            research_row = row_by_label[("五、专项调研", "")]
            self.assertEqual(research_row.actual_count, 0)
            self.assertEqual(research_row.group_counts, {"1-2月": 0, "3月": 0})

    def test_build_sample_table_rows_uses_user_supplied_group_labels(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir)
            self._write_raw_sources(input_dir)

            group_specs = parse_sample_group_specs(
                ["1-2月=01-02", "3月=03"],
                default_year="2026",
            )
            result = build_sample_table_rows(
                input_dir=input_dir,
                sample_groups=group_specs,
            )
            row_by_label = {
                (row.category_label, row.display_name): row
                for row in result.rows
            }

            self.assertEqual(result.group_labels, ("1-2月", "3月"))

            organizer_row = row_by_label[("一、会展客户", "展览活动主（承）办")]
            self.assertEqual(organizer_row.group_counts, {"1-2月": 1, "3月": 0})

            guest_row = row_by_label[("六、酒店客户", "酒店散客")]
            self.assertEqual(guest_row.group_counts, {"1-2月": 1, "3月": 1})

    def test_generate_sample_table_report_creates_standalone_workbook_with_formulas(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_dir = temp_path / "inputs"
            output_dir = temp_path / "outputs"
            input_dir.mkdir()
            output_dir.mkdir()
            self._write_raw_sources(input_dir)
            group_specs = parse_sample_group_specs(
                ["1-2月=01-02", "3月=03"],
                default_year="2026",
            )

            output_path = generate_sample_table_report(
                input_dir=input_dir,
                output_dir=output_dir,
                sample_groups=group_specs,
            )

            self.assertTrue(output_path.exists())

            workbook = load_workbook(output_path, data_only=False)
            worksheet = workbook.active

            self.assertEqual(worksheet.title, SAMPLE_TABLE_SHEET_NAME)
            self.assertEqual(worksheet["A1"].value, DEFAULT_SAMPLE_TABLE_TITLE)
            self.assertIn("A1:G1", {str(cell_range) for cell_range in worksheet.merged_cells.ranges})
            self.assertIn("A3:A9", {str(cell_range) for cell_range in worksheet.merged_cells.ranges})

            self.assertEqual(worksheet["A2"].value, "客户大类")
            self.assertEqual(worksheet["B2"].value, "样本类型")
            self.assertEqual(worksheet["C2"].value, "样本量")
            self.assertEqual(worksheet["D2"].value, "样本进度百分比")
            self.assertEqual(worksheet["E2"].value, "总执行样本量")
            self.assertEqual(worksheet["F2"].value, "1-2月")
            self.assertEqual(worksheet["G2"].value, "3月")

            self.assertEqual(worksheet["B3"].value, "展览活动主（承）办")
            self.assertEqual(worksheet["C3"].value, 38)
            self.assertEqual(worksheet["D3"].value, "=IFERROR(E3/C3,0)")
            self.assertEqual(worksheet["E3"].value, 1)
            self.assertEqual(worksheet["F3"].value, 1)
            self.assertEqual(worksheet["G3"].value, 0)

            self.assertEqual(worksheet["A11"].value, "二、酒店暗访（次）")
            self.assertIsNone(worksheet["B11"].value)
            self.assertEqual(worksheet["C11"].value, 4)
            self.assertEqual(worksheet["D11"].value, "=IFERROR(E11/C11,0)")
            self.assertEqual(worksheet["E11"].value, 1)
            self.assertEqual(worksheet["F11"].value, 0)
            self.assertEqual(worksheet["G11"].value, 0)

            self.assertEqual(worksheet["B10"].value, "小计")
            self.assertEqual(worksheet["C10"].value, "=SUM(C3:C9)")
            self.assertEqual(worksheet["D10"].value, "=IFERROR(E10/C10,0)")
            self.assertEqual(worksheet["E10"].value, "=SUM(E3:E9)")
            self.assertEqual(worksheet["F10"].value, "=SUM(F3:F9)")
            self.assertEqual(worksheet["G10"].value, "=SUM(G3:G9)")

            self.assertEqual(worksheet["B22"].value, "酒店散客")
            self.assertEqual(worksheet["E22"].value, 2)
            self.assertEqual(worksheet["F22"].value, 1)
            self.assertEqual(worksheet["G22"].value, 1)
            self.assertEqual(worksheet["B23"].value, "酒店住宿团队")
            self.assertEqual(worksheet["E23"].value, 1)
            self.assertEqual(worksheet["F23"].value, 0)
            self.assertEqual(worksheet["G23"].value, 1)
            self.assertEqual(worksheet["B24"].value, "酒店参会客户")
            self.assertEqual(worksheet["E24"].value, 1)
            self.assertEqual(worksheet["F24"].value, 0)
            self.assertEqual(worksheet["G24"].value, 1)
            self.assertEqual(worksheet["B25"].value, "酒店会议活动主（承）办")
            self.assertEqual(worksheet["E25"].value, 1)
            self.assertEqual(worksheet["F25"].value, 0)
            self.assertEqual(worksheet["G25"].value, 1)
            self.assertEqual(worksheet["B26"].value, "酒店餐饮客户")
            self.assertEqual(worksheet["E26"].value, 4)
            self.assertEqual(worksheet["F26"].value, 2)
            self.assertEqual(worksheet["G26"].value, 2)
            self.assertEqual(worksheet["B27"].value, "小计")
            self.assertEqual(worksheet["C27"].value, "=SUM(C22:C26)")
            self.assertEqual(worksheet["E27"].value, "=SUM(E22:E26)")
            self.assertEqual(worksheet["F27"].value, "=SUM(F22:F26)")
            self.assertEqual(worksheet["G27"].value, "=SUM(G22:G26)")

            self.assertEqual(worksheet["A28"].value, "合计")
            self.assertEqual(worksheet["C28"].value, "=SUM(C3:C9,C11:C11,C12:C16,C18:C19,C21:C21,C22:C26)")
            self.assertEqual(worksheet["D28"].value, "=IFERROR(E28/C28,0)")
            self.assertEqual(worksheet["E28"].value, "=SUM(E3:E9,E11:E11,E12:E16,E18:E19,E21:E21,E22:E26)")
            self.assertEqual(worksheet["F28"].value, "=SUM(F3:F9,F11:F11,F12:F16,F18:F19,F21:F21,F22:F26)")
            self.assertEqual(worksheet["G28"].value, "=SUM(G3:G9,G11:G11,G12:G16,G18:G19,G21:G21,G22:G26)")

    def _write_raw_sources(self, input_dir: Path) -> None:
        write_source_workbook(
            input_dir / "展览.xlsx",
            [
                {"C": "展览", "E": "展览主承办", "月份": "01-02"},
                {"C": "展览", "E": "参展商", "年份": "2026", "月份": "03"},
                {"C": "展览", "E": "参展商", "年份": "2026", "月份": "03"},
                {"C": "展览", "E": "专业观众", "年份": "2026", "月份": "03"},
            ],
        )
        write_source_workbook(
            input_dir / "会展服务商.xlsx",
            [
                {"D": "会展服务商", "月份": "01-02"},
            ],
        )
        write_source_workbook(
            input_dir / "会议.xlsx",
            [
                {"C": "会议", "E": "会议主承办", "月份": "03"},
                {"C": "会议", "E": "参会人员", "月份": "01-02"},
                {"C": "会议", "E": "参会人员", "月份": "03"},
                {"C": "酒店会议", "E": "酒店会议主承办", "月份": "03"},
                {"C": "酒店会议", "E": "参会人员", "月份": "03"},
            ],
        )
        write_source_workbook(
            input_dir / "餐饮.xlsx",
            [
                {"C": "餐饮", "D": "商务简餐", "月份": "01-02"},
                {"C": "餐饮", "D": "特色美食廊", "月份": "03"},
                {"C": "餐饮", "D": "宴会", "月份": "03"},
                {"C": "餐饮", "D": "自助餐", "月份": "01-02"},
                {"C": "酒店餐饮", "D": "商务简餐", "月份": "01-02"},
                {"C": "酒店餐饮", "D": "宴会", "月份": "03"},
                {"C": "酒店餐饮", "D": "酒店宴会", "月份": "03"},
                {"C": "酒店餐饮", "D": "酒店自助餐", "月份": "01-02"},
            ],
        )
        write_source_workbook(
            input_dir / "旅游.xlsx",
            [
                {"C": "旅行社工作人员", "年份": "2026", "月份": "03"},
                {"C": "游客", "年份": "2026", "月份": "03"},
                {"C": "游客", "年份": "2026", "月份": "03"},
            ],
        )
        write_source_workbook(
            input_dir / "酒店.xlsx",
            [
                {"C": "散客", "月份": "01-02"},
                {"C": "散客", "月份": "03"},
                {"C": "住宿团队", "月份": "03"},
            ],
        )


if __name__ == "__main__":
    unittest.main()
