from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from survey_stats import (
    DEFAULT_SHEET_NAME,
    EXHIBITOR_ROLE_NAME,
    EXHIBITOR_TEMPLATE,
    ORGANIZER_ROLE_NAME,
    ORGANIZER_TEMPLATE,
    OVERALL_FILL,
    SECTION_FILL,
    SERVICE_PROVIDER_ROLE_NAME,
    SERVICE_PROVIDER_TEMPLATE,
    VISITOR_ROLE_NAME,
    VISITOR_TEMPLATE,
    build_output_path,
    build_result_dataframe,
    compute_role_stats,
    excel_column_to_index,
    excel_round,
    generate_role_report,
    load_batch_config,
)


def build_mock_dataframe(role_name: str, role_column: str = "E") -> pd.DataFrame:
    column_count = excel_column_to_index("CB") + 1
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
        "BU",
        "BV",
        "BX",
        "BY",
        "CA",
        "CB",
        "K",
        "M",
        "N",
        "O",
        "P",
        "Q",
        "R",
        "S",
        "T",
        "U",
        "W",
        "X",
        "Y",
        "Z",
    ):
        rows[0][excel_column_to_index(column_name)] = 9
        rows[1][excel_column_to_index(column_name)] = 1

    return pd.DataFrame(rows, columns=columns)


class SurveyStatsTest(unittest.TestCase):
    def test_excel_round_matches_excel_style(self) -> None:
        self.assertEqual(excel_round(9.125), 9.13)

    def test_organizer_stats_follow_template_mapping(self) -> None:
        df = build_mock_dataframe(ORGANIZER_ROLE_NAME)
        stats = compute_role_stats(df, ORGANIZER_TEMPLATE)
        result_df = build_result_dataframe(stats)

        self.assertEqual(result_df.iloc[0]["指标"], ORGANIZER_ROLE_NAME)
        self.assertEqual(result_df.iloc[0]["满意度"], 9.0)
        self.assertEqual(result_df.iloc[0]["重要性"], 9.0)

        section_row = result_df[result_df["指标"] == "会展服务"].iloc[0]
        self.assertEqual(section_row["满意度"], 9.0)
        self.assertEqual(section_row["重要性"], 9.0)

    def test_visitor_stats_keep_original_formula_quirks(self) -> None:
        df = build_mock_dataframe(VISITOR_ROLE_NAME)
        df.iloc[0, excel_column_to_index("W")] = 7
        df.iloc[0, excel_column_to_index("X")] = 3
        df.iloc[0, excel_column_to_index("BK")] = 8
        df.iloc[0, excel_column_to_index("BL")] = 2

        stats = compute_role_stats(df, VISITOR_TEMPLATE)
        result_df = build_result_dataframe(stats)

        facility_row = result_df[result_df["指标"] == "设施设备齐全"].iloc[0]
        security_row = result_df[result_df["指标"] == "安保服务"].iloc[0]

        self.assertEqual(facility_row["满意度"], 7.0)
        self.assertEqual(facility_row["重要性"], 7.0)
        self.assertEqual(security_row["满意度"], 8.0)
        self.assertEqual(security_row["重要性"], 8.0)

    def test_service_provider_stats_keep_original_formula_quirks(self) -> None:
        df = build_mock_dataframe(SERVICE_PROVIDER_ROLE_NAME, role_column="D")
        df.iloc[0, excel_column_to_index("AH")] = 9
        df.iloc[0, excel_column_to_index("AL")] = 12
        df.iloc[0, excel_column_to_index("AV")] = 7
        df.iloc[0, excel_column_to_index("AW")] = 3
        df.iloc[0, excel_column_to_index("AX")] = 10

        stats = compute_role_stats(df, SERVICE_PROVIDER_TEMPLATE)
        result_df = build_result_dataframe(stats)

        service_attitude_row = result_df[result_df["指标"] == "工作人员服务态度"].iloc[0]
        revisit_row = result_df[result_df["指标"] == "展后回访"].iloc[0]

        self.assertTrue(pd.isna(service_attitude_row["满意度"]))
        self.assertEqual(service_attitude_row["重要性"], 9.0)
        self.assertEqual(revisit_row["满意度"], 7.0)
        self.assertEqual(revisit_row["重要性"], 10.0)

    def test_generate_role_report_saves_named_file_with_colors(self) -> None:
        df = build_mock_dataframe(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "input.xlsx"
            output_path = build_output_path(temp_path / "outputs", EXHIBITOR_ROLE_NAME)

            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            result_df, saved_path = generate_role_report(
                input_path=input_file,
                role_definition=EXHIBITOR_TEMPLATE,
                output_path=output_path,
                sheet_name=DEFAULT_SHEET_NAME,
                sheet_title=EXHIBITOR_ROLE_NAME,
            )

            self.assertEqual(saved_path, output_path)
            self.assertTrue(output_path.exists())
            self.assertEqual(result_df.iloc[0]["指标"], EXHIBITOR_ROLE_NAME)

            exported_df = pd.read_excel(output_path, sheet_name=EXHIBITOR_ROLE_NAME)
            self.assertEqual(exported_df.iloc[0]["指标"], EXHIBITOR_ROLE_NAME)

            workbook = load_workbook(output_path)
            worksheet = workbook[EXHIBITOR_ROLE_NAME]
            self.assertEqual(worksheet["A2"].fill.start_color.rgb, OVERALL_FILL.start_color.rgb)
            self.assertEqual(worksheet["A3"].fill.start_color.rgb, SECTION_FILL.start_color.rgb)
            self.assertEqual(worksheet["B3"].fill.start_color.rgb, SECTION_FILL.start_color.rgb)

    def test_load_batch_config_reads_sources_and_jobs(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            config_path = temp_path / "jobs.toml"
            source_file = temp_path / "source.xlsx"
            source_df = build_mock_dataframe(ORGANIZER_ROLE_NAME)
            with pd.ExcelWriter(source_file, engine="openpyxl") as writer:
                source_df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"

[[jobs]]
name = "展览主承办"
path = "source.xlsx"
sheet = "问卷数据"
template = "organizer"
role_name = "展览主承办"

[[jobs]]
name = "会展服务商"
path = "source.xlsx"
sheet = "数据"
template = "service_provider"
role_name = "会展服务商"
""".strip(),
                encoding="utf-8",
            )

            config = load_batch_config(config_path)
            self.assertEqual(config.output_dir, (temp_path / "exports").resolve())
            self.assertEqual(config.jobs[0].path, source_file.resolve())
            self.assertEqual(config.jobs[1].sheet_name, "数据")
            self.assertEqual(config.jobs[0].name, "展览主承办")
            self.assertEqual(config.jobs[0].template_name, "organizer")
            self.assertEqual(config.jobs[1].template_name, "service_provider")


if __name__ == "__main__":
    unittest.main()
