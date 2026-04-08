from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from survey_stats import (
    CATERING_BUFFET_ROLE_NAME,
    CATERING_BUFFET_TEMPLATE,
    CATERING_FOOD_HALL_TEMPLATE,
    CATERING_HOTEL_BUFFET_ROLE_NAME,
    CATERING_HOTEL_BUFFET_TEMPLATE,
    CATERING_WEDDING_BANQUET_ROLE_NAME,
    CATERING_WEDDING_BANQUET_TEMPLATE,
    DEFAULT_SHEET_NAME,
    EXHIBITOR_ROLE_NAME,
    EXHIBITOR_TEMPLATE,
    HOTEL_MEETING_ATTENDEE_ROLE_NAME,
    HOTEL_MEETING_ATTENDEE_TEMPLATE,
    HOTEL_MEETING_ORGANIZER_ROLE_NAME,
    HOTEL_MEETING_ORGANIZER_TEMPLATE,
    MEETING_ATTENDEE_ROLE_NAME,
    MEETING_ATTENDEE_TEMPLATE,
    MEETING_ORGANIZER_ROLE_NAME,
    MEETING_ORGANIZER_TEMPLATE,
    MissingGroupNotice,
    ORGANIZER_ROLE_NAME,
    ORGANIZER_TEMPLATE,
    OVERALL_FILL,
    SECTION_FILL,
    SERVICE_PROVIDER_ROLE_NAME,
    SERVICE_PROVIDER_TEMPLATE,
    TEMPLATE_DEFINITIONS,
    VISITOR_ROLE_NAME,
    VISITOR_TEMPLATE,
    build_missing_group_summary,
    build_output_path,
    build_result_dataframe,
    compute_role_stats,
    excel_column_to_index,
    excel_round,
    generate_role_report,
    load_batch_config,
)


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

    def test_catering_wedding_banquet_importance_uses_importance_columns(self) -> None:
        df = build_mock_dataframe(CATERING_WEDDING_BANQUET_ROLE_NAME, role_column="D")
        df.iloc[0, excel_column_to_index("AQ")] = 9
        df.iloc[0, excel_column_to_index("AR")] = 3
        df.iloc[0, excel_column_to_index("AO")] = 8
        df.iloc[0, excel_column_to_index("AP")] = 4
        df.iloc[0, excel_column_to_index("AS")] = 7
        df.iloc[0, excel_column_to_index("AT")] = 5
        df.iloc[0, excel_column_to_index("W")] = 6
        df.iloc[0, excel_column_to_index("X")] = 2
        df.iloc[0, excel_column_to_index("AK")] = 10
        df.iloc[0, excel_column_to_index("AL")] = 1

        stats = compute_role_stats(df, CATERING_WEDDING_BANQUET_TEMPLATE)
        result_df = build_result_dataframe(stats)

        appearance_row = result_df[result_df["指标"] == "工作人员仪容仪表"].iloc[0]
        attitude_row = result_df[result_df["指标"] == "工作人员服务态度"].iloc[0]
        skill_row = result_df[result_df["指标"] == "工作人员业务技能"].iloc[0]
        tea_break_row = result_df[result_df["指标"] == "婚宴茶歇"].iloc[0]
        temperature_row = result_df[result_df["指标"] == "菜品温度"].iloc[0]

        self.assertEqual(appearance_row["满意度"], 9.0)
        self.assertEqual(appearance_row["重要性"], 3.0)
        self.assertEqual(attitude_row["满意度"], 8.0)
        self.assertEqual(attitude_row["重要性"], 4.0)
        self.assertEqual(skill_row["满意度"], 7.0)
        self.assertEqual(skill_row["重要性"], 5.0)
        self.assertEqual(tea_break_row["满意度"], 6.0)
        self.assertEqual(tea_break_row["重要性"], 2.0)
        self.assertEqual(temperature_row["满意度"], 10.0)
        self.assertEqual(temperature_row["重要性"], 1.0)

    def test_catering_buffet_templates_keep_navigation_and_car_finder_mapping(self) -> None:
        for role_name, template in (
            (CATERING_BUFFET_ROLE_NAME, CATERING_BUFFET_TEMPLATE),
            (CATERING_HOTEL_BUFFET_ROLE_NAME, CATERING_HOTEL_BUFFET_TEMPLATE),
        ):
            df = build_mock_dataframe(role_name, role_column="D")
            df.iloc[0, excel_column_to_index("AY")] = 4
            df.iloc[0, excel_column_to_index("AZ")] = 5
            df.iloc[0, excel_column_to_index("BB")] = 8
            df.iloc[0, excel_column_to_index("BC")] = 6

            stats = compute_role_stats(df, template)
            result_df = build_result_dataframe(stats)

            navigation_row = result_df[result_df["指标"] == "室内导航系统"].iloc[0]
            car_finder_row = result_df[result_df["指标"] == "寻车系统"].iloc[0]

            self.assertEqual(navigation_row["满意度"], 4.0)
            self.assertEqual(navigation_row["重要性"], 5.0)
            self.assertEqual(car_finder_row["满意度"], 8.0)
            self.assertEqual(car_finder_row["重要性"], 6.0)

    def test_meeting_organizer_template_uses_fixed_skill_and_car_finder_columns(self) -> None:
        for role_name, template in (
            (MEETING_ORGANIZER_ROLE_NAME, MEETING_ORGANIZER_TEMPLATE),
            (HOTEL_MEETING_ORGANIZER_ROLE_NAME, HOTEL_MEETING_ORGANIZER_TEMPLATE),
        ):
            df = build_mock_dataframe(role_name)
            df.iloc[0, excel_column_to_index("AG")] = 4
            df.iloc[0, excel_column_to_index("AH")] = 3
            df.iloc[0, excel_column_to_index("AI")] = 8
            df.iloc[0, excel_column_to_index("AJ")] = 7
            df.iloc[0, excel_column_to_index("CB")] = 6
            df.iloc[0, excel_column_to_index("CC")] = 5
            df.iloc[0, excel_column_to_index("BY")] = 2
            df.iloc[0, excel_column_to_index("BZ")] = 1

            stats = compute_role_stats(df, template)
            result_df = build_result_dataframe(stats)

            skill_row = result_df[result_df["指标"] == "工作人员业务技能"].iloc[0]
            car_finder_row = result_df[result_df["指标"] == "寻车系统"].iloc[0]

            self.assertEqual(skill_row["满意度"], 4.0)
            self.assertEqual(skill_row["重要性"], 3.0)
            self.assertEqual(car_finder_row["满意度"], 6.0)
            self.assertEqual(car_finder_row["重要性"], 5.0)

    def test_meeting_attendee_templates_use_fixed_skill_and_parking_columns(self) -> None:
        for role_name, template in (
            (HOTEL_MEETING_ATTENDEE_ROLE_NAME, HOTEL_MEETING_ATTENDEE_TEMPLATE),
            (MEETING_ATTENDEE_ROLE_NAME, MEETING_ATTENDEE_TEMPLATE),
        ):
            df = build_mock_dataframe(role_name)
            df.iloc[0, excel_column_to_index("AG")] = 5
            df.iloc[0, excel_column_to_index("AH")] = 4
            df.iloc[0, excel_column_to_index("AI")] = 9
            df.iloc[0, excel_column_to_index("AJ")] = 8
            df.iloc[0, excel_column_to_index("Q")] = 3
            df.iloc[0, excel_column_to_index("R")] = 2
            df.iloc[0, excel_column_to_index("AA")] = 7
            df.iloc[0, excel_column_to_index("AB")] = 6

            stats = compute_role_stats(df, template)
            result_df = build_result_dataframe(stats)

            skill_row = result_df[result_df["指标"] == "工作人员业务技能"].iloc[0]
            parking_row = result_df[result_df["指标"] == "园区停车方便"].iloc[0]

            self.assertEqual(skill_row["满意度"], 5.0)
            self.assertEqual(skill_row["重要性"], 4.0)
            self.assertEqual(parking_row["满意度"], 3.0)
            self.assertEqual(parking_row["重要性"], 2.0)

    def test_missing_group_outputs_blank_statistics_without_error(self) -> None:
        df = build_mock_dataframe("不存在的餐饮群体", role_column="D")
        stats = compute_role_stats(df, CATERING_FOOD_HALL_TEMPLATE)
        result_df = build_result_dataframe(stats)

        self.assertEqual(stats.matched_row_count, 0)
        self.assertEqual(result_df.iloc[0]["指标"], CATERING_FOOD_HALL_TEMPLATE.role_name)
        self.assertTrue(pd.isna(result_df.iloc[0]["满意度"]))
        self.assertTrue(pd.isna(result_df.iloc[0]["重要性"]))

        metric_row = result_df[result_df["指标"] == "菜肴品质"].iloc[0]
        self.assertTrue(pd.isna(metric_row["满意度"]))
        self.assertTrue(pd.isna(metric_row["重要性"]))

    def test_build_missing_group_summary_lists_missing_jobs(self) -> None:
        summary = build_missing_group_summary(
            [
                MissingGroupNotice("特色美食廊", Path("/tmp/catering.xlsx"), DEFAULT_SHEET_NAME),
                MissingGroupNotice("婚宴", Path("/tmp/catering.xlsx"), DEFAULT_SHEET_NAME),
            ]
        )

        self.assertIsNotNone(summary)
        self.assertIn("以下指定的客户分组在来源数据中未找到任何匹配记录", summary)
        self.assertIn("特色美食廊 [catering.xlsx / 问卷数据]", summary)
        self.assertIn("婚宴 [catering.xlsx / 问卷数据]", summary)

    def test_template_role_names_are_unique(self) -> None:
        role_names = [role_definition.role_name for role_definition in TEMPLATE_DEFINITIONS.values()]
        self.assertEqual(len(role_names), len(set(role_names)))

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
