from __future__ import annotations

import argparse
import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

import pandas as pd
from openpyxl import load_workbook

from phase_column_preprocess import preprocess_phase_column_if_needed
from survey_stats import (
    CATERING_BUFFET_ROLE_NAME,
    CATERING_BUFFET_TEMPLATE,
    CATERING_BANQUET_ROLE_NAME,
    CATERING_BUSINESS_MEAL_ROLE_NAME,
    CATERING_FOOD_HALL_TEMPLATE,
    CATERING_HOTEL_BANQUET_ROLE_NAME,
    CATERING_HOTEL_BANQUET_TEMPLATE,
    CATERING_HOTEL_BUFFET_ROLE_NAME,
    CATERING_HOTEL_BUFFET_TEMPLATE,
    CATERING_TOUR_MEAL_ROLE_NAME,
    CATERING_WEDDING_BANQUET_ROLE_NAME,
    CATERING_WEDDING_BANQUET_TEMPLATE,
    DEFAULT_SHEET_NAME,
    DirectoryDiscoveryResult,
    DIRECTORY_NOTICE_REASON_MISSING_ROLE_DATA,
    DIRECTORY_NOTICE_REASON_MISSING_SOURCE_FILE,
    EXHIBITOR_ROLE_NAME,
    EXHIBITOR_TEMPLATE,
    HOTEL_MEETING_ATTENDEE_ROLE_NAME,
    HOTEL_MEETING_ATTENDEE_TEMPLATE,
    HOTEL_MEETING_ORGANIZER_ROLE_NAME,
    HOTEL_MEETING_ORGANIZER_TEMPLATE,
    HOTEL_GROUP_GUEST_ROLE_NAME,
    HOTEL_GROUP_GUEST_TEMPLATE,
    HOTEL_INDIVIDUAL_GUEST_ROLE_NAME,
    HOTEL_INDIVIDUAL_GUEST_TEMPLATE,
    MEETING_ATTENDEE_ROLE_NAME,
    MEETING_ATTENDEE_TEMPLATE,
    MEETING_ORGANIZER_ROLE_NAME,
    MEETING_ORGANIZER_TEMPLATE,
    MissingCustomerTypeNotice,
    MissingGroupNotice,
    ORGANIZER_ROLE_NAME,
    ORGANIZER_TEMPLATE,
    OVERALL_FILL,
    SECTION_FILL,
    SERVICE_PROVIDER_ROLE_NAME,
    SERVICE_PROVIDER_TEMPLATE,
    TEMPLATE_DEFINITIONS,
    TOURIST_ROLE_NAME,
    TOURIST_TEMPLATE,
    TRAVEL_STAFF_ROLE_NAME,
    TRAVEL_STAFF_TEMPLATE,
    UnmappedCustomerCategoryNotice,
    VISITOR_ROLE_NAME,
    VISITOR_TEMPLATE,
    build_missing_customer_type_summary,
    build_missing_group_summary,
    build_output_path,
    build_result_dataframe,
    build_unmapped_customer_category_summary,
    compute_role_stats,
    compute_metric_average,
    excel_column_to_index,
    excel_round,
    generate_customer_category_report_bundle,
    generate_role_report,
    generate_role_report_bundle,
    load_batch_config,
    mean_ignore_empty,
    run_directory_batch,
    run_config_mode,
    run_single_mode,
)


def build_mock_dataframe(role_name: str, role_column: str = "E") -> pd.DataFrame:
    column_count = excel_column_to_index("CF") + 1
    columns = [f"col_{index + 1}" for index in range(column_count)]
    rows = [[None for _ in range(column_count)] for _ in range(2)]

    rows[0][excel_column_to_index(role_column)] = role_name
    rows[1][excel_column_to_index(role_column)] = "其他身份"

    default_category_by_role_name = {
        ORGANIZER_ROLE_NAME: "展览",
        EXHIBITOR_ROLE_NAME: "展览",
        VISITOR_ROLE_NAME: "展览",
        MEETING_ORGANIZER_ROLE_NAME: "会议",
        MEETING_ATTENDEE_ROLE_NAME: "会议",
        HOTEL_MEETING_ORGANIZER_ROLE_NAME: "酒店会议",
        HOTEL_MEETING_ATTENDEE_ROLE_NAME: "酒店会议",
        CATERING_FOOD_HALL_TEMPLATE.role_name: "餐饮",
        CATERING_BUSINESS_MEAL_ROLE_NAME: "餐饮",
        CATERING_TOUR_MEAL_ROLE_NAME: "餐饮",
        CATERING_BANQUET_ROLE_NAME: "餐饮",
        CATERING_WEDDING_BANQUET_ROLE_NAME: "餐饮",
        CATERING_BUFFET_ROLE_NAME: "餐饮",
        CATERING_HOTEL_BANQUET_ROLE_NAME: "酒店餐饮",
        CATERING_HOTEL_BUFFET_ROLE_NAME: "酒店餐饮",
    }
    default_category = default_category_by_role_name.get(role_name)
    if default_category is not None:
        rows[0][excel_column_to_index("C")] = default_category
        rows[1][excel_column_to_index("C")] = f"其他{default_category}"

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


def build_meeting_category_split_dataframe(role_name: str) -> pd.DataFrame:
    df = build_mock_dataframe(role_name)
    category_index = excel_column_to_index("C")
    role_index = excel_column_to_index("E")
    df.iloc[0, category_index] = "会议"
    df.iloc[1, category_index] = "酒店会议"
    df.iloc[1, role_index] = role_name
    return df


def build_directory_mode_meeting_dataframe() -> pd.DataFrame:
    organizer_df = build_mock_dataframe(MEETING_ORGANIZER_ROLE_NAME)
    organizer_df.iloc[0, excel_column_to_index("C")] = "会议"
    organizer_df.iloc[1, excel_column_to_index("C")] = "酒店会议"
    organizer_df.iloc[1, excel_column_to_index("E")] = HOTEL_MEETING_ORGANIZER_ROLE_NAME

    attendee_df = build_mock_dataframe(MEETING_ATTENDEE_ROLE_NAME)
    attendee_df.iloc[0, excel_column_to_index("C")] = "会议"
    attendee_df.iloc[1, excel_column_to_index("C")] = "酒店会议"
    attendee_df.iloc[1, excel_column_to_index("E")] = "参会人员"
    return pd.concat([organizer_df, attendee_df], ignore_index=True)


def build_catering_category_split_dataframe() -> pd.DataFrame:
    banquet_df = build_mock_dataframe(CATERING_HOTEL_BANQUET_ROLE_NAME, role_column="D")
    banquet_df.iloc[0, excel_column_to_index("C")] = "酒店餐饮"
    banquet_df.iloc[1, excel_column_to_index("C")] = "餐饮"
    banquet_df.iloc[1, excel_column_to_index("D")] = CATERING_HOTEL_BANQUET_ROLE_NAME

    buffet_df = build_mock_dataframe(CATERING_HOTEL_BUFFET_ROLE_NAME, role_column="D")
    buffet_df.iloc[0, excel_column_to_index("C")] = "酒店餐饮"
    buffet_df.iloc[1, excel_column_to_index("C")] = "餐饮"
    buffet_df.iloc[1, excel_column_to_index("D")] = CATERING_HOTEL_BUFFET_ROLE_NAME

    business_meal_df = build_mock_dataframe(CATERING_BUSINESS_MEAL_ROLE_NAME, role_column="D")
    business_meal_df.iloc[0, excel_column_to_index("C")] = "酒店餐饮"
    business_meal_df.iloc[1, excel_column_to_index("C")] = "餐饮"
    business_meal_df.iloc[1, excel_column_to_index("D")] = CATERING_BUSINESS_MEAL_ROLE_NAME

    standard_banquet_df = build_mock_dataframe(CATERING_BANQUET_ROLE_NAME, role_column="D")
    standard_banquet_df.iloc[0, excel_column_to_index("C")] = "酒店餐饮"
    standard_banquet_df.iloc[1, excel_column_to_index("C")] = "餐饮"
    standard_banquet_df.iloc[1, excel_column_to_index("D")] = CATERING_BANQUET_ROLE_NAME

    return pd.concat(
        [banquet_df, buffet_df, business_meal_df, standard_banquet_df],
        ignore_index=True,
    )


def build_directory_mode_tourism_dataframe() -> pd.DataFrame:
    tourist_df = build_mock_dataframe(TOURIST_ROLE_NAME, role_column="C")
    travel_staff_df = build_mock_dataframe(TRAVEL_STAFF_ROLE_NAME, role_column="C")
    return pd.concat([tourist_df, travel_staff_df], ignore_index=True)


def build_template_matching_dataframe(role_definition) -> pd.DataFrame:
    df = build_mock_dataframe(
        role_definition.role_match_value or role_definition.role_name,
        role_column=role_definition.role_column,
    )
    for row_condition in role_definition.row_conditions:
        df.iloc[0, excel_column_to_index(row_condition.column)] = row_condition.expected_value
        df.iloc[1, excel_column_to_index(row_condition.column)] = f"其他{row_condition.expected_value}"
    return df


class SurveyStatsTest(unittest.TestCase):
    def test_excel_round_matches_excel_style(self) -> None:
        self.assertEqual(excel_round(9.125), 9.13)

    def test_mean_ignore_empty_avoids_float_boundary_rounding_error(self) -> None:
        self.assertEqual(mean_ignore_empty([9.48, 9.52, 9.53, 9.49]), 9.51)

    def test_compute_metric_average_avoids_float_boundary_rounding_error(self) -> None:
        df = pd.DataFrame(
            {
                "role": [SERVICE_PROVIDER_ROLE_NAME] * 4,
                "score": [9.48, 9.52, 9.53, 9.49],
            }
        )
        role_mask = df["role"].eq(SERVICE_PROVIDER_ROLE_NAME)

        self.assertEqual(compute_metric_average(df, role_mask, "B"), 9.51)

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

    def test_summary_mode_rebuilds_organizer_sections_to_match_summary_dimensions(self) -> None:
        df = build_mock_dataframe(ORGANIZER_ROLE_NAME)
        stats = compute_role_stats(df, ORGANIZER_TEMPLATE, calculation_mode="summary")
        result_df = build_result_dataframe(stats)

        section_names = result_df["指标"].tolist()
        self.assertIn("产品服务", section_names)
        self.assertIn("硬件设施", section_names)
        self.assertIn("配套服务", section_names)
        self.assertIn("智慧场馆/服务", section_names)
        self.assertIn("餐饮服务", section_names)
        self.assertNotIn("会展服务", section_names)

        product_row = result_df[result_df["指标"] == "产品服务"].iloc[0]
        dining_row = result_df[result_df["指标"] == "餐饮服务"].iloc[0]

        self.assertEqual(product_row["满意度"], 9.0)
        self.assertEqual(product_row["重要性"], 9.0)
        self.assertEqual(dining_row["满意度"], 9.0)
        self.assertEqual(dining_row["重要性"], 9.0)

    def test_summary_mode_rebuilds_hotel_guest_sections_and_omits_grey_dimensions(self) -> None:
        df = build_mock_dataframe(HOTEL_INDIVIDUAL_GUEST_ROLE_NAME, role_column="C")
        stats = compute_role_stats(
            df,
            HOTEL_INDIVIDUAL_GUEST_TEMPLATE,
            calculation_mode="summary",
        )
        result_df = build_result_dataframe(stats)

        section_names = result_df["指标"].tolist()
        self.assertIn("产品服务", section_names)
        self.assertIn("硬件设施", section_names)
        self.assertIn("智慧场馆/服务", section_names)
        self.assertIn("餐饮服务", section_names)
        self.assertNotIn("入住服务", section_names)
        self.assertNotIn("配套服务", section_names)

    def test_summary_mode_splits_dining_out_of_support_section_for_organizer(self) -> None:
        df = build_mock_dataframe(ORGANIZER_ROLE_NAME)
        df.iloc[0, excel_column_to_index("BG")] = 3
        df.iloc[0, excel_column_to_index("BH")] = 3

        stats = compute_role_stats(df, ORGANIZER_TEMPLATE, calculation_mode="summary")
        result_df = build_result_dataframe(stats)

        support_row = result_df[result_df["指标"] == "配套服务"].iloc[0]
        dining_row = result_df[
            (result_df["指标"] == "餐饮服务") & result_df["满意度"].notna()
        ].iloc[0]

        self.assertEqual(support_row["满意度"], 9.0)
        self.assertEqual(support_row["重要性"], 9.0)
        self.assertEqual(dining_row["满意度"], 3.0)
        self.assertEqual(dining_row["重要性"], 3.0)

    def test_organizer_template_uses_corrected_importance_columns(self) -> None:
        df = build_mock_dataframe(ORGANIZER_ROLE_NAME)
        df.iloc[0, excel_column_to_index("AW")] = 9
        df.iloc[0, excel_column_to_index("AX")] = 2
        df.iloc[0, excel_column_to_index("AY")] = 8
        df.iloc[0, excel_column_to_index("K")] = 7
        df.iloc[0, excel_column_to_index("L")] = 3
        df.iloc[0, excel_column_to_index("U")] = 6
        df.iloc[0, excel_column_to_index("V")] = 4
        df.iloc[0, excel_column_to_index("W")] = 5
        df.iloc[0, excel_column_to_index("X")] = 1
        df.iloc[0, excel_column_to_index("BK")] = 8
        df.iloc[0, excel_column_to_index("BL")] = 2

        stats = compute_role_stats(df, ORGANIZER_TEMPLATE)
        result_df = build_result_dataframe(stats)

        report_process_row = result_df[result_df["指标"] == "报馆流程及服务"].iloc[0]
        traffic_flow_row = result_df[result_df["指标"] == "交通流线"].iloc[0]
        cargo_route_row = result_df[result_df["指标"] == "货运通道"].iloc[0]
        facility_row = result_df[result_df["指标"] == "设施设备齐全"].iloc[0]
        security_row = result_df[result_df["指标"] == "安保服务"].iloc[0]

        self.assertEqual(report_process_row["满意度"], 9.0)
        self.assertEqual(report_process_row["重要性"], 2.0)
        self.assertEqual(traffic_flow_row["满意度"], 7.0)
        self.assertEqual(traffic_flow_row["重要性"], 3.0)
        self.assertEqual(cargo_route_row["满意度"], 6.0)
        self.assertEqual(cargo_route_row["重要性"], 4.0)
        self.assertEqual(facility_row["满意度"], 5.0)
        self.assertEqual(facility_row["重要性"], 1.0)
        self.assertEqual(security_row["满意度"], 8.0)
        self.assertEqual(security_row["重要性"], 2.0)

    def test_visitor_stats_use_corrected_importance_columns(self) -> None:
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
        self.assertEqual(facility_row["重要性"], 3.0)
        self.assertEqual(security_row["满意度"], 8.0)
        self.assertEqual(security_row["重要性"], 2.0)

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
            df = build_template_matching_dataframe(template)
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
            df = build_template_matching_dataframe(template)
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

    def test_meeting_attendee_templates_split_rows_by_category_column(self) -> None:
        df = build_meeting_category_split_dataframe(MEETING_ATTENDEE_ROLE_NAME)
        df.iloc[0, excel_column_to_index("Q")] = 3
        df.iloc[0, excel_column_to_index("R")] = 2
        df.iloc[1, excel_column_to_index("Q")] = 8
        df.iloc[1, excel_column_to_index("R")] = 7
        df.iloc[1, excel_column_to_index("E")] = HOTEL_MEETING_ATTENDEE_ROLE_NAME

        meeting_stats = compute_role_stats(df, MEETING_ATTENDEE_TEMPLATE)
        hotel_stats = compute_role_stats(df, HOTEL_MEETING_ATTENDEE_TEMPLATE)

        meeting_df = build_result_dataframe(meeting_stats)
        hotel_df = build_result_dataframe(hotel_stats)
        meeting_parking_row = meeting_df[meeting_df["指标"] == "园区停车方便"].iloc[0]
        hotel_parking_row = hotel_df[hotel_df["指标"] == "园区停车方便"].iloc[0]

        self.assertEqual(meeting_stats.matched_row_count, 1)
        self.assertEqual(hotel_stats.matched_row_count, 1)
        self.assertEqual(meeting_parking_row["满意度"], 3.0)
        self.assertEqual(meeting_parking_row["重要性"], 2.0)
        self.assertEqual(hotel_parking_row["满意度"], 8.0)
        self.assertEqual(hotel_parking_row["重要性"], 7.0)

    def test_generate_customer_category_report_bundle_accepts_hotel_meeting_attendee_shared_label(self) -> None:
        df = build_template_matching_dataframe(HOTEL_MEETING_ATTENDEE_TEMPLATE)
        df.iloc[0, excel_column_to_index("E")] = "参会人员"
        df.iloc[0, excel_column_to_index("Q")] = 6
        df.iloc[0, excel_column_to_index("R")] = 5

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "会议.xlsx"
            output_file = temp_path / "酒店参会客户.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            from survey_customer_category_rules import CUSTOMER_CATEGORY_RULE_BY_NAME

            report = generate_customer_category_report_bundle(
                input_path=input_file,
                category_rule=CUSTOMER_CATEGORY_RULE_BY_NAME["酒店参会客户"],
                output_path=output_file,
                sheet_name=DEFAULT_SHEET_NAME,
                dry_run=True,
            )

        parking_row = report.result_df[report.result_df["指标"] == "园区停车方便"].iloc[0]
        self.assertEqual(report.stats.role_name, "酒店参会客户")
        self.assertEqual(report.stats.matched_row_count, 1)
        self.assertEqual(parking_row["满意度"], 6.0)
        self.assertEqual(parking_row["重要性"], 5.0)

    def test_meeting_organizer_templates_split_rows_by_category_column(self) -> None:
        df = build_meeting_category_split_dataframe(MEETING_ORGANIZER_ROLE_NAME)
        df.iloc[0, excel_column_to_index("Q")] = 4
        df.iloc[0, excel_column_to_index("R")] = 3
        df.iloc[1, excel_column_to_index("Q")] = 9
        df.iloc[1, excel_column_to_index("R")] = 8
        df.iloc[1, excel_column_to_index("E")] = HOTEL_MEETING_ORGANIZER_ROLE_NAME

        meeting_stats = compute_role_stats(df, MEETING_ORGANIZER_TEMPLATE)
        hotel_stats = compute_role_stats(df, HOTEL_MEETING_ORGANIZER_TEMPLATE)

        meeting_df = build_result_dataframe(meeting_stats)
        hotel_df = build_result_dataframe(hotel_stats)
        meeting_parking_row = meeting_df[meeting_df["指标"] == "园区停车方便"].iloc[0]
        hotel_parking_row = hotel_df[hotel_df["指标"] == "园区停车方便"].iloc[0]

        self.assertEqual(meeting_stats.matched_row_count, 1)
        self.assertEqual(hotel_stats.matched_row_count, 1)
        self.assertEqual(meeting_parking_row["满意度"], 4.0)
        self.assertEqual(meeting_parking_row["重要性"], 3.0)
        self.assertEqual(hotel_parking_row["满意度"], 9.0)
        self.assertEqual(hotel_parking_row["重要性"], 8.0)

    def test_hotel_catering_templates_require_hotel_catering_auxiliary_category(self) -> None:
        df = build_catering_category_split_dataframe()
        df.iloc[0, excel_column_to_index("I")] = 9
        df.iloc[0, excel_column_to_index("J")] = 8
        df.iloc[1, excel_column_to_index("I")] = 3
        df.iloc[1, excel_column_to_index("J")] = 2
        df.iloc[2, excel_column_to_index("I")] = 7
        df.iloc[2, excel_column_to_index("J")] = 6
        df.iloc[3, excel_column_to_index("I")] = 4
        df.iloc[3, excel_column_to_index("J")] = 1

        banquet_stats = compute_role_stats(df, CATERING_HOTEL_BANQUET_TEMPLATE)
        buffet_stats = compute_role_stats(df, CATERING_HOTEL_BUFFET_TEMPLATE)

        banquet_df = build_result_dataframe(banquet_stats)
        buffet_df = build_result_dataframe(buffet_stats)

        self.assertEqual(banquet_stats.matched_row_count, 1)
        self.assertEqual(buffet_stats.matched_row_count, 1)
        self.assertEqual(
            banquet_df[banquet_df["指标"] == "园区停车方便"].iloc[0]["满意度"],
            9.0,
        )
        self.assertEqual(
            buffet_df[buffet_df["指标"] == "园区停车方便"].iloc[0]["满意度"],
            7.0,
        )

    def test_generate_customer_category_report_bundle_merges_hotel_catering_components(self) -> None:
        df = build_catering_category_split_dataframe()
        df.iloc[0, excel_column_to_index("I")] = 10
        df.iloc[0, excel_column_to_index("J")] = 8
        df.iloc[2, excel_column_to_index("I")] = 6
        df.iloc[2, excel_column_to_index("J")] = 4
        df.iloc[4, excel_column_to_index("I")] = 8
        df.iloc[4, excel_column_to_index("J")] = 6
        df.iloc[6, excel_column_to_index("I")] = 4
        df.iloc[6, excel_column_to_index("J")] = 2

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "餐饮.xlsx"
            output_file = temp_path / "酒店餐饮客户.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            from survey_customer_category_rules import CUSTOMER_CATEGORY_RULE_BY_NAME

            report = generate_customer_category_report_bundle(
                input_path=input_file,
                category_rule=CUSTOMER_CATEGORY_RULE_BY_NAME["酒店餐饮客户"],
                output_path=output_file,
                sheet_name=DEFAULT_SHEET_NAME,
                dry_run=True,
            )

        parking_row = report.result_df[report.result_df["指标"] == "园区停车方便"].iloc[0]
        self.assertEqual(report.stats.role_name, "酒店餐饮客户")
        self.assertEqual(report.stats.matched_row_count, 4)
        self.assertEqual(parking_row["满意度"], 7.0)
        self.assertEqual(parking_row["重要性"], 5.0)

    def test_hotel_guest_templates_use_role_column_c_and_fixed_importance_columns(self) -> None:
        for role_name, template in (
            (HOTEL_INDIVIDUAL_GUEST_ROLE_NAME, HOTEL_INDIVIDUAL_GUEST_TEMPLATE),
            (HOTEL_GROUP_GUEST_ROLE_NAME, HOTEL_GROUP_GUEST_TEMPLATE),
        ):
            df = build_mock_dataframe(role_name, role_column="C")
            df.iloc[0, excel_column_to_index("Q")] = 4
            df.iloc[0, excel_column_to_index("R")] = 3
            df.iloc[0, excel_column_to_index("AA")] = 8
            df.iloc[0, excel_column_to_index("AB")] = 7
            df.iloc[0, excel_column_to_index("Y")] = 6
            df.iloc[0, excel_column_to_index("Z")] = 5
            df.iloc[0, excel_column_to_index("AI")] = 2
            df.iloc[0, excel_column_to_index("AJ")] = 1

            stats = compute_role_stats(df, template)
            result_df = build_result_dataframe(stats)

            room_facility_row = result_df[result_df["指标"] == "客房设施设备"].iloc[0]
            checkin_row = result_df[result_df["指标"] == "入住登记"].iloc[0]
            appearance_row = result_df[result_df["指标"] == "工作人员仪容仪表"].iloc[0]

            self.assertEqual(stats.matched_row_count, 1)
            self.assertEqual(room_facility_row["满意度"], 4.0)
            self.assertEqual(room_facility_row["重要性"], 3.0)
            self.assertEqual(checkin_row["满意度"], 6.0)
            self.assertEqual(checkin_row["重要性"], 5.0)
            self.assertEqual(appearance_row["满意度"], 2.0)
            self.assertEqual(appearance_row["重要性"], 1.0)

    def test_tourism_templates_use_role_column_c_and_expected_sections(self) -> None:
        for role_name, template in (
            (TRAVEL_STAFF_ROLE_NAME, TRAVEL_STAFF_TEMPLATE),
            (TOURIST_ROLE_NAME, TOURIST_TEMPLATE),
        ):
            df = build_mock_dataframe(role_name, role_column="C")
            df.iloc[0, excel_column_to_index("T")] = 2
            df.iloc[0, excel_column_to_index("U")] = 1
            df.iloc[0, excel_column_to_index("P")] = 4
            df.iloc[0, excel_column_to_index("Q")] = 3
            df.iloc[0, excel_column_to_index("AI")] = 6
            df.iloc[0, excel_column_to_index("AJ")] = 5
            df.iloc[0, excel_column_to_index("AR")] = 8
            df.iloc[0, excel_column_to_index("AS")] = 7

            stats = compute_role_stats(df, template)
            result_df = build_result_dataframe(stats)

            ticket_row = result_df[result_df["指标"] == "售票/销售服务"].iloc[0]
            parking_row = result_df[result_df["指标"] == "园区停车方便"].iloc[0]
            voice_row = result_df[result_df["指标"] == "语音导览"].iloc[0]
            mall_row = result_df[result_df["指标"] == "线上商城"].iloc[0]

            self.assertEqual(stats.matched_row_count, 1)
            self.assertEqual(ticket_row["满意度"], 2.0)
            self.assertEqual(ticket_row["重要性"], 1.0)
            self.assertEqual(parking_row["满意度"], 4.0)
            self.assertEqual(parking_row["重要性"], 3.0)
            self.assertEqual(voice_row["满意度"], 6.0)
            self.assertEqual(voice_row["重要性"], 5.0)
            self.assertEqual(mall_row["满意度"], 8.0)
            self.assertEqual(mall_row["重要性"], 7.0)

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

    def test_build_missing_customer_type_summary_groups_reasons(self) -> None:
        summary = build_missing_customer_type_summary(
            [
                MissingCustomerTypeNotice(
                    "会展服务商",
                    "会展服务商.xlsx",
                    DEFAULT_SHEET_NAME,
                    DIRECTORY_NOTICE_REASON_MISSING_SOURCE_FILE,
                ),
                MissingCustomerTypeNotice(
                    "专业观众",
                    "展览.xlsx",
                    DEFAULT_SHEET_NAME,
                    DIRECTORY_NOTICE_REASON_MISSING_ROLE_DATA,
                ),
            ]
        )

        self.assertIsNotNone(summary)
        self.assertIn("以下客户类型因缺少来源数据被跳过", summary)
        self.assertIn("[缺少来源文件]", summary)
        self.assertIn("会展服务商 [会展服务商.xlsx / 问卷数据]", summary)
        self.assertIn("[来源文件存在但未找到匹配身份值]", summary)
        self.assertIn("专业观众 [展览.xlsx / 问卷数据]", summary)

    def test_build_unmapped_customer_category_summary_lists_excluded_combinations(self) -> None:
        summary = build_unmapped_customer_category_summary(
            [
                UnmappedCustomerCategoryNotice(
                    source_file_name="会议.xlsx",
                    auxiliary_value="会议",
                    data_value="酒店参会客户",
                    row_count=1,
                ),
            ]
        )

        self.assertIsNotNone(summary)
        self.assertIn("未纳入 V1.0 客户类别口径", summary)
        self.assertIn("会议.xlsx: 会议 + 酒店参会客户（1 行）", summary)

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

    def test_preprocess_phase_column_moves_third_column_to_end_and_saves_workbook(self) -> None:
        df = build_shifted_dataframe_with_phase_column(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_file = Path(temp_dir) / "shifted.xlsx"
            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            notice = preprocess_phase_column_if_needed(input_file, DEFAULT_SHEET_NAME)

            self.assertIsNotNone(notice)
            reloaded_df = pd.read_excel(input_file, sheet_name=DEFAULT_SHEET_NAME)
            self.assertEqual(reloaded_df.iloc[0, excel_column_to_index("E")], EXHIBITOR_ROLE_NAME)
            self.assertEqual(reloaded_df.iloc[0, reloaded_df.shape[1] - 1], "一期")
            self.assertEqual(reloaded_df.iloc[1, reloaded_df.shape[1] - 1], "二期")

    def test_run_single_mode_prints_notice_when_phase_column_preprocessed(self) -> None:
        df = build_shifted_dataframe_with_phase_column(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "shifted.xlsx"
            output_file = temp_path / "result.xlsx"

            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            buffer = io.StringIO()
            with redirect_stdout(buffer):
                run_single_mode(
                    argparse.Namespace(
                        input=input_file,
                        template="exhibitor",
                        role_name=EXHIBITOR_ROLE_NAME,
                        output=output_file,
                        sheet_name=DEFAULT_SHEET_NAME,
                        calculation_mode="template",
                        dry_run=True,
                    )
                )

            output = buffer.getvalue()
            self.assertIn("已执行输入文件预处理", output)
            report = generate_role_report_bundle(
                input_path=input_file,
                role_definition=EXHIBITOR_TEMPLATE,
                output_path=output_file,
                sheet_name=DEFAULT_SHEET_NAME,
                sheet_title=EXHIBITOR_ROLE_NAME,
                calculation_mode="template",
                dry_run=True,
            )
            self.assertEqual(report.stats.matched_row_count, 1)
            self.assertEqual(report.result_df.iloc[0]["满意度"], 9.0)

    def test_run_single_mode_only_prints_file_progress_without_result_table(self) -> None:
        df = build_mock_dataframe(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_file = temp_path / "input.xlsx"
            output_file = temp_path / "result.xlsx"

            with pd.ExcelWriter(input_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            buffer = io.StringIO()
            with redirect_stdout(buffer):
                run_single_mode(
                    argparse.Namespace(
                        input=input_file,
                        template="exhibitor",
                        role_name=EXHIBITOR_ROLE_NAME,
                        output=output_file,
                        sheet_name=DEFAULT_SHEET_NAME,
                        calculation_mode="template",
                        dry_run=True,
                    )
                )

            output = buffer.getvalue()
            self.assertIn("[1/1] 正在处理文件：input.xlsx（参展商）", output)
            self.assertIn("[1/1] 已完成校验：input.xlsx（参展商）", output)
            self.assertNotIn("## 参展商", output)
            self.assertNotIn("| 指标 |", output)

    def test_run_config_mode_only_prints_file_progress_without_result_table(self) -> None:
        df = build_mock_dataframe(EXHIBITOR_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            source_file = temp_path / "source.xlsx"
            config_path = temp_path / "jobs.toml"

            with pd.ExcelWriter(source_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"

[[jobs]]
name = "参展商-一批"
path = "source.xlsx"
sheet = "问卷数据"
template = "exhibitor"
role_name = "参展商"
output_name = "参展商-一批"

[[jobs]]
name = "参展商-二批"
path = "source.xlsx"
sheet = "问卷数据"
template = "exhibitor"
role_name = "参展商"
output_name = "参展商-二批"
""".strip(),
                encoding="utf-8",
            )

            buffer = io.StringIO()
            with redirect_stdout(buffer):
                run_config_mode(
                    argparse.Namespace(
                        config=config_path,
                        job=[],
                        dry_run=True,
                        sheet_name=DEFAULT_SHEET_NAME,
                        output_format=None,
                        calculation_mode=None,
                        output_dir=None,
                    )
                )

            output = buffer.getvalue()
            self.assertIn("[1/2] 正在处理文件：source.xlsx（参展商-一批）", output)
            self.assertIn("[2/2] 正在处理文件：source.xlsx（参展商-二批）", output)
            self.assertIn("[1/2] 已完成校验：source.xlsx（参展商-一批）", output)
            self.assertIn("[2/2] 已完成校验：source.xlsx（参展商-二批）", output)
            self.assertNotIn("## 参展商-一批", output)
            self.assertNotIn("| 指标 |", output)

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

    def test_load_batch_config_reads_summary_calculation_mode(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            config_path = temp_path / "jobs.toml"

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"
calculation_mode = "summary"

[[jobs]]
name = "展览主承办"
path = "source.xlsx"
sheet = "问卷数据"
template = "organizer"
role_name = "展览主承办"
""".strip(),
                encoding="utf-8",
            )

            config = load_batch_config(config_path)
            self.assertEqual(config.calculation_mode, "summary")

    def test_load_batch_config_reads_input_dir_mode_and_source_file_overrides(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            config_path = temp_path / "jobs.toml"

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"
sheet_name = "数据"
input_dir = "datas"

[source_file_overrides]
"会展服务商.xlsx" = "自定义会展服务商.xlsx"
""".strip(),
                encoding="utf-8",
            )

            config = load_batch_config(config_path)
            self.assertEqual(config.output_dir, (temp_path / "exports").resolve())
            self.assertEqual(config.sheet_name, "数据")
            self.assertEqual(config.input_dir, (temp_path / "datas").resolve())
            self.assertEqual(config.jobs, ())
            self.assertEqual(len(config.source_file_overrides), 1)
            self.assertEqual(config.source_file_overrides[0].standard_file_name, "会展服务商.xlsx")
            self.assertEqual(config.source_file_overrides[0].actual_file_name, "自定义会展服务商.xlsx")

    def test_load_batch_config_rejects_jobs_and_input_dir_together(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            config_path = temp_path / "jobs.toml"

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"
input_dir = "datas"

[[jobs]]
name = "展览主承办"
path = "source.xlsx"
template = "organizer"
""".strip(),
                encoding="utf-8",
            )

            with self.assertRaisesRegex(ValueError, "不能同时包含 input_dir 和 \\[\\[jobs\\]\\]"):
                load_batch_config(config_path)

    def test_run_config_mode_directory_mode_only_reports_missing_customer_types_at_end(self) -> None:
        df = build_mock_dataframe(ORGANIZER_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            data_dir = temp_path / "datas"
            data_dir.mkdir()
            source_file = data_dir / "展览.xlsx"
            config_path = temp_path / "jobs.toml"

            with pd.ExcelWriter(source_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"
input_dir = "datas"
""".strip(),
                encoding="utf-8",
            )

            buffer = io.StringIO()
            with redirect_stdout(buffer):
                run_config_mode(
                    argparse.Namespace(
                        config=config_path,
                        job=[],
                        dry_run=True,
                        sheet_name=DEFAULT_SHEET_NAME,
                        output_format=None,
                        calculation_mode=None,
                        output_dir=None,
                    )
                )

            output = buffer.getvalue()
            self.assertIn("[1/1] 正在处理文件：展览.xlsx（展览主承办）", output)
            self.assertIn("[1/1] 已完成校验：展览.xlsx（展览主承办）", output)
            self.assertIn("以下客户类型因缺少来源数据被跳过，未生成统计结果", output)
            self.assertIn("[缺少来源文件]", output)
            self.assertIn("会展服务商 [会展服务商.xlsx / 问卷数据]", output)
            self.assertIn("[来源文件存在但未找到匹配身份值]", output)
            self.assertIn("参展商 [展览.xlsx / 问卷数据]", output)
            self.assertIn("专业观众 [展览.xlsx / 问卷数据]", output)
            self.assertNotIn("展览主承办 [展览.xlsx / 问卷数据]", output)

    def test_run_config_mode_directory_mode_skips_empty_exports(self) -> None:
        df = build_mock_dataframe(ORGANIZER_ROLE_NAME)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            data_dir = temp_path / "datas"
            data_dir.mkdir()
            source_file = data_dir / "展览.xlsx"
            config_path = temp_path / "jobs.toml"

            with pd.ExcelWriter(source_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"
input_dir = "datas"
""".strip(),
                encoding="utf-8",
            )

            with redirect_stdout(io.StringIO()):
                run_config_mode(
                    argparse.Namespace(
                        config=config_path,
                        job=[],
                        dry_run=False,
                        sheet_name=DEFAULT_SHEET_NAME,
                        output_format=None,
                        calculation_mode=None,
                        output_dir=None,
                    )
                )

            self.assertTrue((temp_path / "exports" / "展览主承办.xlsx").exists())
            self.assertFalse((temp_path / "exports" / "参展商.xlsx").exists())
            self.assertFalse((temp_path / "exports" / "专业观众.xlsx").exists())

    def test_run_config_mode_directory_mode_discovers_hotel_meeting_customer_types(self) -> None:
        df = build_directory_mode_meeting_dataframe()

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            data_dir = temp_path / "datas"
            data_dir.mkdir()
            source_file = data_dir / "会议.xlsx"
            config_path = temp_path / "jobs.toml"

            with pd.ExcelWriter(source_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"
input_dir = "datas"
""".strip(),
                encoding="utf-8",
            )

            with redirect_stdout(io.StringIO()):
                run_config_mode(
                    argparse.Namespace(
                        config=config_path,
                        job=[],
                        dry_run=False,
                        sheet_name=DEFAULT_SHEET_NAME,
                        output_format=None,
                        calculation_mode=None,
                        output_dir=None,
                    )
                )

            self.assertTrue((temp_path / "exports" / "会议主承办.xlsx").exists())
            self.assertTrue((temp_path / "exports" / "酒店会议主承办.xlsx").exists())
            self.assertTrue((temp_path / "exports" / "参会人员.xlsx").exists())
            self.assertTrue((temp_path / "exports" / "酒店参会客户.xlsx").exists())

    def test_run_config_mode_directory_mode_discovers_tourism_customer_types(self) -> None:
        df = build_directory_mode_tourism_dataframe()

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            data_dir = temp_path / "datas"
            data_dir.mkdir()
            source_file = data_dir / "旅游.xlsx"
            config_path = temp_path / "jobs.toml"

            with pd.ExcelWriter(source_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"
input_dir = "datas"
""".strip(),
                encoding="utf-8",
            )

            with redirect_stdout(io.StringIO()):
                run_config_mode(
                    argparse.Namespace(
                        config=config_path,
                        job=[],
                        dry_run=False,
                        sheet_name=DEFAULT_SHEET_NAME,
                        output_format=None,
                        calculation_mode=None,
                        output_dir=None,
                    )
                )

            self.assertTrue((temp_path / "exports" / "旅行社工作人员.xlsx").exists())
            self.assertTrue((temp_path / "exports" / "游客.xlsx").exists())

    def test_run_config_mode_directory_mode_generates_hotel_catering_customer(self) -> None:
        df = build_catering_category_split_dataframe()

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            data_dir = temp_path / "datas"
            data_dir.mkdir()
            source_file = data_dir / "餐饮.xlsx"
            config_path = temp_path / "jobs.toml"

            with pd.ExcelWriter(source_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

            config_path.write_text(
                """
output_dir = "exports"
output_format = "xlsx"
input_dir = "datas"
""".strip(),
                encoding="utf-8",
            )

            with redirect_stdout(io.StringIO()):
                run_config_mode(
                    argparse.Namespace(
                        config=config_path,
                        job=["酒店餐饮客户"],
                        dry_run=False,
                        sheet_name=DEFAULT_SHEET_NAME,
                        output_format=None,
                        calculation_mode=None,
                        output_dir=None,
                    )
                )

            self.assertTrue((temp_path / "exports" / "酒店餐饮客户.xlsx").exists())
            self.assertFalse((temp_path / "exports" / "酒店宴会.xlsx").exists())
            self.assertFalse((temp_path / "exports" / "酒店自助餐.xlsx").exists())

    def test_run_directory_batch_uses_programmatic_directory_mode_without_config_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_dir = root / "input_dir"
            output_dir = root / "output_dir"
            input_dir.mkdir()
            output_dir.mkdir()

            with mock.patch("survey_stats.discover_directory_jobs") as mock_discover, mock.patch(
                "survey_stats.generate_customer_category_report_bundle"
            ) as mock_generate:
                mock_discover.return_value = DirectoryDiscoveryResult(
                    jobs=(),
                    missing_customer_type_notices=(),
                    preprocess_notices=(),
                    unmapped_customer_category_notices=(),
                )

                run_directory_batch(
                    input_dir=input_dir,
                    output_dir=output_dir,
                )

            mock_discover.assert_called_once()
            mock_generate.assert_not_called()


if __name__ == "__main__":
    unittest.main()
