from __future__ import annotations

import unittest

from survey_customer_mappings import (
    CUSTOMER_TYPE_MAPPING_BY_TEMPLATE,
    DOCUMENT_DISPLAY_NAME_TO_TEMPLATE_NAMES,
    SOURCE_FILE_TO_TEMPLATE_NAMES,
    STANDARD_CUSTOMER_TYPE_MAPPINGS,
)
from survey_stats import TEMPLATE_DEFINITIONS


class SurveyCustomerMappingsTest(unittest.TestCase):
    def test_standard_customer_type_mappings_cover_all_supported_templates(self) -> None:
        mapped_template_names = {mapping.template_name for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS}

        self.assertEqual(mapped_template_names, set(TEMPLATE_DEFINITIONS))
        self.assertEqual(len(STANDARD_CUSTOMER_TYPE_MAPPINGS), len(TEMPLATE_DEFINITIONS))

    def test_standard_customer_type_mappings_use_current_template_role_names(self) -> None:
        for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS:
            self.assertEqual(
                mapping.template_role_name,
                TEMPLATE_DEFINITIONS[mapping.template_name].role_name,
            )

    def test_source_file_to_template_names_matches_standard_buckets(self) -> None:
        self.assertEqual(
            SOURCE_FILE_TO_TEMPLATE_NAMES,
            {
                "展览.xlsx": ("organizer", "exhibitor", "visitor"),
                "会展服务商.xlsx": ("service_provider",),
                "会议.xlsx": (
                    "meeting_organizer",
                    "hotel_meeting_organizer",
                    "hotel_meeting_attendee",
                    "meeting_attendee",
                ),
                "酒店.xlsx": ("hotel_individual_guest", "hotel_group_guest"),
                "餐饮.xlsx": (
                    "catering_food_hall",
                    "catering_business_meal",
                    "catering_tour_meal",
                    "catering_banquet",
                    "catering_wedding_banquet",
                    "catering_buffet",
                    "catering_hotel_banquet",
                    "catering_hotel_buffet",
                ),
            },
        )

    def test_special_document_display_name_cases_are_preserved(self) -> None:
        self.assertIsNone(CUSTOMER_TYPE_MAPPING_BY_TEMPLATE["catering_tour_meal"].document_display_name)
        self.assertIn(
            "未单列",
            CUSTOMER_TYPE_MAPPING_BY_TEMPLATE["catering_tour_meal"].note or "",
        )
        self.assertEqual(
            DOCUMENT_DISPLAY_NAME_TO_TEMPLATE_NAMES["餐饮客户"],
            ("catering_hotel_banquet", "catering_hotel_buffet"),
        )


if __name__ == "__main__":
    unittest.main()
