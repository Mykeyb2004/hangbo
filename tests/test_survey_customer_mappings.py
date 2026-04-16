from __future__ import annotations

import unittest

from survey_customer_category_rules import (
    CUSTOMER_CATEGORY_RULE_BY_NAME,
    CUSTOMER_CATEGORY_RULES,
    DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES,
    SOURCE_FILE_TO_CATEGORY_RULE_NAMES,
)
from survey_stats import TEMPLATE_DEFINITIONS


class SurveyCustomerMappingsTest(unittest.TestCase):
    def test_directory_rules_only_reference_known_templates_when_template_is_declared(self) -> None:
        declared_template_names = {
            rule.template_name
            for rule in CUSTOMER_CATEGORY_RULES
            if rule.template_name is not None
        }
        self.assertTrue(declared_template_names.issubset(TEMPLATE_DEFINITIONS))

        for rule in CUSTOMER_CATEGORY_RULES:
            if rule.template_name is None:
                continue
            self.assertEqual(rule.name, TEMPLATE_DEFINITIONS[rule.template_name].role_name)

    def test_source_file_to_category_rule_names_matches_directory_rules(self) -> None:
        self.assertEqual(
            SOURCE_FILE_TO_CATEGORY_RULE_NAMES,
            {
                "展览.xlsx": ("展览主承办", "参展商", "专业观众"),
                "会展服务商.xlsx": ("会展服务商",),
                "会议.xlsx": (
                    "会议主承办",
                    "参会人员",
                    "酒店参会客户",
                    "酒店会议主承办",
                ),
                "旅游.xlsx": ("游客", "旅行社工作人员"),
                "酒店.xlsx": ("散客", "住宿团队"),
                "餐饮.xlsx": (
                    "商务简餐",
                    "特色美食廊",
                    "宴会",
                    "婚宴",
                    "自助餐",
                    "酒店餐饮客户",
                ),
            },
        )

    def test_display_rules_cover_directory_rules_in_sequence_order(self) -> None:
        self.assertEqual(tuple(rule.name for rule in DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES), (
            "展览主承办",
            "参展商",
            "专业观众",
            "会展服务商",
            "会议主承办",
            "参会人员",
            "商务简餐",
            "特色美食廊",
            "婚宴",
            "自助餐",
            "宴会",
            "游客",
            "旅行社工作人员",
            "散客",
            "住宿团队",
            "酒店参会客户",
            "酒店会议主承办",
            "酒店餐饮客户",
        ))

    def test_hotel_catering_aggregate_rule_metadata_is_preserved(self) -> None:
        aggregate_rule = CUSTOMER_CATEGORY_RULE_BY_NAME["酒店餐饮客户"]

        self.assertIsNone(aggregate_rule.template_name)
        self.assertEqual(
            aggregate_rule.aggregate_rule_names,
            ("酒店宴会", "酒店自助餐", "酒店餐饮-商务简餐", "酒店餐饮-宴会"),
        )


if __name__ == "__main__":
    unittest.main()
