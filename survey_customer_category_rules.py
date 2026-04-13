from __future__ import annotations

from dataclasses import dataclass
from types import MappingProxyType


@dataclass(frozen=True)
class CustomerCategoryRule:
    name: str
    customer_group: str
    customer_category: str
    source_file_name: str
    sequence_number: int | None = None
    template_name: str | None = None
    data_column: str | None = None
    data_values: tuple[str, ...] = ()
    auxiliary_column: str | None = None
    auxiliary_values: tuple[str, ...] = ()
    aggregate_rule_names: tuple[str, ...] = ()
    directory_enabled: bool = True
    note: str | None = None

    @property
    def is_aggregate(self) -> bool:
        return bool(self.aggregate_rule_names)


ALL_CUSTOMER_CATEGORY_RULES: tuple[CustomerCategoryRule, ...] = (
    CustomerCategoryRule(
        name="展览主承办",
        customer_group="一、会展客户",
        customer_category="展览活动主（承）办",
        source_file_name="展览.xlsx",
        sequence_number=1,
        template_name="organizer",
        data_column="E",
        data_values=("展览主承办",),
        auxiliary_column="C",
        auxiliary_values=("展览",),
    ),
    CustomerCategoryRule(
        name="参展商",
        customer_group="一、会展客户",
        customer_category="参展商",
        source_file_name="展览.xlsx",
        sequence_number=2,
        template_name="exhibitor",
        data_column="E",
        data_values=("参展商",),
        auxiliary_column="C",
        auxiliary_values=("展览",),
    ),
    CustomerCategoryRule(
        name="专业观众",
        customer_group="一、会展客户",
        customer_category="专业观众",
        source_file_name="展览.xlsx",
        sequence_number=3,
        template_name="visitor",
        data_column="E",
        data_values=("专业观众",),
        auxiliary_column="C",
        auxiliary_values=("展览",),
    ),
    CustomerCategoryRule(
        name="会展服务商",
        customer_group="一、会展客户",
        customer_category="会展服务商",
        source_file_name="会展服务商.xlsx",
        sequence_number=4,
        template_name="service_provider",
        data_column="D",
        data_values=("会展服务商",),
    ),
    CustomerCategoryRule(
        name="会议主承办",
        customer_group="一、会展客户",
        customer_category="会议活动主（承）办",
        source_file_name="会议.xlsx",
        sequence_number=5,
        template_name="meeting_organizer",
        data_column="E",
        data_values=("会议主承办",),
        auxiliary_column="C",
        auxiliary_values=("会议",),
    ),
    CustomerCategoryRule(
        name="参会人员",
        customer_group="一、会展客户",
        customer_category="参会客户",
        source_file_name="会议.xlsx",
        sequence_number=6,
        template_name="meeting_attendee",
        data_column="E",
        data_values=("参会人员",),
        auxiliary_column="C",
        auxiliary_values=("会议",),
    ),
    CustomerCategoryRule(
        name="商务简餐",
        customer_group="二、餐饮客户",
        customer_category="商务简餐",
        source_file_name="餐饮.xlsx",
        sequence_number=7,
        template_name="catering_business_meal",
        data_column="D",
        data_values=("商务简餐",),
        auxiliary_column="C",
        auxiliary_values=("餐饮",),
    ),
    CustomerCategoryRule(
        name="特色美食廊",
        customer_group="二、餐饮客户",
        customer_category="特色美食廊",
        source_file_name="餐饮.xlsx",
        sequence_number=8,
        template_name="catering_food_hall",
        data_column="D",
        data_values=("特色美食廊",),
        auxiliary_column="C",
        auxiliary_values=("餐饮",),
    ),
    CustomerCategoryRule(
        name="宴会",
        customer_group="二、餐饮客户",
        customer_category="宴会",
        source_file_name="餐饮.xlsx",
        sequence_number=9,
        template_name="catering_banquet",
        data_column="D",
        data_values=("宴会",),
        auxiliary_column="C",
        auxiliary_values=("餐饮",),
    ),
    CustomerCategoryRule(
        name="婚宴",
        customer_group="二、餐饮客户",
        customer_category="婚宴",
        source_file_name="餐饮.xlsx",
        sequence_number=10,
        template_name="catering_wedding_banquet",
        data_column="D",
        data_values=("婚宴",),
        auxiliary_column="C",
        auxiliary_values=("餐饮",),
    ),
    CustomerCategoryRule(
        name="自助餐",
        customer_group="二、餐饮客户",
        customer_category="自助餐",
        source_file_name="餐饮.xlsx",
        sequence_number=11,
        template_name="catering_buffet",
        data_column="D",
        data_values=("自助餐",),
        auxiliary_column="C",
        auxiliary_values=("餐饮",),
    ),
    CustomerCategoryRule(
        name="旅行社工作人员",
        customer_group="三、G20峰会体验馆",
        customer_category="旅行社工作人员",
        source_file_name="旅游.xlsx",
        sequence_number=12,
        template_name="travel_staff",
        data_column="C",
        data_values=("旅行社工作人员",),
    ),
    CustomerCategoryRule(
        name="游客",
        customer_group="三、G20峰会体验馆",
        customer_category="游客",
        source_file_name="旅游.xlsx",
        sequence_number=13,
        template_name="tourist",
        data_column="C",
        data_values=("游客",),
    ),
    CustomerCategoryRule(
        name="散客",
        customer_group="五、酒店客户",
        customer_category="散客",
        source_file_name="酒店.xlsx",
        sequence_number=14,
        template_name="hotel_individual_guest",
        data_column="C",
        data_values=("散客",),
    ),
    CustomerCategoryRule(
        name="住宿团队",
        customer_group="五、酒店客户",
        customer_category="住宿团队",
        source_file_name="酒店.xlsx",
        sequence_number=15,
        template_name="hotel_group_guest",
        data_column="C",
        data_values=("住宿团队",),
    ),
    CustomerCategoryRule(
        name="酒店会议主承办",
        customer_group="五、酒店客户",
        customer_category="酒店会议活动主（承）办",
        source_file_name="会议.xlsx",
        sequence_number=16,
        template_name="hotel_meeting_organizer",
        data_column="E",
        data_values=("酒店会议主承办",),
        auxiliary_column="C",
        auxiliary_values=("酒店会议",),
    ),
    CustomerCategoryRule(
        name="酒店参会客户",
        customer_group="五、酒店客户",
        customer_category="酒店参会客户",
        source_file_name="会议.xlsx",
        sequence_number=17,
        template_name="hotel_meeting_attendee",
        data_column="E",
        data_values=("酒店参会客户", "参会人员"),
        auxiliary_column="C",
        auxiliary_values=("酒店会议",),
    ),
    CustomerCategoryRule(
        name="酒店宴会",
        customer_group="五、酒店客户",
        customer_category="酒店餐饮客户",
        source_file_name="餐饮.xlsx",
        template_name="catering_hotel_banquet",
        data_column="D",
        data_values=("酒店宴会",),
        auxiliary_column="C",
        auxiliary_values=("酒店餐饮",),
        directory_enabled=False,
        note="内部组件规则，用于生成酒店餐饮客户聚合结果。",
    ),
    CustomerCategoryRule(
        name="酒店自助餐",
        customer_group="五、酒店客户",
        customer_category="酒店餐饮客户",
        source_file_name="餐饮.xlsx",
        template_name="catering_hotel_buffet",
        data_column="D",
        data_values=("酒店自助餐",),
        auxiliary_column="C",
        auxiliary_values=("酒店餐饮",),
        directory_enabled=False,
        note="内部组件规则，用于生成酒店餐饮客户聚合结果。",
    ),
    CustomerCategoryRule(
        name="酒店餐饮-商务简餐",
        customer_group="五、酒店客户",
        customer_category="酒店餐饮客户",
        source_file_name="餐饮.xlsx",
        template_name="catering_business_meal",
        data_column="D",
        data_values=("商务简餐",),
        auxiliary_column="C",
        auxiliary_values=("酒店餐饮",),
        directory_enabled=False,
        note="内部组件规则，用于生成酒店餐饮客户聚合结果。",
    ),
    CustomerCategoryRule(
        name="酒店餐饮-宴会",
        customer_group="五、酒店客户",
        customer_category="酒店餐饮客户",
        source_file_name="餐饮.xlsx",
        template_name="catering_banquet",
        data_column="D",
        data_values=("宴会",),
        auxiliary_column="C",
        auxiliary_values=("酒店餐饮",),
        directory_enabled=False,
        note="内部组件规则，用于生成酒店餐饮客户聚合结果。",
    ),
    CustomerCategoryRule(
        name="酒店餐饮客户",
        customer_group="五、酒店客户",
        customer_category="酒店餐饮客户",
        source_file_name="餐饮.xlsx",
        sequence_number=18,
        data_column="D",
        data_values=("酒店自助餐", "酒店宴会", "商务简餐", "宴会"),
        auxiliary_column="C",
        auxiliary_values=("酒店餐饮",),
        aggregate_rule_names=("酒店宴会", "酒店自助餐", "酒店餐饮-商务简餐", "酒店餐饮-宴会"),
    ),
)


CUSTOMER_CATEGORY_RULES = tuple(
    rule for rule in ALL_CUSTOMER_CATEGORY_RULES if rule.directory_enabled
)


DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES = tuple(
    sorted(
        (
            rule
            for rule in CUSTOMER_CATEGORY_RULES
            if rule.sequence_number is not None
        ),
        key=lambda rule: rule.sequence_number,
    )
)


CUSTOMER_CATEGORY_RULE_BY_NAME = MappingProxyType(
    {rule.name: rule for rule in ALL_CUSTOMER_CATEGORY_RULES}
)


SOURCE_FILE_TO_CATEGORY_RULE_NAMES = MappingProxyType(
    {
        source_file_name: tuple(
            rule.name
            for rule in CUSTOMER_CATEGORY_RULES
            if rule.source_file_name == source_file_name
        )
        for source_file_name in dict.fromkeys(rule.source_file_name for rule in CUSTOMER_CATEGORY_RULES)
    }
)


__all__ = [
    "CustomerCategoryRule",
    "ALL_CUSTOMER_CATEGORY_RULES",
    "CUSTOMER_CATEGORY_RULES",
    "DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES",
    "CUSTOMER_CATEGORY_RULE_BY_NAME",
    "SOURCE_FILE_TO_CATEGORY_RULE_NAMES",
]
