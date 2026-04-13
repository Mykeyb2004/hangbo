from __future__ import annotations

from dataclasses import dataclass
from types import MappingProxyType


@dataclass(frozen=True)
class CustomerTypeMapping:
    template_name: str
    template_role_name: str
    document_display_name: str | None
    source_file_name: str
    note: str | None = None


STANDARD_CUSTOMER_TYPE_MAPPINGS: tuple[CustomerTypeMapping, ...] = (
    CustomerTypeMapping("organizer", "展览主承办", "展览活动主（承）办", "展览.xlsx"),
    CustomerTypeMapping("exhibitor", "参展商", "参展商", "展览.xlsx"),
    CustomerTypeMapping("visitor", "专业观众", "专业观众", "展览.xlsx"),
    CustomerTypeMapping("service_provider", "会展服务商", "会展服务商", "会展服务商.xlsx"),
    CustomerTypeMapping("meeting_organizer", "会议主承办", "会议活动主（承）办", "会议.xlsx"),
    CustomerTypeMapping(
        "hotel_meeting_organizer",
        "酒店会议主承办",
        "酒店会议活动主（承）办",
        "会议.xlsx",
    ),
    CustomerTypeMapping("hotel_meeting_attendee", "酒店参会客户", "酒店参会客户", "会议.xlsx"),
    CustomerTypeMapping("meeting_attendee", "参会人员", "参会客户", "会议.xlsx"),
    CustomerTypeMapping("travel_staff", "旅行社工作人员", "旅行社工作人员", "旅游.xlsx"),
    CustomerTypeMapping("tourist", "游客", "游客", "旅游.xlsx"),
    CustomerTypeMapping("hotel_individual_guest", "散客", "散客", "酒店.xlsx"),
    CustomerTypeMapping("hotel_group_guest", "住宿团队", "住宿团队", "酒店.xlsx"),
    CustomerTypeMapping("catering_food_hall", "特色美食廊", "特色美食廊", "餐饮.xlsx"),
    CustomerTypeMapping("catering_business_meal", "商务简餐", "商务简餐", "餐饮.xlsx"),
    CustomerTypeMapping(
        "catering_tour_meal",
        "旅游团餐",
        None,
        "餐饮.xlsx",
        "当前文档汇总表未单列该客户类型，但 survey_stats.py 已支持。",
    ),
    CustomerTypeMapping("catering_banquet", "宴会", "宴会", "餐饮.xlsx"),
    CustomerTypeMapping("catering_wedding_banquet", "婚宴", "婚宴", "餐饮.xlsx"),
    CustomerTypeMapping("catering_buffet", "自助餐", "自助餐", "餐饮.xlsx"),
    CustomerTypeMapping(
        "catering_hotel_banquet",
        "酒店宴会",
        "餐饮客户",
        "餐饮.xlsx",
        "文档中并入“餐饮客户”合并行，不单独展示。",
    ),
    CustomerTypeMapping(
        "catering_hotel_buffet",
        "酒店自助餐",
        "餐饮客户",
        "餐饮.xlsx",
        "文档中并入“餐饮客户”合并行，不单独展示。",
    ),
)


CUSTOMER_TYPE_MAPPING_BY_TEMPLATE = MappingProxyType(
    {mapping.template_name: mapping for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS}
)


SOURCE_FILE_TO_TEMPLATE_NAMES = MappingProxyType(
    {
        source_file_name: tuple(
            mapping.template_name
            for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS
            if mapping.source_file_name == source_file_name
        )
        for source_file_name in dict.fromkeys(
            mapping.source_file_name for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS
        )
    }
)


DOCUMENT_DISPLAY_NAME_TO_TEMPLATE_NAMES = MappingProxyType(
    {
        document_display_name: tuple(
            mapping.template_name
            for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS
            if mapping.document_display_name == document_display_name
        )
        for document_display_name in dict.fromkeys(
            mapping.document_display_name
            for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS
            if mapping.document_display_name is not None
        )
    }
)


__all__ = [
    "CustomerTypeMapping",
    "STANDARD_CUSTOMER_TYPE_MAPPINGS",
    "CUSTOMER_TYPE_MAPPING_BY_TEMPLATE",
    "SOURCE_FILE_TO_TEMPLATE_NAMES",
    "DOCUMENT_DISPLAY_NAME_TO_TEMPLATE_NAMES",
]
