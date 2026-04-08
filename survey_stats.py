from __future__ import annotations

import argparse
import tomllib
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

DEFAULT_SHEET_NAME = "问卷数据"
DEFAULT_OUTPUT_FORMAT = "xlsx"
DEFAULT_ROLE_COLUMN = "E"
VALID_OUTPUT_FORMATS = ("xlsx", "csv", "md")

ORGANIZER_ROLE_NAME = "展览主承办"
EXHIBITOR_ROLE_NAME = "参展商"
VISITOR_ROLE_NAME = "专业观众"
SERVICE_PROVIDER_ROLE_NAME = "会展服务商"
MEETING_ORGANIZER_ROLE_NAME = "会议主承办"
HOTEL_MEETING_ORGANIZER_ROLE_NAME = "酒店会议主承办"
HOTEL_MEETING_ATTENDEE_ROLE_NAME = "酒店参会客户"
MEETING_ATTENDEE_ROLE_NAME = "参会人员"
CATERING_FOOD_HALL_ROLE_NAME = "特色美食廊"
CATERING_BUSINESS_MEAL_ROLE_NAME = "商务简餐"
CATERING_TOUR_MEAL_ROLE_NAME = "旅游团餐"
CATERING_BANQUET_ROLE_NAME = "宴会"
CATERING_WEDDING_BANQUET_ROLE_NAME = "婚宴"
CATERING_BUFFET_ROLE_NAME = "自助餐"
CATERING_HOTEL_BANQUET_ROLE_NAME = "酒店宴会"
CATERING_HOTEL_BUFFET_ROLE_NAME = "酒店自助餐"
CATERING_ROLE_COLUMN = "D"

OVERALL_FILL = PatternFill(fill_type="solid", start_color="F4B183", end_color="F4B183")
SECTION_FILL = PatternFill(fill_type="solid", start_color="C6EFD1", end_color="C6EFD1")
EMPHASIS_FONT = Font(bold=True)
CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center")


@dataclass(frozen=True)
class MetricDefinition:
    name: str
    satisfaction_column: str
    importance_column: str
    satisfaction_gt_zero_column: str | None = None
    satisfaction_lt_eleven_column: str | None = None
    importance_gt_zero_column: str | None = None
    importance_lt_eleven_column: str | None = None


@dataclass(frozen=True)
class SectionDefinition:
    name: str
    metrics: tuple[MetricDefinition, ...]


@dataclass(frozen=True)
class RoleDefinition:
    role_name: str
    role_column: str
    sections: tuple[SectionDefinition, ...]


@dataclass(frozen=True)
class JobConfig:
    name: str
    path: Path
    sheet_name: str
    template_name: str
    role_name: str
    output_name: str
    output_format: str | None = None


@dataclass(frozen=True)
class BatchConfig:
    config_path: Path
    output_dir: Path
    output_format: str
    jobs: tuple[JobConfig, ...]


@dataclass(frozen=True)
class MetricResult:
    name: str
    satisfaction: float | None
    importance: float | None


@dataclass(frozen=True)
class SectionResult:
    name: str
    satisfaction: float | None
    importance: float | None
    metrics: tuple[MetricResult, ...]


@dataclass(frozen=True)
class SurveyStatistics:
    role_name: str
    satisfaction: float | None
    importance: float | None
    sections: tuple[SectionResult, ...]
    matched_row_count: int


@dataclass(frozen=True)
class GeneratedReport:
    result_df: pd.DataFrame
    output_path: Path
    stats: SurveyStatistics


@dataclass(frozen=True)
class MissingGroupNotice:
    job_name: str
    input_path: Path
    sheet_name: str


ORGANIZER_TEMPLATE = RoleDefinition(
    role_name=ORGANIZER_ROLE_NAME,
    role_column=DEFAULT_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "会展服务",
            (
                MetricDefinition("销售经理服务态度", "AE", "AF"),
                MetricDefinition("销售经理业务能力", "AG", "AH"),
                MetricDefinition("现场对接协调沟通", "AI", "AJ"),
                MetricDefinition("现场对接服务效率", "AK", "AL"),
                MetricDefinition("工作人员仪容仪表", "AM", "AN"),
                MetricDefinition("工作人员服务态度", "AO", "AP"),
                MetricDefinition("工作人员业务技能", "AQ", "AR"),
                MetricDefinition("报馆流程及服务", "AW", "AY"),
                MetricDefinition("现场布置与搭建", "BA", "BB"),
                MetricDefinition("在线服务渠道便捷", "AY", "AZ"),
                MetricDefinition("撤展服务", "BC", "BD"),
                MetricDefinition("展后回访", "BE", "BF"),
                MetricDefinition("整体协调和配合", "AU", "AV"),
                MetricDefinition("服务响应及时性", "AS", "AT"),
                MetricDefinition("设备租赁", "AA", "AB"),
            ),
        ),
        SectionDefinition(
            "硬件设施",
            (
                MetricDefinition("园区停车方便", "O", "P"),
                MetricDefinition("交通便利，容易到达", "M", "N"),
                MetricDefinition("交通流线", "K", "K"),
                MetricDefinition("货运通道", "U", "U"),
                MetricDefinition("标识标牌清晰", "S", "T"),
                MetricDefinition("设施设备齐全", "W", "W"),
                MetricDefinition("展厅使用情况", "Y", "Z"),
            ),
        ),
        SectionDefinition(
            "配套服务",
            (
                MetricDefinition("餐饮服务", "BG", "BH"),
                MetricDefinition("客房服务", "BI", "BJ"),
                MetricDefinition("安保服务", "BK", "BK"),
                MetricDefinition("保洁服务", "BM", "BN"),
            ),
        ),
        SectionDefinition(
            "智慧场馆",
            (
                MetricDefinition("杭州国博APP", "BR", "BS"),
                MetricDefinition("室内导航系统", "BU", "BV"),
                MetricDefinition("寻车系统", "BX", "BY"),
                MetricDefinition("云上看馆", "CA", "CB"),
            ),
        ),
    ),
)

EXHIBITOR_TEMPLATE = RoleDefinition(
    role_name=EXHIBITOR_ROLE_NAME,
    role_column=DEFAULT_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "会场服务",
            (
                MetricDefinition("工作人员仪容仪表", "AM", "AN"),
                MetricDefinition("工作人员服务态度", "AO", "AP"),
                MetricDefinition("工作人员业务技能", "AQ", "AR"),
                MetricDefinition("现场布置与搭建", "BA", "BB"),
                MetricDefinition("接待引导服务", "BO", "BP"),
            ),
        ),
        SectionDefinition(
            "硬件设施",
            (
                MetricDefinition("园区停车方便", "O", "P"),
                MetricDefinition("交通便利，容易到达", "M", "N"),
                MetricDefinition("标识标牌清晰", "S", "T"),
                MetricDefinition("设施设备齐全", "W", "X"),
                MetricDefinition("展厅使用情况", "Y", "Z"),
                MetricDefinition("参展环境", "AC", "AD"),
            ),
        ),
        SectionDefinition(
            "配套服务",
            (
                MetricDefinition("餐饮服务", "BG", "BH"),
                MetricDefinition("客房服务", "BI", "BJ"),
                MetricDefinition("安保服务", "BK", "BL"),
                MetricDefinition("保洁服务", "BM", "BN"),
            ),
        ),
        SectionDefinition(
            "智慧场馆",
            (
                MetricDefinition("杭州国博APP", "BR", "BS"),
                MetricDefinition("室内导航系统", "BU", "BV"),
                MetricDefinition("寻车系统", "BX", "BY"),
                MetricDefinition("云上看馆", "CA", "CB"),
            ),
        ),
    ),
)

VISITOR_TEMPLATE = RoleDefinition(
    role_name=VISITOR_ROLE_NAME,
    role_column=DEFAULT_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "会展服务",
            (
                MetricDefinition("工作人员仪容仪表", "AM", "AN"),
                MetricDefinition("工作人员服务态度", "AO", "AP"),
                MetricDefinition("工作人员业务技能", "AQ", "AR"),
                MetricDefinition("接待引导服务", "BO", "BP"),
            ),
        ),
        SectionDefinition(
            "硬件设施",
            (
                MetricDefinition("展会路线安排", "Q", "R"),
                MetricDefinition("园区停车方便", "O", "P"),
                MetricDefinition("交通便利，容易到达", "M", "N"),
                MetricDefinition("标识标牌清晰", "S", "T"),
                MetricDefinition("设施设备齐全", "W", "W"),
                MetricDefinition("展厅使用情况", "Y", "Z"),
                MetricDefinition("参展环境", "AC", "AD"),
            ),
        ),
        SectionDefinition(
            "配套服务",
            (
                MetricDefinition("餐饮服务", "BG", "BH"),
                MetricDefinition("客房服务", "BI", "BJ"),
                MetricDefinition("安保服务", "BK", "BK"),
                MetricDefinition("保洁服务", "BM", "BN"),
            ),
        ),
        SectionDefinition(
            "智慧场馆",
            (
                MetricDefinition("杭州国博APP", "BR", "BS"),
                MetricDefinition("室内导航系统", "BU", "BV"),
                MetricDefinition("寻车系统", "BX", "BY"),
                MetricDefinition("云上看馆", "CA", "CB"),
            ),
        ),
    ),
)

SERVICE_PROVIDER_TEMPLATE = RoleDefinition(
    role_name=SERVICE_PROVIDER_ROLE_NAME,
    role_column="D",
    sections=(
        SectionDefinition(
            "会展服务",
            (
                MetricDefinition("现场对接协调沟通", "AN", "AO"),
                MetricDefinition("现场对接服务效率", "AL", "AM"),
                MetricDefinition("工作人员仪容仪表", "AF", "AG"),
                MetricDefinition(
                    "工作人员服务态度",
                    "AH",
                    "AI",
                    satisfaction_gt_zero_column="AH",
                    satisfaction_lt_eleven_column="AL",
                ),
                MetricDefinition("工作人员业务技能", "AJ", "AK"),
                MetricDefinition("现场布置与搭建", "Z", "AA"),
                MetricDefinition("报馆流程及服务", "AB", "AC"),
                MetricDefinition("在线服务渠道便捷", "AD", "AE"),
                MetricDefinition("工程设备服务", "X", "Y"),
                MetricDefinition("整体协调和配合", "AR", "AS"),
                MetricDefinition("服务响应及时性", "AP", "AQ"),
                MetricDefinition("撤展服务", "AT", "AU"),
                MetricDefinition("展后回访", "AV", "AX"),
            ),
        ),
        SectionDefinition(
            "硬件设施",
            (
                MetricDefinition("交通便利，容易到达", "H", "I"),
                MetricDefinition("标识标牌清晰", "R", "S"),
                MetricDefinition("交通流线", "L", "M"),
                MetricDefinition("园区停车便利", "J", "K"),
                MetricDefinition("货运通道", "N", "O"),
                MetricDefinition("快递物流", "P", "Q"),
                MetricDefinition("休息空间", "V", "W"),
                MetricDefinition("设施设备齐全", "T", "U"),
            ),
        ),
        SectionDefinition(
            "配套服务",
            (
                MetricDefinition("餐饮服务", "AX", "AY"),
                MetricDefinition("客房服务", "AZ", "BA"),
                MetricDefinition("安保服务", "BB", "BC"),
                MetricDefinition("保洁服务", "BD", "BE"),
            ),
        ),
        SectionDefinition(
            "智慧场馆",
            (
                MetricDefinition("室内导航系统", "BJ", "BK"),
                MetricDefinition("寻车系统", "BM", "BN"),
                MetricDefinition("杭州国博APP", "BG", "BH"),
                MetricDefinition("云上看馆", "BP", "BQ"),
            ),
        ),
    ),
)

MEETING_ORGANIZER_SERVICE_SECTION = SectionDefinition(
    "会展服务",
    (
        MetricDefinition("销售经理服务态度", "AS", "AT"),
        MetricDefinition("销售经理业务能力", "AU", "AV"),
        MetricDefinition("现场对接协调沟通", "AW", "AX"),
        MetricDefinition("现场对接服务效率", "AY", "AZ"),
        MetricDefinition("工作人员仪容仪表", "AE", "AF"),
        MetricDefinition("工作人员服务态度", "AI", "AJ"),
        MetricDefinition("工作人员业务技能", "AG", "AH"),
        MetricDefinition("现场布置与搭建", "BC", "BD"),
        MetricDefinition("茶歇服务与品质", "BG", "BH"),
        MetricDefinition("在线服务渠道便捷", "BI", "BJ"),
        MetricDefinition("报馆流程及服务", "BA", "BB"),
        MetricDefinition("工程设施设备", "BE", "BF"),
        MetricDefinition("撤会服务", "BO", "BP"),
        MetricDefinition("会后回访", "BQ", "BR"),
        MetricDefinition("整体协调和配合", "BM", "BN"),
        MetricDefinition("服务响应及时性", "BK", "BL"),
    ),
)

MEETING_HARDWARE_SECTION = SectionDefinition(
    "硬件设施",
    (
        MetricDefinition("交通便利，容易到达", "O", "P"),
        MetricDefinition("标识标牌清晰", "S", "T"),
        MetricDefinition("会议厅室匹配性，方便性", "W", "X"),
        MetricDefinition("设施设备齐全", "U", "V"),
        MetricDefinition("休息空间", "AA", "AB"),
        MetricDefinition("园区停车方便", "Q", "R"),
        MetricDefinition("AV效果", "Y", "Z"),
        MetricDefinition("交通流线", "M", "N"),
    ),
)

MEETING_SUPPORT_SECTION = SectionDefinition(
    "配套服务",
    (
        MetricDefinition("餐饮服务", "AK", "AL"),
        MetricDefinition("客房服务", "AM", "AN"),
        MetricDefinition("安保服务", "AQ", "AR"),
        MetricDefinition("保洁服务", "AO", "AP"),
    ),
)

MEETING_SMART_SECTION = SectionDefinition(
    "智慧场馆",
    (
        MetricDefinition("杭州国博APP", "BV", "BW"),
        MetricDefinition("室内导航系统", "BY", "BZ"),
        MetricDefinition("寻车系统", "CB", "CC"),
        MetricDefinition("云上看馆", "CE", "CF"),
    ),
)

MEETING_ATTENDEE_SERVICE_SECTION = SectionDefinition(
    "会展服务",
    (
        MetricDefinition("工作人员仪容仪表", "AE", "AF"),
        MetricDefinition("工作人员业务技能", "AG", "AH"),
        MetricDefinition("工作人员服务态度", "AI", "AJ"),
        MetricDefinition("接待引导服务", "BS", "BT"),
        MetricDefinition("茶歇服务品质", "BG", "BH"),
    ),
)

MEETING_ATTENDEE_HARDWARE_SECTION = SectionDefinition(
    "硬件设施",
    (
        MetricDefinition("交通便利，容易到达", "O", "P"),
        MetricDefinition("标识标牌清晰", "S", "T"),
        MetricDefinition("会议厅室匹配性，方便性", "W", "X"),
        MetricDefinition("设施设备齐全", "U", "V"),
        MetricDefinition("休息空间", "AA", "AB"),
        MetricDefinition("园区停车方便", "Q", "R"),
        MetricDefinition("AV效果", "Y", "Z"),
        MetricDefinition("参会环境", "AC", "AD"),
    ),
)

MEETING_ORGANIZER_TEMPLATE = RoleDefinition(
    role_name=MEETING_ORGANIZER_ROLE_NAME,
    role_column=DEFAULT_ROLE_COLUMN,
    sections=(
        MEETING_ORGANIZER_SERVICE_SECTION,
        MEETING_HARDWARE_SECTION,
        MEETING_SUPPORT_SECTION,
        MEETING_SMART_SECTION,
    ),
)

HOTEL_MEETING_ORGANIZER_TEMPLATE = RoleDefinition(
    role_name=HOTEL_MEETING_ORGANIZER_ROLE_NAME,
    role_column=DEFAULT_ROLE_COLUMN,
    sections=(
        MEETING_ORGANIZER_SERVICE_SECTION,
        MEETING_HARDWARE_SECTION,
        MEETING_SUPPORT_SECTION,
        MEETING_SMART_SECTION,
    ),
)

HOTEL_MEETING_ATTENDEE_TEMPLATE = RoleDefinition(
    role_name=HOTEL_MEETING_ATTENDEE_ROLE_NAME,
    role_column=DEFAULT_ROLE_COLUMN,
    sections=(
        MEETING_ATTENDEE_SERVICE_SECTION,
        MEETING_ATTENDEE_HARDWARE_SECTION,
        MEETING_SUPPORT_SECTION,
        MEETING_SMART_SECTION,
    ),
)

MEETING_ATTENDEE_TEMPLATE = RoleDefinition(
    role_name=MEETING_ATTENDEE_ROLE_NAME,
    role_column=DEFAULT_ROLE_COLUMN,
    sections=(
        MEETING_ATTENDEE_SERVICE_SECTION,
        MEETING_ATTENDEE_HARDWARE_SECTION,
        MEETING_SUPPORT_SECTION,
        MEETING_SMART_SECTION,
    ),
)

CATERING_BASIC_HARDWARE_SECTION = SectionDefinition(
    "硬件设施",
    (
        MetricDefinition("园区停车方便", "I", "J"),
        MetricDefinition("交通便利，容易到达", "G", "H"),
        MetricDefinition("标识标牌清晰", "K", "L"),
    ),
)

CATERING_STANDARD_SMART_SECTION = SectionDefinition(
    "智慧场馆",
    (
        MetricDefinition("杭州国博APP", "AV", "AW"),
        MetricDefinition("寻车系统", "BB", "BC"),
        MetricDefinition("室内导航系统", "AY", "AZ"),
        MetricDefinition("云上看馆", "BE", "BF"),
    ),
)

CATERING_BUFFET_SMART_SECTION = SectionDefinition(
    "智慧场馆",
    (
        MetricDefinition("杭州国博APP", "AV", "AW"),
        MetricDefinition("室内导航系统", "AY", "AZ"),
        MetricDefinition("寻车系统", "BB", "BC"),
        MetricDefinition("云上看馆", "BE", "BF"),
    ),
)

CATERING_BANQUET_HARDWARE_SECTION = SectionDefinition(
    "硬件设施",
    (
        MetricDefinition("标识标牌清晰", "K", "L"),
        MetricDefinition("宴会视听设备", "Q", "R"),
        MetricDefinition("园区停车方便", "I", "J"),
        MetricDefinition("交通便利，容易到达", "G", "H"),
    ),
)

CATERING_WEDDING_HARDWARE_SECTION = SectionDefinition(
    "硬件设施",
    (
        MetricDefinition("园区停车方便", "I", "J"),
        MetricDefinition("交通便利，容易到达", "G", "H"),
        MetricDefinition("标识标牌清晰", "K", "L"),
        MetricDefinition("宴会视听设备", "Q", "R"),
    ),
)

CATERING_FOOD_HALL_TEMPLATE = RoleDefinition(
    role_name=CATERING_FOOD_HALL_ROLE_NAME,
    role_column=CATERING_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "餐饮服务",
            (
                MetricDefinition("菜肴品质", "Y", "Z"),
                MetricDefinition("菜品口味口相", "AA", "AB"),
                MetricDefinition("菜品地方特色", "U", "V"),
                MetricDefinition("菜肴份量", "AI", "AJ"),
                MetricDefinition("工作人员服务品质", "AM", "AN"),
                MetricDefinition("就餐区域温度", "M", "N"),
                MetricDefinition("就餐区域卫生", "O", "P"),
            ),
        ),
        CATERING_BASIC_HARDWARE_SECTION,
        CATERING_STANDARD_SMART_SECTION,
    ),
)

CATERING_BUSINESS_MEAL_TEMPLATE = RoleDefinition(
    role_name=CATERING_BUSINESS_MEAL_ROLE_NAME,
    role_column=CATERING_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "餐饮服务",
            (
                MetricDefinition("菜品温度", "AK", "AL"),
                MetricDefinition("菜品分量", "AI", "AJ"),
                MetricDefinition("供应速度", "AC", "AD"),
                MetricDefinition("菜品卫生", "AE", "AF"),
                MetricDefinition("工作人员服务品质", "AM", "AN"),
                MetricDefinition("就餐区域温度", "M", "N"),
                MetricDefinition("就餐区域卫生", "O", "P"),
            ),
        ),
        CATERING_BASIC_HARDWARE_SECTION,
        CATERING_STANDARD_SMART_SECTION,
    ),
)

CATERING_TOUR_MEAL_TEMPLATE = RoleDefinition(
    role_name=CATERING_TOUR_MEAL_ROLE_NAME,
    role_column=CATERING_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "餐饮服务",
            (
                MetricDefinition("菜品温度", "AK", "AL"),
                MetricDefinition("菜品分量", "AI", "AJ"),
                MetricDefinition("供应速度", "AC", "AD"),
                MetricDefinition("菜品卫生", "AE", "AF"),
                MetricDefinition("工作人员服务品质", "AM", "AN"),
                MetricDefinition("就餐区域温度", "M", "N"),
                MetricDefinition("就餐区域卫生", "O", "P"),
            ),
        ),
        CATERING_BASIC_HARDWARE_SECTION,
        CATERING_STANDARD_SMART_SECTION,
    ),
)

CATERING_BANQUET_TEMPLATE = RoleDefinition(
    role_name=CATERING_BANQUET_ROLE_NAME,
    role_column=CATERING_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "餐饮服务",
            (
                MetricDefinition("菜品温度", "AK", "AL"),
                MetricDefinition("菜肴品种", "AG", "AH"),
                MetricDefinition("菜品供应速度", "AC", "AD"),
                MetricDefinition("菜品口味品相", "AA", "AB"),
                MetricDefinition("菜品地方特色", "U", "V"),
                MetricDefinition("接待指引", "S", "T"),
                MetricDefinition("就餐区域温度", "M", "N"),
                MetricDefinition("就餐区域卫生", "O", "P"),
                MetricDefinition("工作人员服务品质", "AM", "AN"),
            ),
        ),
        CATERING_BANQUET_HARDWARE_SECTION,
        CATERING_STANDARD_SMART_SECTION,
    ),
)

CATERING_WEDDING_BANQUET_TEMPLATE = RoleDefinition(
    role_name=CATERING_WEDDING_BANQUET_ROLE_NAME,
    role_column=CATERING_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "餐饮服务",
            (
                MetricDefinition("接待指引", "S", "T"),
                MetricDefinition("菜品口味品相", "AA", "AB"),
                MetricDefinition("菜品地方特色", "U", "V"),
                MetricDefinition("菜品供应速度", "AC", "AD"),
                MetricDefinition("就餐区域温度", "M", "N"),
                MetricDefinition("就餐区域卫生", "O", "P"),
                MetricDefinition("工作人员仪容仪表", "AQ", "AR"),
                MetricDefinition("工作人员服务态度", "AO", "AP"),
                MetricDefinition("工作人员业务技能", "AS", "AT"),
                MetricDefinition("婚宴茶歇", "W", "X"),
                MetricDefinition("菜品温度", "AK", "AL"),
            ),
        ),
        CATERING_WEDDING_HARDWARE_SECTION,
        CATERING_STANDARD_SMART_SECTION,
    ),
)

CATERING_BUFFET_TEMPLATE = RoleDefinition(
    role_name=CATERING_BUFFET_ROLE_NAME,
    role_column=CATERING_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "餐饮服务",
            (
                MetricDefinition("菜肴品种", "AG", "AH"),
                MetricDefinition("菜品供应速度", "AC", "AD"),
                MetricDefinition("菜品口味口相", "AA", "AB"),
                MetricDefinition("菜品地方特色", "U", "V"),
                MetricDefinition("接待指引", "S", "T"),
                MetricDefinition("工作人员服务品质", "AM", "AN"),
                MetricDefinition("就餐区域温度", "M", "N"),
                MetricDefinition("就餐区域卫生", "O", "P"),
            ),
        ),
        CATERING_BASIC_HARDWARE_SECTION,
        CATERING_BUFFET_SMART_SECTION,
    ),
)

CATERING_HOTEL_BANQUET_TEMPLATE = RoleDefinition(
    role_name=CATERING_HOTEL_BANQUET_ROLE_NAME,
    role_column=CATERING_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "餐饮服务",
            (
                MetricDefinition("菜品温度", "AK", "AL"),
                MetricDefinition("菜肴品种", "AG", "AH"),
                MetricDefinition("菜品供应速度", "AC", "AD"),
                MetricDefinition("菜品口味品相", "AA", "AB"),
                MetricDefinition("菜品地方特色", "U", "V"),
                MetricDefinition("接待指引", "S", "T"),
                MetricDefinition("就餐区域温度", "M", "N"),
                MetricDefinition("就餐区域卫生", "O", "P"),
                MetricDefinition("工作人员服务品质", "AM", "AN"),
            ),
        ),
        CATERING_BANQUET_HARDWARE_SECTION,
        CATERING_STANDARD_SMART_SECTION,
    ),
)

CATERING_HOTEL_BUFFET_TEMPLATE = RoleDefinition(
    role_name=CATERING_HOTEL_BUFFET_ROLE_NAME,
    role_column=CATERING_ROLE_COLUMN,
    sections=(
        SectionDefinition(
            "餐饮服务",
            (
                MetricDefinition("菜肴品种", "AG", "AH"),
                MetricDefinition("菜品供应速度", "AC", "AD"),
                MetricDefinition("菜品口味口相", "AA", "AB"),
                MetricDefinition("菜品地方特色", "U", "V"),
                MetricDefinition("接待指引", "S", "T"),
                MetricDefinition("工作人员服务品质", "AM", "AN"),
                MetricDefinition("就餐区域温度", "M", "N"),
                MetricDefinition("就餐区域卫生", "O", "P"),
            ),
        ),
        CATERING_BASIC_HARDWARE_SECTION,
        CATERING_BUFFET_SMART_SECTION,
    ),
)

TEMPLATE_DEFINITIONS: dict[str, RoleDefinition] = {
    "organizer": ORGANIZER_TEMPLATE,
    "exhibitor": EXHIBITOR_TEMPLATE,
    "visitor": VISITOR_TEMPLATE,
    "service_provider": SERVICE_PROVIDER_TEMPLATE,
    "meeting_organizer": MEETING_ORGANIZER_TEMPLATE,
    "hotel_meeting_organizer": HOTEL_MEETING_ORGANIZER_TEMPLATE,
    "hotel_meeting_attendee": HOTEL_MEETING_ATTENDEE_TEMPLATE,
    "meeting_attendee": MEETING_ATTENDEE_TEMPLATE,
    "catering_food_hall": CATERING_FOOD_HALL_TEMPLATE,
    "catering_business_meal": CATERING_BUSINESS_MEAL_TEMPLATE,
    "catering_tour_meal": CATERING_TOUR_MEAL_TEMPLATE,
    "catering_banquet": CATERING_BANQUET_TEMPLATE,
    "catering_wedding_banquet": CATERING_WEDDING_BANQUET_TEMPLATE,
    "catering_buffet": CATERING_BUFFET_TEMPLATE,
    "catering_hotel_banquet": CATERING_HOTEL_BANQUET_TEMPLATE,
    "catering_hotel_buffet": CATERING_HOTEL_BUFFET_TEMPLATE,
}


def excel_column_to_index(column_name: str) -> int:
    index = 0
    for char in column_name.upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(f"非法 Excel 列名: {column_name}")
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def excel_round(value: float | int | None, digits: int = 2) -> float | None:
    if value is None or pd.isna(value):
        return None

    quantizer = Decimal("1").scaleb(-digits)
    rounded = Decimal(str(value)).quantize(quantizer, rounding=ROUND_HALF_UP)
    return float(rounded)


def mean_ignore_empty(values: list[float | None]) -> float | None:
    defined_values = [value for value in values if value is not None and not pd.isna(value)]
    if not defined_values:
        return None
    return excel_round(sum(defined_values) / len(defined_values))


def format_value(value: float | None) -> str:
    if value is None or pd.isna(value):
        return ""
    text = f"{value:.2f}"
    return text.rstrip("0").rstrip(".")


def resolve_role_definition(template_name: str, role_name: str | None = None) -> RoleDefinition:
    template = TEMPLATE_DEFINITIONS.get(template_name)
    if template is None:
        supported = ", ".join(sorted(TEMPLATE_DEFINITIONS))
        raise ValueError(f"未知模板: {template_name}，支持的模板有: {supported}")
    return RoleDefinition(
        role_name=role_name or template.role_name,
        role_column=template.role_column,
        sections=template.sections,
    )


def required_columns(role_definition: RoleDefinition) -> set[str]:
    columns = {role_definition.role_column}
    for section in role_definition.sections:
        for metric in section.metrics:
            columns.add(metric.satisfaction_column)
            columns.add(metric.importance_column)
            columns.add(metric.satisfaction_gt_zero_column or metric.satisfaction_column)
            columns.add(metric.satisfaction_lt_eleven_column or metric.satisfaction_column)
            columns.add(metric.importance_gt_zero_column or metric.importance_column)
            columns.add(metric.importance_lt_eleven_column or metric.importance_column)
    return columns


def validate_dataframe(df: pd.DataFrame, role_definition: RoleDefinition) -> None:
    max_required_index = max(
        excel_column_to_index(column_name) for column_name in required_columns(role_definition)
    )
    if len(df.columns) <= max_required_index:
        required_max = max(required_columns(role_definition), key=excel_column_to_index)
        raise ValueError(
            f"{role_definition.role_name} 所需的来源 sheet 列数不足，"
            f" 需要至少覆盖到 Excel 列 {required_max}。"
        )


def load_survey_dataframe(
    input_path: Path,
    role_definition: RoleDefinition,
    sheet_name: str = DEFAULT_SHEET_NAME,
) -> pd.DataFrame:
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    validate_dataframe(df, role_definition)
    return df


def compute_metric_average(
    df: pd.DataFrame,
    role_mask: pd.Series,
    column_name: str,
    gt_zero_column_name: str | None = None,
    lt_eleven_column_name: str | None = None,
) -> float | None:
    series = pd.to_numeric(df.iloc[:, excel_column_to_index(column_name)], errors="coerce")
    gt_zero_series = pd.to_numeric(
        df.iloc[:, excel_column_to_index(gt_zero_column_name or column_name)],
        errors="coerce",
    )
    lt_eleven_series = pd.to_numeric(
        df.iloc[:, excel_column_to_index(lt_eleven_column_name or column_name)],
        errors="coerce",
    )
    valid_series = series[role_mask & gt_zero_series.gt(0) & lt_eleven_series.lt(11)]
    if valid_series.empty:
        return None
    return excel_round(valid_series.mean())


def compute_role_stats(
    df: pd.DataFrame,
    role_definition: RoleDefinition,
) -> SurveyStatistics:
    role_series = (
        df.iloc[:, excel_column_to_index(role_definition.role_column)]
        .astype("string")
        .fillna("")
        .str.strip()
    )
    role_mask = role_series.eq(role_definition.role_name)
    matched_row_count = int(role_mask.sum())

    section_results: list[SectionResult] = []
    for section in role_definition.sections:
        metric_results: list[MetricResult] = []
        for metric in section.metrics:
            satisfaction = compute_metric_average(
                df,
                role_mask,
                metric.satisfaction_column,
                gt_zero_column_name=metric.satisfaction_gt_zero_column,
                lt_eleven_column_name=metric.satisfaction_lt_eleven_column,
            )
            importance = compute_metric_average(
                df,
                role_mask,
                metric.importance_column,
                gt_zero_column_name=metric.importance_gt_zero_column,
                lt_eleven_column_name=metric.importance_lt_eleven_column,
            )
            metric_results.append(
                MetricResult(
                    name=metric.name,
                    satisfaction=satisfaction,
                    importance=importance,
                )
            )

        section_results.append(
            SectionResult(
                name=section.name,
                satisfaction=mean_ignore_empty([metric.satisfaction for metric in metric_results]),
                importance=mean_ignore_empty([metric.importance for metric in metric_results]),
                metrics=tuple(metric_results),
            )
        )

    overall_satisfaction = mean_ignore_empty([section.satisfaction for section in section_results])
    overall_importance = mean_ignore_empty([section.importance for section in section_results])
    return SurveyStatistics(
        role_name=role_definition.role_name,
        satisfaction=overall_satisfaction,
        importance=overall_importance,
        sections=tuple(section_results),
        matched_row_count=matched_row_count,
    )


def build_result_dataframe(stats: SurveyStatistics) -> pd.DataFrame:
    rows: list[dict[str, float | str | None]] = [
        {
            "指标": stats.role_name,
            "满意度": stats.satisfaction,
            "重要性": stats.importance,
        }
    ]

    for section in stats.sections:
        rows.append(
            {
                "指标": section.name,
                "满意度": section.satisfaction,
                "重要性": section.importance,
            }
        )
        for metric in section.metrics:
            rows.append(
                {
                    "指标": metric.name,
                    "满意度": metric.satisfaction,
                    "重要性": metric.importance,
                }
            )

    return pd.DataFrame(rows, columns=["指标", "满意度", "重要性"])


def render_markdown_table(df: pd.DataFrame) -> str:
    headers = df.columns.tolist()
    rows = [
        [row["指标"], format_value(row["满意度"]), format_value(row["重要性"])]
        for _, row in df.iterrows()
    ]

    widths = []
    for index, header in enumerate(headers):
        width = len(str(header))
        for row in rows:
            width = max(width, len(str(row[index])))
        widths.append(width)

    def build_row(values: list[str]) -> str:
        cells = [f" {str(value).ljust(widths[index])} " for index, value in enumerate(values)]
        return "|" + "|".join(cells) + "|"

    separator = "|" + "|".join("-" * (width + 2) for width in widths) + "|"

    lines = [build_row(headers), separator]
    lines.extend(build_row(row) for row in rows)
    return "\n".join(lines)


def style_worksheet(worksheet, role_definition: RoleDefinition) -> None:
    section_names = {section.name for section in role_definition.sections}

    for cell in worksheet[1]:
        cell.alignment = CENTER_ALIGNMENT
        cell.font = EMPHASIS_FONT

    for row_index in range(2, worksheet.max_row + 1):
        indicator_name = worksheet.cell(row=row_index, column=1).value
        row_fill = None
        if indicator_name == role_definition.role_name:
            row_fill = OVERALL_FILL
        elif indicator_name in section_names:
            row_fill = SECTION_FILL

        for column_index in range(1, 4):
            cell = worksheet.cell(row=row_index, column=column_index)
            cell.alignment = CENTER_ALIGNMENT
            if row_fill is not None:
                cell.fill = row_fill
                cell.font = EMPHASIS_FONT

    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 12
    worksheet.column_dimensions["C"].width = 12


def normalize_output_dir(output_dir: Path) -> Path:
    if output_dir.exists() and output_dir.is_dir():
        return output_dir

    if output_dir.exists() and output_dir.is_file():
        return output_dir.parent / f"{output_dir.stem}_outputs"

    if output_dir.suffix:
        return output_dir.with_suffix("")

    return output_dir


def build_output_path(
    output_dir: Path,
    output_name: str,
    output_format: str = DEFAULT_OUTPUT_FORMAT,
) -> Path:
    return normalize_output_dir(output_dir) / f"{output_name}.{output_format}"


def save_results(
    df: pd.DataFrame,
    output_path: Path,
    role_definition: RoleDefinition,
    sheet_title: str,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    suffix = output_path.suffix.lower()

    if suffix == ".xlsx":
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_title, index=False)
            style_worksheet(writer.sheets[sheet_title], role_definition)
    elif suffix == ".csv":
        df.to_csv(output_path, index=False, encoding="utf-8-sig")
    elif suffix in {".md", ".markdown"}:
        output_path.write_text(render_markdown_table(df), encoding="utf-8")
    else:
        raise ValueError("仅支持输出为 .xlsx、.csv、.md 或 .markdown 文件。")


def default_config_output_dir(config_path: Path) -> Path:
    return config_path.parent / f"{config_path.stem}_outputs"


def load_batch_config(config_path: Path, default_sheet_name: str = DEFAULT_SHEET_NAME) -> BatchConfig:
    with config_path.open("rb") as file:
        raw_config = tomllib.load(file)

    raw_jobs = raw_config.get("jobs")
    if not isinstance(raw_jobs, list) or not raw_jobs:
        raise ValueError("配置文件必须包含至少一个 [[jobs]] 定义。")

    output_dir_value = raw_config.get("output_dir")
    if output_dir_value is None:
        output_dir = default_config_output_dir(config_path)
    else:
        output_dir = (config_path.parent / str(output_dir_value)).resolve()

    output_format = str(raw_config.get("output_format", DEFAULT_OUTPUT_FORMAT)).lower()
    if output_format not in VALID_OUTPUT_FORMATS:
        raise ValueError(f"配置文件 output_format 仅支持: {', '.join(VALID_OUTPUT_FORMATS)}")

    jobs: list[JobConfig] = []
    for index, job_data in enumerate(raw_jobs, start=1):
        if not isinstance(job_data, dict):
            raise ValueError(f"jobs 第 {index} 项必须是表结构。")
        name = str(job_data.get("name", "")).strip()
        path_value = job_data.get("path")
        template_name = str(job_data.get("template", "")).strip()
        if not name or not path_value or not template_name:
            raise ValueError(f"jobs 第 {index} 项必须包含 name/path/template。")
        default_role_name = resolve_role_definition(template_name).role_name
        role_name = str(job_data.get("role_name", default_role_name)).strip() or default_role_name
        output_name = str(job_data.get("output_name", name)).strip() or name
        sheet_name = str(job_data.get("sheet", default_sheet_name)).strip() or default_sheet_name
        job_path = (config_path.parent / str(path_value)).resolve()
        job_output_format_value = job_data.get("output_format")
        job_output_format = None
        if job_output_format_value is not None:
            job_output_format = str(job_output_format_value).lower()
            if job_output_format not in VALID_OUTPUT_FORMATS:
                raise ValueError(
                    f"jobs 第 {index} 项的 output_format 仅支持: {', '.join(VALID_OUTPUT_FORMATS)}"
                )
        jobs.append(
            JobConfig(
                name=name,
                path=job_path,
                sheet_name=sheet_name,
                template_name=template_name,
                role_name=role_name,
                output_name=output_name,
                output_format=job_output_format,
            )
        )

    return BatchConfig(
        config_path=config_path.resolve(),
        output_dir=output_dir,
        output_format=output_format,
        jobs=tuple(jobs),
    )


def select_jobs(
    jobs: tuple[JobConfig, ...],
    job_filters: list[str],
) -> tuple[JobConfig, ...]:
    selected_jobs = jobs
    if job_filters:
        selected_job_names = set(job_filters)
        selected_jobs = tuple(job for job in selected_jobs if job.name in selected_job_names)
    return selected_jobs


def generate_role_report_bundle(
    input_path: Path,
    role_definition: RoleDefinition,
    output_path: Path,
    sheet_name: str = DEFAULT_SHEET_NAME,
    sheet_title: str | None = None,
    dry_run: bool = False,
) -> GeneratedReport:
    survey_df = load_survey_dataframe(input_path, role_definition, sheet_name=sheet_name)
    stats = compute_role_stats(survey_df, role_definition)
    result_df = build_result_dataframe(stats)
    final_sheet_title = sheet_title or role_definition.role_name
    if not dry_run:
        save_results(result_df, output_path, role_definition, final_sheet_title)
    return GeneratedReport(result_df=result_df, output_path=output_path, stats=stats)


def generate_role_report(
    input_path: Path,
    role_definition: RoleDefinition,
    output_path: Path,
    sheet_name: str = DEFAULT_SHEET_NAME,
    sheet_title: str | None = None,
    dry_run: bool = False,
) -> tuple[pd.DataFrame, Path]:
    report = generate_role_report_bundle(
        input_path=input_path,
        role_definition=role_definition,
        output_path=output_path,
        sheet_name=sheet_name,
        sheet_title=sheet_title,
        dry_run=dry_run,
    )
    return report.result_df, report.output_path


def build_missing_group_summary(notices: list[MissingGroupNotice]) -> str | None:
    if not notices:
        return None

    lines = ["以下指定的客户分组在来源数据中未找到任何匹配记录，已输出空白统计结果："]
    for notice in notices:
        lines.append(f"- {notice.job_name} [{notice.input_path.name} / {notice.sheet_name}]")
    return "\n".join(lines)


def print_missing_group_summary(notices: list[MissingGroupNotice]) -> None:
    summary = build_missing_group_summary(notices)
    if summary is not None:
        print(f"\n{summary}")


def print_report(title: str, result_df: pd.DataFrame, output_path: Path, dry_run: bool = False) -> None:
    print(f"\n## {title}")
    print(render_markdown_table(result_df))
    if dry_run:
        print(f"\n[DRY RUN] 将输出到: {output_path}")
    else:
        print(f"\n结果已保存到: {output_path}")


def run_single_mode(args: argparse.Namespace) -> None:
    if not all([args.input, args.template, args.role_name, args.output]):
        raise ValueError("--input、--template、--role-name、--output 必须同时提供。")

    role_definition = resolve_role_definition(args.template, args.role_name)
    report = generate_role_report_bundle(
        input_path=args.input,
        role_definition=role_definition,
        output_path=args.output,
        sheet_name=args.sheet_name,
        sheet_title=args.output.stem,
        dry_run=args.dry_run,
    )
    print_report(args.role_name, report.result_df, report.output_path, dry_run=args.dry_run)
    if report.stats.matched_row_count == 0:
        print_missing_group_summary(
            [MissingGroupNotice(args.role_name, args.input, args.sheet_name)]
        )


def run_legacy_batch_mode(args: argparse.Namespace) -> None:
    if not all([args.organizer_input, args.exhibitor_input, args.visitor_input, args.output_dir]):
        raise ValueError(
            "--organizer-input、--exhibitor-input、--visitor-input、--output-dir 必须同时提供。"
        )

    jobs = (
        (ORGANIZER_TEMPLATE, args.organizer_input),
        (EXHIBITOR_TEMPLATE, args.exhibitor_input),
        (VISITOR_TEMPLATE, args.visitor_input),
    )
    output_dir = normalize_output_dir(args.output_dir)
    output_format = args.output_format or DEFAULT_OUTPUT_FORMAT
    missing_group_notices: list[MissingGroupNotice] = []

    for role_definition, input_path in jobs:
        output_path = build_output_path(output_dir, role_definition.role_name, output_format)
        report = generate_role_report_bundle(
            input_path=input_path,
            role_definition=role_definition,
            output_path=output_path,
            sheet_name=args.sheet_name,
            sheet_title=role_definition.role_name,
            dry_run=args.dry_run,
        )
        print_report(
            role_definition.role_name,
            report.result_df,
            report.output_path,
            dry_run=args.dry_run,
        )
        if report.stats.matched_row_count == 0:
            missing_group_notices.append(
                MissingGroupNotice(role_definition.role_name, input_path, args.sheet_name)
            )

    print_missing_group_summary(missing_group_notices)


def run_config_mode(args: argparse.Namespace) -> None:
    config = load_batch_config(args.config, default_sheet_name=args.sheet_name)
    selected_jobs = select_jobs(config.jobs, args.job)
    if not selected_jobs:
        raise ValueError("筛选后没有可运行的 jobs。")

    output_dir = normalize_output_dir(args.output_dir or config.output_dir)
    global_output_format = args.output_format or config.output_format
    missing_group_notices: list[MissingGroupNotice] = []

    for job in selected_jobs:
        role_definition = resolve_role_definition(job.template_name, job.role_name)
        output_format = job.output_format or global_output_format
        output_path = build_output_path(output_dir, job.output_name, output_format)
        report = generate_role_report_bundle(
            input_path=job.path,
            role_definition=role_definition,
            output_path=output_path,
            sheet_name=job.sheet_name,
            sheet_title=job.name,
            dry_run=args.dry_run,
        )
        title = f"{job.name} ({job.template_name})"
        print_report(title, report.result_df, report.output_path, dry_run=args.dry_run)
        if report.stats.matched_row_count == 0:
            missing_group_notices.append(MissingGroupNotice(job.name, job.path, job.sheet_name))

    print_missing_group_summary(missing_group_notices)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按会展问卷模板批量或单次计算统计结果，并导出 Excel/CSV/Markdown 文件。"
    )
    parser.add_argument("--config", type=Path, help="批量模式：TOML 配置文件路径")
    parser.add_argument("--job", action="append", default=[], help="批量模式：只运行某个 job 名称，可重复传入")
    parser.add_argument("--dry-run", action="store_true", help="只校验并展示结果，不实际写文件")
    parser.add_argument("--sheet-name", default=DEFAULT_SHEET_NAME, help=f"默认 sheet 名，默认 {DEFAULT_SHEET_NAME}")
    parser.add_argument(
        "--output-format",
        choices=VALID_OUTPUT_FORMATS,
        help=f"覆盖输出格式，支持 {', '.join(VALID_OUTPUT_FORMATS)}",
    )
    parser.add_argument("--output-dir", type=Path, help="批量模式输出目录，或覆盖配置里的 output_dir")

    parser.add_argument("--input", type=Path, help="单任务模式：来源 Excel 文件")
    parser.add_argument(
        "--template",
        choices=sorted(TEMPLATE_DEFINITIONS),
        help="单任务模式：模板类型",
    )
    parser.add_argument("--role-name", help="单任务模式：按该身份分组")
    parser.add_argument("--output", type=Path, help="单任务模式：输出文件路径")

    parser.add_argument("--organizer-input", type=Path, help="兼容模式：展览主承办来源 Excel")
    parser.add_argument("--exhibitor-input", type=Path, help="兼容模式：参展商来源 Excel")
    parser.add_argument("--visitor-input", type=Path, help="兼容模式：专业观众来源 Excel")

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.config:
        run_config_mode(args)
        return

    if any([args.input, args.template, args.role_name, args.output]):
        run_single_mode(args)
        return

    if any([args.organizer_input, args.exhibitor_input, args.visitor_input]):
        run_legacy_batch_mode(args)
        return

    raise SystemExit(
        "请使用 --config 批量模式，或提供单任务参数 --input/--template/--role-name/--output。"
    )


if __name__ == "__main__":
    main()
