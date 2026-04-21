from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sample_table import find_column, normalize_month_value, normalize_year_value
from survey_customer_category_rules import CUSTOMER_CATEGORY_RULE_BY_NAME
from survey_stats import (
    TEMPLATE_DEFINITIONS,
    build_role_mask,
    compute_metric_average,
    compute_role_stats,
    excel_column_to_index,
    get_effective_role_definition,
)

EXHIBITION_RULE_NAMES = (
    "展览主承办",
    "参展商",
    "专业观众",
    "会展服务商",
    "会议主承办",
    "参会人员",
)
PERIOD_DIRS: dict[str, Path] = {
    "1-2月": ROOT / "datas" / "1-2月",
    "3月": ROOT / "datas" / "3月",
    "Q1": ROOT / "datas" / "Q1",
}
COMMENT_THEMES: dict[str, tuple[str, ...]] = {
    "餐饮": ("餐饮", "吃", "菜", "盒饭", "茶歇", "咖啡"),
    "停车": ("停车", "停车场"),
    "标识导视": ("标识", "导视", "指引", "路牌", "导航"),
    "动线交通": ("动线", "路线", "交通流线", "路线安排", "交通"),
    "展厅场地": ("展厅", "场地", "场馆", "会场"),
    "休息空间": ("休息", "座位"),
    "温度空调": ("空调", "温度", "太热", "太冷"),
    "服务响应": ("服务", "响应", "沟通", "协调", "对接"),
}


def find_column_by_keywords(df: pd.DataFrame, *keywords: str) -> str | None:
    for column in df.columns:
        text = str(column)
        if all(keyword in text for keyword in keywords):
            return text
    return None


def compute_valid_count(
    df: pd.DataFrame,
    role_mask: pd.Series,
    column_name: str,
    gt_zero_column_name: str | None = None,
    lt_eleven_column_name: str | None = None,
) -> int:
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
    return int(valid_series.count())


def safe_event_name(value: object) -> str:
    if value is None or pd.isna(value):
        return "未填写活动名称"
    text = str(value).strip()
    return text or "未填写活动名称"


def role_period_analysis(rule_name: str, period_label: str) -> dict[str, object]:
    rule = CUSTOMER_CATEGORY_RULE_BY_NAME[rule_name]
    role_definition = TEMPLATE_DEFINITIONS[rule.template_name]
    summary_role_definition = get_effective_role_definition(role_definition, calculation_mode="summary")
    source_path = PERIOD_DIRS[period_label] / rule.source_file_name
    if not source_path.exists():
        return {
            "role_row": {
                "period": period_label,
                "role": rule.customer_category,
                "sample_count": 0,
                "score": None,
                "product_service": None,
                "hardware": None,
                "support": None,
                "smart": None,
                "dining": None,
            },
            "metric_rows": [],
            "event_rows": [],
            "comment_rows": [],
        }

    period_df = pd.read_excel(source_path, sheet_name="问卷数据")
    stats = compute_role_stats(period_df, role_definition, calculation_mode="summary")
    role_mask = build_role_mask(period_df, summary_role_definition)

    section_score_map = {section.name: section.satisfaction for section in stats.sections}
    metric_rows: list[dict[str, object]] = []
    for section in summary_role_definition.sections:
        for metric in section.metrics:
            metric_rows.append(
                {
                    "period": period_label,
                    "role": rule.customer_category,
                    "section": section.name,
                    "metric": metric.name,
                    "score": compute_metric_average(
                        period_df,
                        role_mask,
                        metric.satisfaction_column,
                        gt_zero_column_name=metric.satisfaction_gt_zero_column,
                        lt_eleven_column_name=metric.satisfaction_lt_eleven_column,
                    ),
                    "valid_count": compute_valid_count(
                        period_df,
                        role_mask,
                        metric.satisfaction_column,
                        gt_zero_column_name=metric.satisfaction_gt_zero_column,
                        lt_eleven_column_name=metric.satisfaction_lt_eleven_column,
                    ),
                }
            )

    event_column = find_column_by_keywords(period_df, "活动名称")
    event_rows: list[dict[str, object]] = []
    if event_column is not None and int(role_mask.sum()) > 0:
        role_df = period_df.loc[role_mask].copy()
        role_df["__event_name__"] = role_df[event_column].map(safe_event_name)
        for event_name, event_group in role_df.groupby("__event_name__"):
            event_stats = compute_role_stats(event_group, role_definition, calculation_mode="summary")
            event_rows.append(
                {
                    "period": period_label,
                    "role": rule.customer_category,
                    "event_name": event_name,
                    "sample_count": int(len(event_group)),
                    "score": event_stats.satisfaction,
                }
            )

    comment_column = find_column_by_keywords(period_df, "评价或建议")
    comment_rows: list[dict[str, object]] = []
    if comment_column is not None and int(role_mask.sum()) > 0:
        comments = (
            period_df.loc[role_mask, comment_column]
            .astype("string")
            .fillna("")
            .str.replace(r"\s+", "", regex=True)
            .str.strip()
        )
        comments = comments[comments.ne("")]
        for theme_name, keywords in COMMENT_THEMES.items():
            matched_comments = [text for text in comments.tolist() if any(keyword in text for keyword in keywords)]
            if matched_comments:
                comment_rows.append(
                    {
                        "period": period_label,
                        "role": rule.customer_category,
                        "theme": theme_name,
                        "comment_count": len(matched_comments),
                        "sample_comment": matched_comments[0][:80],
                    }
                )

    return {
        "role_row": {
            "period": period_label,
            "role": rule.customer_category,
            "sample_count": stats.matched_row_count,
            "score": stats.satisfaction,
            "product_service": section_score_map.get("产品服务"),
            "hardware": section_score_map.get("硬件设施"),
            "support": section_score_map.get("配套服务"),
            "smart": section_score_map.get("智慧场馆/服务"),
            "dining": section_score_map.get("餐饮服务"),
        },
        "metric_rows": metric_rows,
        "event_rows": event_rows,
        "comment_rows": comment_rows,
    }


def build_category_summary(role_df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for period, group in role_df.groupby("period"):
        valid_scores = group["score"].dropna()
        total_samples = int(group["sample_count"].sum())
        weighted_score = (
            (group["score"] * group["sample_count"]).sum() / group.loc[group["score"].notna(), "sample_count"].sum()
            if group.loc[group["score"].notna(), "sample_count"].sum() > 0
            else None
        )
        rows.append(
            {
                "period": period,
                "role_count": int(valid_scores.count()),
                "sample_count": total_samples,
                "simple_avg_score": round(float(valid_scores.mean()), 4) if not valid_scores.empty else None,
                "weighted_avg_score": round(float(weighted_score), 4) if weighted_score is not None else None,
                "weighted_product_service": round(
                    float((group["product_service"] * group["sample_count"]).sum() / group.loc[group["product_service"].notna(), "sample_count"].sum()),
                    4,
                ) if group.loc[group["product_service"].notna(), "sample_count"].sum() > 0 else None,
                "weighted_hardware": round(
                    float((group["hardware"] * group["sample_count"]).sum() / group.loc[group["hardware"].notna(), "sample_count"].sum()),
                    4,
                ) if group.loc[group["hardware"].notna(), "sample_count"].sum() > 0 else None,
                "weighted_support": round(
                    float((group["support"] * group["sample_count"]).sum() / group.loc[group["support"].notna(), "sample_count"].sum()),
                    4,
                ) if group.loc[group["support"].notna(), "sample_count"].sum() > 0 else None,
                "weighted_dining": round(
                    float((group["dining"] * group["sample_count"]).sum() / group.loc[group["dining"].notna(), "sample_count"].sum()),
                    4,
                ) if group.loc[group["dining"].notna(), "sample_count"].sum() > 0 else None,
            }
        )
    return pd.DataFrame(rows)


def build_gap_contribution(role_df: pd.DataFrame, period: str) -> pd.DataFrame:
    group = role_df[role_df["period"].eq(period) & role_df["score"].notna()].copy()
    total_samples = group["sample_count"].sum()
    group["sample_share"] = group["sample_count"] / total_samples
    group["gap_to_10"] = 10 - group["score"]
    group["weighted_gap_contribution"] = group["gap_to_10"] * group["sample_share"]
    return group[["role", "sample_count", "sample_share", "score", "gap_to_10", "weighted_gap_contribution"]].sort_values(
        ["weighted_gap_contribution", "sample_count"], ascending=[False, False]
    )


def build_delta_contribution(role_df: pd.DataFrame) -> pd.DataFrame:
    pivot = role_df.pivot(index="role", columns="period", values=["score", "sample_count"])
    pivot.columns = [f"{metric}_{period}" for metric, period in pivot.columns]
    pivot = pivot.reset_index()
    pivot = pivot[pivot["score_1-2月"].notna() & pivot["score_3月"].notna()].copy()
    total_samples_3 = pivot["sample_count_3月"].sum()
    pivot["score_delta"] = pivot["score_3月"] - pivot["score_1-2月"]
    pivot["sample_share_3月"] = pivot["sample_count_3月"] / total_samples_3
    pivot["weighted_delta_contribution"] = pivot["score_delta"] * pivot["sample_share_3月"]
    return pivot[[
        "role",
        "score_1-2月",
        "sample_count_1-2月",
        "score_3月",
        "sample_count_3月",
        "score_delta",
        "sample_share_3月",
        "weighted_delta_contribution",
    ]].sort_values("weighted_delta_contribution")


def format_frame(df: pd.DataFrame, max_rows: int | None = None) -> str:
    if max_rows is not None:
        df = df.head(max_rows)
    if df.empty:
        return "<empty>"
    return df.to_string(index=False)


def main() -> None:
    role_rows: list[dict[str, object]] = []
    metric_rows: list[dict[str, object]] = []
    event_rows: list[dict[str, object]] = []
    comment_rows: list[dict[str, object]] = []

    for period in PERIOD_DIRS:
        for rule_name in EXHIBITION_RULE_NAMES:
            result = role_period_analysis(rule_name, period)
            role_rows.append(result["role_row"])
            metric_rows.extend(result["metric_rows"])
            event_rows.extend(result["event_rows"])
            comment_rows.extend(result["comment_rows"])

    role_df = pd.DataFrame(role_rows)
    metric_df = pd.DataFrame(metric_rows)
    event_df = pd.DataFrame(event_rows)
    comment_df = pd.DataFrame(comment_rows)

    print("\n=== 会展客户角色样本与得分 ===")
    print(format_frame(role_df.sort_values(["period", "sample_count"], ascending=[True, False])))

    print("\n=== 会展客户类别总览 ===")
    print(format_frame(build_category_summary(role_df).sort_values("period")))

    print("\n=== Q1 样本加权缺口贡献（越高越拖累总体）===")
    print(format_frame(build_gap_contribution(role_df, "Q1"), max_rows=6))

    print("\n=== 3月相对1-2月的样本加权变化贡献（越负越拖累3月）===")
    print(format_frame(build_delta_contribution(role_df), max_rows=6))

    q1_low_metrics = metric_df[(metric_df["period"].eq("Q1")) & metric_df["score"].notna() & metric_df["valid_count"].ge(5)].copy()
    q1_low_metrics = q1_low_metrics.sort_values(["score", "valid_count"]).head(20)
    print("\n=== Q1 会展客户低分指标（valid_count>=5）===")
    print(format_frame(q1_low_metrics[["role", "section", "metric", "score", "valid_count"]]))

    march_low_metrics = metric_df[(metric_df["period"].eq("3月")) & metric_df["score"].notna() & metric_df["valid_count"].ge(5)].copy()
    march_low_metrics = march_low_metrics.sort_values(["score", "valid_count"]).head(20)
    print("\n=== 3月 会展客户低分指标（valid_count>=5）===")
    print(format_frame(march_low_metrics[["role", "section", "metric", "score", "valid_count"]]))

    low_events = event_df[(event_df["period"].eq("3月")) & event_df["sample_count"].ge(5) & event_df["score"].notna()].copy()
    low_events = low_events.sort_values(["score", "sample_count"]).head(15)
    print("\n=== 3月 低分活动（样本>=5）===")
    print(format_frame(low_events[["role", "event_name", "sample_count", "score"]]))

    comment_summary = comment_df.groupby(["period", "role", "theme"], as_index=False).agg(
        comment_count=("comment_count", "sum"),
        sample_comment=("sample_comment", "first"),
    )
    comment_summary = comment_summary.sort_values(["period", "comment_count"], ascending=[True, False])
    print("\n=== 开放题主题命中（会展客户）===")
    print(format_frame(comment_summary, max_rows=30))


if __name__ == "__main__":
    main()
