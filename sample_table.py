from __future__ import annotations

import argparse
import tomllib
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from summary_table import (
    SUMMARY_BODY_FILL,
    SUMMARY_BORDER,
    SUMMARY_BODY_FONT,
    SUMMARY_CENTER_ALIGNMENT,
    SUMMARY_HEADER_FILL,
    SUMMARY_HEADER_FONT,
    SUMMARY_SIDE_FILL,
    SUMMARY_SIDE_FONT,
    SUMMARY_TITLE_FONT,
    apply_common_style,
    set_text_cell,
)
from survey_customer_category_rules import CUSTOMER_CATEGORY_RULE_BY_NAME
from survey_stats import (
    DEFAULT_SHEET_NAME,
    build_customer_category_rule_mask,
    normalize_output_dir,
)

DEFAULT_SAMPLE_TABLE_TITLE = "杭博客户类型样本统计表"
DEFAULT_SAMPLE_TABLE_OUTPUT_NAME = "客户类型样本统计表.xlsx"
DEFAULT_SAMPLE_TABLE_CONFIG_PATH = Path(__file__).with_name("sample_table.default.toml")
SAMPLE_TABLE_SHEET_NAME = "样本统计"

COUNT_NUMBER_FORMAT = "0"
PERCENT_NUMBER_FORMAT = "0.00%"


@dataclass(frozen=True)
class SampleTableRowConfig:
    category_label: str
    display_name: str
    target_sample_size: int
    rule_name: str | None = None
    actual_count_override: int | None = None


@dataclass(frozen=True)
class SampleTableConfig:
    title: str
    sheet_name: str
    output_name: str
    rows: tuple[SampleTableRowConfig, ...]


@dataclass(frozen=True)
class SampleTableRowResult:
    category_label: str
    display_name: str
    target_sample_size: int
    actual_count: int
    month_text: str


@dataclass(frozen=True)
class RowGroup:
    category_label: str
    rows: tuple[SampleTableRowResult, ...]


def load_sample_table_config(
    config_path: Path = DEFAULT_SAMPLE_TABLE_CONFIG_PATH,
) -> SampleTableConfig:
    data = tomllib.loads(config_path.read_text(encoding="utf-8"))
    rows_data = data.get("rows")
    if not isinstance(rows_data, list) or not rows_data:
        raise ValueError("样本统计配置缺少 rows 列表。")

    rows: list[SampleTableRowConfig] = []
    for row_data in rows_data:
        rows.append(
            SampleTableRowConfig(
                category_label=str(row_data["category_label"]).strip(),
                display_name=str(row_data.get("display_name", "")).strip(),
                target_sample_size=int(row_data["target_sample_size"]),
                rule_name=(
                    str(row_data["rule_name"]).strip()
                    if row_data.get("rule_name") is not None
                    else None
                ),
                actual_count_override=(
                    int(row_data["actual_count_override"])
                    if row_data.get("actual_count_override") is not None
                    else None
                ),
            )
        )

    return SampleTableConfig(
        title=str(data.get("title", DEFAULT_SAMPLE_TABLE_TITLE)).strip() or DEFAULT_SAMPLE_TABLE_TITLE,
        sheet_name=str(data.get("sheet_name", SAMPLE_TABLE_SHEET_NAME)).strip() or SAMPLE_TABLE_SHEET_NAME,
        output_name=str(data.get("output_name", DEFAULT_SAMPLE_TABLE_OUTPUT_NAME)).strip()
        or DEFAULT_SAMPLE_TABLE_OUTPUT_NAME,
        rows=tuple(rows),
    )


def normalize_month_value(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0"):
        numeric_text = text[:-2]
        if numeric_text.isdigit():
            return numeric_text
    return text


def find_month_column(df: pd.DataFrame) -> object | None:
    for column_name in df.columns:
        if str(column_name).strip() == "月份":
            return column_name
    return None


def summarize_month_values(df: pd.DataFrame, mask: pd.Series) -> str:
    month_column = find_month_column(df)
    if month_column is None:
        return ""

    ordered_months: list[str] = []
    seen_months: set[str] = set()
    for value in df.loc[mask, month_column].tolist():
        month_text = normalize_month_value(value)
        if not month_text or month_text in seen_months:
            continue
        ordered_months.append(month_text)
        seen_months.add(month_text)
    return "、".join(ordered_months)


def combine_month_texts(month_texts: list[str]) -> str:
    ordered_months: list[str] = []
    seen_months: set[str] = set()
    for month_text in month_texts:
        for value in month_text.split("、"):
            normalized_value = value.strip()
            if not normalized_value or normalized_value in seen_months:
                continue
            ordered_months.append(normalized_value)
            seen_months.add(normalized_value)
    return "、".join(ordered_months)


def load_source_dataframe(
    input_dir: Path,
    source_file_name: str,
    *,
    sheet_name: str,
    dataframe_cache: dict[Path, pd.DataFrame | None],
) -> pd.DataFrame | None:
    input_path = input_dir / source_file_name
    if input_path in dataframe_cache:
        return dataframe_cache[input_path]

    if not input_path.exists() or not input_path.is_file():
        dataframe_cache[input_path] = None
        return None

    dataframe_cache[input_path] = pd.read_excel(input_path, sheet_name=sheet_name)
    return dataframe_cache[input_path]


def build_sample_table_rows(
    input_dir: Path,
    config: SampleTableConfig | None = None,
    *,
    source_sheet_name: str = DEFAULT_SHEET_NAME,
) -> tuple[SampleTableRowResult, ...]:
    resolved_config = config or load_sample_table_config()
    dataframe_cache: dict[Path, pd.DataFrame | None] = {}
    row_results: list[SampleTableRowResult] = []

    for row_config in resolved_config.rows:
        actual_count = row_config.actual_count_override or 0
        month_text = ""

        if row_config.actual_count_override is None and row_config.rule_name is not None:
            rule = CUSTOMER_CATEGORY_RULE_BY_NAME.get(row_config.rule_name)
            if rule is None:
                raise ValueError(f"样本统计配置引用了未知映射规则：{row_config.rule_name}")

            dataframe = load_source_dataframe(
                input_dir,
                rule.source_file_name,
                sheet_name=source_sheet_name,
                dataframe_cache=dataframe_cache,
            )
            if dataframe is not None:
                mask = build_customer_category_rule_mask(dataframe, rule)
                actual_count = int(mask.sum())
                month_text = summarize_month_values(dataframe, mask)

        row_results.append(
            SampleTableRowResult(
                category_label=row_config.category_label,
                display_name=row_config.display_name,
                target_sample_size=row_config.target_sample_size,
                actual_count=actual_count,
                month_text=month_text,
            )
        )

    return tuple(row_results)


def build_row_groups(rows: tuple[SampleTableRowResult, ...]) -> tuple[RowGroup, ...]:
    if not rows:
        return ()

    groups: list[RowGroup] = []
    current_category = rows[0].category_label
    current_rows: list[SampleTableRowResult] = []

    for row in rows:
        if row.category_label != current_category:
            groups.append(RowGroup(category_label=current_category, rows=tuple(current_rows)))
            current_category = row.category_label
            current_rows = []
        current_rows.append(row)

    groups.append(RowGroup(category_label=current_category, rows=tuple(current_rows)))
    return tuple(groups)


def set_count_cell(cell, value: int | str | None) -> None:
    cell.value = value
    apply_common_style(cell, fill=SUMMARY_BODY_FILL, font=SUMMARY_BODY_FONT)
    cell.number_format = COUNT_NUMBER_FORMAT


def set_percent_cell(cell, value: str) -> None:
    cell.value = value
    apply_common_style(cell, fill=SUMMARY_BODY_FILL, font=SUMMARY_BODY_FONT)
    cell.number_format = PERCENT_NUMBER_FORMAT


def set_month_cell(cell, value: str) -> None:
    cell.value = value or None
    apply_common_style(cell, fill=SUMMARY_BODY_FILL, font=SUMMARY_BODY_FONT)


def merge_category_column(worksheet, start_row: int, end_row: int) -> None:
    if start_row >= end_row:
        return
    worksheet.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)


def ensure_output_path(output_dir: Path, output_name: str) -> Path:
    final_output_dir = normalize_output_dir(output_dir)
    file_name = output_name if output_name.lower().endswith(".xlsx") else f"{output_name}.xlsx"
    return final_output_dir / file_name


def style_sample_table_worksheet(
    worksheet,
    rows: tuple[SampleTableRowResult, ...],
    *,
    title: str,
) -> None:
    worksheet.merge_cells("A1:F1")
    worksheet.sheet_view.showGridLines = False
    worksheet.row_dimensions[1].height = 36
    worksheet.row_dimensions[2].height = 28
    set_text_cell(worksheet["A1"], title, fill=SUMMARY_HEADER_FILL, font=SUMMARY_TITLE_FONT)
    for column_index in range(2, 7):
        apply_common_style(
            worksheet.cell(row=1, column=column_index),
            fill=SUMMARY_HEADER_FILL,
            font=SUMMARY_TITLE_FONT,
        )

    header_names = ("客户大类", "样本类型", "样本量", "样本进度百分比", "总执行样本量", "月份")
    for column_index, column_name in enumerate(header_names, start=1):
        set_text_cell(
            worksheet.cell(row=2, column=column_index),
            column_name,
            fill=SUMMARY_HEADER_FILL,
            font=SUMMARY_HEADER_FONT,
        )

    data_row_ranges: list[tuple[int, int]] = []
    all_month_texts: list[str] = []
    current_excel_row = 3

    for group in build_row_groups(rows):
        group_start_row = current_excel_row
        group_month_texts: list[str] = []

        for row in group.rows:
            worksheet.row_dimensions[current_excel_row].height = 24
            set_text_cell(
                worksheet.cell(row=current_excel_row, column=1),
                row.category_label,
                fill=SUMMARY_SIDE_FILL,
                font=SUMMARY_SIDE_FONT,
            )
            set_text_cell(
                worksheet.cell(row=current_excel_row, column=2),
                row.display_name or None,
                fill=SUMMARY_SIDE_FILL,
                font=SUMMARY_SIDE_FONT,
            )
            set_count_cell(worksheet.cell(row=current_excel_row, column=3), row.target_sample_size)
            set_percent_cell(
                worksheet.cell(row=current_excel_row, column=4),
                f"=IFERROR(E{current_excel_row}/C{current_excel_row},0)",
            )
            set_count_cell(worksheet.cell(row=current_excel_row, column=5), row.actual_count)
            set_month_cell(worksheet.cell(row=current_excel_row, column=6), row.month_text)

            group_month_texts.append(row.month_text)
            all_month_texts.append(row.month_text)
            current_excel_row += 1

        group_end_row = current_excel_row - 1
        data_row_ranges.append((group_start_row, group_end_row))
        merge_category_column(worksheet, group_start_row, group_end_row)

        if len(group.rows) > 1:
            worksheet.row_dimensions[current_excel_row].height = 24
            set_text_cell(
                worksheet.cell(row=current_excel_row, column=1),
                None,
                fill=SUMMARY_SIDE_FILL,
                font=SUMMARY_SIDE_FONT,
            )
            set_text_cell(
                worksheet.cell(row=current_excel_row, column=2),
                "小计",
                fill=SUMMARY_SIDE_FILL,
                font=SUMMARY_SIDE_FONT,
            )
            set_count_cell(
                worksheet.cell(row=current_excel_row, column=3),
                f"=SUM(C{group_start_row}:C{group_end_row})",
            )
            set_percent_cell(
                worksheet.cell(row=current_excel_row, column=4),
                f"=IFERROR(E{current_excel_row}/C{current_excel_row},0)",
            )
            set_count_cell(
                worksheet.cell(row=current_excel_row, column=5),
                f"=SUM(E{group_start_row}:E{group_end_row})",
            )
            set_month_cell(
                worksheet.cell(row=current_excel_row, column=6),
                combine_month_texts(group_month_texts),
            )
            current_excel_row += 1

    total_row_index = current_excel_row
    worksheet.row_dimensions[total_row_index].height = 28
    worksheet.merge_cells(start_row=total_row_index, start_column=1, end_row=total_row_index, end_column=2)
    set_text_cell(
        worksheet.cell(row=total_row_index, column=1),
        "合计",
        fill=SUMMARY_HEADER_FILL,
        font=SUMMARY_HEADER_FONT,
    )
    apply_common_style(
        worksheet.cell(row=total_row_index, column=2),
        fill=SUMMARY_HEADER_FILL,
        font=SUMMARY_HEADER_FONT,
    )

    count_ranges = ",".join(f"C{start_row}:C{end_row}" for start_row, end_row in data_row_ranges)
    actual_ranges = ",".join(f"E{start_row}:E{end_row}" for start_row, end_row in data_row_ranges)
    set_count_cell(worksheet.cell(row=total_row_index, column=3), f"=SUM({count_ranges})")
    set_percent_cell(
        worksheet.cell(row=total_row_index, column=4),
        f"=IFERROR(E{total_row_index}/C{total_row_index},0)",
    )
    set_count_cell(worksheet.cell(row=total_row_index, column=5), f"=SUM({actual_ranges})")
    set_month_cell(
        worksheet.cell(row=total_row_index, column=6),
        combine_month_texts(all_month_texts),
    )

    worksheet.freeze_panes = "C3"
    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 30
    worksheet.column_dimensions["C"].width = 12
    worksheet.column_dimensions["D"].width = 16
    worksheet.column_dimensions["E"].width = 14
    worksheet.column_dimensions["F"].width = 18


def generate_sample_table_report(
    input_dir: Path,
    output_dir: Path,
    *,
    output_name: str | None = None,
    config_path: Path = DEFAULT_SAMPLE_TABLE_CONFIG_PATH,
    source_sheet_name: str = DEFAULT_SHEET_NAME,
) -> Path:
    config = load_sample_table_config(config_path)
    rows = build_sample_table_rows(
        input_dir=input_dir,
        config=config,
        source_sheet_name=source_sheet_name,
    )
    output_path = ensure_output_path(output_dir, output_name or config.output_name)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = config.sheet_name
    style_sample_table_worksheet(worksheet, rows, title=config.title)
    workbook.save(output_path)
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按客户类型映射规则统计原始问卷样本量，生成样本统计表。")
    parser.add_argument("--input-dir", type=Path, required=True, help="输入目录：原始问卷 xlsx 所在目录")
    parser.add_argument("--output-dir", type=Path, required=True, help="输出目录")
    parser.add_argument("--output-name", help="输出文件名，默认读取配置中的 output_name")
    parser.add_argument(
        "--config",
        type=Path,
        default=DEFAULT_SAMPLE_TABLE_CONFIG_PATH,
        help=f"样本统计配置文件，默认 {DEFAULT_SAMPLE_TABLE_CONFIG_PATH.name}",
    )
    parser.add_argument(
        "--source-sheet-name",
        default=DEFAULT_SHEET_NAME,
        help=f"原始数据 sheet 名称，默认 {DEFAULT_SHEET_NAME}",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_path = generate_sample_table_report(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        output_name=args.output_name,
        config_path=args.config,
        source_sheet_name=args.source_sheet_name,
    )
    print(f"样本统计表已保存到: {output_path}")


if __name__ == "__main__":
    main()
