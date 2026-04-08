from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

DEFAULT_SHEET_NAME = "问卷数据"
DEFAULT_FIELD_NAME = "开始填表时间"


@dataclass(frozen=True)
class FileMonthSummary:
    path: Path
    status: str
    row_count: int
    non_empty_count: int
    valid_count: int
    invalid_count: int
    months: tuple[str, ...]
    error_message: str | None = None

    @property
    def empty_count(self) -> int:
        return max(self.row_count - self.non_empty_count, 0)


@dataclass(frozen=True)
class DirectoryMonthSummary:
    input_dir: Path
    file_summaries: tuple[FileMonthSummary, ...]
    detected_months: tuple[str, ...]
    all_valid_values_in_one_month: bool
    all_files_in_one_month: bool


def iter_excel_paths(input_dir: Path, recursive: bool = False) -> list[Path]:
    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    paths = sorted(path for path in input_dir.glob(pattern) if path.is_file())
    return [
        path
        for path in paths
        if not path.name.startswith("~$")
        and not path.name.startswith("._")
    ]


def is_blank(value: object) -> bool:
    if pd.isna(value):
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def summarize_workbook(
    workbook_path: Path,
    *,
    sheet_name: str = DEFAULT_SHEET_NAME,
    field_name: str = DEFAULT_FIELD_NAME,
) -> FileMonthSummary:
    try:
        with pd.ExcelFile(workbook_path) as excel_file:
            if sheet_name not in excel_file.sheet_names:
                return FileMonthSummary(
                    path=workbook_path,
                    status="missing_sheet",
                    row_count=0,
                    non_empty_count=0,
                    valid_count=0,
                    invalid_count=0,
                    months=(),
                )

            dataframe = excel_file.parse(sheet_name=sheet_name, dtype=object)
    except Exception as exc:  # pragma: no cover - 防御性分支
        return FileMonthSummary(
            path=workbook_path,
            status="read_error",
            row_count=0,
            non_empty_count=0,
            valid_count=0,
            invalid_count=0,
            months=(),
            error_message=str(exc),
        )
    row_count = len(dataframe.index)

    if field_name not in dataframe.columns:
        return FileMonthSummary(
            path=workbook_path,
            status="missing_field",
            row_count=row_count,
            non_empty_count=0,
            valid_count=0,
            invalid_count=0,
            months=(),
        )

    raw_values = [value for value in dataframe[field_name].tolist() if not is_blank(value)]
    non_empty_count = len(raw_values)

    if non_empty_count == 0:
        return FileMonthSummary(
            path=workbook_path,
            status="no_valid_values",
            row_count=row_count,
            non_empty_count=0,
            valid_count=0,
            invalid_count=0,
            months=(),
        )

    parsed_series = pd.to_datetime(pd.Series(raw_values, dtype=object), errors="coerce")
    valid_series = parsed_series.dropna()
    valid_count = len(valid_series.index)
    invalid_count = non_empty_count - valid_count
    months = tuple(sorted({timestamp.strftime("%Y-%m") for timestamp in valid_series}))

    status = "ok" if valid_count > 0 else "no_valid_values"

    return FileMonthSummary(
        path=workbook_path,
        status=status,
        row_count=row_count,
        non_empty_count=non_empty_count,
        valid_count=valid_count,
        invalid_count=invalid_count,
        months=months,
    )


def analyze_directory(
    input_dir: Path,
    *,
    recursive: bool = False,
    sheet_name: str = DEFAULT_SHEET_NAME,
    field_name: str = DEFAULT_FIELD_NAME,
) -> DirectoryMonthSummary:
    file_summaries = tuple(
        summarize_workbook(path, sheet_name=sheet_name, field_name=field_name)
        for path in iter_excel_paths(input_dir, recursive=recursive)
    )
    detected_months = tuple(sorted({month for item in file_summaries for month in item.months}))
    all_valid_values_in_one_month = len(detected_months) == 1 and any(
        item.valid_count > 0 for item in file_summaries
    )
    all_files_in_one_month = (
        bool(file_summaries)
        and len(detected_months) == 1
        and all(item.status == "ok" and len(item.months) == 1 for item in file_summaries)
    )
    return DirectoryMonthSummary(
        input_dir=input_dir.resolve(),
        file_summaries=file_summaries,
        detected_months=detected_months,
        all_valid_values_in_one_month=all_valid_values_in_one_month,
        all_files_in_one_month=all_files_in_one_month,
    )


def describe_file_summary(item: FileMonthSummary) -> str:
    if item.status == "missing_sheet":
        return f"缺少 {DEFAULT_SHEET_NAME} sheet"

    if item.status == "missing_field":
        return f"字段缺失（{DEFAULT_SHEET_NAME} sheet 中未找到 {DEFAULT_FIELD_NAME}，共 {item.row_count} 行）"

    if item.status == "read_error":
        return f"读取失败（{item.error_message or '未知错误'}）"

    if item.status == "no_valid_values":
        if item.non_empty_count == 0:
            return f"无有效时间值（空值 {item.empty_count} 条）"
        return (
            "无有效时间值"
            f"（非空 {item.non_empty_count} 条，无法解析 {item.invalid_count} 条）"
        )

    months_text = ", ".join(item.months)
    if len(item.months) > 1:
        return (
            f"{months_text}（跨月；有效 {item.valid_count} 条，"
            f"空值 {item.empty_count} 条，无法解析 {item.invalid_count} 条）"
        )

    return (
        f"{months_text}（有效 {item.valid_count} 条，"
        f"空值 {item.empty_count} 条，无法解析 {item.invalid_count} 条）"
    )


def format_directory_summary(summary: DirectoryMonthSummary) -> str:
    lines = [
        f"目录: {summary.input_dir}",
        f"扫描文件数: {len(summary.file_summaries)}",
        "",
        "整体结论：",
        f"- 所有有效“{DEFAULT_FIELD_NAME}”是否都属于同一个月: {'是' if summary.all_valid_values_in_one_month else '否'}",
        f"- 检测到的月份: {', '.join(summary.detected_months) if summary.detected_months else '未检测到'}",
        f"- 是否每个文件都能确认且只属于同一个月: {'是' if summary.all_files_in_one_month else '否'}",
        "",
        "文件明细：",
    ]

    if not summary.file_summaries:
        lines.append("- 未找到可扫描的 xlsx 文件")
        return "\n".join(lines)

    for item in summary.file_summaries:
        lines.append(f"- {item.path.name}: {describe_file_summary(item)}")
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="遍历指定目录下的 xlsx 文件，检查问卷数据 sheet 中“开始填表时间”是否都属于同一个月。"
    )
    parser.add_argument("--input-dir", required=True, help="要扫描的目录")
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="是否递归扫描子目录中的 xlsx 文件",
    )
    parser.add_argument(
        "--sheet-name",
        default=DEFAULT_SHEET_NAME,
        help=f"要读取的 sheet 名，默认 {DEFAULT_SHEET_NAME}",
    )
    parser.add_argument(
        "--field-name",
        default=DEFAULT_FIELD_NAME,
        help=f"要检查的字段名，默认 {DEFAULT_FIELD_NAME}",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    summary = analyze_directory(
        Path(args.input_dir),
        recursive=args.recursive,
        sheet_name=args.sheet_name,
        field_name=args.field_name,
    )
    print(format_directory_summary(summary))


if __name__ == "__main__":
    main()
