from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from fill_year_month_columns import apply_year_month_to_directory
from merge_questionnaire_workbooks import (
    MergeSummary,
    format_merge_summary,
    merge_workbooks_by_filename,
)
from pipeline_paths import parse_single_month_batch
from pipeline_precheck import workbook_has_year_month_headers
from sample_table import generate_sample_table_report


class BatchNameError(ValueError):
    pass


class MixedSourceYearMonthError(ValueError):
    pass


class SourcePreparationError(ValueError):
    pass


@dataclass(frozen=True)
class MergeSamplePaths:
    year: str
    batch_name: str
    data_root: Path
    raw_year_dir: Path
    merged_raw_dir: Path
    sample_summary_dir: Path
    sample_summary_path: Path


@dataclass(frozen=True)
class MergeSampleRunConfig:
    year: str
    batch_name: str
    selected_dirs: tuple[Path, ...]
    data_root: str | Path
    sheet_name: str
    sample_config_path: Path
    overwrite: bool = False


@dataclass(frozen=True)
class MergeSampleRunResult:
    paths: MergeSamplePaths
    merge_summary: MergeSummary
    sample_summary_path: Path


def build_merge_sample_paths(
    *,
    year: str,
    batch_name: str,
    data_root: str | Path = Path("data"),
) -> MergeSamplePaths:
    year = year.strip()
    batch_name = batch_name.strip()
    data_root = Path(data_root)
    raw_year_dir = data_root / "raw" / year
    merged_raw_dir = raw_year_dir / batch_name
    sample_summary_dir = data_root / "sample_summary" / year / batch_name
    sample_summary_path = sample_summary_dir / f"{batch_name}客户类型样本统计表.xlsx"

    return MergeSamplePaths(
        year=year,
        batch_name=batch_name,
        data_root=data_root,
        raw_year_dir=raw_year_dir,
        merged_raw_dir=merged_raw_dir,
        sample_summary_dir=sample_summary_dir,
        sample_summary_path=sample_summary_path,
    )


def discover_source_directories(raw_year_dir: Path) -> tuple[Path, ...]:
    if not raw_year_dir.is_dir():
        raise FileNotFoundError(f"年份原始数据目录不存在: {raw_year_dir}")

    source_dirs = tuple(
        sorted(
            (item for item in raw_year_dir.iterdir() if item.is_dir()),
            key=lambda item: item.name,
        )
    )
    if not source_dirs:
        raise ValueError(f"没有可选择的来源目录: {raw_year_dir}")

    return source_dirs


def parse_number_selection(raw_value: str, *, item_count: int) -> tuple[int, ...]:
    value = raw_value.strip()
    if not value:
        raise ValueError("至少选择一个来源目录")

    selected: list[int] = []
    seen: set[int] = set()
    for raw_part in value.split(","):
        part = raw_part.strip()
        if not part:
            raise ValueError("选择编号不能为空")

        if "-" in part:
            start_text, end_text = (item.strip() for item in part.split("-", maxsplit=1))
            start = _parse_selection_number(start_text)
            end = _parse_selection_number(end_text)
            if start > end:
                raise ValueError("范围起点不能大于终点")
            numbers = range(start, end + 1)
        else:
            numbers = (_parse_selection_number(part),)

        for number in numbers:
            if number < 1 or number > item_count:
                raise ValueError(f"选择编号超出范围: {number}")
            index = number - 1
            if index not in seen:
                selected.append(index)
                seen.add(index)

    if not selected:
        raise ValueError("至少选择一个来源目录")

    return tuple(selected)


def validate_batch_name(raw_name: str, selected_dirs: tuple[Path, ...]) -> str:
    batch_name = raw_name.strip()
    if not batch_name:
        raise BatchNameError("批次名称不能为空")
    if set(batch_name) == {"."}:
        raise BatchNameError("批次名称不能为当前或上级目录")
    if "/" in batch_name or "\\" in batch_name:
        raise BatchNameError("批次名称不能包含路径分隔符")
    if any(part == ".." for part in Path(batch_name).parts):
        raise BatchNameError("批次名称不能包含上级目录引用")
    if batch_name in {source_dir.name for source_dir in selected_dirs}:
        raise BatchNameError("批次名称不能与来源目录名称相同")

    return batch_name


def iter_source_excel_paths(source_dir: Path) -> tuple[Path, ...]:
    return tuple(
        sorted(
            (
                item
                for item in source_dir.iterdir()
                if item.is_file()
                and item.suffix == ".xlsx"
                and not item.name.startswith(("~$", "._"))
            ),
            key=lambda item: item.name,
        )
    )


def check_mixed_source_year_month_headers(source_dir: Path, *, sheet_name: str) -> None:
    workbook_paths = iter_source_excel_paths(source_dir)
    if not workbook_paths:
        raise MixedSourceYearMonthError(f"混合来源目录 {source_dir} 没有可用的 Excel 文件")

    for workbook_path in workbook_paths:
        if not workbook_has_year_month_headers(workbook_path, sheet_name):
            raise MixedSourceYearMonthError(
                f"混合来源目录 {source_dir} 中的文件 {workbook_path.name} 缺少“年份”/“月份”列"
            )


def prepare_source_directories(
    selected_dirs: tuple[Path, ...],
    *,
    year: str,
    sheet_name: str,
) -> None:
    for source_dir in selected_dirs:
        single_month = parse_single_month_batch(source_dir.name)
        if single_month is not None:
            summary = apply_year_month_to_directory(
                source_dir,
                year=str(year),
                month=str(single_month),
                sheet_name=sheet_name,
            )
            skipped_results = tuple(
                result for result in summary.file_results if result.status != "updated"
            )
            if skipped_results:
                first_result = skipped_results[0]
                raise SourcePreparationError(
                    f"单月来源目录 {source_dir} 年月填充失败: "
                    f"{first_result.path.name} status={first_result.status}"
                )
            continue

        check_mixed_source_year_month_headers(source_dir, sheet_name=sheet_name)


def merge_summary_has_failures(summary: MergeSummary) -> bool:
    return not summary.results or any(result.status != "merged" for result in summary.results)


def clear_generated_outputs(paths: MergeSamplePaths) -> None:
    if paths.merged_raw_dir.exists():
        for workbook_path in paths.merged_raw_dir.glob("*.xlsx"):
            if workbook_path.is_file():
                workbook_path.unlink()

    if paths.sample_summary_path.exists():
        paths.sample_summary_path.unlink()


def run_merge_sample_summary(config: MergeSampleRunConfig) -> MergeSampleRunResult:
    paths = build_merge_sample_paths(
        year=config.year,
        batch_name=config.batch_name,
        data_root=config.data_root,
    )

    if config.overwrite:
        clear_generated_outputs(paths)

    prepare_source_directories(
        config.selected_dirs,
        year=config.year,
        sheet_name=config.sheet_name,
    )
    merge_summary = merge_workbooks_by_filename(
        config.selected_dirs,
        output_dir=paths.merged_raw_dir,
        sheet_name=config.sheet_name,
    )
    if merge_summary_has_failures(merge_summary):
        raise RuntimeError(
            "合并阶段存在失败项，已停止生成样本统计表。\n"
            f"{format_merge_summary(merge_summary, sheet_name=config.sheet_name)}"
        )

    sample_summary_path = generate_sample_table_report(
        input_dir=paths.merged_raw_dir,
        output_dir=paths.sample_summary_dir,
        output_name=paths.sample_summary_path.name,
        config_path=config.sample_config_path,
        source_sheet_name=config.sheet_name,
        default_year=config.year,
    )

    return MergeSampleRunResult(
        paths=paths,
        merge_summary=merge_summary,
        sample_summary_path=sample_summary_path,
    )


def _parse_selection_number(raw_value: str) -> int:
    try:
        return int(raw_value)
    except ValueError as exc:
        raise ValueError(f"无效的选择编号: {raw_value}") from exc
