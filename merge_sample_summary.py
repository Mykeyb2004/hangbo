from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


class BatchNameError(ValueError):
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


def build_merge_sample_paths(
    *,
    year: str,
    batch_name: str,
    data_root: Path = Path("data"),
) -> MergeSamplePaths:
    year = year.strip()
    batch_name = batch_name.strip()
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
    if batch_name in {".", ".."}:
        raise BatchNameError("批次名称不能为当前或上级目录")
    if "/" in batch_name or "\\" in batch_name:
        raise BatchNameError("批次名称不能包含路径分隔符")
    if any(part == ".." for part in Path(batch_name).parts):
        raise BatchNameError("批次名称不能包含上级目录引用")
    if batch_name in {source_dir.name for source_dir in selected_dirs}:
        raise BatchNameError("批次名称不能与来源目录名称相同")

    return batch_name


def _parse_selection_number(raw_value: str) -> int:
    try:
        return int(raw_value)
    except ValueError as exc:
        raise ValueError(f"无效的选择编号: {raw_value}") from exc
