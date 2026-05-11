from __future__ import annotations

import argparse
from pathlib import Path

from hangbo.merge.questionnaire_workbooks import format_merge_summary
from hangbo.merge.sample_summary import (
    MergeSampleRunConfig,
    build_merge_sample_paths,
    confirm_overwrite_if_needed,
    discover_source_directories,
    prompt_batch_name,
    run_merge_sample_summary,
    select_directories,
)
from hangbo.pipeline.config import load_pipeline_defaults


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="合并多个原始数据目录并生成客户类型样本统计表。")
    parser.add_argument("--year", required=True, help="年份，例如 2026")
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("pipeline.defaults.toml"),
        help="流水线默认配置文件路径",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    defaults = load_pipeline_defaults(args.config)
    raw_year_dir = Path("data") / "raw" / args.year
    source_dirs = discover_source_directories(raw_year_dir)
    selected_dirs = select_directories(source_dirs)
    batch_name = prompt_batch_name(selected_dirs)
    paths = build_merge_sample_paths(
        year=args.year,
        batch_name=batch_name,
        data_root=Path("data"),
    )
    if not confirm_overwrite_if_needed(paths):
        print("已取消。")
        return

    result = run_merge_sample_summary(
        MergeSampleRunConfig(
            year=args.year,
            batch_name=batch_name,
            selected_dirs=selected_dirs,
            data_root=Path("data"),
            sheet_name=defaults.sheet_name,
            sample_config_path=defaults.sample_config_path,
            overwrite=True,
        )
    )
    print(format_merge_summary(result.merge_summary, sheet_name=defaults.sheet_name))
    print(f"样本统计表：{result.sample_summary_path}")


if __name__ == "__main__":
    main()
