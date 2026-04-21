from __future__ import annotations

import argparse
from pathlib import Path

from pipeline_config import load_pipeline_defaults
from pipeline_paths import build_pipeline_paths, parse_single_month_batch
from pipeline_runtime import run_pipeline


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按固定目录约定执行主流程：预查错、分项统计、汇总、样本、PPT。"
    )
    parser.add_argument("--year", required=True, help="年份目录，例如 2026")
    parser.add_argument("--batch", required=True, help="批次目录，例如 3月 / 1-2月 / Q1")
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("pipeline.defaults.toml"),
        help="全局默认配置文件，默认 pipeline.defaults.toml",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    defaults = load_pipeline_defaults(args.config)
    paths = build_pipeline_paths(args.year, args.batch)
    single_month = parse_single_month_batch(args.batch)
    run_pipeline(paths=paths, defaults=defaults, single_month=single_month)


if __name__ == "__main__":
    main()
