from __future__ import annotations

import re
from pathlib import Path

from pipeline_models import BatchRef, PipelinePaths

STANDARD_SOURCE_FILE_NAMES: tuple[str, ...] = (
    "展览.xlsx",
    "会议.xlsx",
    "酒店.xlsx",
    "餐饮.xlsx",
    "会展服务商.xlsx",
    "旅游.xlsx",
)

SINGLE_MONTH_BATCH_RE = re.compile(r"^(0?[1-9]|1[0-2])月$")


def parse_single_month_batch(batch: str) -> int | None:
    normalized_batch = str(batch).strip()
    match = SINGLE_MONTH_BATCH_RE.fullmatch(normalized_batch)
    if match is None:
        return None
    return int(match.group(1))


def build_pipeline_paths(
    year: str,
    batch: str,
    *,
    data_root: Path = Path("data"),
    logs_root: Path = Path("logs/pipeline"),
) -> PipelinePaths:
    batch_ref = BatchRef(year=str(year).strip(), batch=str(batch).strip())
    raw_dir = data_root / "raw" / batch_ref.year / batch_ref.batch
    satisfaction_detail_dir = data_root / "satisfaction_detail" / batch_ref.year / batch_ref.batch
    satisfaction_summary_dir = data_root / "satisfaction_summary" / batch_ref.year / batch_ref.batch
    sample_summary_dir = data_root / "sample_summary" / batch_ref.year / batch_ref.batch
    ppt_dir = data_root / "ppt" / batch_ref.year / batch_ref.batch
    logs_dir = logs_root / batch_ref.year / batch_ref.batch

    return PipelinePaths(
        batch_ref=batch_ref,
        data_root=data_root,
        logs_root=logs_root,
        raw_dir=raw_dir,
        satisfaction_detail_dir=satisfaction_detail_dir,
        satisfaction_summary_dir=satisfaction_summary_dir,
        sample_summary_dir=sample_summary_dir,
        ppt_dir=ppt_dir,
        logs_dir=logs_dir,
        summary_workbook_path=satisfaction_summary_dir / f"{batch_ref.batch}客户类型满意度汇总表.xlsx",
        sample_workbook_path=sample_summary_dir / f"{batch_ref.batch}客户类型样本统计表.xlsx",
        ppt_path=ppt_dir / f"{batch_ref.batch}满意度报告.pptx",
        precheck_log_path=logs_dir / "precheck.log",
        pipeline_log_path=logs_dir / "pipeline.log",
        unmapped_log_path=logs_dir / "unmapped_customer_records.log",
        standard_source_paths=tuple(raw_dir / file_name for file_name in STANDARD_SOURCE_FILE_NAMES),
    )
