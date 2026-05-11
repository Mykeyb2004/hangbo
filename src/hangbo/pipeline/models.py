from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal


@dataclass(frozen=True)
class BatchRef:
    year: str
    batch: str


@dataclass(frozen=True)
class PipelinePaths:
    batch_ref: BatchRef
    data_root: Path
    logs_root: Path
    raw_dir: Path
    satisfaction_detail_dir: Path
    satisfaction_summary_dir: Path
    sample_summary_dir: Path
    ppt_dir: Path
    logs_dir: Path
    summary_workbook_path: Path
    sample_workbook_path: Path
    ppt_path: Path
    precheck_log_path: Path
    pipeline_log_path: Path
    unmapped_log_path: Path
    standard_source_paths: tuple[Path, ...]


@dataclass(frozen=True)
class PipelineIssue:
    severity: Literal["blocking", "warning"]
    code: str
    message: str
    path: Path | None = None


@dataclass(frozen=True)
class PrecheckResult:
    blocking_issues: tuple[PipelineIssue, ...]
    warning_issues: tuple[PipelineIssue, ...]
    should_autofill_year_month: bool
