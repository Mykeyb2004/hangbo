# Main Pipeline Default Paths Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a new CLI-first main pipeline that runs by `year + batch`, uses fixed `data/...` directories plus one global defaults file, pauses on blocking precheck issues, and continues to full outputs after user confirmation.

**Architecture:** Keep the existing business engines (`survey_stats.py`, `summary_table.py`, `sample_table.py`, `generate_ppt.py`, `check_unmapped_customer_records.py`) intact as much as possible. Add a thin orchestration layer that resolves batch paths, loads one defaults file, classifies precheck results, loops on confirmation, and then calls the engines through programmatic APIs instead of subprocess-heavy chaining.

**Tech Stack:** Python 3.11, `tomllib`, `dataclasses`, `pathlib`, `openpyxl`, `pandas`, existing project modules, `unittest`, `unittest.mock`, `uv run pytest`

---

## File Map

### New Files

- `pipeline_models.py`
  - Shared dataclasses for batch identity, resolved paths, pipeline issues, precheck results, and runtime summaries.
- `pipeline_paths.py`
  - Fixed directory contract, standard raw source filenames, batch month parsing, and output filename derivation.
- `pipeline_config.py`
  - Loader and validators for `pipeline.defaults.toml`, plus config dataclasses for pipeline and PPT defaults.
- `pipeline_precheck.py`
  - Raw directory checks, standard file checks, year/month checks, and unmapped-record audit wrapping.
- `pipeline_runtime.py`
  - Main orchestration loop, user confirmation loop, autofill trigger for single-month batches, and engine execution order.
- `main_pipeline.py`
  - CLI shell that parses args, resolves defaults and paths, and calls `pipeline_runtime.run_pipeline`.
- `pipeline.defaults.toml`
  - One global defaults file for sheet name, calculation mode, sample config path, and PPT defaults.
- `tests/test_pipeline_paths.py`
  - Unit tests for path resolution and batch parsing.
- `tests/test_pipeline_config.py`
  - Unit tests for defaults loading and relative-path resolution.
- `tests/test_pipeline_precheck.py`
  - Unit tests for blocking/warning classification.
- `tests/test_pipeline_runtime.py`
  - Orchestration tests for recheck loop, autofill, and engine call order.
- `tests/test_main_pipeline.py`
  - CLI parse and entry wiring tests.

### Modified Files

- `survey_stats.py`
  - Extract a programmatic directory-run entrypoint that does not require batch TOML files.
- `README.md`
  - Add the new recommended `main_pipeline.py` workflow and explain compatibility with old config-driven flow.
- `docs/workflow.md`
  - Update the workflow doc to describe the new CLI-first main pipeline entry.
- `docs/从原始数据到PPT全流程命令.md`
  - Add the new `main_pipeline.py` command and explain the confirm-to-continue precheck stage.

### Existing Files Reused As-Is

- `summary_table.py`
- `sample_table.py`
- `generate_ppt.py`
- `check_unmapped_customer_records.py`
- `fill_year_month_columns.py`
- `sample_table.default.toml`

---

### Task 1: Add Path Models And Fixed Directory Resolution

**Files:**
- Create: `pipeline_models.py`
- Create: `pipeline_paths.py`
- Test: `tests/test_pipeline_paths.py`

- [ ] **Step 1: Write the failing path-resolution tests**

```python
from __future__ import annotations

from pathlib import Path
import unittest

from pipeline_paths import (
    STANDARD_SOURCE_FILE_NAMES,
    build_pipeline_paths,
    parse_single_month_batch,
)


class PipelinePathsTest(unittest.TestCase):
    def test_build_pipeline_paths_uses_fixed_directory_contract(self) -> None:
        paths = build_pipeline_paths(
            year="2026",
            batch="3月",
            data_root=Path("data"),
            logs_root=Path("logs/pipeline"),
        )

        self.assertEqual(paths.raw_dir, Path("data/raw/2026/3月"))
        self.assertEqual(paths.satisfaction_detail_dir, Path("data/satisfaction_detail/2026/3月"))
        self.assertEqual(paths.satisfaction_summary_dir, Path("data/satisfaction_summary/2026/3月"))
        self.assertEqual(paths.sample_summary_dir, Path("data/sample_summary/2026/3月"))
        self.assertEqual(paths.ppt_dir, Path("data/ppt/2026/3月"))
        self.assertEqual(paths.logs_dir, Path("logs/pipeline/2026/3月"))
        self.assertEqual(paths.summary_workbook_path.name, "3月客户类型满意度汇总表.xlsx")
        self.assertEqual(paths.sample_workbook_path.name, "3月客户类型样本统计表.xlsx")
        self.assertEqual(paths.ppt_path.name, "3月满意度报告.pptx")

    def test_build_pipeline_paths_exposes_standard_source_file_paths(self) -> None:
        paths = build_pipeline_paths(year="2026", batch="Q1")
        self.assertEqual(
            tuple(paths.standard_source_paths),
            tuple(Path("data/raw/2026/Q1") / file_name for file_name in STANDARD_SOURCE_FILE_NAMES),
        )

    def test_parse_single_month_batch_returns_month_number(self) -> None:
        self.assertEqual(parse_single_month_batch("3月"), 3)
        self.assertEqual(parse_single_month_batch("03月"), 3)

    def test_parse_single_month_batch_returns_none_for_combined_batches(self) -> None:
        self.assertIsNone(parse_single_month_batch("1-2月"))
        self.assertIsNone(parse_single_month_batch("Q1"))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
uv run pytest tests/test_pipeline_paths.py -q
```

Expected:

```text
E   ModuleNotFoundError: No module named 'pipeline_paths'
```

- [ ] **Step 3: Write the minimal path models and resolver**

`pipeline_models.py`

```python
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


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
```

`pipeline_paths.py`

```python
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
    match = SINGLE_MONTH_BATCH_RE.fullmatch(str(batch).strip())
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
    standard_source_paths = tuple(raw_dir / file_name for file_name in STANDARD_SOURCE_FILE_NAMES)
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
        standard_source_paths=standard_source_paths,
    )
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
uv run pytest tests/test_pipeline_paths.py -q
```

Expected:

```text
4 passed
```

- [ ] **Step 5: Commit**

```bash
git add pipeline_models.py pipeline_paths.py tests/test_pipeline_paths.py
git commit -m "feat: add fixed pipeline path resolver"
```

### Task 2: Add Global Defaults Loader

**Files:**
- Create: `pipeline_config.py`
- Create: `pipeline.defaults.toml`
- Test: `tests/test_pipeline_config.py`

- [ ] **Step 1: Write the failing defaults-loader tests**

```python
from __future__ import annotations

from pathlib import Path
import tempfile
import textwrap
import unittest

from pipeline_config import load_pipeline_defaults


class PipelineConfigTest(unittest.TestCase):
    def test_load_pipeline_defaults_reads_required_fields_and_resolves_relative_paths(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            config_path = root / "pipeline.defaults.toml"
            config_path.write_text(
                textwrap.dedent(
                    """
                    sheet_name = "问卷数据"
                    calculation_mode = "template"
                    sample_config_path = "sample_table.default.toml"

                    [ppt]
                    template_path = "templates/template.pptx"
                    sheet_name_mode = "first"
                    section_mode = "auto"
                    blank_display = ""
                    title_suffix = ""
                    max_single_table_rows = 18
                    max_split_table_rows = 19
                    sort_files = true
                    body_font_size_pt = 10.5
                    header_font_size_pt = 11.0
                    summary_font_size_pt = 12.0
                    template_slide_index = 0

                    [ppt.chart_page]
                    enabled = true
                    placeholder_text = "图表分析内容待补充。"
                    image_dpi = 220

                    [ppt.llm_notes]
                    enabled = false
                    env_path = ".env"
                    system_role_path = "system_role.md"
                    target_chars = 300
                    temperature = 0.4
                    max_tokens = 500
                    checkpoint_chars = 80
                    """
                ).strip()
                + "\n",
                encoding="utf-8",
            )

            defaults = load_pipeline_defaults(config_path)

            self.assertEqual(defaults.sheet_name, "问卷数据")
            self.assertEqual(defaults.calculation_mode, "template")
            self.assertEqual(defaults.sample_config_path, root / "sample_table.default.toml")
            self.assertEqual(defaults.ppt.template_path, root / "templates/template.pptx")
            self.assertEqual(defaults.ppt.llm_notes.env_path, root / ".env")

    def test_load_pipeline_defaults_rejects_invalid_calculation_mode(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            config_path = root / "pipeline.defaults.toml"
            config_path.write_text(
                textwrap.dedent(
                    """
                    sheet_name = "问卷数据"
                    calculation_mode = "bad-mode"
                    sample_config_path = "sample_table.default.toml"

                    [ppt]
                    template_path = "templates/template.pptx"
                    sheet_name_mode = "first"
                    section_mode = "auto"
                    blank_display = ""
                    title_suffix = ""
                    max_single_table_rows = 18
                    max_split_table_rows = 19
                    sort_files = true
                    body_font_size_pt = 10.5
                    header_font_size_pt = 11.0
                    summary_font_size_pt = 12.0
                    template_slide_index = 0

                    [ppt.chart_page]
                    enabled = false
                    placeholder_text = "图表分析内容待补充。"
                    image_dpi = 220

                    [ppt.llm_notes]
                    enabled = false
                    env_path = ".env"
                    system_role_path = "system_role.md"
                    target_chars = 300
                    temperature = 0.4
                    max_tokens = 500
                    checkpoint_chars = 80
                    """
                ).strip()
                + "\n",
                encoding="utf-8",
            )

            with self.assertRaises(ValueError):
                load_pipeline_defaults(config_path)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
uv run pytest tests/test_pipeline_config.py -q
```

Expected:

```text
E   ModuleNotFoundError: No module named 'pipeline_config'
```

- [ ] **Step 3: Write the minimal defaults dataclasses, loader, and global config file**

`pipeline_config.py`

```python
from __future__ import annotations

import tomllib
from dataclasses import dataclass
from pathlib import Path

from generate_ppt import normalize_section_mode
from survey_stats import normalize_calculation_mode


@dataclass(frozen=True)
class PipelineLlmNotesDefaults:
    enabled: bool
    env_path: Path
    system_role_path: Path
    target_chars: int
    temperature: float
    max_tokens: int
    checkpoint_chars: int


@dataclass(frozen=True)
class PipelineChartPageDefaults:
    enabled: bool
    placeholder_text: str
    image_dpi: int


@dataclass(frozen=True)
class PipelinePptDefaults:
    template_path: Path
    sheet_name_mode: str
    section_mode: str
    blank_display: str
    title_suffix: str
    max_single_table_rows: int
    max_split_table_rows: int
    sort_files: bool
    body_font_size_pt: float
    header_font_size_pt: float
    summary_font_size_pt: float
    template_slide_index: int
    chart_page: PipelineChartPageDefaults
    llm_notes: PipelineLlmNotesDefaults


@dataclass(frozen=True)
class PipelineDefaults:
    sheet_name: str
    calculation_mode: str
    sample_config_path: Path
    ppt: PipelinePptDefaults


def resolve_config_path(config_dir: Path, raw_path: str | Path) -> Path:
    path = Path(raw_path)
    if path.is_absolute():
        return path
    return config_dir / path


def load_pipeline_defaults(config_path: Path = Path("pipeline.defaults.toml")) -> PipelineDefaults:
    resolved_config_path = config_path.resolve()
    config_dir = resolved_config_path.parent
    raw = tomllib.loads(resolved_config_path.read_text(encoding="utf-8"))
    ppt_raw = raw["ppt"]
    chart_raw = ppt_raw["chart_page"]
    llm_raw = ppt_raw["llm_notes"]
    return PipelineDefaults(
        sheet_name=str(raw.get("sheet_name", "问卷数据")).strip() or "问卷数据",
        calculation_mode=normalize_calculation_mode(raw.get("calculation_mode", "template")),
        sample_config_path=resolve_config_path(config_dir, raw.get("sample_config_path", "sample_table.default.toml")),
        ppt=PipelinePptDefaults(
            template_path=resolve_config_path(config_dir, ppt_raw["template_path"]),
            sheet_name_mode=str(ppt_raw.get("sheet_name_mode", "first")),
            section_mode=normalize_section_mode(ppt_raw.get("section_mode", "auto")),
            blank_display=str(ppt_raw.get("blank_display", "")),
            title_suffix=str(ppt_raw.get("title_suffix", "")),
            max_single_table_rows=int(ppt_raw.get("max_single_table_rows", 18)),
            max_split_table_rows=int(ppt_raw.get("max_split_table_rows", 19)),
            sort_files=bool(ppt_raw.get("sort_files", True)),
            body_font_size_pt=float(ppt_raw.get("body_font_size_pt", 10.5)),
            header_font_size_pt=float(ppt_raw.get("header_font_size_pt", 11.0)),
            summary_font_size_pt=float(ppt_raw.get("summary_font_size_pt", 12.0)),
            template_slide_index=int(ppt_raw.get("template_slide_index", 0)),
            chart_page=PipelineChartPageDefaults(
                enabled=bool(chart_raw.get("enabled", True)),
                placeholder_text=str(chart_raw.get("placeholder_text", "图表分析内容待补充。")),
                image_dpi=int(chart_raw.get("image_dpi", 220)),
            ),
            llm_notes=PipelineLlmNotesDefaults(
                enabled=bool(llm_raw.get("enabled", False)),
                env_path=resolve_config_path(config_dir, llm_raw.get("env_path", ".env")),
                system_role_path=resolve_config_path(config_dir, llm_raw.get("system_role_path", "system_role.md")),
                target_chars=int(llm_raw.get("target_chars", 300)),
                temperature=float(llm_raw.get("temperature", 0.4)),
                max_tokens=int(llm_raw.get("max_tokens", 500)),
                checkpoint_chars=int(llm_raw.get("checkpoint_chars", 80)),
            ),
        ),
    )
```

`pipeline.defaults.toml`

```toml
sheet_name = "问卷数据"
calculation_mode = "template"
sample_config_path = "sample_table.default.toml"

[ppt]
template_path = "templates/template.pptx"
sheet_name_mode = "first"
section_mode = "auto"
blank_display = ""
title_suffix = ""
max_single_table_rows = 18
max_split_table_rows = 19
sort_files = true
body_font_size_pt = 10.5
header_font_size_pt = 11.0
summary_font_size_pt = 12.0
template_slide_index = 0

[ppt.chart_page]
enabled = true
placeholder_text = "图表分析内容待补充。后续将在此处补充该客户分组二级指标的整体解读、优势项与待提升项。"
image_dpi = 220

[ppt.llm_notes]
enabled = true
env_path = ".env"
system_role_path = "system_role.md"
target_chars = 300
temperature = 0.6
max_tokens = 500
checkpoint_chars = 80
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
uv run pytest tests/test_pipeline_config.py -q
```

Expected:

```text
2 passed
```

- [ ] **Step 5: Commit**

```bash
git add pipeline_config.py pipeline.defaults.toml tests/test_pipeline_config.py
git commit -m "feat: add pipeline defaults loader"
```

### Task 3: Build Precheck Classification

**Files:**
- Create: `pipeline_precheck.py`
- Modify: `pipeline_models.py`
- Test: `tests/test_pipeline_precheck.py`

- [ ] **Step 1: Write the failing precheck tests**

```python
from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
from unittest import mock

from openpyxl import Workbook

from pipeline_paths import STANDARD_SOURCE_FILE_NAMES, build_pipeline_paths
from pipeline_precheck import run_precheck


def write_workbook(path: Path, headers: list[str], *, sheet_name: str = "问卷数据") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(headers)
    worksheet.append(["value" for _ in headers])
    workbook.save(path)


class PipelinePrecheckTest(unittest.TestCase):
    def test_run_precheck_marks_missing_raw_dir_as_blocking(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths("2026", "3月", data_root=Path(tmp_dir) / "data", logs_root=Path(tmp_dir) / "logs")
            result = run_precheck(paths, sheet_name="问卷数据", single_month=3)
            self.assertEqual(len(result.blocking_issues), 1)
            self.assertIn("原始批次目录不存在", result.blocking_issues[0].message)

    def test_run_precheck_marks_missing_year_month_as_warning_for_single_month_batch(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths("2026", "3月", data_root=Path(tmp_dir) / "data", logs_root=Path(tmp_dir) / "logs")
            for file_name in STANDARD_SOURCE_FILE_NAMES:
                write_workbook(paths.raw_dir / file_name, ["姓名", "开始填表时间"])
            result = run_precheck(paths, sheet_name="问卷数据", single_month=3)
            self.assertTrue(result.should_autofill_year_month)
            self.assertFalse(result.blocking_issues)
            self.assertEqual(len(result.warning_issues), 1)

    def test_run_precheck_marks_missing_year_month_as_blocking_for_combined_batch(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths("2026", "Q1", data_root=Path(tmp_dir) / "data", logs_root=Path(tmp_dir) / "logs")
            for file_name in STANDARD_SOURCE_FILE_NAMES:
                write_workbook(paths.raw_dir / file_name, ["姓名", "开始填表时间"])
            result = run_precheck(paths, sheet_name="问卷数据", single_month=None)
            self.assertFalse(result.should_autofill_year_month)
            self.assertTrue(result.blocking_issues)
            self.assertIn("缺少“年份”/“月份”列", result.blocking_issues[0].message)

    def test_run_precheck_marks_missing_sheet_as_blocking(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths("2026", "3月", data_root=Path(tmp_dir) / "data", logs_root=Path(tmp_dir) / "logs")
            for file_name in STANDARD_SOURCE_FILE_NAMES:
                write_workbook(paths.raw_dir / file_name, ["年份", "月份"], sheet_name="其他sheet")
            result = run_precheck(paths, sheet_name="问卷数据", single_month=3)
            self.assertTrue(result.blocking_issues)
            self.assertIn("缺少 sheet", result.blocking_issues[0].message)

    @mock.patch("pipeline_precheck.run_unmapped_audit")
    def test_run_precheck_turns_unmapped_records_into_blocking_issue(self, mock_run_unmapped: mock.Mock) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths("2026", "3月", data_root=Path(tmp_dir) / "data", logs_root=Path(tmp_dir) / "logs")
            for file_name in STANDARD_SOURCE_FILE_NAMES:
                write_workbook(paths.raw_dir / file_name, ["年份", "月份", "开始填表时间"])
            mock_run_unmapped.return_value = (2, paths.unmapped_log_path)
            result = run_precheck(paths, sheet_name="问卷数据", single_month=3)
            self.assertEqual(len(result.blocking_issues), 1)
            self.assertIn("未映射标签", result.blocking_issues[0].message)

    @mock.patch("pipeline_precheck.run_unmapped_audit")
    def test_run_precheck_turns_audit_error_into_blocking_issue(self, mock_run_unmapped: mock.Mock) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths("2026", "3月", data_root=Path(tmp_dir) / "data", logs_root=Path(tmp_dir) / "logs")
            for file_name in STANDARD_SOURCE_FILE_NAMES:
                write_workbook(paths.raw_dir / file_name, ["年份", "月份", "开始填表时间"])
            mock_run_unmapped.side_effect = RuntimeError("读取客户映射失败")
            result = run_precheck(paths, sheet_name="问卷数据", single_month=3)
            self.assertEqual(len(result.blocking_issues), 1)
            self.assertIn("预查错过程失败", result.blocking_issues[0].message)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
uv run pytest tests/test_pipeline_precheck.py -q
```

Expected:

```text
E   ModuleNotFoundError: No module named 'pipeline_precheck'
```

- [ ] **Step 3: Implement issue models and precheck logic**

In `pipeline_models.py`, keep the existing `BatchRef` and `PipelinePaths` classes from Task 1. Add `Literal` to imports:

```python
from typing import Literal
```

Then append these classes after `PipelinePaths`:

```python
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
```

`pipeline_precheck.py`

```python
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from check_unmapped_customer_records import (
    format_directory_audit_report,
    run_directory_audit,
    write_audit_log,
)
from pipeline_models import PipelineIssue, PrecheckResult
from pipeline_paths import STANDARD_SOURCE_FILE_NAMES


def workbook_has_year_month_headers(workbook_path: Path, sheet_name: str) -> bool:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"缺少 sheet：{sheet_name}")
        worksheet = workbook[sheet_name]
        headers = {
            str(cell).strip()
            for cell in next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
            if cell is not None
        }
        return "年份" in headers and "月份" in headers
    finally:
        workbook.close()


def run_unmapped_audit(input_dir: Path, *, sheet_name: str, log_path: Path) -> tuple[int, Path]:
    report = run_directory_audit(input_dir, sheet_name=sheet_name)
    report_text = format_directory_audit_report(report, log_path=log_path)
    write_audit_log(report_text, log_path)
    return report.total_unmapped_records, log_path


def run_precheck(paths, *, sheet_name: str, single_month: int | None) -> PrecheckResult:
    blocking: list[PipelineIssue] = []
    warning: list[PipelineIssue] = []
    should_autofill_year_month = False

    if not paths.raw_dir.exists() or not paths.raw_dir.is_dir():
        return PrecheckResult(
            blocking_issues=(
                PipelineIssue("blocking", "missing_raw_dir", f"原始批次目录不存在：{paths.raw_dir}", paths.raw_dir),
            ),
            warning_issues=(),
            should_autofill_year_month=False,
        )

    present_source_paths = [paths.raw_dir / file_name for file_name in STANDARD_SOURCE_FILE_NAMES if (paths.raw_dir / file_name).exists()]
    if not present_source_paths:
        blocking.append(
            PipelineIssue("blocking", "missing_standard_sources", f"标准来源文件缺失：{paths.raw_dir}", paths.raw_dir)
        )
    for file_name in STANDARD_SOURCE_FILE_NAMES:
        source_path = paths.raw_dir / file_name
        if source_path.exists():
            continue
        blocking.append(PipelineIssue("blocking", "missing_source_file", f"缺少标准来源文件：{file_name}", source_path))

    if not blocking:
        files_missing_year_month = []
        for source_path in present_source_paths:
            try:
                has_year_month_headers = workbook_has_year_month_headers(source_path, sheet_name)
            except ValueError as exc:
                blocking.append(
                    PipelineIssue(
                        "blocking",
                        "missing_sheet",
                        str(exc),
                        source_path,
                    )
                )
                continue
            if not has_year_month_headers:
                files_missing_year_month.append(source_path)
        if files_missing_year_month and single_month is not None:
            should_autofill_year_month = True
            warning.append(
                PipelineIssue(
                    "warning",
                    "autofill_year_month",
                    f"单月批次缺少“年份”/“月份”列，将在继续执行前自动补写：{len(files_missing_year_month)} 个文件",
                )
            )
        elif files_missing_year_month:
            blocking.append(
                PipelineIssue(
                    "blocking",
                    "missing_year_month_columns",
                    f"组合批次缺少“年份”/“月份”列，需先修正原始数据：{files_missing_year_month[0].name}",
                    files_missing_year_month[0],
                )
            )

    if not blocking:
        try:
            total_unmapped_records, _ = run_unmapped_audit(
                paths.raw_dir,
                sheet_name=sheet_name,
                log_path=paths.unmapped_log_path,
            )
        except Exception as exc:
            blocking.append(
                PipelineIssue(
                    "blocking",
                    "precheck_error",
                    f"预查错过程失败，请先处理后重试：{exc}",
                    paths.raw_dir,
                )
            )
        else:
            if total_unmapped_records:
                blocking.append(
                    PipelineIssue(
                        "blocking",
                        "unmapped_customer_records",
                        f"存在未映射标签或标签组合，请先修正原始数据（共 {total_unmapped_records} 条）",
                        paths.unmapped_log_path,
                    )
                )

    return PrecheckResult(
        blocking_issues=tuple(blocking),
        warning_issues=tuple(warning),
        should_autofill_year_month=should_autofill_year_month,
    )
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
uv run pytest tests/test_pipeline_precheck.py -q
```

Expected:

```text
6 passed
```

- [ ] **Step 5: Commit**

```bash
git add pipeline_models.py pipeline_precheck.py tests/test_pipeline_precheck.py
git commit -m "feat: add pipeline precheck classification"
```

### Task 4: Add Programmatic Engine Entrypoints And Runtime Orchestration

**Files:**
- Create: `pipeline_runtime.py`
- Modify: `survey_stats.py`
- Test: `tests/test_pipeline_runtime.py`
- Test: `tests/test_survey_stats.py`

- [ ] **Step 1: Write the failing orchestration and programmatic-entry tests**

`tests/test_pipeline_runtime.py`

```python
from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
from unittest import mock

from pipeline_config import (
    PipelineChartPageDefaults,
    PipelineDefaults,
    PipelineLlmNotesDefaults,
    PipelinePptDefaults,
)
from pipeline_models import PipelineIssue, PrecheckResult
from pipeline_paths import build_pipeline_paths
from pipeline_runtime import run_pipeline, wait_for_confirmation


def build_defaults() -> PipelineDefaults:
    return PipelineDefaults(
        sheet_name="问卷数据",
        calculation_mode="template",
        sample_config_path=Path("sample_table.default.toml"),
        ppt=PipelinePptDefaults(
            template_path=Path("templates/template.pptx"),
            sheet_name_mode="first",
            section_mode="auto",
            blank_display="",
            title_suffix="",
            max_single_table_rows=18,
            max_split_table_rows=19,
            sort_files=True,
            body_font_size_pt=10.5,
            header_font_size_pt=11.0,
            summary_font_size_pt=12.0,
            template_slide_index=0,
            chart_page=PipelineChartPageDefaults(True, "图表分析内容待补充。", 220),
            llm_notes=PipelineLlmNotesDefaults(False, Path(".env"), Path("system_role.md"), 300, 0.4, 500, 80),
        ),
    )


class PipelineRuntimeTest(unittest.TestCase):
    def test_wait_for_confirmation_reprompts_until_y_yes_or_continue(self) -> None:
        prompts = iter(["不是", "y"])
        outputs: list[str] = []
        wait_for_confirmation(
            input_func=lambda prompt: next(prompts),
            output_func=outputs.append,
        )
        self.assertEqual(outputs, ["未识别的输入：不是"])

    @mock.patch("pipeline_runtime.generate_presentation")
    @mock.patch("pipeline_runtime.generate_sample_table_report")
    @mock.patch("pipeline_runtime.generate_summary_report")
    @mock.patch("pipeline_runtime.run_directory_batch")
    @mock.patch("pipeline_runtime.apply_year_month_to_directory")
    @mock.patch("pipeline_runtime.run_precheck")
    def test_run_pipeline_rechecks_after_confirmation_before_running_engines(
        self,
        mock_run_precheck: mock.Mock,
        mock_apply_year_month: mock.Mock,
        mock_run_directory_batch: mock.Mock,
        mock_generate_summary: mock.Mock,
        mock_generate_sample: mock.Mock,
        mock_generate_presentation: mock.Mock,
    ) -> None:
        blocking_issue = PipelineIssue("blocking", "unmapped", "存在未映射标签")
        mock_run_precheck.side_effect = [
            PrecheckResult((blocking_issue,), (), False),
            PrecheckResult((), (), True),
        ]
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=root / "data",
                logs_root=root / "logs/pipeline",
            )
            outputs: list[str] = []

            run_pipeline(
                paths=paths,
                defaults=build_defaults(),
                single_month=3,
                input_func=lambda prompt: "继续",
                output_func=outputs.append,
            )

            self.assertTrue(paths.pipeline_log_path.exists())
            self.assertTrue(paths.precheck_log_path.exists())
            self.assertTrue(paths.satisfaction_detail_dir.exists())
            self.assertTrue(paths.satisfaction_summary_dir.exists())
            self.assertTrue(paths.sample_summary_dir.exists())
            self.assertTrue(paths.ppt_dir.exists())
            self.assertTrue(any("输出目录不存在，已创建" in message for message in outputs))

        self.assertEqual(mock_run_precheck.call_count, 2)
        mock_apply_year_month.assert_called_once()
        mock_run_directory_batch.assert_called_once()
        mock_generate_summary.assert_called_once()
        mock_generate_sample.assert_called_once()
        mock_generate_presentation.assert_called_once()

    @mock.patch("pipeline_runtime.run_precheck")
    def test_run_pipeline_keeps_waiting_when_recheck_still_blocks(self, mock_run_precheck: mock.Mock) -> None:
        blocking_issue = PipelineIssue("blocking", "unmapped", "存在未映射标签")
        mock_run_precheck.side_effect = [
            PrecheckResult((blocking_issue,), (), False),
            PrecheckResult((blocking_issue,), (), False),
        ]
        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths(
                "2026",
                "Q1",
                data_root=Path(tmp_dir) / "data",
                logs_root=Path(tmp_dir) / "logs/pipeline",
            )
            prompts = iter(["yes", "stop"])

            with self.assertRaises(SystemExit):
                run_pipeline(
                    paths=paths,
                    defaults=build_defaults(),
                    single_month=None,
                    input_func=lambda prompt: next(prompts),
                    output_func=lambda message: None,
                )

    @mock.patch("pipeline_runtime.generate_summary_report")
    @mock.patch("pipeline_runtime.run_directory_batch")
    @mock.patch("pipeline_runtime.run_precheck")
    def test_run_pipeline_stops_when_engine_step_fails(
        self,
        mock_run_precheck: mock.Mock,
        mock_run_directory_batch: mock.Mock,
        mock_generate_summary: mock.Mock,
    ) -> None:
        mock_run_precheck.return_value = PrecheckResult((), (), False)
        mock_run_directory_batch.side_effect = RuntimeError("满意度分项统计失败")
        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(tmp_dir) / "data",
                logs_root=Path(tmp_dir) / "logs/pipeline",
            )
            with self.assertRaisesRegex(RuntimeError, "满意度分项统计失败"):
                run_pipeline(
                    paths=paths,
                    defaults=build_defaults(),
                    single_month=3,
                    output_func=lambda message: None,
                )

        mock_generate_summary.assert_not_called()
```

`tests/test_survey_stats.py` append:

First update imports:

```python
from unittest import mock

from survey_stats import (
    DirectoryDiscoveryResult,
    run_directory_batch,
)
```

Then append the test method inside `SurveyStatsTest`:

```python
    def test_run_directory_batch_uses_programmatic_directory_mode_without_config_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            input_dir = root / "data/raw/2026/3月"
            output_dir = root / "data/satisfaction_detail/2026/3月"
            input_dir.mkdir(parents=True)

            with mock.patch("survey_stats.discover_directory_jobs") as mock_discover, mock.patch(
                "survey_stats.generate_customer_category_report_bundle"
            ) as mock_generate:
                mock_discover.return_value = DirectoryDiscoveryResult(
                    jobs=(),
                    missing_customer_type_notices=(),
                    preprocess_notices=(),
                    unmapped_customer_category_notices=(),
                )

                run_directory_batch(
                    input_dir=input_dir,
                    output_dir=output_dir,
                    sheet_name="问卷数据",
                    output_format="xlsx",
                    calculation_mode="template",
                )

                mock_discover.assert_called_once()
                self.assertFalse(mock_generate.called)
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
uv run pytest tests/test_pipeline_runtime.py tests/test_survey_stats.py -q
```

Expected:

```text
E   ModuleNotFoundError: No module named 'pipeline_runtime'
E   ImportError: cannot import name 'run_directory_batch' from 'survey_stats'
```

- [ ] **Step 3: Extract the programmatic directory runner and implement the orchestration loop**

`survey_stats.py` add:

```python
def run_batch_config(
    config: BatchConfig,
    *,
    output_dir_override: Path | None = None,
    output_format_override: str | None = None,
    calculation_mode_override: str | None = None,
    selected_job_names: list[str] | tuple[str, ...] = (),
    dry_run: bool = False,
) -> None:
    output_dir = normalize_output_dir(output_dir_override or config.output_dir)
    global_output_format = output_format_override or config.output_format
    calculation_mode = normalize_calculation_mode(calculation_mode_override or config.calculation_mode)
    missing_group_notices: list[MissingGroupNotice] = []
    missing_customer_type_notices: list[MissingCustomerTypeNotice] = []
    unmapped_customer_category_notices: list[UnmappedCustomerCategoryNotice] = []
    preprocess_notice_lookup: dict[Path, str] = {}

    if config.input_dir is None:
        selected_jobs = select_jobs(config.jobs, list(selected_job_names))
        if not selected_jobs:
            raise ValueError("筛选后没有可运行的 jobs。")
    else:
        discovery_result = discover_directory_jobs(config)
        selected_jobs = select_jobs(discovery_result.jobs, list(selected_job_names))
        missing_customer_type_notices = list(
            select_missing_customer_type_notices(
                discovery_result.missing_customer_type_notices,
                list(selected_job_names),
            )
        )
        preprocess_notice_lookup = {
            record.input_path: record.notice for record in discovery_result.preprocess_notices
        }
        unmapped_customer_category_notices = list(discovery_result.unmapped_customer_category_notices)
        if not selected_jobs and not missing_customer_type_notices:
            raise ValueError("筛选后没有可运行的 jobs。")

    total_jobs = len(selected_jobs)
    for index, job in enumerate(selected_jobs, start=1):
        output_format = job.output_format or global_output_format
        output_path = build_output_path(output_dir, job.output_name, output_format)
        print_file_progress_start(index, total_jobs, job.path, job.name)
        if job.category_rule_name is None:
            role_definition = resolve_role_definition(job.template_name, job.role_name)
            report = generate_role_report_bundle(
                input_path=job.path,
                role_definition=role_definition,
                output_path=output_path,
                sheet_name=job.sheet_name,
                sheet_title=job.name,
                calculation_mode=calculation_mode,
                dry_run=dry_run,
                save_empty_report=config.input_dir is None,
            )
        else:
            category_rule = CUSTOMER_CATEGORY_RULE_BY_NAME[job.category_rule_name]
            report = generate_customer_category_report_bundle(
                input_path=job.path,
                category_rule=category_rule,
                output_path=output_path,
                sheet_name=job.sheet_name,
                calculation_mode=calculation_mode,
                dry_run=dry_run,
                save_empty_report=config.input_dir is None,
            )
        preprocess_notice = report.preprocess_notice or preprocess_notice_lookup.pop(job.path, None)
        print_preprocess_notice(index, total_jobs, preprocess_notice)
        if config.input_dir is not None and report.stats.matched_row_count == 0:
            missing_customer_type_notices.append(
                MissingCustomerTypeNotice(
                    customer_type_name=job.name,
                    source_reference=job.path.name,
                    sheet_name=job.sheet_name,
                    reason=DIRECTORY_NOTICE_REASON_MISSING_ROLE_DATA,
                )
            )
            continue
        print_file_progress_result(index, total_jobs, job.path, job.name, report.output_path, dry_run=dry_run)
        if report.stats.matched_row_count == 0:
            missing_group_notices.append(MissingGroupNotice(job.name, job.path, job.sheet_name))

    if config.input_dir is None:
        print_missing_group_summary(missing_group_notices)
    else:
        print_missing_customer_type_summary(missing_customer_type_notices)
        print_unmapped_customer_category_summary(unmapped_customer_category_notices)


def run_directory_batch(
    *,
    input_dir: Path,
    output_dir: Path,
    sheet_name: str = DEFAULT_SHEET_NAME,
    output_format: str = DEFAULT_OUTPUT_FORMAT,
    calculation_mode: str = DEFAULT_CALCULATION_MODE,
    job_filters: list[str] | tuple[str, ...] = (),
    dry_run: bool = False,
) -> None:
    config = BatchConfig(
        config_path=Path("<programmatic-directory-mode>"),
        output_dir=output_dir,
        output_format=output_format,
        calculation_mode=normalize_calculation_mode(calculation_mode),
        sheet_name=sheet_name,
        input_dir=input_dir,
    )
    run_batch_config(
        config,
        output_dir_override=output_dir,
        output_format_override=output_format,
        calculation_mode_override=calculation_mode,
        selected_job_names=job_filters,
        dry_run=dry_run,
    )
```

`pipeline_runtime.py`

```python
from __future__ import annotations

from pathlib import Path

from fill_year_month_columns import apply_year_month_to_directory
from generate_ppt import (
    ChartPageConfig,
    LlmNotesConfig,
    PptBatchConfig,
    PptLayoutConfig,
    generate_presentation,
)
from pipeline_precheck import run_precheck
from sample_table import generate_sample_table_report
from summary_table import generate_summary_report
from survey_stats import run_directory_batch


def append_log_line(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as file:
        file.write(message + "\n")


def emit(paths, output_func, message: str) -> None:
    output_func(message)
    append_log_line(paths.pipeline_log_path, message)


def write_precheck_log(paths, precheck) -> None:
    lines = ["预查错结果"]
    if precheck.blocking_issues:
        lines.append("阻断问题：")
        lines.extend(f"- {issue.message}" for issue in precheck.blocking_issues)
    if precheck.warning_issues:
        lines.append("警告：")
        lines.extend(f"- {issue.message}" for issue in precheck.warning_issues)
    if not precheck.blocking_issues and not precheck.warning_issues:
        lines.append("未发现问题。")
    paths.precheck_log_path.parent.mkdir(parents=True, exist_ok=True)
    paths.precheck_log_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def wait_for_confirmation(*, input_func=input, output_func=print) -> None:
    while True:
        answer = input_func("修改完成后输入 y / yes / 继续；输入 stop 终止：").strip().lower()
        if answer in {"y", "yes", "继续"}:
            return
        if answer in {"stop", "quit", "exit"}:
            raise SystemExit("用户取消主流程。")
        output_func(f"未识别的输入：{answer}")


def ensure_output_directories(paths, output_func=print) -> None:
    output_dirs = (
        paths.satisfaction_detail_dir,
        paths.satisfaction_summary_dir,
        paths.sample_summary_dir,
        paths.ppt_dir,
    )
    for output_dir in output_dirs:
        if output_dir.exists():
            continue
        output_dir.mkdir(parents=True, exist_ok=True)
        emit(paths, output_func, f"[WARN] 输出目录不存在，已创建：{output_dir}")


def build_ppt_batch_config(paths, defaults) -> PptBatchConfig:
    return PptBatchConfig(
        template_path=defaults.ppt.template_path,
        input_dir=paths.satisfaction_detail_dir,
        output_ppt=paths.ppt_path,
        sheet_name_mode=defaults.ppt.sheet_name_mode,
        blank_display=defaults.ppt.blank_display,
        title_suffix=defaults.ppt.title_suffix,
        section_mode=defaults.ppt.section_mode,
        max_single_table_rows=defaults.ppt.max_single_table_rows,
        max_split_table_rows=defaults.ppt.max_split_table_rows,
        sort_files=defaults.ppt.sort_files,
        layout=PptLayoutConfig(),
        llm_notes=LlmNotesConfig(
            enabled=defaults.ppt.llm_notes.enabled,
            env_path=defaults.ppt.llm_notes.env_path,
            system_role_path=defaults.ppt.llm_notes.system_role_path,
            target_chars=defaults.ppt.llm_notes.target_chars,
            temperature=defaults.ppt.llm_notes.temperature,
            max_tokens=defaults.ppt.llm_notes.max_tokens,
            checkpoint_chars=defaults.ppt.llm_notes.checkpoint_chars,
        ),
        body_font_size_pt=defaults.ppt.body_font_size_pt,
        header_font_size_pt=defaults.ppt.header_font_size_pt,
        summary_font_size_pt=defaults.ppt.summary_font_size_pt,
        template_slide_index=defaults.ppt.template_slide_index,
        chart_page=ChartPageConfig(
            enabled=defaults.ppt.chart_page.enabled,
            placeholder_text=defaults.ppt.chart_page.placeholder_text,
            image_dpi=defaults.ppt.chart_page.image_dpi,
        ),
    )


def run_pipeline(
    *,
    paths,
    defaults,
    single_month: int | None,
    input_func=input,
    output_func=print,
) -> None:
    while True:
        precheck = run_precheck(paths, sheet_name=defaults.sheet_name, single_month=single_month)
        write_precheck_log(paths, precheck)
        for issue in precheck.warning_issues:
            emit(paths, output_func, f"[WARN] {issue.message}")
        if not precheck.blocking_issues:
            break
        emit(paths, output_func, "[PRECHECK] 发现阻断问题：")
        for issue in precheck.blocking_issues:
            emit(paths, output_func, f"- {issue.message}")
        emit(paths, output_func, f"请先修改原始数据目录：{paths.raw_dir}")
        wait_for_confirmation(input_func=input_func, output_func=output_func)
        emit(paths, output_func, "[PRECHECK] 正在重新检查...")

    ensure_output_directories(paths, output_func)

    if precheck.should_autofill_year_month and single_month is not None:
        emit(paths, output_func, "[PIPELINE] 检测为单月批次，正在补写年份/月份...")
        apply_year_month_to_directory(
            paths.raw_dir,
            year=paths.batch_ref.year,
            month=str(single_month),
            sheet_name=defaults.sheet_name,
        )

    emit(paths, output_func, "[PIPELINE] 开始生成满意度分项统计...")
    run_directory_batch(
        input_dir=paths.raw_dir,
        output_dir=paths.satisfaction_detail_dir,
        sheet_name=defaults.sheet_name,
        output_format="xlsx",
        calculation_mode=defaults.calculation_mode,
    )
    emit(paths, output_func, "[PIPELINE] 开始生成满意度汇总表...")
    generate_summary_report(
        input_dir=paths.satisfaction_detail_dir,
        output_dir=paths.satisfaction_summary_dir,
        output_name=paths.summary_workbook_path.name,
    )
    emit(paths, output_func, "[PIPELINE] 开始生成样本统计表...")
    generate_sample_table_report(
        input_dir=paths.raw_dir,
        output_dir=paths.sample_summary_dir,
        output_name=paths.sample_workbook_path.name,
        config_path=defaults.sample_config_path,
        source_sheet_name=defaults.sheet_name,
        default_year=paths.batch_ref.year,
    )
    emit(paths, output_func, "[PIPELINE] 开始生成 PPT...")
    generate_presentation(build_ppt_batch_config(paths, defaults))
    emit(paths, output_func, f"[PIPELINE] PPT 已生成：{paths.ppt_path}")
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
uv run pytest tests/test_pipeline_runtime.py tests/test_survey_stats.py -q
```

Expected:

```text
all selected tests passed
```

- [ ] **Step 5: Commit**

```bash
git add pipeline_runtime.py survey_stats.py tests/test_pipeline_runtime.py tests/test_survey_stats.py
git commit -m "feat: add programmatic pipeline orchestration"
```

### Task 5: Add CLI Entry And Update Docs

**Files:**
- Create: `main_pipeline.py`
- Test: `tests/test_main_pipeline.py`
- Modify: `README.md`
- Modify: `docs/workflow.md`
- Modify: `docs/从原始数据到PPT全流程命令.md`

- [ ] **Step 1: Write the failing CLI entry tests**

```python
from __future__ import annotations

import unittest
from unittest import mock

from main_pipeline import main, parse_args


class MainPipelineCliTest(unittest.TestCase):
    def test_parse_args_reads_year_batch_and_config(self) -> None:
        args = parse_args(["--year", "2026", "--batch", "3月", "--config", "pipeline.defaults.toml"])
        self.assertEqual(args.year, "2026")
        self.assertEqual(args.batch, "3月")
        self.assertEqual(str(args.config), "pipeline.defaults.toml")

    @mock.patch("main_pipeline.run_pipeline")
    @mock.patch("main_pipeline.load_pipeline_defaults")
    @mock.patch("main_pipeline.build_pipeline_paths")
    def test_main_builds_paths_and_invokes_runtime(
        self,
        mock_build_pipeline_paths: mock.Mock,
        mock_load_pipeline_defaults: mock.Mock,
        mock_run_pipeline: mock.Mock,
    ) -> None:
        main(["--year", "2026", "--batch", "3月"])
        mock_build_pipeline_paths.assert_called_once_with("2026", "3月")
        mock_load_pipeline_defaults.assert_called_once()
        mock_run_pipeline.assert_called_once()


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
uv run pytest tests/test_main_pipeline.py -q
```

Expected:

```text
E   ModuleNotFoundError: No module named 'main_pipeline'
```

- [ ] **Step 3: Implement the CLI shell and update the docs**

`main_pipeline.py`

```python
from __future__ import annotations

import argparse
from pathlib import Path

from pipeline_config import load_pipeline_defaults
from pipeline_paths import build_pipeline_paths, parse_single_month_batch
from pipeline_runtime import run_pipeline


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按固定目录约定执行主流程：预查错、分项统计、汇总、样本、PPT。")
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
    run_pipeline(
        paths=paths,
        defaults=defaults,
        single_month=single_month,
    )


if __name__ == "__main__":
    main()
```

`README.md` add:

````markdown
## 新主流程入口

推荐使用新的主流程 CLI：

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

它会按固定目录约定读取：

- `data/raw/{year}/{batch}`

并自动输出到：

- `data/satisfaction_detail/{year}/{batch}`
- `data/satisfaction_summary/{year}/{batch}`
- `data/sample_summary/{year}/{batch}`
- `data/ppt/{year}/{batch}`

若预查错发现阻断问题，程序会停下等待人工修正，确认后重新检查并继续执行。
````

`docs/workflow.md` add one new top-level section describing:

````markdown
## 新 CLI 主流程

当前推荐入口：

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

执行顺序：

1. 预查错
2. 人工修正后确认继续
3. 满意度分项统计
4. 满意度汇总
5. 样本汇总
6. PPT 生成
````

`docs/从原始数据到PPT全流程命令.md` replace the “three separate commands” recommendation with:

````markdown
## 推荐主流程命令

```bash
uv run python main_pipeline.py --year 2026 --batch 1-2月
uv run python main_pipeline.py --year 2026 --batch 3月
uv run python main_pipeline.py --year 2026 --batch Q1
```

说明：

- 若预查错发现阻断问题，程序会暂停
- 修正原始数据后，在终端确认继续
- 通过后自动完成分项统计、汇总表、样本表和 PPT
````

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
uv run pytest tests/test_main_pipeline.py -q
```

Expected:

```text
2 passed
```

Then run a focused regression sweep:

```bash
uv run pytest tests/test_pipeline_paths.py tests/test_pipeline_config.py tests/test_pipeline_precheck.py tests/test_pipeline_runtime.py tests/test_main_pipeline.py tests/test_survey_stats.py -q
```

Expected:

```text
all selected tests passed
```

- [ ] **Step 5: Commit**

```bash
git add main_pipeline.py README.md docs/workflow.md docs/从原始数据到PPT全流程命令.md tests/test_main_pipeline.py
git commit -m "feat: add main pipeline cli entry"
```

## Self-Review Checklist

- Spec coverage:
  - Fixed `data/...` directory contract: covered by Task 1
  - One global defaults file: covered by Task 2
  - Blocking/warning precheck split: covered by Task 3
  - Confirm-then-recheck loop: covered by Task 4
  - CLI entry `--year --batch`: covered by Task 5
  - Single-month autofill vs combined-batch blocking: covered by Tasks 3 and 4
  - Precheck exception as blocking issue: covered by Task 3
  - Output directory auto-create warning: covered by Task 4
  - Formal engine failure stops pipeline: covered by Task 4
  - Pipeline, precheck, and unmapped logs: covered by Tasks 3 and 4
  - Keep old engines and config flow compatible: covered by Task 4 and docs in Task 5
- Placeholder scan:
  - No placeholder markers, copy-forward shortcuts, or unspecified commands remain
  - All code-changing steps include concrete code blocks
  - All test steps include exact `uv run pytest` commands
- Type and naming consistency:
  - `BatchRef`, `PipelinePaths`, `PipelineDefaults`, `PrecheckResult`, `PipelineIssue`
  - `build_pipeline_paths`, `load_pipeline_defaults`, `run_precheck`, `run_directory_batch`, `run_pipeline`
