# Merge Sample Summary Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `merge_sample_summary.py`, an interactive CLI that lets users select raw source folders, merges them, and generates only the customer-type sample summary workbook.

**Architecture:** Add one focused script that owns interaction, validation, and orchestration while reusing existing business engines. Pure helper functions stay separately testable; terminal TUI code is thin and falls back to numbered selection.

**Tech Stack:** Python 3.11+, standard library `argparse`, `curses`, `dataclasses`, `pathlib`; existing `openpyxl`, `fill_year_month_columns.py`, `merge_questionnaire_workbooks.py`, `sample_table.py`, `pipeline_paths.py`, and `pipeline_config.py`.

---

## File Structure

- Create `merge_sample_summary.py`
  - CLI parsing and `main()`.
  - Source folder discovery.
  - Terminal multi-select UI with numbered fallback.
  - Batch name and overwrite validation.
  - Single-month autofill, mixed-source year/month checks, merge orchestration, and sample summary generation.
- Create `tests/test_merge_sample_summary.py`
  - Unit tests for pure helpers.
  - Orchestration tests with mocked engine functions.
  - Light integration-style tests using temporary Excel files for mixed-source year/month checks.
- Modify `README.md`
  - Add the new command as a narrow utility for generating a merged sample summary only.
- Modify `docs/新数据分析流程说明.md`
  - Document the new workflow separately from `main_pipeline.py`.
- Do not modify `main_pipeline.py`, `pipeline_runtime.py`, `survey_stats.py`, `summary_table.py`, or `generate_ppt.py`.

---

### Task 1: Pure Path, Selection, and Validation Helpers

**Files:**
- Create: `merge_sample_summary.py`
- Test: `tests/test_merge_sample_summary.py`

- [ ] **Step 1: Write failing tests for directory discovery, numbered selection parsing, batch name validation, and output paths**

Add this file:

```python
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from merge_sample_summary import (
    BatchNameError,
    build_merge_sample_paths,
    discover_source_directories,
    parse_number_selection,
    validate_batch_name,
)


class MergeSampleSummaryHelpersTest(unittest.TestCase):
    def test_discover_source_directories_lists_only_direct_folders_sorted_by_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            raw_year_dir = Path(temp_dir) / "data" / "raw" / "2026"
            raw_year_dir.mkdir(parents=True)
            (raw_year_dir / "3月").mkdir()
            (raw_year_dir / "1-2月").mkdir()
            (raw_year_dir / "说明.txt").write_text("ignore", encoding="utf-8")
            (raw_year_dir / "Q1").mkdir()
            (raw_year_dir / "Q1" / "nested").mkdir()

            result = discover_source_directories(raw_year_dir)

        self.assertEqual([item.name for item in result], ["1-2月", "3月", "Q1"])

    def test_discover_source_directories_rejects_missing_year_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaisesRegex(FileNotFoundError, "年份原始数据目录不存在"):
                discover_source_directories(Path(temp_dir) / "data" / "raw" / "2026")

    def test_discover_source_directories_rejects_empty_year_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            raw_year_dir = Path(temp_dir) / "data" / "raw" / "2026"
            raw_year_dir.mkdir(parents=True)

            with self.assertRaisesRegex(ValueError, "没有可选择的来源目录"):
                discover_source_directories(raw_year_dir)

    def test_parse_number_selection_supports_commas_and_ranges(self) -> None:
        result = parse_number_selection("1, 3-5, 2", item_count=5)

        self.assertEqual(result, (0, 2, 3, 4, 1))

    def test_parse_number_selection_rejects_out_of_range_and_empty_values(self) -> None:
        with self.assertRaisesRegex(ValueError, "至少选择一个来源目录"):
            parse_number_selection("", item_count=3)
        with self.assertRaisesRegex(ValueError, "超出范围"):
            parse_number_selection("4", item_count=3)
        with self.assertRaisesRegex(ValueError, "范围起点不能大于终点"):
            parse_number_selection("3-1", item_count=3)

    def test_validate_batch_name_rejects_empty_separator_and_source_conflict(self) -> None:
        selected_dirs = (Path("data/raw/2026/1-2月"), Path("data/raw/2026/3月"))

        self.assertEqual(validate_batch_name(" Q1 ", selected_dirs), "Q1")
        for raw_name in ("", "  ", "../Q1", "Q1/backup", "1-2月"):
            with self.assertRaises(BatchNameError):
                validate_batch_name(raw_name, selected_dirs)

    def test_build_merge_sample_paths_uses_existing_directory_contract(self) -> None:
        paths = build_merge_sample_paths(
            year="2026",
            batch_name="Q1",
            data_root=Path("data"),
        )

        self.assertEqual(paths.raw_year_dir, Path("data/raw/2026"))
        self.assertEqual(paths.merged_raw_dir, Path("data/raw/2026/Q1"))
        self.assertEqual(paths.sample_summary_dir, Path("data/sample_summary/2026/Q1"))
        self.assertEqual(
            paths.sample_summary_path,
            Path("data/sample_summary/2026/Q1/Q1客户类型样本统计表.xlsx"),
        )


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests and verify they fail because the module does not exist**

Run:

```bash
uv run python -m unittest tests/test_merge_sample_summary.py
```

Expected: FAIL with `ModuleNotFoundError: No module named 'merge_sample_summary'`.

- [ ] **Step 3: Create `merge_sample_summary.py` with the minimal helper implementation**

Create:

```python
from __future__ import annotations

import argparse
import curses
from dataclasses import dataclass
from pathlib import Path

from fill_year_month_columns import apply_year_month_to_directory
from merge_questionnaire_workbooks import (
    DEFAULT_SHEET_NAME,
    MergeSummary,
    format_merge_summary,
    merge_workbooks_by_filename,
)
from pipeline_config import load_pipeline_defaults
from pipeline_paths import parse_single_month_batch
from pipeline_precheck import workbook_has_year_month_headers
from sample_table import generate_sample_table_report


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
    normalized_year = str(year).strip()
    normalized_batch = str(batch_name).strip()
    raw_year_dir = data_root / "raw" / normalized_year
    merged_raw_dir = raw_year_dir / normalized_batch
    sample_summary_dir = data_root / "sample_summary" / normalized_year / normalized_batch
    return MergeSamplePaths(
        year=normalized_year,
        batch_name=normalized_batch,
        data_root=data_root,
        raw_year_dir=raw_year_dir,
        merged_raw_dir=merged_raw_dir,
        sample_summary_dir=sample_summary_dir,
        sample_summary_path=sample_summary_dir / f"{normalized_batch}客户类型样本统计表.xlsx",
    )


def discover_source_directories(raw_year_dir: Path) -> tuple[Path, ...]:
    if not raw_year_dir.exists() or not raw_year_dir.is_dir():
        raise FileNotFoundError(f"年份原始数据目录不存在：{raw_year_dir}")

    directories = tuple(sorted(
        (path for path in raw_year_dir.iterdir() if path.is_dir()),
        key=lambda item: item.name,
    ))
    if not directories:
        raise ValueError(f"年份目录下没有可选择的来源目录：{raw_year_dir}")
    return directories


def parse_number_selection(raw_value: str, *, item_count: int) -> tuple[int, ...]:
    text = str(raw_value).strip()
    if not text:
        raise ValueError("至少选择一个来源目录。")

    selected_indexes: list[int] = []
    for part in text.split(","):
        token = part.strip()
        if not token:
            continue
        if "-" in token:
            start_text, end_text = (piece.strip() for piece in token.split("-", 1))
            if not start_text.isdigit() or not end_text.isdigit():
                raise ValueError(f"非法范围：{token}")
            start_number = int(start_text)
            end_number = int(end_text)
            if start_number > end_number:
                raise ValueError("范围起点不能大于终点。")
            numbers = range(start_number, end_number + 1)
        else:
            if not token.isdigit():
                raise ValueError(f"非法编号：{token}")
            numbers = (int(token),)

        for number in numbers:
            if number < 1 or number > item_count:
                raise ValueError(f"编号超出范围：{number}")
            selected_indexes.append(number - 1)

    if not selected_indexes:
        raise ValueError("至少选择一个来源目录。")
    return tuple(selected_indexes)


def validate_batch_name(raw_name: str, selected_dirs: tuple[Path, ...]) -> str:
    batch_name = str(raw_name).strip()
    if not batch_name:
        raise BatchNameError("输出批次名不能为空。")
    if Path(batch_name).name != batch_name or "/" in batch_name or "\\" in batch_name:
        raise BatchNameError("输出批次名不能包含路径分隔符。")
    if batch_name in {path.name for path in selected_dirs}:
        raise BatchNameError("输出批次名不能与已选择的来源目录同名。")
    return batch_name
```

- [ ] **Step 4: Run helper tests and verify they pass**

Run:

```bash
uv run python -m unittest tests/test_merge_sample_summary.py
```

Expected: PASS.

- [ ] **Step 5: Commit helper foundation**

```bash
git add merge_sample_summary.py tests/test_merge_sample_summary.py
git commit -m "Add merge sample summary helpers"
```

---

### Task 2: Source Preparation and Mixed Directory Year/Month Checks

**Files:**
- Modify: `merge_sample_summary.py`
- Modify: `tests/test_merge_sample_summary.py`

- [ ] **Step 1: Add failing tests for single-month autofill and mixed-source blocking**

Append these imports to `tests/test_merge_sample_summary.py`:

```python
from unittest import mock
from openpyxl import Workbook

from merge_sample_summary import (
    MixedSourceYearMonthError,
    prepare_source_directories,
)
```

Append this helper and test class:

```python
def write_questionnaire_workbook(
    output_path: Path,
    headers: list[str],
    rows: list[list[object]],
    *,
    sheet_name: str = "问卷数据",
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


class MergeSampleSummaryPreparationTest(unittest.TestCase):
    @mock.patch("merge_sample_summary.apply_year_month_to_directory")
    def test_prepare_source_directories_autofills_only_single_month_dirs(
        self,
        mock_apply_year_month: mock.Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            month_dir = root / "3月"
            mixed_dir = root / "1-2月"
            month_dir.mkdir()
            mixed_dir.mkdir()
            write_questionnaire_workbook(
                mixed_dir / "展览.xlsx",
                ["姓名", "年份", "月份"],
                [["张三", "2026", "1-2"]],
            )

            prepare_source_directories(
                (month_dir, mixed_dir),
                year="2026",
                sheet_name="问卷数据",
            )

        mock_apply_year_month.assert_called_once_with(
            month_dir,
            year="2026",
            month="3",
            sheet_name="问卷数据",
        )

    def test_prepare_source_directories_blocks_mixed_dir_missing_year_month_headers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            mixed_dir = root / "1-2月"
            mixed_dir.mkdir()
            write_questionnaire_workbook(
                mixed_dir / "展览.xlsx",
                ["姓名", "月份"],
                [["张三", "1-2"]],
            )

            with self.assertRaisesRegex(MixedSourceYearMonthError, "缺少“年份”/“月份”列"):
                prepare_source_directories(
                    (mixed_dir,),
                    year="2026",
                    sheet_name="问卷数据",
                )
```

- [ ] **Step 2: Run the targeted tests and verify they fail because functions are missing**

Run:

```bash
uv run python -m unittest tests.test_merge_sample_summary.MergeSampleSummaryPreparationTest
```

Expected: FAIL with import errors for `MixedSourceYearMonthError` and `prepare_source_directories`.

- [ ] **Step 3: Implement source preparation helpers**

Add to `merge_sample_summary.py`:

```python
class MixedSourceYearMonthError(ValueError):
    pass


def iter_source_excel_paths(source_dir: Path) -> tuple[Path, ...]:
    return tuple(sorted(
        path
        for path in source_dir.glob("*.xlsx")
        if path.is_file()
        and not path.name.startswith("~$")
        and not path.name.startswith("._")
    ))


def check_mixed_source_year_month_headers(source_dir: Path, *, sheet_name: str) -> None:
    missing_paths: list[Path] = []
    for workbook_path in iter_source_excel_paths(source_dir):
        if not workbook_has_year_month_headers(workbook_path, sheet_name):
            missing_paths.append(workbook_path)

    if missing_paths:
        first_path = missing_paths[0]
        raise MixedSourceYearMonthError(
            f"混合来源目录缺少“年份”/“月份”列：{source_dir}，首个问题文件：{first_path.name}"
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
            apply_year_month_to_directory(
                source_dir,
                year=str(year),
                month=str(single_month),
                sheet_name=sheet_name,
            )
        else:
            check_mixed_source_year_month_headers(source_dir, sheet_name=sheet_name)
```

- [ ] **Step 4: Run preparation tests and verify they pass**

Run:

```bash
uv run python -m unittest tests.test_merge_sample_summary.MergeSampleSummaryPreparationTest
```

Expected: PASS.

- [ ] **Step 5: Run the full new test file**

Run:

```bash
uv run python -m unittest tests/test_merge_sample_summary.py
```

Expected: PASS.

- [ ] **Step 6: Commit source preparation**

```bash
git add merge_sample_summary.py tests/test_merge_sample_summary.py
git commit -m "Prepare merge sample source directories"
```

---

### Task 3: Merge and Sample Summary Orchestration

**Files:**
- Modify: `merge_sample_summary.py`
- Modify: `tests/test_merge_sample_summary.py`

- [ ] **Step 1: Add failing orchestration tests**

Append imports:

```python
from merge_questionnaire_workbooks import MergeResult, MergeSummary

from merge_sample_summary import (
    MergeSampleRunConfig,
    run_merge_sample_summary,
)
```

Append tests:

```python
class MergeSampleSummaryRunTest(unittest.TestCase):
    @mock.patch("merge_sample_summary.generate_sample_table_report")
    @mock.patch("merge_sample_summary.merge_workbooks_by_filename")
    @mock.patch("merge_sample_summary.prepare_source_directories")
    def test_run_merge_sample_summary_merges_and_generates_only_sample_summary(
        self,
        mock_prepare: mock.Mock,
        mock_merge: mock.Mock,
        mock_generate_sample: mock.Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_dir = root / "data" / "raw" / "2026" / "3月"
            source_dir.mkdir(parents=True)
            output_dir = root / "data" / "raw" / "2026" / "Q1"
            sample_output = root / "data" / "sample_summary" / "2026" / "Q1" / "Q1客户类型样本统计表.xlsx"
            mock_merge.return_value = MergeSummary(
                input_dirs=(source_dir,),
                output_dir=output_dir,
                results=(
                    MergeResult(
                        file_name="展览.xlsx",
                        source_paths=(source_dir / "展览.xlsx",),
                        status="merged",
                        merged_rows=1,
                        output_path=output_dir / "展览.xlsx",
                    ),
                ),
            )
            mock_generate_sample.return_value = sample_output

            config = MergeSampleRunConfig(
                year="2026",
                batch_name="Q1",
                selected_dirs=(source_dir,),
                data_root=root / "data",
                sheet_name="问卷数据",
                sample_config_path=Path("sample_table.default.toml"),
                overwrite=True,
            )

            result = run_merge_sample_summary(config)

        self.assertEqual(result.sample_summary_path, sample_output)
        mock_prepare.assert_called_once_with(
            (source_dir,),
            year="2026",
            sheet_name="问卷数据",
        )
        mock_merge.assert_called_once_with(
            (source_dir,),
            output_dir=output_dir,
            sheet_name="问卷数据",
        )
        mock_generate_sample.assert_called_once_with(
            input_dir=output_dir,
            output_dir=sample_output.parent,
            output_name=sample_output.name,
            config_path=Path("sample_table.default.toml"),
            source_sheet_name="问卷数据",
            default_year="2026",
        )

    @mock.patch("merge_sample_summary.generate_sample_table_report")
    @mock.patch("merge_sample_summary.merge_workbooks_by_filename")
    @mock.patch("merge_sample_summary.prepare_source_directories")
    def test_run_merge_sample_summary_stops_when_any_merge_result_failed(
        self,
        mock_prepare: mock.Mock,
        mock_merge: mock.Mock,
        mock_generate_sample: mock.Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_dir = root / "data" / "raw" / "2026" / "1-2月"
            source_dir.mkdir(parents=True)
            output_dir = root / "data" / "raw" / "2026" / "Q1"
            mock_merge.return_value = MergeSummary(
                input_dirs=(source_dir,),
                output_dir=output_dir,
                results=(
                    MergeResult(
                        file_name="展览.xlsx",
                        source_paths=(source_dir / "展览.xlsx",),
                        status="missing_sheet",
                    ),
                ),
            )

            config = MergeSampleRunConfig(
                year="2026",
                batch_name="Q1",
                selected_dirs=(source_dir,),
                data_root=root / "data",
                sheet_name="问卷数据",
                sample_config_path=Path("sample_table.default.toml"),
                overwrite=True,
            )

            with self.assertRaisesRegex(RuntimeError, "合并阶段存在失败项"):
                run_merge_sample_summary(config)

        mock_generate_sample.assert_not_called()
```

- [ ] **Step 2: Run orchestration tests and verify they fail because run types are missing**

Run:

```bash
uv run python -m unittest tests.test_merge_sample_summary.MergeSampleSummaryRunTest
```

Expected: FAIL with import errors for `MergeSampleRunConfig` and `run_merge_sample_summary`.

- [ ] **Step 3: Implement orchestration dataclasses and runner**

Add to `merge_sample_summary.py`:

```python
@dataclass(frozen=True)
class MergeSampleRunConfig:
    year: str
    batch_name: str
    selected_dirs: tuple[Path, ...]
    data_root: Path
    sheet_name: str
    sample_config_path: Path
    overwrite: bool = False


@dataclass(frozen=True)
class MergeSampleRunResult:
    paths: MergeSamplePaths
    merge_summary: MergeSummary
    sample_summary_path: Path


def merge_summary_has_failures(summary: MergeSummary) -> bool:
    return any(result.status != "merged" for result in summary.results)


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
            "合并阶段存在失败项，未生成样本统计表。\n"
            + format_merge_summary(merge_summary, sheet_name=config.sheet_name)
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
```

- [ ] **Step 4: Run orchestration tests and verify they pass**

Run:

```bash
uv run python -m unittest tests.test_merge_sample_summary.MergeSampleSummaryRunTest
```

Expected: PASS.

- [ ] **Step 5: Run the full new test file**

Run:

```bash
uv run python -m unittest tests/test_merge_sample_summary.py
```

Expected: PASS.

- [ ] **Step 6: Commit orchestration**

```bash
git add merge_sample_summary.py tests/test_merge_sample_summary.py
git commit -m "Run merged sample summary generation"
```

---

### Task 4: Interactive CLI and Fallback Selection

**Files:**
- Modify: `merge_sample_summary.py`
- Modify: `tests/test_merge_sample_summary.py`

- [ ] **Step 1: Add tests for fallback selection and overwrite confirmation**

Append imports:

```python
from merge_sample_summary import (
    confirm_overwrite_if_needed,
    select_directories_by_number_prompt,
)
```

Append tests:

```python
class MergeSampleSummaryInteractionTest(unittest.TestCase):
    def test_select_directories_by_number_prompt_reprompts_until_valid_selection(self) -> None:
        source_dirs = (Path("1-2月"), Path("3月"), Path("4月"))
        prompts = iter(["bad", "1,3"])
        outputs: list[str] = []

        result = select_directories_by_number_prompt(
            source_dirs,
            input_func=lambda prompt: next(prompts),
            output_func=outputs.append,
        )

        self.assertEqual(result, (Path("1-2月"), Path("4月")))
        self.assertTrue(any("非法编号" in line for line in outputs))

    def test_confirm_overwrite_if_needed_returns_true_when_outputs_do_not_exist(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=Path(temp_dir) / "data",
            )

            self.assertTrue(
                confirm_overwrite_if_needed(
                    paths,
                    input_func=lambda prompt: "n",
                    output_func=lambda message: None,
                )
            )

    def test_confirm_overwrite_if_needed_respects_user_answer_when_outputs_exist(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = build_merge_sample_paths(
                year="2026",
                batch_name="Q1",
                data_root=Path(temp_dir) / "data",
            )
            paths.merged_raw_dir.mkdir(parents=True)

            self.assertFalse(
                confirm_overwrite_if_needed(
                    paths,
                    input_func=lambda prompt: "n",
                    output_func=lambda message: None,
                )
            )
            self.assertTrue(
                confirm_overwrite_if_needed(
                    paths,
                    input_func=lambda prompt: "yes",
                    output_func=lambda message: None,
                )
            )
```

- [ ] **Step 2: Run interaction tests and verify they fail because functions are missing**

Run:

```bash
uv run python -m unittest tests.test_merge_sample_summary.MergeSampleSummaryInteractionTest
```

Expected: FAIL with import errors.

- [ ] **Step 3: Implement fallback selection, overwrite confirmation, TUI selection, and CLI `main()`**

Append to `merge_sample_summary.py`:

```python
def select_directories_by_number_prompt(
    source_dirs: tuple[Path, ...],
    *,
    input_func=input,
    output_func=print,
) -> tuple[Path, ...]:
    while True:
        output_func("请选择要合并的来源目录：")
        for index, source_dir in enumerate(source_dirs, start=1):
            output_func(f"{index}. {source_dir.name}")
        raw_value = input_func("输入编号，支持逗号/范围，例如 1,2 或 1-3：")
        try:
            indexes = parse_number_selection(raw_value, item_count=len(source_dirs))
        except ValueError as exc:
            output_func(str(exc))
            continue
        return tuple(source_dirs[index] for index in indexes)


def select_directories_with_curses(source_dirs: tuple[Path, ...]) -> tuple[Path, ...]:
    def run(screen) -> tuple[Path, ...] | None:
        curses.curs_set(0)
        current_index = 0
        selected_indexes: set[int] = set()
        while True:
            screen.clear()
            screen.addstr(0, 0, "请选择要合并的来源目录（↑/↓移动，空格选择，Enter确认，q退出）")
            for index, source_dir in enumerate(source_dirs):
                marker = "[x]" if index in selected_indexes else "[ ]"
                line = f"{marker} {source_dir.name}"
                attr = curses.A_REVERSE if index == current_index else curses.A_NORMAL
                screen.addstr(index + 2, 0, line, attr)
            screen.refresh()

            key = screen.getch()
            if key in (curses.KEY_UP, ord("k")):
                current_index = max(0, current_index - 1)
            elif key in (curses.KEY_DOWN, ord("j")):
                current_index = min(len(source_dirs) - 1, current_index + 1)
            elif key == ord(" "):
                if current_index in selected_indexes:
                    selected_indexes.remove(current_index)
                else:
                    selected_indexes.add(current_index)
            elif key in (10, 13, curses.KEY_ENTER):
                if selected_indexes:
                    return tuple(source_dirs[index] for index in sorted(selected_indexes))
                screen.addstr(len(source_dirs) + 3, 0, "至少选择一个来源目录。")
                screen.refresh()
                screen.getch()
            elif key in (27, ord("q")):
                return None

    result = curses.wrapper(run)
    if result is None:
        raise SystemExit("用户取消。")
    return result


def select_directories(
    source_dirs: tuple[Path, ...],
    *,
    output_func=print,
) -> tuple[Path, ...]:
    try:
        return select_directories_with_curses(source_dirs)
    except curses.error:
        output_func("当前终端不支持交互选择，已切换为编号输入。")
        return select_directories_by_number_prompt(source_dirs, output_func=output_func)


def prompt_batch_name(
    selected_dirs: tuple[Path, ...],
    *,
    input_func=input,
    output_func=print,
) -> str:
    while True:
        raw_name = input_func("请输入合并后的批次名，例如 Q1：")
        try:
            return validate_batch_name(raw_name, selected_dirs)
        except BatchNameError as exc:
            output_func(str(exc))


def confirm_overwrite_if_needed(
    paths: MergeSamplePaths,
    *,
    input_func=input,
    output_func=print,
) -> bool:
    existing_targets = []
    if paths.merged_raw_dir.exists():
        existing_targets.append(paths.merged_raw_dir)
    if paths.sample_summary_path.exists():
        existing_targets.append(paths.sample_summary_path)
    if not existing_targets:
        return True

    output_func("检测到输出已存在：")
    for target in existing_targets:
        output_func(f"- {target}")
    answer = input_func("是否覆盖本工具生成的合并原始 Excel 和样本统计表？输入 y/yes 确认：").strip().lower()
    return answer in {"y", "yes"}


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="交互式合并原始数据，并只生成客户类型样本统计汇总表。")
    parser.add_argument("--year", required=True, help="年份目录，例如 2026")
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
    raw_year_dir = Path("data") / "raw" / str(args.year).strip()
    source_dirs = discover_source_directories(raw_year_dir)
    selected_dirs = select_directories(source_dirs)
    batch_name = prompt_batch_name(selected_dirs)
    paths = build_merge_sample_paths(year=str(args.year), batch_name=batch_name)
    if not confirm_overwrite_if_needed(paths):
        raise SystemExit("用户取消覆盖，未写入新结果。")

    result = run_merge_sample_summary(
        MergeSampleRunConfig(
            year=str(args.year).strip(),
            batch_name=batch_name,
            selected_dirs=selected_dirs,
            data_root=Path("data"),
            sheet_name=defaults.sheet_name,
            sample_config_path=defaults.sample_config_path,
            overwrite=True,
        )
    )
    print(format_merge_summary(result.merge_summary, sheet_name=defaults.sheet_name))
    print(f"样本统计表已保存到: {result.sample_summary_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run interaction tests and verify they pass**

Run:

```bash
uv run python -m unittest tests.test_merge_sample_summary.MergeSampleSummaryInteractionTest
```

Expected: PASS.

- [ ] **Step 5: Run all merge sample tests**

Run:

```bash
uv run python -m unittest tests/test_merge_sample_summary.py
```

Expected: PASS.

- [ ] **Step 6: Commit CLI interaction**

```bash
git add merge_sample_summary.py tests/test_merge_sample_summary.py
git commit -m "Add interactive merge sample summary CLI"
```

---

### Task 5: Documentation

**Files:**
- Modify: `README.md`
- Modify: `docs/新数据分析流程说明.md`

- [ ] **Step 1: Update README with the narrow utility**

Add this section after the existing “主流程怎么跑” section:

```markdown
## 只生成合并后的分月样本统计表

如果只需要把多个 `data/raw/{year}` 下的来源目录合并，并生成客户类型样本统计汇总表，可以使用独立工具：

```bash
uv run python merge_sample_summary.py --year 2026
```

运行后，终端会列出 `data/raw/2026` 下的所有直接子文件夹。使用方向键移动、空格选择、Enter 确认，然后输入输出批次名，例如 `Q1`。

该工具只会生成：

- 合并原始数据：`data/raw/{year}/{batch}`
- 样本统计表：`data/sample_summary/{year}/{batch}/{batch}客户类型样本统计表.xlsx`

它不会生成满意度分项统计、满意度汇总表或 PPT。
```

- [ ] **Step 2: Update detailed workflow docs**

Add this section near the multi-month merge explanation in `docs/新数据分析流程说明.md`:

```markdown
### 5.5 只生成合并后的分月样本统计表

如果当前任务只需要样本统计表，不需要满意度分项、满意度汇总或 PPT，可以运行：

```bash
uv run python merge_sample_summary.py --year 2026
```

脚本会列出 `data/raw/2026` 下所有直接子文件夹，支持在终端中用方向键移动、空格多选、Enter 确认。确认后输入输出批次名，例如 `Q1`。

处理规则：

- `1月` 到 `12月` 这类单月来源目录会在合并前自动补写 `年份` / `月份`。
- `1-2月` 这类混合来源目录可以选择，但必须已经包含正确的 `年份` / `月份` 列。
- 合并原始数据保存到 `data/raw/{year}/{batch}`。
- 样本统计表保存到 `data/sample_summary/{year}/{batch}/{batch}客户类型样本统计表.xlsx`。
- 不生成满意度分项统计、满意度汇总表或 PPT。
```

- [ ] **Step 3: Run docs grep to ensure the new command appears**

Run:

```bash
rg -n "merge_sample_summary.py|只生成合并后的分月样本统计表" README.md docs/新数据分析流程说明.md
```

Expected: command appears in both files.

- [ ] **Step 4: Commit docs**

```bash
git add README.md docs/新数据分析流程说明.md
git commit -m "Document merge sample summary utility"
```

---

### Task 6: Final Verification

**Files:**
- No new files unless verification uncovers defects.

- [ ] **Step 1: Run focused tests**

Run:

```bash
uv run python -m unittest tests/test_merge_sample_summary.py
```

Expected: PASS.

- [ ] **Step 2: Run related existing tests**

Run:

```bash
uv run python -m unittest \
  tests/test_merge_questionnaire_workbooks.py \
  tests/test_fill_year_month_columns.py \
  tests/test_sample_table.py \
  tests/test_main_pipeline.py
```

Expected: PASS.

- [ ] **Step 3: Run CLI help**

Run:

```bash
uv run python merge_sample_summary.py --help
```

Expected: help text includes `--year` and describes the interactive merged sample summary utility.

- [ ] **Step 4: Check git status**

Run:

```bash
git status --short
```

Expected: clean working tree after task commits, or only intentional uncommitted changes if the user requested no commits during execution.

---

## Self-Review

Spec coverage:

- Interactive folder selection with cursor and Space is covered in Task 4.
- Numbered fallback is covered in Tasks 1 and 4.
- Output batch path and sample summary path are covered in Task 1.
- Single-month autofill and mixed-source blocking are covered in Task 2.
- Merge-only-plus-sample-generation orchestration is covered in Task 3.
- Avoiding satisfaction/PPT generation is covered by Task 3 tests and file structure constraints.
- Documentation is covered in Task 5.
- Verification is covered in Task 6.

Placeholder scan:

- No red-flag placeholder terms or unspecified “add tests” steps remain.

Type consistency:

- `MergeSamplePaths`, `MergeSampleRunConfig`, and `MergeSampleRunResult` are introduced before use in later tasks.
- Helper names used by tests match implementation names in the corresponding tasks.
