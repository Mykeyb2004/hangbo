# Code Structure Reorganization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move reusable business logic from root-level Python files into a standard `src/hangbo` package while preserving the two existing user-facing commands.

**Architecture:** Keep `main_pipeline.py` and `merge_sample_summary.py` at the repository root with their CLI parsing intact. Move the implementation modules into domain packages under `src/hangbo`, then update imports, tests, and GitNexus metadata without changing processing behavior.

**Tech Stack:** Python 3.11+, `uv`, `pytest`, `hatchling` packaging, GitNexus, existing `openpyxl`, `pandas`, `python-pptx`, `matplotlib`, and OpenAI dependencies.

---

## File Structure

- Create `src/hangbo/__init__.py`
  - Package marker for the application.
- Create `src/hangbo/pipeline/`
  - `config.py`: pipeline defaults and TOML config loading.
  - `models.py`: pipeline dataclasses.
  - `paths.py`: batch parsing and path construction.
  - `runtime.py`: main pipeline orchestration.
- Create `src/hangbo/precheck/`
  - `checks.py`: precheck orchestration.
  - `phase_column.py`: phase-marker preprocessing.
  - `unmapped_customers.py`: customer mapping audit.
  - `year_month.py`: year/month column autofill.
- Create `src/hangbo/survey/`
  - `stats.py`: satisfaction report generation.
  - `customer_category_rules.py`: customer category rules.
- Create `src/hangbo/summary/table.py`
  - Summary workbook generation.
- Create `src/hangbo/sample/table.py`
  - Customer-type sample table generation.
- Create `src/hangbo/ppt/`
  - `generator.py`: PPT generation.
  - `chart_renderer.py`: chart image rendering.
- Create `src/hangbo/merge/`
  - `questionnaire_workbooks.py`: raw workbook merging.
  - `sample_summary.py`: merge-and-sample-summary business logic.
- Modify `main_pipeline.py`
  - Keep `parse_args()` and `main()`.
  - Import from `hangbo.pipeline.*`.
- Modify `merge_sample_summary.py`
  - Keep `parse_args()` and `main()`.
  - Import business helpers from `hangbo.merge.sample_summary`.
- Modify `pyproject.toml`
  - Add build backend configuration so `uv run` installs the `src/hangbo` package.
- Modify tests under `tests/`
  - Update imports and mock patch targets to new package paths.
- Modify analysis scripts under `test/`
  - Update imports to new package paths.

Do not keep root-level compatibility wrappers for the moved modules. The only root Python entry scripts after the refactor should be `main_pipeline.py` and `merge_sample_summary.py`.

---

### Task 1: Baseline And Package Scaffold

**Files:**
- Modify: `pyproject.toml`
- Create: `src/hangbo/__init__.py`
- Create: `src/hangbo/pipeline/__init__.py`
- Create: `src/hangbo/precheck/__init__.py`
- Create: `src/hangbo/survey/__init__.py`
- Create: `src/hangbo/summary/__init__.py`
- Create: `src/hangbo/sample/__init__.py`
- Create: `src/hangbo/ppt/__init__.py`
- Create: `src/hangbo/merge/__init__.py`

- [ ] **Step 1: Capture current workspace state**

Run:

```bash
git status --short
```

Expected: note any existing unrelated changes before editing. At the time this plan was written, `AGENTS.md` and `CLAUDE.md` had unrelated GitNexus-generated statistic changes.

- [ ] **Step 2: Run focused baseline tests**

Run:

```bash
PYTHONPATH=. uv run --with pytest pytest \
  tests/test_main_pipeline.py \
  tests/test_merge_sample_summary.py \
  tests/test_pipeline_runtime.py \
  tests/test_pipeline_config.py \
  tests/test_pipeline_paths.py \
  tests/test_pipeline_precheck.py \
  tests/test_survey_stats.py \
  tests/test_generate_ppt.py \
  tests/test_sample_table.py \
  tests/test_summary_table.py
```

Expected: PASS before refactoring. If this fails, stop and resolve the baseline failure before moving files.

- [ ] **Step 3: Add the package scaffold**

Create these empty files:

```text
src/hangbo/__init__.py
src/hangbo/pipeline/__init__.py
src/hangbo/precheck/__init__.py
src/hangbo/survey/__init__.py
src/hangbo/summary/__init__.py
src/hangbo/sample/__init__.py
src/hangbo/ppt/__init__.py
src/hangbo/merge/__init__.py
```

- [ ] **Step 4: Configure package installation in `pyproject.toml`**

Add this section below the existing `[project]` table:

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["src/hangbo"]
```

- [ ] **Step 5: Verify the empty package imports**

Run:

```bash
uv run python -c "import hangbo; print(hangbo.__name__)"
```

Expected output:

```text
hangbo
```

- [ ] **Step 6: Commit the scaffold**

Run:

```bash
git add pyproject.toml src/hangbo
git commit -m "chore: add hangbo package scaffold"
```

---

### Task 2: Move Modules Into Domain Packages

**Files:**
- Move: `pipeline_config.py` -> `src/hangbo/pipeline/config.py`
- Move: `pipeline_models.py` -> `src/hangbo/pipeline/models.py`
- Move: `pipeline_paths.py` -> `src/hangbo/pipeline/paths.py`
- Move: `pipeline_runtime.py` -> `src/hangbo/pipeline/runtime.py`
- Move: `pipeline_precheck.py` -> `src/hangbo/precheck/checks.py`
- Move: `phase_column_preprocess.py` -> `src/hangbo/precheck/phase_column.py`
- Move: `check_unmapped_customer_records.py` -> `src/hangbo/precheck/unmapped_customers.py`
- Move: `fill_year_month_columns.py` -> `src/hangbo/precheck/year_month.py`
- Move: `survey_stats.py` -> `src/hangbo/survey/stats.py`
- Move: `survey_customer_category_rules.py` -> `src/hangbo/survey/customer_category_rules.py`
- Move: `summary_table.py` -> `src/hangbo/summary/table.py`
- Move: `sample_table.py` -> `src/hangbo/sample/table.py`
- Move: `generate_ppt.py` -> `src/hangbo/ppt/generator.py`
- Move: `ppt_chart_renderer.py` -> `src/hangbo/ppt/chart_renderer.py`
- Move: `merge_questionnaire_workbooks.py` -> `src/hangbo/merge/questionnaire_workbooks.py`
- Move: `merge_sample_summary.py` -> `src/hangbo/merge/sample_summary.py`
- Create: `merge_sample_summary.py`

- [ ] **Step 1: Move the implementation modules**

Run:

```bash
git mv pipeline_config.py src/hangbo/pipeline/config.py
git mv pipeline_models.py src/hangbo/pipeline/models.py
git mv pipeline_paths.py src/hangbo/pipeline/paths.py
git mv pipeline_runtime.py src/hangbo/pipeline/runtime.py
git mv pipeline_precheck.py src/hangbo/precheck/checks.py
git mv phase_column_preprocess.py src/hangbo/precheck/phase_column.py
git mv check_unmapped_customer_records.py src/hangbo/precheck/unmapped_customers.py
git mv fill_year_month_columns.py src/hangbo/precheck/year_month.py
git mv survey_stats.py src/hangbo/survey/stats.py
git mv survey_customer_category_rules.py src/hangbo/survey/customer_category_rules.py
git mv summary_table.py src/hangbo/summary/table.py
git mv sample_table.py src/hangbo/sample/table.py
git mv generate_ppt.py src/hangbo/ppt/generator.py
git mv ppt_chart_renderer.py src/hangbo/ppt/chart_renderer.py
git mv merge_questionnaire_workbooks.py src/hangbo/merge/questionnaire_workbooks.py
git mv merge_sample_summary.py src/hangbo/merge/sample_summary.py
```

- [ ] **Step 2: Recreate the root `merge_sample_summary.py` CLI script**

Create `merge_sample_summary.py` with this content:

```python
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
```

- [ ] **Step 3: Remove CLI-only code from `src/hangbo/merge/sample_summary.py`**

In `src/hangbo/merge/sample_summary.py`, remove:

```python
import argparse
```

Remove the complete `parse_args()` function, the complete `main()` function, and the final module entrypoint block `if __name__ == "__main__": main()` from the package module. Those three CLI pieces now live only in the root `merge_sample_summary.py` script.

Also remove the now-unused import:

```python
from pipeline_config import load_pipeline_defaults
```

Keep all dataclasses, validation helpers, selection helpers, source preparation helpers, publishing helpers, and `run_merge_sample_summary()` in `src/hangbo/merge/sample_summary.py`.

- [ ] **Step 4: Update `main_pipeline.py` imports**

Replace the old imports:

```python
from pipeline_config import load_pipeline_defaults
from pipeline_paths import build_pipeline_paths, parse_single_month_batch
from pipeline_runtime import run_pipeline
```

with:

```python
from hangbo.pipeline.config import load_pipeline_defaults
from hangbo.pipeline.paths import build_pipeline_paths, parse_single_month_batch
from hangbo.pipeline.runtime import run_pipeline
```

- [ ] **Step 5: Confirm root Python files are limited to the two entry scripts**

Run:

```bash
find . -maxdepth 1 -type f -name '*.py' -print | sort
```

Expected output:

```text
./main_pipeline.py
./merge_sample_summary.py
```

Do not commit yet; imports are expected to be broken until Task 3 is complete.

---

### Task 3: Update Package Imports

**Files:**
- Modify: `src/hangbo/**/*.py`
- Modify: `main_pipeline.py`
- Modify: `merge_sample_summary.py`
- Modify: `test/*.py`

- [ ] **Step 1: Apply the package import mapping**

Update imports according to this exact mapping:

```text
check_unmapped_customer_records -> hangbo.precheck.unmapped_customers
fill_year_month_columns -> hangbo.precheck.year_month
generate_ppt -> hangbo.ppt.generator
merge_questionnaire_workbooks -> hangbo.merge.questionnaire_workbooks
merge_sample_summary -> hangbo.merge.sample_summary
phase_column_preprocess -> hangbo.precheck.phase_column
pipeline_config -> hangbo.pipeline.config
pipeline_models -> hangbo.pipeline.models
pipeline_paths -> hangbo.pipeline.paths
pipeline_precheck -> hangbo.precheck.checks
pipeline_runtime -> hangbo.pipeline.runtime
ppt_chart_renderer -> hangbo.ppt.chart_renderer
sample_table -> hangbo.sample.table
summary_table -> hangbo.summary.table
survey_customer_category_rules -> hangbo.survey.customer_category_rules
survey_stats -> hangbo.survey.stats
```

The key package imports after this step should include:

```python
# src/hangbo/pipeline/config.py
from hangbo.ppt.generator import normalize_section_mode
from hangbo.survey.stats import normalize_calculation_mode

# src/hangbo/pipeline/paths.py
from hangbo.pipeline.models import BatchRef, PipelinePaths

# src/hangbo/pipeline/runtime.py
from hangbo.precheck.year_month import apply_year_month_to_directory
from hangbo.ppt.generator import (
    CategoryIntroSlideConfig,
    ChartPageConfig,
    LlmNotesConfig,
    PptBatchConfig,
    PptLayoutConfig,
    generate_presentation,
)
from hangbo.precheck.checks import run_precheck
from hangbo.sample.table import generate_sample_table_report
from hangbo.summary.table import generate_summary_report
from hangbo.survey.stats import run_directory_batch

# src/hangbo/precheck/checks.py
from hangbo.precheck.unmapped_customers import (
    format_directory_audit_report,
    run_directory_audit,
    write_audit_log,
)
from hangbo.precheck.phase_column import preprocess_phase_column_if_needed
from hangbo.pipeline.models import PipelineIssue, PipelinePaths, PrecheckResult

# src/hangbo/merge/sample_summary.py
from hangbo.precheck.year_month import apply_year_month_to_directory
from hangbo.merge.questionnaire_workbooks import (
    MergeSummary,
    format_merge_summary,
    merge_workbooks_by_filename,
)
from hangbo.pipeline.paths import parse_single_month_batch
from hangbo.precheck.checks import workbook_has_year_month_headers
from hangbo.sample.table import generate_sample_table_report

# src/hangbo/ppt/generator.py
from hangbo.survey.stats import (
    RoleDefinition,
    TEMPLATE_DEFINITIONS,
    format_value,
    get_effective_role_definition,
)
from hangbo.survey.customer_category_rules import (
    ALL_CUSTOMER_CATEGORY_RULES,
    CUSTOMER_CATEGORY_RULE_BY_NAME,
    DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES,
)
from hangbo.ppt.chart_renderer import ChartPoint, ChartRenderConfig, render_chart_image
from hangbo.summary.table import normalize_text

# src/hangbo/sample/table.py
from hangbo.survey.customer_category_rules import CUSTOMER_CATEGORY_RULE_BY_NAME, CustomerCategoryRule

# src/hangbo/summary/table.py
from hangbo.survey.customer_category_rules import DISPLAY_ORDERED_CUSTOMER_CATEGORY_RULES, CustomerCategoryRule
from hangbo.survey.stats import OVERALL_FILL, SECTION_FILL, excel_round, mean_ignore_empty, normalize_output_dir

# src/hangbo/survey/stats.py
from hangbo.precheck.phase_column import preprocess_phase_column_if_needed
from hangbo.survey.customer_category_rules import (
    CUSTOMER_CATEGORY_RULE_BY_NAME,
    CUSTOMER_CATEGORY_RULES,
    CustomerCategoryRule,
)

# src/hangbo/precheck/unmapped_customers.py
from hangbo.survey.customer_category_rules import CUSTOMER_CATEGORY_RULES
```

- [ ] **Step 2: Update analysis scripts in `test/`**

Replace imports in these files with package paths:

```text
test/analyze_exhibition_customer_satisfaction.py
test/analyze_four_customer_groups_overview.py
test/analyze_group_and_type_dimension_table.py
test/analyze_other_customer_groups.py
test/analyze_dimension_trends_by_group.py
test/build_html_dashboard_data.py
```

Use the same mapping from Step 1. For example:

```python
from sample_table import build_customer_category_rule_mask
```

becomes:

```python
from hangbo.sample.table import build_customer_category_rule_mask
```

- [ ] **Step 3: Verify no old root-module imports remain outside tests that intentionally import entry scripts**

Run:

```bash
rg -n "from (check_unmapped_customer_records|fill_year_month_columns|generate_ppt|merge_questionnaire_workbooks|phase_column_preprocess|pipeline_config|pipeline_models|pipeline_paths|pipeline_precheck|pipeline_runtime|ppt_chart_renderer|sample_table|summary_table|survey_customer_category_rules|survey_stats) import|import (generate_ppt|ppt_chart_renderer|sample_table|summary_table|survey_stats)" src test tests main_pipeline.py merge_sample_summary.py
```

Expected: no output, except no issue if `tests/test_main_pipeline.py` imports `main_pipeline` and if `tests/test_merge_sample_summary.py` imports `merge_sample_summary` only for `parse_args` or root CLI behavior.

- [ ] **Step 4: Run import smoke checks**

Run:

```bash
uv run python - <<'PY'
from hangbo.pipeline.config import load_pipeline_defaults
from hangbo.pipeline.paths import build_pipeline_paths
from hangbo.pipeline.runtime import run_pipeline
from hangbo.merge.sample_summary import run_merge_sample_summary
from hangbo.survey.stats import run_directory_batch
from hangbo.ppt.generator import generate_presentation
print("package imports ok")
PY
```

Expected output:

```text
package imports ok
```

Do not commit yet; tests still need import updates.

---

### Task 4: Update Tests And Mock Patch Targets

**Files:**
- Modify: `tests/test_check_unmapped_customer_records.py`
- Modify: `tests/test_fill_year_month_columns.py`
- Modify: `tests/test_generate_ppt.py`
- Modify: `tests/test_merge_questionnaire_workbooks.py`
- Modify: `tests/test_merge_sample_summary.py`
- Modify: `tests/test_phase_column_preprocess.py`
- Modify: `tests/test_pipeline_config.py`
- Modify: `tests/test_pipeline_paths.py`
- Modify: `tests/test_pipeline_precheck.py`
- Modify: `tests/test_pipeline_runtime.py`
- Modify: `tests/test_ppt_chart_renderer.py`
- Modify: `tests/test_sample_table.py`
- Modify: `tests/test_summary_table.py`
- Modify: `tests/test_survey_customer_category_rules.py`
- Modify: `tests/test_survey_customer_mappings.py`
- Modify: `tests/test_survey_stats.py`

- [ ] **Step 1: Update direct test imports**

Apply these replacements:

```text
from check_unmapped_customer_records import -> from hangbo.precheck.unmapped_customers import
from fill_year_month_columns import -> from hangbo.precheck.year_month import
import generate_ppt as generate_ppt_module -> import hangbo.ppt.generator as generate_ppt_module
from generate_ppt import -> from hangbo.ppt.generator import
from merge_questionnaire_workbooks import -> from hangbo.merge.questionnaire_workbooks import
from merge_sample_summary import -> from hangbo.merge.sample_summary import
from phase_column_preprocess import -> from hangbo.precheck.phase_column import
from pipeline_config import -> from hangbo.pipeline.config import
from pipeline_models import -> from hangbo.pipeline.models import
from pipeline_paths import -> from hangbo.pipeline.paths import
from pipeline_precheck import -> from hangbo.precheck.checks import
from pipeline_runtime import -> from hangbo.pipeline.runtime import
from ppt_chart_renderer import -> from hangbo.ppt.chart_renderer import
from sample_table import -> from hangbo.sample.table import
from summary_table import -> from hangbo.summary.table import
from survey_customer_category_rules import -> from hangbo.survey.customer_category_rules import
from survey_stats import -> from hangbo.survey.stats import
```

Keep this import unchanged in `tests/test_main_pipeline.py`:

```python
from main_pipeline import main, parse_args
```

In `tests/test_merge_sample_summary.py`, import `parse_args` from the root script and import all business helpers from the package:

```python
from merge_sample_summary import parse_args
from hangbo.merge.sample_summary import (
    BatchNameError,
    MergeSampleRunConfig,
    MixedSourceYearMonthError,
    SourcePreparationError,
    build_merge_sample_paths,
    clear_generated_outputs,
    confirm_overwrite_if_needed,
    discover_source_directories,
    iter_source_excel_paths,
    parse_number_selection,
    prompt_batch_name,
    prepare_source_directories,
    run_merge_sample_summary,
    select_directories,
    select_directories_by_number_prompt,
    validate_batch_name,
)
```

- [ ] **Step 2: Update mock patch targets**

Apply these exact patch target replacements:

```text
"pipeline_runtime.generate_presentation" -> "hangbo.pipeline.runtime.generate_presentation"
"pipeline_runtime.generate_sample_table_report" -> "hangbo.pipeline.runtime.generate_sample_table_report"
"pipeline_runtime.generate_summary_report" -> "hangbo.pipeline.runtime.generate_summary_report"
"pipeline_runtime.run_directory_batch" -> "hangbo.pipeline.runtime.run_directory_batch"
"pipeline_runtime.apply_year_month_to_directory" -> "hangbo.pipeline.runtime.apply_year_month_to_directory"
"pipeline_runtime.run_precheck" -> "hangbo.pipeline.runtime.run_precheck"

"pipeline_precheck.run_unmapped_audit" -> "hangbo.precheck.checks.run_unmapped_audit"

"survey_stats.discover_directory_jobs" -> "hangbo.survey.stats.discover_directory_jobs"

"merge_sample_summary.apply_year_month_to_directory" -> "hangbo.merge.sample_summary.apply_year_month_to_directory"
"merge_sample_summary.prepare_source_directories" -> "hangbo.merge.sample_summary.prepare_source_directories"
"merge_sample_summary.merge_workbooks_by_filename" -> "hangbo.merge.sample_summary.merge_workbooks_by_filename"
"merge_sample_summary.generate_sample_table_report" -> "hangbo.merge.sample_summary.generate_sample_table_report"
"merge_sample_summary.select_directories_with_curses" -> "hangbo.merge.sample_summary.select_directories_with_curses"
"merge_sample_summary.select_directories_by_number_prompt" -> "hangbo.merge.sample_summary.select_directories_by_number_prompt"
```

Keep these patch targets unchanged in `tests/test_main_pipeline.py` because they patch the root entry script's imported names:

```text
"main_pipeline.run_pipeline"
"main_pipeline.load_pipeline_defaults"
"main_pipeline.build_pipeline_paths"
```

- [ ] **Step 3: Verify old patch targets are gone**

Run:

```bash
rg -n "patch\\(\"(pipeline_runtime|pipeline_precheck|survey_stats|merge_sample_summary)\\.|mock\\.patch\\(\"(pipeline_runtime|pipeline_precheck|survey_stats|merge_sample_summary)\\." tests
```

Expected: no output. If `merge_sample_summary` appears only for a root CLI-specific patch added later, verify it targets a name imported by the root script.

- [ ] **Step 4: Run the focused test group**

Run:

```bash
PYTHONPATH=. uv run --with pytest pytest \
  tests/test_main_pipeline.py \
  tests/test_merge_sample_summary.py \
  tests/test_pipeline_runtime.py \
  tests/test_pipeline_config.py \
  tests/test_pipeline_paths.py \
  tests/test_pipeline_precheck.py \
  tests/test_survey_stats.py \
  tests/test_generate_ppt.py \
  tests/test_sample_table.py \
  tests/test_summary_table.py
```

Expected: PASS.

- [ ] **Step 5: Commit the move and import updates**

Run:

```bash
git add main_pipeline.py merge_sample_summary.py src test tests pyproject.toml
git commit -m "refactor: move core code into hangbo package"
```

---

### Task 5: Final Verification And GitNexus Refresh

**Files:**
- No planned source edits unless verification finds a refactor bug.

- [ ] **Step 1: Compile root entry scripts, package code, and tests**

Run:

```bash
uv run python -m compileall main_pipeline.py merge_sample_summary.py src tests
```

Expected: command exits 0.

- [ ] **Step 2: Check CLI help still works**

Run:

```bash
uv run python main_pipeline.py --help
uv run python merge_sample_summary.py --help
```

Expected: both commands exit 0 and display their existing help text.

- [ ] **Step 3: Re-run the focused test group**

Run:

```bash
PYTHONPATH=. uv run --with pytest pytest \
  tests/test_main_pipeline.py \
  tests/test_merge_sample_summary.py \
  tests/test_pipeline_runtime.py \
  tests/test_pipeline_config.py \
  tests/test_pipeline_paths.py \
  tests/test_pipeline_precheck.py \
  tests/test_survey_stats.py \
  tests/test_generate_ppt.py \
  tests/test_sample_table.py \
  tests/test_summary_table.py
```

Expected: PASS.

- [ ] **Step 4: Run the full test suite**

Run:

```bash
PYTHONPATH=. uv run --with pytest pytest
```

Expected: PASS. If this fails only because `tests/test_start_time_month_check.py` imports the pre-existing missing module `check_start_time_month`, record it as unrelated to this refactor and keep the focused suite as the refactor gate.

- [ ] **Step 5: Refresh GitNexus**

Run:

```bash
npx gitnexus analyze
```

Expected: repository indexes successfully. Scope extraction or vector-extension warnings may appear, but the command must exit 0.

- [ ] **Step 6: Inspect changed graph impact**

Run GitNexus detect changes for all uncommitted changes:

```text
gitnexus detect_changes(scope="all", repo="hangbo")
```

Expected: changed symbols are the moved modules and entry scripts. Affected processes should match the known pipeline, merge sample summary, survey, summary, sample, and PPT flows.

- [ ] **Step 7: Verify root cleanup**

Run:

```bash
find . -maxdepth 1 -type f -name '*.py' -print | sort
```

Expected output:

```text
./main_pipeline.py
./merge_sample_summary.py
```

- [ ] **Step 8: Review final diff**

Run:

```bash
git status --short
git diff --stat HEAD
```

Expected: only intended refactor changes plus GitNexus-generated doc statistic changes if `npx gitnexus analyze` rewrites them.

- [ ] **Step 9: Commit final verification metadata if needed**

If GitNexus rewrote `AGENTS.md` or `CLAUDE.md`, review the diff. If the only changes are GitNexus statistic updates, commit them separately:

```bash
git add AGENTS.md CLAUDE.md
git commit -m "docs: refresh GitNexus index metadata"
```

If there are no metadata changes, skip this step.
