# Code Structure Reorganization Design

## Goal

Reorganize the project so the root directory keeps only the two user-facing entry scripts while the reusable business logic moves into a standard `src/hangbo` package. The daily commands should remain unchanged:

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
uv run python merge_sample_summary.py --year 2026
```

## Chosen Approach

Use a staged package reorganization without splitting the largest modules yet. This keeps the refactor focused on file boundaries and imports, avoids changing business behavior, and creates a cleaner base for later targeted decomposition of `survey_stats.py` and `generate_ppt.py`.

## Target Layout

```text
main_pipeline.py
merge_sample_summary.py
src/
  hangbo/
    __init__.py
    pipeline/
      __init__.py
      config.py
      models.py
      paths.py
      runtime.py
    precheck/
      __init__.py
      checks.py
      phase_column.py
      unmapped_customers.py
      year_month.py
    survey/
      __init__.py
      stats.py
      customer_category_rules.py
    summary/
      __init__.py
      table.py
    sample/
      __init__.py
      table.py
    ppt/
      __init__.py
      generator.py
      chart_renderer.py
    merge/
      __init__.py
      questionnaire_workbooks.py
      sample_summary.py
```

## File Mapping

| Current file | Target file |
| --- | --- |
| `pipeline_config.py` | `src/hangbo/pipeline/config.py` |
| `pipeline_models.py` | `src/hangbo/pipeline/models.py` |
| `pipeline_paths.py` | `src/hangbo/pipeline/paths.py` |
| `pipeline_runtime.py` | `src/hangbo/pipeline/runtime.py` |
| `pipeline_precheck.py` | `src/hangbo/precheck/checks.py` |
| `phase_column_preprocess.py` | `src/hangbo/precheck/phase_column.py` |
| `check_unmapped_customer_records.py` | `src/hangbo/precheck/unmapped_customers.py` |
| `fill_year_month_columns.py` | `src/hangbo/precheck/year_month.py` |
| `survey_stats.py` | `src/hangbo/survey/stats.py` |
| `survey_customer_category_rules.py` | `src/hangbo/survey/customer_category_rules.py` |
| `summary_table.py` | `src/hangbo/summary/table.py` |
| `sample_table.py` | `src/hangbo/sample/table.py` |
| `generate_ppt.py` | `src/hangbo/ppt/generator.py` |
| `ppt_chart_renderer.py` | `src/hangbo/ppt/chart_renderer.py` |
| `merge_questionnaire_workbooks.py` | `src/hangbo/merge/questionnaire_workbooks.py` |
| `merge_sample_summary.py` business logic | `src/hangbo/merge/sample_summary.py` |

## Entry Script Policy

`main_pipeline.py` and `merge_sample_summary.py` stay in the project root and keep their command-line parsing. They should become the stable user-facing command layer and import implementation code from `hangbo.*`.

`merge_sample_summary.py` should keep `parse_args()` and `main()` in the root script, while moving the dataclasses, validation, directory selection, source preparation, publishing, and `run_merge_sample_summary()` logic into `src/hangbo/merge/sample_summary.py`.

## Packaging

Configure the project as a standard `src` package in `pyproject.toml` so `uv run python main_pipeline.py` and tests can import `hangbo.*` without modifying `sys.path` inside scripts.

The current root-level module names should not be kept as compatibility wrappers except for the two approved entry scripts. Tests should import from the new package paths, with entry-script tests continuing to import `main_pipeline` and root `merge_sample_summary`.

## Testing Strategy

After the move, run focused tests for the affected import graph:

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

Then run the full test suite if the focused set passes.

## Non-Goals

This refactor should not change data processing behavior, output file formats, PPT rendering, command-line flags, config file names, or the user-facing run commands.

This refactor should not split `survey_stats.py` or `generate_ppt.py` into smaller internal modules yet. Those files can be decomposed in later focused changes once the package boundaries are stable.

## Risks And Mitigations

The main risk is broken imports across tests and scripts. Mitigate this by moving files in small groups, updating imports immediately, and running focused tests after each group.

The second risk is accidentally changing CLI behavior while slimming the root scripts. Mitigate this by preserving the public `parse_args()` and `main()` behavior in both root entry files and keeping existing CLI tests.

The third risk is stale GitNexus references after moving files. Mitigate this by running `npx gitnexus analyze` after the refactor and using `gitnexus detect_changes` to inspect the affected graph.
