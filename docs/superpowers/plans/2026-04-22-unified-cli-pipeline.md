# Unified CLI Pipeline Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `main_pipeline.py --year --batch` the only recommended workflow, delete GUI code, and remove per-batch TOML entry points.

**Architecture:** Keep `pipeline_paths.py` as the single path derivation layer and `pipeline.defaults.toml` as the single shared defaults file. Remove files and tests that created or validated per-data-source runtime configs, then update docs to describe only the CLI pipeline.

**Tech Stack:** Python 3.11+, `unittest`, `uv run`, TOML defaults, Markdown docs.

---

## File Structure

- Modify: `tests/test_main_pipeline.py` — make default-config CLI behavior explicit.
- Create: `tests/test_unified_cli_contract.py` — guard against GUI/legacy TOML/doc regressions.
- Modify: `tests/test_generate_ppt.py` — remove dependency on deleted `report_jobs.*.toml`.
- Delete: `hangbo_gui.py` — GUI is fully removed.
- Delete: `tests/test_hangbo_gui.py` — GUI tests are fully removed.
- Delete: `job.toml`, `job01-02.toml`, `job03.toml`, `job_Q1.toml` — old stats configs are removed.
- Delete: `report_jobs.1-2月.toml`, `report_jobs.3月.toml`, `report_jobs.Q1.toml`, `report_jobs.example.toml`, `report_jobs.directory.example.toml` — old batch/report configs are removed.
- Delete: `ppt_job.example.toml` — old standalone PPT config example is removed.
- Delete: `docs/UI适配新数据分析流程计划.md` — obsolete GUI planning doc is removed.
- Modify: `README.md` — remove GUI section and old explicit default-config guidance.
- Modify: `docs/README.md` — keep index focused on CLI docs.
- Modify: `docs/PPT生成说明.md` — describe PPT through main pipeline and direct CLI args, not config examples.
- Modify: `docs/用户运行用例故事.md` — remove GUI/old config references and show default CLI commands.
- Modify: `docs/统计口径与结果说明.md` — remove `job.toml` examples and use direct CLI args for advanced single-script use.
- Modify: `docs/新数据分析流程说明.md`, `docs/数据准备与预查错.md` — ensure main pipeline examples omit explicit `--config pipeline.defaults.toml`.

---

### Task 1: Lock CLI Default Config Contract

**Files:**
- Modify: `tests/test_main_pipeline.py`

- [ ] **Step 1: Write the default-config test first**

Replace the first parse-args test with two tests that make the expected CLI contract explicit:

```python
def test_parse_args_uses_default_config_when_omitted(self) -> None:
    args = parse_args(["--year", "2026", "--batch", "3月"])

    self.assertEqual(args.year, "2026")
    self.assertEqual(args.batch, "3月")
    self.assertEqual(args.config, Path("pipeline.defaults.toml"))

def test_parse_args_keeps_config_override_for_advanced_use(self) -> None:
    args = parse_args(
        ["--year", "2026", "--batch", "3月", "--config", "pipeline.no-llm.toml"]
    )

    self.assertEqual(args.year, "2026")
    self.assertEqual(args.batch, "3月")
    self.assertEqual(args.config, Path("pipeline.no-llm.toml"))
```

- [ ] **Step 2: Run test to verify current behavior**

Run:

```bash
uv run python -m unittest tests/test_main_pipeline.py
```

Expected: PASS. This behavior already exists; the test locks the contract before deletions.

- [ ] **Step 3: Keep production code unchanged**

No production change is needed in `main_pipeline.py`; it already defaults `--config` to `pipeline.defaults.toml`.

- [ ] **Step 4: Re-run main pipeline tests**

Run:

```bash
uv run python -m unittest tests/test_main_pipeline.py
```

Expected: PASS.

---

### Task 2: Add Legacy Removal Guard

**Files:**
- Create: `tests/test_unified_cli_contract.py`

- [ ] **Step 1: Write failing removal tests**

Create `tests/test_unified_cli_contract.py` with:

```python
from __future__ import annotations

import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

LEGACY_ENTRY_FILES = (
    "hangbo_gui.py",
    "tests/test_hangbo_gui.py",
    "job.toml",
    "job01-02.toml",
    "job03.toml",
    "job_Q1.toml",
    "report_jobs.1-2月.toml",
    "report_jobs.3月.toml",
    "report_jobs.Q1.toml",
    "report_jobs.example.toml",
    "report_jobs.directory.example.toml",
    "ppt_job.example.toml",
    "docs/UI适配新数据分析流程计划.md",
)

DOC_PATHS = (
    PROJECT_ROOT / "README.md",
    PROJECT_ROOT / "docs" / "README.md",
    PROJECT_ROOT / "docs" / "PPT生成说明.md",
    PROJECT_ROOT / "docs" / "用户运行用例故事.md",
    PROJECT_ROOT / "docs" / "统计口径与结果说明.md",
    PROJECT_ROOT / "docs" / "新数据分析流程说明.md",
    PROJECT_ROOT / "docs" / "数据准备与预查错.md",
)

STALE_DOC_PATTERNS = (
    "uv run python hangbo_gui.py",
    "--config pipeline.defaults.toml",
    "report_jobs.",
    "ppt_job.example.toml",
    "job.toml",
    "job03.toml",
    "job_Q1.toml",
    "job01-02.toml",
    "GUI 入口",
    "GUI 工作台",
)


class UnifiedCliContractTest(unittest.TestCase):
    def test_legacy_gui_and_batch_config_entry_files_are_removed(self) -> None:
        existing_paths = [
            relative_path
            for relative_path in LEGACY_ENTRY_FILES
            if (PROJECT_ROOT / relative_path).exists()
        ]

        self.assertEqual(existing_paths, [])

    def test_user_docs_do_not_recommend_legacy_entry_points(self) -> None:
        matches: list[str] = []
        for path in DOC_PATHS:
            content = path.read_text(encoding="utf-8")
            for pattern in STALE_DOC_PATTERNS:
                if pattern in content:
                    matches.append(f"{path.relative_to(PROJECT_ROOT)} contains {pattern!r}")

        self.assertEqual(matches, [])

    def test_readme_recommends_default_cli_without_explicit_config(self) -> None:
        content = (PROJECT_ROOT / "README.md").read_text(encoding="utf-8")

        self.assertIn("uv run python main_pipeline.py --year 2026 --batch 3月", content)
        self.assertNotIn(
            "uv run python main_pipeline.py --year 2026 --batch 3月 --config pipeline.defaults.toml",
            content,
        )


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
uv run python -m unittest tests/test_unified_cli_contract.py
```

Expected: FAIL because legacy files and stale doc references still exist.

---

### Task 3: Delete GUI and Legacy TOML Files

**Files:**
- Delete: `hangbo_gui.py`
- Delete: `tests/test_hangbo_gui.py`
- Delete: `job.toml`
- Delete: `job01-02.toml`
- Delete: `job03.toml`
- Delete: `job_Q1.toml`
- Delete: `report_jobs.1-2月.toml`
- Delete: `report_jobs.3月.toml`
- Delete: `report_jobs.Q1.toml`
- Delete: `report_jobs.example.toml`
- Delete: `report_jobs.directory.example.toml`
- Delete: `ppt_job.example.toml`
- Delete: `docs/UI适配新数据分析流程计划.md`

- [ ] **Step 1: Remove files with `apply_patch`**

Delete each listed file using `apply_patch` delete hunks.

- [ ] **Step 2: Run removal guard**

Run:

```bash
uv run python -m unittest tests/test_unified_cli_contract.py
```

Expected: FAIL only on stale documentation patterns. The file-removal test should PASS.

- [ ] **Step 3: Run broad import-sensitive tests**

Run:

```bash
uv run python -m unittest tests/test_main_pipeline.py tests/test_pipeline_runtime.py
```

Expected: PASS because main pipeline does not import GUI or old TOML files.

---

### Task 4: Remove `report_jobs.*` Test Dependency

**Files:**
- Modify: `tests/test_generate_ppt.py`
- Modify: `tests/test_pipeline_config.py` if shared expectations need to move there

- [ ] **Step 1: Run targeted PPT tests after deleting old configs**

Run:

```bash
uv run python -m unittest tests/test_generate_ppt.py
```

Expected: FAIL in `test_project_batch_configs_use_balanced_llm_notes_settings` because `report_jobs.*.toml` files were removed.

- [ ] **Step 2: Replace old-config test with pipeline-defaults test**

In `tests/test_generate_ppt.py`, replace:

```python
def test_project_batch_configs_use_balanced_llm_notes_settings(self) -> None:
    repo_root = Path(__file__).resolve().parents[1]

    for config_name in ("report_jobs.3月.toml", "report_jobs.Q1.toml", "report_jobs.1-2月.toml"):
        config = load_batch_config(repo_root / config_name)

        self.assertEqual(config.llm_notes.target_chars, 120, msg=config_name)
        self.assertEqual(config.llm_notes.temperature, 0.4, msg=config_name)
        self.assertEqual(config.llm_notes.max_tokens, 200, msg=config_name)
```

with a focused test that validates the PPT config generated by the pipeline defaults loader:

```python
def test_project_pipeline_defaults_use_balanced_ppt_llm_notes_settings(self) -> None:
    from pipeline_config import load_pipeline_defaults

    repo_root = Path(__file__).resolve().parents[1]
    defaults = load_pipeline_defaults(repo_root / "pipeline.defaults.toml")

    self.assertEqual(defaults.ppt.llm_notes.target_chars, 120)
    self.assertEqual(defaults.ppt.llm_notes.temperature, 0.4)
    self.assertEqual(defaults.ppt.llm_notes.max_tokens, 200)
```

- [ ] **Step 3: Run targeted PPT tests**

Run:

```bash
uv run python -m unittest tests/test_generate_ppt.py
```

Expected: PASS.

---

### Task 5: Update User Documentation

**Files:**
- Modify: `README.md`
- Modify: `docs/README.md`
- Modify: `docs/PPT生成说明.md`
- Modify: `docs/用户运行用例故事.md`
- Modify: `docs/统计口径与结果说明.md`
- Modify: `docs/新数据分析流程说明.md`
- Modify: `docs/数据准备与预查错.md`

- [ ] **Step 1: Run docs guard to confirm stale docs still fail**

Run:

```bash
uv run python -m unittest tests/test_unified_cli_contract.py
```

Expected: FAIL on `test_user_docs_do_not_recommend_legacy_entry_points`.

- [ ] **Step 2: Update `README.md`**

Remove the `GUI 入口` section. Ensure the main command appears as:

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

Keep `pipeline.defaults.toml` described as the default global settings file, not something users must pass every run.

- [ ] **Step 3: Update `docs/PPT生成说明.md`**

Replace examples that use:

```bash
uv run python generate_ppt.py --config ppt_job.example.toml
```

with direct-argument examples:

```bash
uv run python generate_ppt.py \
  --template-path templates/template.pptx \
  --input-dir data/satisfaction_detail/2026/3月 \
  --output-ppt data/ppt/2026/3月/3月满意度报告.pptx
```

Keep the main recommendation as running PPT through:

```bash
uv run python main_pipeline.py --year 2026 --batch 3月
```

- [ ] **Step 4: Update `docs/用户运行用例故事.md`**

Remove `ppt_job.example.toml` references. Keep advanced no-LLM override as:

```bash
cp pipeline.defaults.toml pipeline.4月.no-llm.toml
uv run python main_pipeline.py --year 2026 --batch 4月 --config pipeline.4月.no-llm.toml
```

Do not show `--config pipeline.defaults.toml` for normal runs.

- [ ] **Step 5: Update `docs/统计口径与结果说明.md`**

Replace:

```bash
uv run python survey_stats.py --config job.toml --calculation-mode summary
```

with the recommended main-pipeline command:

```bash
uv run python main_pipeline.py --year 2026 --batch Q1
```

Remove standalone `survey_stats.py --config` examples from user-facing docs because the repo no longer maintains old batch TOML files.

- [ ] **Step 6: Update remaining docs**

Remove stale mentions in:

- `docs/README.md`
- `docs/新数据分析流程说明.md`
- `docs/数据准备与预查错.md`

Ensure no user-facing docs contain `--config pipeline.defaults.toml`, `report_jobs.`, `ppt_job.example.toml`, `job.toml`, or GUI instructions.

- [ ] **Step 7: Run docs guard**

Run:

```bash
uv run python -m unittest tests/test_unified_cli_contract.py
```

Expected: PASS.

---

### Task 6: Full Relevant Validation

**Files:**
- No file edits unless validation exposes an issue from this change.

- [ ] **Step 1: Run focused pipeline/config tests**

Run:

```bash
uv run python -m unittest \
  tests/test_main_pipeline.py \
  tests/test_pipeline_paths.py \
  tests/test_pipeline_config.py \
  tests/test_pipeline_runtime.py \
  tests/test_generate_ppt.py \
  tests/test_unified_cli_contract.py
```

Expected: PASS.

- [ ] **Step 2: Run repository tests**

Run:

```bash
uv run python -m unittest discover tests
```

Expected: PASS. If unrelated failures appear, record them and do not fix unrelated code.

- [ ] **Step 3: Search for stale references**

Run:

```bash
rg -n "hangbo_gui|uv run python hangbo_gui.py|--config pipeline\\.defaults\\.toml|report_jobs\\.|ppt_job\\.example\\.toml|job03\\.toml|job_Q1\\.toml|job01-02\\.toml|job\\.toml" README.md docs tests *.py *.toml
```

Expected: no stale user-facing references. References inside the approved design or implementation plan are acceptable if the final search includes `docs/superpowers`; otherwise run the search excluding `docs/superpowers`.

---

## Self-Review

- Spec coverage: The plan covers default CLI usage, GUI deletion, old TOML deletion, test cleanup, docs cleanup, and validation.
- Placeholder scan: No `TBD` or unresolved implementation placeholders remain.
- Type consistency: New tests use `unittest`, `Path`, and existing project root conventions consistently.
