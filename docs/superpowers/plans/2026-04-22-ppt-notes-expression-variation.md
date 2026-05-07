# PPT Notes Expression Variation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Improve PPT notes wording by adding configurable highlight gating, banned-term cleanup, and batch-level expression variation.

**Architecture:** Keep LLM analysis generation in `generate_ppt.py`, but add a small batch state object that feeds expression guidance into each prompt and records used openings after each page. Configuration flows from `pipeline.defaults.toml` through `pipeline_config.py` and `pipeline_runtime.py` into `LlmNotesConfig`.

**Tech Stack:** Python dataclasses, `unittest`, TOML defaults, `uv run` test execution.

---

### Task 1: Configurable Highlight Threshold

**Files:**
- Modify: `generate_ppt.py`
- Modify: `pipeline_config.py`
- Modify: `pipeline_runtime.py`
- Modify: `pipeline.defaults.toml`
- Test: `tests/test_generate_ppt.py`
- Test: `tests/test_pipeline_config.py`
- Test: `tests/test_pipeline_runtime.py`

- [ ] Write failing tests asserting `highlight_threshold` defaults to `9.6`, loads from TOML, and flows into `PptBatchConfig.llm_notes`.
- [ ] Run targeted tests with `uv run python -m unittest tests.test_generate_ppt tests.test_pipeline_config tests.test_pipeline_runtime`.
- [ ] Add `highlight_threshold` fields and default parsing.
- [ ] Re-run targeted tests.

### Task 2: Prompt and Post-Processing Rules

**Files:**
- Modify: `generate_ppt.py`
- Modify: `system_role.md`
- Test: `tests/test_generate_ppt.py`

- [ ] Write failing tests for prompt wording, banned term cleanup, and low-score removal of `亮点：`.
- [ ] Run `uv run python -m unittest tests.test_generate_ppt`.
- [ ] Replace “拖累” prompt language with “弱势项/相对偏弱” language.
- [ ] Add final notes normalization that replaces banned terms and removes `亮点：` below threshold.
- [ ] Re-run `uv run python -m unittest tests.test_generate_ppt`.

### Task 3: Batch-Level Expression Variation

**Files:**
- Modify: `generate_ppt.py`
- Test: `tests/test_generate_ppt.py`

- [ ] Write failing tests for expression guidance rotation and used-opening injection across pages.
- [ ] Run `uv run python -m unittest tests.test_generate_ppt`.
- [ ] Add expression style constants and batch state tracking.
- [ ] Pass batch state through `generate_presentation()` into each notes request.
- [ ] Re-run `uv run python -m unittest tests.test_generate_ppt`.

### Task 4: Documentation and Verification

**Files:**
- Modify: `docs/PPT生成说明.md`
- Modify: `docs/用户运行用例故事.md`

- [ ] Document `highlight_threshold` and batch-level expression variation.
- [ ] Run full relevant tests with `uv run python -m unittest tests.test_generate_ppt tests.test_pipeline_config tests.test_pipeline_runtime`.
- [ ] Review diff for unrelated changes.
