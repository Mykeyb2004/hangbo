# PPT Table Fonts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce a Q1 observation PPT where numeric tables use 15pt text, Chinese labels use KaiTi, numeric values use Times New Roman, and default table regions are slightly taller.

**Architecture:** Keep all changes inside the existing PPT generation pipeline. Extend `render_table()`/`set_cell_text()` with explicit table font selection and update default PPT layout/font constants in config-loading paths.

**Tech Stack:** Python 3.11, `python-pptx`, `uv run`, `unittest`.

---

### Task 1: Table Font Tests

**Files:**
- Modify: `tests/test_generate_ppt.py`

- [ ] Add tests proving pipeline defaults use 15pt and the slightly taller layout.
- [ ] Add tests proving rendered table cells use `楷体` for Chinese cells and `Times New Roman` for numeric cells.
- [ ] Run: `uv run python -m unittest tests.test_generate_ppt.GeneratePptTest.test_ppt_batch_config_defaults_use_14pt_table_fonts tests.test_generate_ppt.GeneratePptTest.test_render_table_uses_kaiti_for_text_and_times_for_numbers -v`
- [ ] Expected before implementation: failures showing old defaults or missing font names.

### Task 2: Minimal Implementation B

**Files:**
- Modify: `generate_ppt.py`
- Modify: `pipeline.defaults.toml`
- Modify: `pipeline_config.py`

- [ ] Keep table font constants for `楷体` and `Times New Roman`, and update pipeline default table font sizes to `15.0`.
- [ ] Increase default summary/detail table heights without crossing the 7.5in slide boundary.
- [ ] Move default summary/detail table regions upward to restore bottom safety margin after the 15pt size increase.
- [ ] Pass an explicit font name into `set_cell_text()`.
- [ ] Use `apply_run_font_name()` for table cells.
- [ ] Tighten top/bottom table cell margins to `1pt`.

### Task 3: Verification and Q1 Output

**Files:**
- Generate: `data/ppt/2026/Q1/Q1满意度报告-14号表格观察版.pptx`

- [ ] Run focused tests with `uv run python -m unittest tests.test_generate_ppt.GeneratePptTest.test_ppt_batch_config_defaults_use_14pt_table_fonts tests.test_generate_ppt.GeneratePptTest.test_render_table_uses_kaiti_for_text_and_times_for_numbers -v`.
- [ ] Generate Q1 observation PPT using existing `data/satisfaction_detail/2026/Q1` input and chart placeholders/notes as configured for local output.
- [ ] Open the generated PPT with `python-pptx` and verify at least one table page has 14pt text and the expected fonts.
