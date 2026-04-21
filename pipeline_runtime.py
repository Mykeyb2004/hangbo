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
        precheck = run_precheck(
            paths,
            sheet_name=defaults.sheet_name,
            single_month=single_month,
        )
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
