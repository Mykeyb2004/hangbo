from __future__ import annotations

import queue
import shlex
import subprocess
import sys
import threading
import tomllib
from dataclasses import dataclass, replace
from enum import Enum
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext
import tkinter as tk
from tkinter import ttk

from survey_customer_mappings import STANDARD_CUSTOMER_TYPE_MAPPINGS
from survey_stats import (
    BatchConfig,
    DIRECTORY_NOTICE_REASON_MISSING_ROLE_DATA,
    DIRECTORY_NOTICE_REASON_MISSING_SOURCE_FILE,
    discover_directory_jobs,
)

PROJECT_ROOT = Path(__file__).resolve().parent
LOG_DIR = PROJECT_ROOT / "logs"
GUI_RUNTIME_DIR = LOG_DIR / "gui_runtime"
GUI_PROFILE_DIR = LOG_DIR / "gui_profiles"
GUI_BATCH_DIR = GUI_PROFILE_DIR / "batches"
GUI_SESSION_PATH = GUI_PROFILE_DIR / "last_session.toml"
DEFAULT_SHEET_NAME = "问卷数据"


class WorkflowMode(str, Enum):
    SINGLE = "single"
    MERGED = "merged"


class CustomerTypePreviewStatus(str, Enum):
    READY = "ready"
    MISSING_SOURCE_FILE = "missing_source_file"
    MISSING_ROLE_DATA = "missing_role_data"
    MISSING_INPUT_DIR = "missing_input_dir"


class WorkflowRunStatus(str, Enum):
    IDLE = "idle"
    RUNNING = "running"
    CANCELLING = "cancelling"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass(frozen=True)
class GuiBatchConfig:
    batch_name: str = "2026年3月"
    workflow_mode: WorkflowMode = WorkflowMode.SINGLE
    single_input_dir: Path = PROJECT_ROOT / "datas" / "3月"
    merge_input_dirs: tuple[Path, ...] = ()
    merge_output_dir: Path = PROJECT_ROOT / "datas" / "合并结果"
    sheet_name: str = DEFAULT_SHEET_NAME
    year_value: str = ""
    month_value: str = ""
    stats_output_dir: Path = PROJECT_ROOT / "输出结果" / "3月"
    calculation_mode: str = "template"
    output_format: str = "xlsx"
    summary_output_dir: Path = PROJECT_ROOT / "汇总结果" / "3月"
    summary_output_name: str = "3月客户类型满意度汇总表.xlsx"
    ppt_template_path: Path = PROJECT_ROOT / "templates" / "template.pptx"
    output_ppt_path: Path = PROJECT_ROOT / "输出结果" / "3月满意度报告.pptx"
    ppt_section_mode: str = "auto"
    ppt_blank_display: str = ""

    def effective_input_dir(self) -> Path:
        if self.workflow_mode == WorkflowMode.MERGED:
            return self.merge_output_dir
        return self.single_input_dir


DEFAULT_GUI_BATCH_CONFIG = GuiBatchConfig()


@dataclass(frozen=True)
class MainWorkflowSelection:
    include_merge: bool = False
    include_phase_preprocess: bool = True
    include_fill_year_month: bool = False
    include_survey_stats: bool = True
    include_summary: bool = True
    include_ppt: bool = True


@dataclass(frozen=True)
class TaskCommand:
    key: str
    title: str
    command: tuple[str, ...]


@dataclass(frozen=True)
class CustomerTypePreviewRow:
    template_name: str
    customer_type_name: str
    document_display_name: str | None
    source_file_name: str
    output_name: str
    status: CustomerTypePreviewStatus
    detail: str
    note: str | None = None


@dataclass(frozen=True)
class StatsPreviewSummary:
    rows: tuple[CustomerTypePreviewRow, ...]
    input_dir: Path

    @property
    def ready_count(self) -> int:
        return sum(1 for row in self.rows if row.status == CustomerTypePreviewStatus.READY)

    @property
    def missing_source_count(self) -> int:
        return sum(
            1 for row in self.rows if row.status == CustomerTypePreviewStatus.MISSING_SOURCE_FILE
        )

    @property
    def missing_role_count(self) -> int:
        return sum(
            1 for row in self.rows if row.status == CustomerTypePreviewStatus.MISSING_ROLE_DATA
        )


@dataclass(frozen=True)
class SavedBatchProfile:
    batch_name: str
    path: Path
    config: GuiBatchConfig


def default_selected_customer_types(summary: StatsPreviewSummary) -> frozenset[str]:
    return frozenset(
        row.customer_type_name
        for row in summary.rows
        if row.status == CustomerTypePreviewStatus.READY
    )


def page_key_for_task(task_key: str) -> str:
    return {
        "merge_workbooks": "data_source",
        "phase_preprocess": "preprocess",
        "fill_year_month": "preprocess",
        "survey_stats": "stats",
        "summary_table": "summary",
        "generate_ppt": "ppt",
    }.get(task_key, "dashboard")


def customer_type_preview_status_text(status: CustomerTypePreviewStatus) -> str:
    return {
        CustomerTypePreviewStatus.READY: "可生成",
        CustomerTypePreviewStatus.MISSING_SOURCE_FILE: "缺少来源文件",
        CustomerTypePreviewStatus.MISSING_ROLE_DATA: "未匹配身份值",
        CustomerTypePreviewStatus.MISSING_INPUT_DIR: "输入目录不存在",
    }[status]


def ordered_selected_customer_types(
    rows: tuple[CustomerTypePreviewRow, ...],
    selected_customer_types: frozenset[str],
) -> tuple[str, ...]:
    return tuple(
        row.customer_type_name
        for row in rows
        if row.customer_type_name in selected_customer_types
    )


def build_stats_preview_summary_text(
    summary: StatsPreviewSummary,
    selected_customer_types: frozenset[str],
) -> str:
    selected_count = len(ordered_selected_customer_types(summary.rows, selected_customer_types))
    return (
        f"输入目录：{summary.input_dir}；"
        f"已选 {selected_count} 个；"
        f"可生成 {summary.ready_count} 个；"
        f"缺少来源 {summary.missing_source_count} 个；"
        f"缺少身份值 {summary.missing_role_count} 个"
    )


def build_workflow_status_text(
    controller: WorkflowRunController,
    task_title_lookup: dict[str, str],
) -> str:
    total = len(controller.planned_step_keys)
    completed = len(controller.completed_step_keys)
    active_key = controller.active_step_key
    active_title = task_title_lookup.get(active_key or "", active_key or "等待启动")
    failed_key = controller.failed_step_key
    failed_title = task_title_lookup.get(failed_key or "", failed_key or "未知步骤")

    if controller.status == WorkflowRunStatus.IDLE:
        return "未开始"
    if controller.status == WorkflowRunStatus.RUNNING:
        return f"执行中 {completed + 1}/{max(total, 1)}：{active_title}"
    if controller.status == WorkflowRunStatus.CANCELLING:
        return f"正在终止：{active_title}"
    if controller.status == WorkflowRunStatus.SUCCEEDED:
        return "执行完成"
    if controller.status == WorkflowRunStatus.CANCELLED:
        return "已取消"
    return f"执行失败：{failed_title}"


def build_gui_batch_config_text(
    config: GuiBatchConfig,
    *,
    active_saved_batch_name: str | None = None,
) -> str:
    lines = [
        f"batch_name = {toml_quote(config.batch_name)}",
        f'workflow_mode = "{config.workflow_mode.value}"',
        f"single_input_dir = {toml_quote(config.single_input_dir)}",
        f"merge_output_dir = {toml_quote(config.merge_output_dir)}",
        f"sheet_name = {toml_quote(config.sheet_name)}",
        f"year_value = {toml_quote(config.year_value)}",
        f"month_value = {toml_quote(config.month_value)}",
        f"stats_output_dir = {toml_quote(config.stats_output_dir)}",
        f'calculation_mode = "{config.calculation_mode}"',
        f'output_format = "{config.output_format}"',
        f"summary_output_dir = {toml_quote(config.summary_output_dir)}",
        f"summary_output_name = {toml_quote(config.summary_output_name)}",
        f"ppt_template_path = {toml_quote(config.ppt_template_path)}",
        f"output_ppt_path = {toml_quote(config.output_ppt_path)}",
        f'ppt_section_mode = "{config.ppt_section_mode}"',
        f"ppt_blank_display = {toml_quote(config.ppt_blank_display)}",
    ]
    if active_saved_batch_name is not None:
        lines.append(f"active_saved_batch_name = {toml_quote(active_saved_batch_name)}")
    if config.merge_input_dirs:
        lines.append("")
        lines.append("merge_input_dirs = [")
        for path in config.merge_input_dirs:
            lines.append(f"  {toml_quote(path)},")
        lines.append("]")
    return "\n".join(lines) + "\n"


def _string_value(raw_data: dict[str, object], key: str, default: str) -> str:
    value = raw_data.get(key)
    if value is None:
        return default
    return str(value)


def _path_value(raw_data: dict[str, object], key: str, default: Path) -> Path:
    value = raw_data.get(key)
    if value is None:
        return default
    return Path(str(value)).resolve()


def parse_gui_batch_config(raw_data: dict[str, object]) -> GuiBatchConfig:
    raw_merge_input_dirs = raw_data.get("merge_input_dirs", [])
    if raw_merge_input_dirs is None:
        raw_merge_input_dirs = []
    if not isinstance(raw_merge_input_dirs, list):
        raise ValueError("merge_input_dirs 必须是列表。")
    merge_input_dirs = tuple(Path(str(value)).resolve() for value in raw_merge_input_dirs)

    raw_workflow_mode = _string_value(
        raw_data,
        "workflow_mode",
        DEFAULT_GUI_BATCH_CONFIG.workflow_mode.value,
    ).strip() or DEFAULT_GUI_BATCH_CONFIG.workflow_mode.value

    return GuiBatchConfig(
        batch_name=_string_value(raw_data, "batch_name", DEFAULT_GUI_BATCH_CONFIG.batch_name).strip()
        or DEFAULT_GUI_BATCH_CONFIG.batch_name,
        workflow_mode=WorkflowMode(raw_workflow_mode),
        single_input_dir=_path_value(
            raw_data,
            "single_input_dir",
            DEFAULT_GUI_BATCH_CONFIG.single_input_dir,
        ),
        merge_input_dirs=merge_input_dirs,
        merge_output_dir=_path_value(
            raw_data,
            "merge_output_dir",
            DEFAULT_GUI_BATCH_CONFIG.merge_output_dir,
        ),
        sheet_name=_string_value(raw_data, "sheet_name", DEFAULT_GUI_BATCH_CONFIG.sheet_name)
        or DEFAULT_GUI_BATCH_CONFIG.sheet_name,
        year_value=_string_value(raw_data, "year_value", DEFAULT_GUI_BATCH_CONFIG.year_value),
        month_value=_string_value(raw_data, "month_value", DEFAULT_GUI_BATCH_CONFIG.month_value),
        stats_output_dir=_path_value(
            raw_data,
            "stats_output_dir",
            DEFAULT_GUI_BATCH_CONFIG.stats_output_dir,
        ),
        calculation_mode=_string_value(
            raw_data,
            "calculation_mode",
            DEFAULT_GUI_BATCH_CONFIG.calculation_mode,
        )
        or DEFAULT_GUI_BATCH_CONFIG.calculation_mode,
        output_format=_string_value(raw_data, "output_format", DEFAULT_GUI_BATCH_CONFIG.output_format)
        or DEFAULT_GUI_BATCH_CONFIG.output_format,
        summary_output_dir=_path_value(
            raw_data,
            "summary_output_dir",
            DEFAULT_GUI_BATCH_CONFIG.summary_output_dir,
        ),
        summary_output_name=_string_value(
            raw_data,
            "summary_output_name",
            DEFAULT_GUI_BATCH_CONFIG.summary_output_name,
        )
        or DEFAULT_GUI_BATCH_CONFIG.summary_output_name,
        ppt_template_path=_path_value(
            raw_data,
            "ppt_template_path",
            DEFAULT_GUI_BATCH_CONFIG.ppt_template_path,
        ),
        output_ppt_path=_path_value(
            raw_data,
            "output_ppt_path",
            DEFAULT_GUI_BATCH_CONFIG.output_ppt_path,
        ),
        ppt_section_mode=_string_value(
            raw_data,
            "ppt_section_mode",
            DEFAULT_GUI_BATCH_CONFIG.ppt_section_mode,
        )
        or DEFAULT_GUI_BATCH_CONFIG.ppt_section_mode,
        ppt_blank_display=_string_value(
            raw_data,
            "ppt_blank_display",
            DEFAULT_GUI_BATCH_CONFIG.ppt_blank_display,
        ),
    )


def load_gui_batch_config(profile_path: Path) -> GuiBatchConfig:
    raw_data = tomllib.loads(profile_path.read_text(encoding="utf-8"))
    return parse_gui_batch_config(raw_data)


def load_gui_session(session_path: Path = GUI_SESSION_PATH) -> tuple[GuiBatchConfig, str | None]:
    raw_data = tomllib.loads(session_path.read_text(encoding="utf-8"))
    config = parse_gui_batch_config(raw_data)
    active_saved_batch_name = str(raw_data.get("active_saved_batch_name", "")).strip() or None
    return config, active_saved_batch_name


def batch_profile_storage_path(
    batch_name: str,
    batch_dir: Path = GUI_BATCH_DIR,
) -> Path:
    encoded_name = batch_name.encode("utf-8").hex()
    return batch_dir / f"batch_{encoded_name}.toml"


def save_batch_profile(
    config: GuiBatchConfig,
    *,
    batch_dir: Path = GUI_BATCH_DIR,
) -> Path:
    batch_dir.mkdir(parents=True, exist_ok=True)
    profile_path = batch_profile_storage_path(config.batch_name, batch_dir)
    profile_path.write_text(build_gui_batch_config_text(config), encoding="utf-8")
    return profile_path


def load_saved_batch_profiles(
    batch_dir: Path = GUI_BATCH_DIR,
) -> tuple[SavedBatchProfile, ...]:
    if not batch_dir.exists():
        return ()

    profiles: list[SavedBatchProfile] = []
    for profile_path in sorted(batch_dir.glob("batch_*.toml")):
        if not profile_path.is_file():
            continue
        try:
            config = load_gui_batch_config(profile_path)
        except Exception:
            continue
        profiles.append(
            SavedBatchProfile(
                batch_name=config.batch_name,
                path=profile_path,
                config=config,
            )
        )
    return tuple(sorted(profiles, key=lambda profile: profile.batch_name))


def delete_batch_profile(
    batch_name: str,
    *,
    batch_dir: Path = GUI_BATCH_DIR,
) -> bool:
    profile_path = batch_profile_storage_path(batch_name, batch_dir)
    if not profile_path.exists():
        return False
    profile_path.unlink()
    return True


def save_gui_session(
    config: GuiBatchConfig,
    *,
    active_saved_batch_name: str | None,
    session_path: Path = GUI_SESSION_PATH,
) -> Path:
    session_path.parent.mkdir(parents=True, exist_ok=True)
    session_path.write_text(
        build_gui_batch_config_text(
            config,
            active_saved_batch_name=active_saved_batch_name or "",
        ),
        encoding="utf-8",
    )
    return session_path


class WorkflowRunController:
    def __init__(self) -> None:
        self.status = WorkflowRunStatus.IDLE
        self.planned_step_keys: tuple[str, ...] = ()
        self.active_step_key: str | None = None
        self.completed_step_keys: tuple[str, ...] = ()
        self.failed_step_key: str | None = None
        self.cancel_requested = False

    @property
    def start_enabled(self) -> bool:
        return self.status in {
            WorkflowRunStatus.IDLE,
            WorkflowRunStatus.SUCCEEDED,
            WorkflowRunStatus.FAILED,
            WorkflowRunStatus.CANCELLED,
        }

    @property
    def terminate_enabled(self) -> bool:
        return self.status == WorkflowRunStatus.RUNNING

    def begin(self, step_keys: tuple[str, ...]) -> None:
        self.status = WorkflowRunStatus.RUNNING
        self.planned_step_keys = step_keys
        self.active_step_key = None
        self.completed_step_keys = ()
        self.failed_step_key = None
        self.cancel_requested = False

    def mark_task_started(self, step_key: str) -> None:
        self.active_step_key = step_key
        if self.status != WorkflowRunStatus.CANCELLING:
            self.status = WorkflowRunStatus.RUNNING

    def mark_task_finished(self, step_key: str, success: bool) -> None:
        if success:
            if step_key not in self.completed_step_keys:
                self.completed_step_keys = (*self.completed_step_keys, step_key)
        else:
            self.failed_step_key = step_key
        self.active_step_key = None

    def request_cancel(self) -> None:
        if self.status in {WorkflowRunStatus.RUNNING, WorkflowRunStatus.CANCELLING}:
            self.cancel_requested = True
            self.status = WorkflowRunStatus.CANCELLING

    def finish_run(self, success: bool) -> None:
        self.active_step_key = None
        if self.cancel_requested and not success:
            self.status = WorkflowRunStatus.CANCELLED
            return
        if success:
            self.status = WorkflowRunStatus.SUCCEEDED
            return
        self.status = WorkflowRunStatus.FAILED


@dataclass(frozen=True)
class ThemePalette:
    background: str = "#f6f1ea"
    surface: str = "#fffaf6"
    surface_alt: str = "#f1e3d4"
    border: str = "#d7c2ad"
    primary: str = "#8b3c2c"
    primary_hover: str = "#6f2d21"
    accent: str = "#b8644c"
    text: str = "#2d241f"
    muted_text: str = "#74675d"
    success: str = "#2f7d4a"
    warning: str = "#b06a1c"


def build_directory_batch_config(config: GuiBatchConfig) -> BatchConfig:
    return BatchConfig(
        config_path=(PROJECT_ROOT / "hangbo_gui.py").resolve(),
        output_dir=config.stats_output_dir,
        output_format=config.output_format,
        calculation_mode=config.calculation_mode,
        sheet_name=config.sheet_name,
        input_dir=config.effective_input_dir(),
    )


def build_stats_preview_summary(config: GuiBatchConfig) -> StatsPreviewSummary:
    input_dir = config.effective_input_dir()
    if not input_dir.exists() or not input_dir.is_dir():
        return StatsPreviewSummary(
            rows=tuple(
                CustomerTypePreviewRow(
                    template_name=mapping.template_name,
                    customer_type_name=mapping.template_role_name,
                    document_display_name=mapping.document_display_name,
                    source_file_name=mapping.source_file_name,
                    output_name=f"{mapping.template_role_name}.xlsx",
                    status=CustomerTypePreviewStatus.MISSING_INPUT_DIR,
                    detail=f"输入目录不存在：{input_dir}",
                    note=mapping.note,
                )
                for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS
            ),
            input_dir=input_dir,
        )

    discovery = discover_directory_jobs(build_directory_batch_config(config))
    job_lookup = {job.name: job for job in discovery.jobs}
    missing_notice_lookup = {
        notice.customer_type_name: notice for notice in discovery.missing_customer_type_notices
    }
    rows: list[CustomerTypePreviewRow] = []
    for mapping in STANDARD_CUSTOMER_TYPE_MAPPINGS:
        job = job_lookup.get(mapping.template_role_name)
        if job is not None:
            rows.append(
                CustomerTypePreviewRow(
                    template_name=mapping.template_name,
                    customer_type_name=mapping.template_role_name,
                    document_display_name=mapping.document_display_name,
                    source_file_name=mapping.source_file_name,
                    output_name=f"{job.output_name}.{config.output_format}",
                    status=CustomerTypePreviewStatus.READY,
                    detail=f"来源文件：{job.path.name}",
                    note=mapping.note,
                )
            )
            continue

        notice = missing_notice_lookup.get(mapping.template_role_name)
        if notice is not None and notice.reason == DIRECTORY_NOTICE_REASON_MISSING_ROLE_DATA:
            status = CustomerTypePreviewStatus.MISSING_ROLE_DATA
            detail = f"来源文件存在，但未找到身份值：{notice.source_reference}"
        else:
            status = CustomerTypePreviewStatus.MISSING_SOURCE_FILE
            source_reference = notice.source_reference if notice is not None else mapping.source_file_name
            detail = f"缺少来源文件：{source_reference}"
        rows.append(
            CustomerTypePreviewRow(
                template_name=mapping.template_name,
                customer_type_name=mapping.template_role_name,
                document_display_name=mapping.document_display_name,
                source_file_name=mapping.source_file_name,
                output_name=f"{mapping.template_role_name}.{config.output_format}",
                status=status,
                detail=detail,
                note=mapping.note,
            )
        )

    return StatsPreviewSummary(rows=tuple(rows), input_dir=input_dir)


def build_main_workflow_step_keys(
    config: GuiBatchConfig,
    selection: MainWorkflowSelection,
) -> tuple[str, ...]:
    step_keys: list[str] = []
    if config.workflow_mode == WorkflowMode.MERGED and selection.include_merge:
        step_keys.append("merge_workbooks")
    if selection.include_phase_preprocess:
        step_keys.append("phase_preprocess")
    if selection.include_fill_year_month:
        step_keys.append("fill_year_month")
    if selection.include_survey_stats:
        step_keys.append("survey_stats")
    if selection.include_summary:
        step_keys.append("summary_table")
    if selection.include_ppt:
        step_keys.append("generate_ppt")
    return tuple(step_keys)


def toml_quote(value: str | Path) -> str:
    text = str(value)
    escaped = text.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def build_survey_stats_config_text(config: GuiBatchConfig) -> str:
    return "\n".join(
        [
            f"output_dir = {toml_quote(config.stats_output_dir)}",
            f'output_format = "{config.output_format}"',
            f'calculation_mode = "{config.calculation_mode}"',
            f"input_dir = {toml_quote(config.effective_input_dir())}",
            f"sheet_name = {toml_quote(config.sheet_name)}",
            "",
        ]
    )


def build_ppt_config_text(config: GuiBatchConfig) -> str:
    return "\n".join(
        [
            f"template_path = {toml_quote(config.ppt_template_path)}",
            f"input_dir = {toml_quote(config.stats_output_dir)}",
            f"output_ppt = {toml_quote(config.output_ppt_path)}",
            'file_pattern = "*.xlsx"',
            'sheet_name_mode = "first"',
            f'section_mode = "{config.ppt_section_mode}"',
            f'blank_display = "{config.ppt_blank_display}"',
            "sort_files = true",
            "",
        ]
    )


def write_runtime_config(prefix: str, content: str) -> Path:
    GUI_RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    safe_prefix = prefix.replace(" ", "_").replace("/", "_")
    existing_count = len(list(GUI_RUNTIME_DIR.glob(f"{safe_prefix}_*.toml")))
    path = GUI_RUNTIME_DIR / f"{safe_prefix}_{existing_count + 1:03d}.toml"
    path.write_text(content, encoding="utf-8")
    return path


def discover_excel_files(input_dir: Path) -> tuple[Path, ...]:
    if not input_dir.exists() or not input_dir.is_dir():
        return ()
    return tuple(
        sorted(
            path
            for path in input_dir.glob("*.xlsx")
            if path.is_file()
            and not path.name.startswith("~$")
            and not path.name.startswith("._")
        )
    )


def build_python_command(script_name: str, *args: str) -> list[str]:
    return [sys.executable, str(PROJECT_ROOT / script_name), *args]


def build_merge_command(config: GuiBatchConfig) -> list[str] | None:
    if config.workflow_mode != WorkflowMode.MERGED:
        return None
    if not config.merge_input_dirs:
        raise ValueError("多月合并模式下，至少需要一个输入目录。")
    command = build_python_command("merge_questionnaire_workbooks.py")
    for input_dir in config.merge_input_dirs:
        command.extend(["--input-dir", str(input_dir)])
    command.extend(["--output-dir", str(config.merge_output_dir), "--sheet-name", config.sheet_name])
    return command


def build_phase_preprocess_command(config: GuiBatchConfig) -> list[str]:
    excel_files = discover_excel_files(config.effective_input_dir())
    if not excel_files:
        raise ValueError(f"未在 {config.effective_input_dir()} 下找到可处理的 xlsx 文件。")
    command = build_python_command("phase_column_preprocess.py", *(str(path) for path in excel_files))
    command.extend(["--sheet-name", config.sheet_name])
    return command


def build_fill_year_month_command(config: GuiBatchConfig) -> list[str]:
    if not config.year_value.strip():
        raise ValueError("请先填写年份。")
    if not config.month_value.strip():
        raise ValueError("请先填写月份。")
    return build_python_command(
        "fill_year_month_columns.py",
        "--input-dir",
        str(config.effective_input_dir()),
        "--year",
        config.year_value.strip(),
        "--month",
        config.month_value.strip(),
        "--sheet-name",
        config.sheet_name,
    )


def build_survey_stats_command(
    config: GuiBatchConfig,
    *,
    dry_run: bool = False,
    selected_job_names: tuple[str, ...] = (),
) -> list[str]:
    config_path = write_runtime_config("survey_stats", build_survey_stats_config_text(config))
    command = build_python_command("survey_stats.py", "--config", str(config_path))
    for job_name in selected_job_names:
        command.extend(["--job", job_name])
    if dry_run:
        command.append("--dry-run")
    return command


def build_summary_command(config: GuiBatchConfig) -> list[str]:
    return build_python_command(
        "summary_table.py",
        "--input-dir",
        str(config.stats_output_dir),
        "--output-dir",
        str(config.summary_output_dir),
        "--output-name",
        config.summary_output_name,
    )


def build_ppt_command(config: GuiBatchConfig, *, dry_run: bool = False) -> list[str]:
    config_path = write_runtime_config("generate_ppt", build_ppt_config_text(config))
    command = build_python_command("generate_ppt.py", "--config", str(config_path))
    if dry_run:
        command.append("--dry-run")
    return command


def build_task_commands(
    config: GuiBatchConfig,
    selection: MainWorkflowSelection,
    *,
    selected_stats_job_names: tuple[str, ...] = (),
) -> tuple[TaskCommand, ...]:
    commands: list[TaskCommand] = []
    for step_key in build_main_workflow_step_keys(config, selection):
        if step_key == "merge_workbooks":
            merge_command = build_merge_command(config)
            if merge_command is not None:
                commands.append(TaskCommand(step_key, "合并多月问卷", tuple(merge_command)))
        elif step_key == "phase_preprocess":
            commands.append(
                TaskCommand(step_key, "修正期次列", tuple(build_phase_preprocess_command(config)))
            )
        elif step_key == "fill_year_month":
            commands.append(
                TaskCommand(step_key, "补写年份月份", tuple(build_fill_year_month_command(config)))
            )
        elif step_key == "survey_stats":
            commands.append(
                TaskCommand(
                    step_key,
                    "生成分项统计",
                    tuple(
                        build_survey_stats_command(
                            config,
                            selected_job_names=selected_stats_job_names,
                        )
                    ),
                )
            )
        elif step_key == "summary_table":
            commands.append(
                TaskCommand(step_key, "生成汇总表", tuple(build_summary_command(config)))
            )
        elif step_key == "generate_ppt":
            commands.append(
                TaskCommand(step_key, "生成PPT", tuple(build_ppt_command(config)))
            )
    return tuple(commands)


class BackgroundTaskRunner:
    def __init__(self, root: tk.Misc, app: "SurveyPlatformApp") -> None:
        self.root = root
        self.app = app
        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._process: subprocess.Popen[str] | None = None
        self.root.after(120, self._poll)

    @property
    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def run_tasks(self, tasks: tuple[TaskCommand, ...]) -> None:
        if self.is_running:
            raise RuntimeError("当前已有任务正在执行。")
        self._thread = threading.Thread(target=self._run_worker, args=(tasks,), daemon=True)
        self._thread.start()

    def terminate(self) -> None:
        if self._process is not None and self._process.poll() is None:
            self._process.terminate()

    def _run_worker(self, tasks: tuple[TaskCommand, ...]) -> None:
        success = True
        for task in tasks:
            self._queue.put(("task_started", task))
            try:
                self._process = subprocess.Popen(
                    list(task.command),
                    cwd=PROJECT_ROOT,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )
            except OSError as exc:
                self._queue.put(("log", f"[ERROR] 无法启动任务 {task.title}：{exc}"))
                self._queue.put(("task_finished", (task, False)))
                success = False
                break

            assert self._process.stdout is not None
            for line in self._process.stdout:
                self._queue.put(("log", line.rstrip("\n")))
            return_code = self._process.wait()
            if return_code != 0:
                self._queue.put(("log", f"[ERROR] 任务失败：{task.title}（退出码 {return_code}）"))
                self._queue.put(("task_finished", (task, False)))
                success = False
                break
            self._queue.put(("task_finished", (task, True)))
        self._process = None
        self._queue.put(("all_done", success))

    def _poll(self) -> None:
        while True:
            try:
                event, payload = self._queue.get_nowait()
            except queue.Empty:
                break
            self._dispatch(event, payload)
        self.root.after(120, self._poll)

    def _dispatch(self, event: str, payload: object) -> None:
        if event == "log":
            self.app.append_log(str(payload))
            return
        if event == "task_started":
            task = payload
            assert isinstance(task, TaskCommand)
            self.app.on_task_started(task)
            return
        if event == "task_finished":
            task, success = payload
            assert isinstance(task, TaskCommand)
            assert isinstance(success, bool)
            self.app.on_task_finished(task, success)
            return
        if event == "all_done":
            assert isinstance(payload, bool)
            self.app.on_all_tasks_finished(payload)


class SurveyPlatformApp(tk.Tk):
    PAGE_TITLES = {
        "dashboard": "工作台总览",
        "data_source": "数据源管理",
        "preprocess": "预处理",
        "stats": "分项统计",
        "summary": "汇总统计",
        "ppt": "PPT生成",
        "logs": "任务日志",
    }

    def __init__(self) -> None:
        super().__init__()
        self.title("杭博问卷统计平台")
        self.geometry("1520x920")
        self.minsize(1320, 820)
        self.protocol("WM_DELETE_WINDOW", self.on_app_close)
        self.palette = ThemePalette()
        self.workflow_controller = WorkflowRunController()
        self.active_saved_batch_name: str | None = None
        self._saved_batch_profiles: tuple[SavedBatchProfile, ...] = ()
        self._start_action_buttons: list[ttk.Button] = []
        self._terminate_action_buttons: list[ttk.Button] = []
        self._create_variables()
        self._step_status_vars = {
            "data_source": tk.StringVar(value="待设置"),
            "preprocess": tk.StringVar(value="待执行"),
            "stats": tk.StringVar(value="待执行"),
            "summary": tk.StringVar(value="待执行"),
            "ppt": tk.StringVar(value="待执行"),
        }
        self._task_status_label_var = tk.StringVar(value="未开始")
        self._next_action_var = tk.StringVar(value="请先确认数据源")
        self._raw_file_count_var = tk.StringVar(value="0")
        self._stats_file_count_var = tk.StringVar(value="0")
        self._summary_file_var = tk.StringVar(value="未生成")
        self._ppt_file_var = tk.StringVar(value="未生成")
        self._current_page_title_var = tk.StringVar(value=self.PAGE_TITLES["dashboard"])
        self.stats_preview_summary_var = tk.StringVar(value="未扫描客群")
        self._task_title_lookup: dict[str, str] = {}
        self.stats_preview_summary: StatsPreviewSummary | None = None
        self.stats_preview_rows: tuple[CustomerTypePreviewRow, ...] = ()
        self.selected_customer_types: frozenset[str] = frozenset()
        self._workflow_progress_dialog: tk.Toplevel | None = None
        self._workflow_progress_tree: ttk.Treeview | None = None
        self._workflow_progress_summary_var = tk.StringVar(value="未启动主流程")
        self._workflow_progress_close_var = tk.StringVar(value="关闭")
        self._workflow_progress_item_ids: dict[str, str] = {}
        self._workflow_progress_last_page_key: str | None = None
        self._apply_theme()
        self._build_layout()
        self._restore_persistent_state()
        self.runner = BackgroundTaskRunner(self, self)
        self.refresh_all_status_views()
        self._sync_action_button_states()
        self.show_page("dashboard")

    def _create_variables(self) -> None:
        self.batch_name_var = tk.StringVar(value="2026年3月")
        self.workflow_mode_var = tk.StringVar(value=WorkflowMode.SINGLE.value)
        self.single_input_dir_var = tk.StringVar(value=str(PROJECT_ROOT / "datas" / "3月"))
        self.merge_output_dir_var = tk.StringVar(value=str(PROJECT_ROOT / "datas" / "合并结果"))
        self.sheet_name_var = tk.StringVar(value=DEFAULT_SHEET_NAME)
        self.year_value_var = tk.StringVar(value="2026")
        self.month_value_var = tk.StringVar(value="03")
        self.stats_output_dir_var = tk.StringVar(value=str(PROJECT_ROOT / "输出结果" / "3月"))
        self.calculation_mode_var = tk.StringVar(value="template")
        self.output_format_var = tk.StringVar(value="xlsx")
        self.summary_output_dir_var = tk.StringVar(value=str(PROJECT_ROOT / "汇总结果" / "3月"))
        self.summary_output_name_var = tk.StringVar(value="3月客户类型满意度汇总表.xlsx")
        self.ppt_template_path_var = tk.StringVar(value=str(PROJECT_ROOT / "templates" / "template.pptx"))
        self.output_ppt_path_var = tk.StringVar(value=str(PROJECT_ROOT / "输出结果" / "3月满意度报告.pptx"))
        self.ppt_section_mode_var = tk.StringVar(value="auto")
        self.ppt_blank_display_var = tk.StringVar(value="")
        self.saved_batch_var = tk.StringVar(value="")
        self.merge_input_list: list[str] = []

    def _apply_theme(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        self.configure(bg=self.palette.background)
        style.configure(".", background=self.palette.background, foreground=self.palette.text)
        style.configure("Root.TFrame", background=self.palette.background)
        style.configure("Surface.TFrame", background=self.palette.surface)
        style.configure("Card.TFrame", background=self.palette.surface, relief="flat")
        style.configure(
            "Sidebar.TFrame",
            background=self.palette.surface_alt,
            bordercolor=self.palette.border,
            relief="solid",
        )
        style.configure("Header.TLabel", background=self.palette.background, foreground=self.palette.text, font=("PingFang SC", 18, "bold"))
        style.configure("SubHeader.TLabel", background=self.palette.surface, foreground=self.palette.text, font=("PingFang SC", 12, "bold"))
        style.configure("Body.TLabel", background=self.palette.surface, foreground=self.palette.text, font=("PingFang SC", 10))
        style.configure("Muted.TLabel", background=self.palette.surface, foreground=self.palette.muted_text, font=("PingFang SC", 10))
        style.configure("Sidebar.TButton", background=self.palette.surface_alt, foreground=self.palette.text, padding=(16, 10), anchor="w", relief="flat")
        style.map("Sidebar.TButton", background=[("active", self.palette.surface), ("pressed", self.palette.surface)])
        style.configure("Primary.TButton", background=self.palette.primary, foreground="#ffffff", padding=(14, 8), relief="flat")
        style.map(
            "Primary.TButton",
            background=[("active", self.palette.primary_hover), ("pressed", self.palette.primary_hover)],
            foreground=[("disabled", "#f5e8e3")],
        )
        style.configure("Secondary.TButton", background=self.palette.surface, foreground=self.palette.primary, padding=(12, 8), relief="flat", borderwidth=1)
        style.map("Secondary.TButton", background=[("active", self.palette.surface_alt)])
        style.configure("Treeview", background="#fffdfb", fieldbackground="#fffdfb", rowheight=28, foreground=self.palette.text, bordercolor=self.palette.border)
        style.configure("Treeview.Heading", background=self.palette.surface_alt, foreground=self.palette.text, font=("PingFang SC", 10, "bold"))
        style.configure("TEntry", fieldbackground="#fffdfb", bordercolor=self.palette.border)
        style.configure("TCombobox", fieldbackground="#fffdfb")
        style.configure("TLabelframe", background=self.palette.surface, foreground=self.palette.text)
        style.configure("TLabelframe.Label", background=self.palette.surface, foreground=self.palette.text, font=("PingFang SC", 10, "bold"))

    def _build_layout(self) -> None:
        root = ttk.Frame(self, style="Root.TFrame", padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(1, weight=1)

        header = ttk.Frame(root, style="Root.TFrame")
        header.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        header.columnconfigure(1, weight=1)
        header.columnconfigure(2, weight=0)
        ttk.Label(header, text="杭博问卷统计平台", style="Header.TLabel").grid(row=0, column=0, sticky="w")

        batch_manager = ttk.Frame(header, style="Root.TFrame")
        batch_manager.grid(row=0, column=1, sticky="ew", padx=(20, 12))
        batch_manager.columnconfigure(1, weight=1)
        batch_manager.columnconfigure(3, weight=1)
        ttk.Label(batch_manager, text="批次名称", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(batch_manager, textvariable=self.batch_name_var, width=18).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(8, 12),
        )
        ttk.Label(batch_manager, text="已保存批次", style="Body.TLabel").grid(row=0, column=2, sticky="w")
        self.saved_batch_combobox = ttk.Combobox(
            batch_manager,
            textvariable=self.saved_batch_var,
            state="readonly",
            values=(),
            width=22,
        )
        self.saved_batch_combobox.grid(row=0, column=3, sticky="ew", padx=(8, 8))
        ttk.Button(batch_manager, text="加载", style="Secondary.TButton", command=self.load_selected_batch).grid(row=0, column=4, padx=(0, 8))
        ttk.Button(batch_manager, text="新建", style="Secondary.TButton", command=self.create_new_batch).grid(row=0, column=5)

        header_meta = ttk.Frame(header, style="Root.TFrame")
        header_meta.grid(row=0, column=2, sticky="e")
        ttk.Button(header_meta, text="保存批次", style="Secondary.TButton", command=self.save_current_batch).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(header_meta, text="删除批次", style="Secondary.TButton", command=self.delete_selected_batch).grid(row=0, column=1, padx=(0, 8))
        terminate_button = ttk.Button(header_meta, text="终止当前任务", style="Secondary.TButton", command=self.terminate_current_task)
        terminate_button.grid(row=0, column=2, padx=(0, 8))
        self._register_terminate_button(terminate_button)
        main_workflow_button = ttk.Button(header_meta, text="一键执行主流程", style="Primary.TButton", command=self.open_main_workflow_dialog)
        main_workflow_button.grid(row=0, column=3)
        self._register_start_button(main_workflow_button)

        sidebar = ttk.Frame(root, style="Sidebar.TFrame", padding=12)
        sidebar.grid(row=1, column=0, sticky="nsw")

        content_wrap = ttk.Frame(root, style="Root.TFrame")
        content_wrap.grid(row=1, column=1, sticky="nsew", padx=(12, 12))
        content_wrap.columnconfigure(0, weight=1)
        content_wrap.rowconfigure(1, weight=1)

        log_wrap = ttk.Frame(root, style="Surface.TFrame", padding=12)
        log_wrap.grid(row=1, column=2, sticky="nsew")
        root.columnconfigure(2, weight=0)

        page_header = ttk.Frame(content_wrap, style="Root.TFrame")
        page_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        page_header.columnconfigure(0, weight=1)
        ttk.Label(page_header, textvariable=self._current_page_title_var, style="SubHeader.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(page_header, textvariable=self._task_status_label_var, style="Muted.TLabel").grid(row=0, column=1, sticky="e")

        self.page_container = ttk.Frame(content_wrap, style="Surface.TFrame", padding=12)
        self.page_container.grid(row=1, column=0, sticky="nsew")
        self.page_container.columnconfigure(0, weight=1)
        self.page_container.rowconfigure(0, weight=1)

        self._build_sidebar(sidebar)
        self._build_log_panel(log_wrap)
        self._build_pages()

    def _build_sidebar(self, parent: ttk.Frame) -> None:
        nav_items = (
            ("dashboard", "1. 工作台总览"),
            ("data_source", "2. 数据源管理"),
            ("preprocess", "3. 预处理"),
            ("stats", "4. 分项统计"),
            ("summary", "5. 汇总统计"),
            ("ppt", "6. PPT生成"),
            ("logs", "7. 任务日志"),
        )
        for index, (page_key, label) in enumerate(nav_items):
            ttk.Button(
                parent,
                text=label,
                style="Sidebar.TButton",
                command=lambda value=page_key: self.show_page(value),
            ).grid(row=index, column=0, sticky="ew", pady=4)
        parent.columnconfigure(0, weight=1)

    def _build_log_panel(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        ttk.Label(parent, text="运行日志", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.log_text = scrolledtext.ScrolledText(
            parent,
            wrap="word",
            width=42,
            font=("SF Mono", 10),
            bg="#fffdfb",
            fg=self.palette.text,
            relief="flat",
            padx=10,
            pady=10,
        )
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")
        ttk.Button(parent, text="清空日志", style="Secondary.TButton", command=self.clear_log).grid(row=2, column=0, sticky="e", pady=(8, 0))

    def _build_pages(self) -> None:
        self.pages: dict[str, ttk.Frame] = {}
        builders = {
            "dashboard": self._build_dashboard_page,
            "data_source": self._build_data_source_page,
            "preprocess": self._build_preprocess_page,
            "stats": self._build_stats_page,
            "summary": self._build_summary_page,
            "ppt": self._build_ppt_page,
            "logs": self._build_logs_page,
        }
        for key, builder in builders.items():
            frame = builder(self.page_container)
            frame.grid(row=0, column=0, sticky="nsew")
            self.pages[key] = frame

    def _register_start_button(self, button: ttk.Button) -> ttk.Button:
        self._start_action_buttons.append(button)
        return button

    def _register_terminate_button(self, button: ttk.Button) -> ttk.Button:
        self._terminate_action_buttons.append(button)
        return button

    def _sync_action_button_states(self) -> None:
        start_state = "normal" if self.workflow_controller.start_enabled else "disabled"
        terminate_state = "normal" if self.workflow_controller.terminate_enabled else "disabled"
        active_start_buttons: list[ttk.Button] = []
        for button in self._start_action_buttons:
            if not button.winfo_exists():
                continue
            button.configure(state=start_state)
            active_start_buttons.append(button)
        self._start_action_buttons = active_start_buttons
        active_terminate_buttons: list[ttk.Button] = []
        for button in self._terminate_action_buttons:
            if not button.winfo_exists():
                continue
            button.configure(state=terminate_state)
            active_terminate_buttons.append(button)
        self._terminate_action_buttons = active_terminate_buttons

    def _build_dashboard_page(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame")
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text="当前批次与主流程状态", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))

        top_info = ttk.Frame(frame, style="Surface.TFrame")
        top_info.grid(row=1, column=0, sticky="ew")
        top_info.columnconfigure(1, weight=1)
        info_pairs = (
            ("批次名称", self.batch_name_var),
            ("数据模式", self.workflow_mode_var),
            ("原始数据", self.single_input_dir_var),
            ("分项输出", self.stats_output_dir_var),
            ("汇总输出", self.summary_output_dir_var),
            ("PPT输出", self.output_ppt_path_var),
        )
        for index, (label, variable) in enumerate(info_pairs):
            ttk.Label(top_info, text=f"{label}：", style="Body.TLabel").grid(row=index, column=0, sticky="nw", pady=2)
            ttk.Label(top_info, textvariable=variable, style="Muted.TLabel").grid(row=index, column=1, sticky="nw", pady=2)

        step_frame = ttk.Frame(frame, style="Surface.TFrame")
        step_frame.grid(row=2, column=0, sticky="ew", pady=(18, 12))
        steps = (
            ("数据源", "data_source"),
            ("预处理", "preprocess"),
            ("分项统计", "stats"),
            ("汇总表", "summary"),
            ("PPT", "ppt"),
        )
        for index, (label, key) in enumerate(steps):
            card = ttk.Frame(step_frame, style="Card.TFrame", padding=12)
            card.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 8, 0))
            ttk.Label(card, text=label, style="SubHeader.TLabel").pack(anchor="w")
            ttk.Label(card, textvariable=self._step_status_vars[key], style="Muted.TLabel").pack(anchor="w", pady=(4, 0))
            step_frame.columnconfigure(index, weight=1)

        card_row = ttk.Frame(frame, style="Surface.TFrame")
        card_row.grid(row=3, column=0, sticky="ew", pady=(6, 12))
        card_row.columnconfigure((0, 1, 2, 3), weight=1)
        self._build_metric_card(card_row, 0, "原始文件数", self._raw_file_count_var)
        self._build_metric_card(card_row, 1, "分项文件数", self._stats_file_count_var)
        self._build_metric_card(card_row, 2, "汇总表状态", self._summary_file_var)
        self._build_metric_card(card_row, 3, "PPT状态", self._ppt_file_var)

        action_box = ttk.Frame(frame, style="Card.TFrame", padding=12)
        action_box.grid(row=4, column=0, sticky="ew")
        action_box.columnconfigure(1, weight=1)
        ttk.Label(action_box, text="当前建议操作", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(action_box, textvariable=self._next_action_var, style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 10))
        action_buttons = (
            ("导入数据源", lambda: self.show_page("data_source")),
            ("执行预处理", lambda: self.show_page("preprocess")),
            ("生成分项统计", lambda: self.show_page("stats")),
            ("生成汇总表", lambda: self.show_page("summary")),
            ("生成PPT", lambda: self.show_page("ppt")),
        )
        button_row = ttk.Frame(action_box, style="Card.TFrame")
        button_row.grid(row=2, column=0, sticky="ew")
        for index, (label, callback) in enumerate(action_buttons):
            style_name = "Primary.TButton" if label == "生成分项统计" else "Secondary.TButton"
            ttk.Button(button_row, text=label, style=style_name, command=callback).grid(row=0, column=index, padx=(0, 8))
        return frame

    def _build_metric_card(self, parent: ttk.Frame, column: int, title: str, variable: tk.StringVar) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=12)
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        ttk.Label(card, text=title, style="SubHeader.TLabel").pack(anchor="w")
        ttk.Label(card, textvariable=variable, style="Muted.TLabel").pack(anchor="w", pady=(6, 0))

    def _build_data_source_page(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame")
        frame.columnconfigure(0, weight=1)

        mode_box = ttk.LabelFrame(frame, text="处理方式", padding=12)
        mode_box.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ttk.Radiobutton(mode_box, text="单个月份 / 单个批次", value=WorkflowMode.SINGLE.value, variable=self.workflow_mode_var, command=self.refresh_all_status_views).grid(row=0, column=0, sticky="w", pady=2)
        ttk.Radiobutton(mode_box, text="合并多个月份后处理", value=WorkflowMode.MERGED.value, variable=self.workflow_mode_var, command=self.refresh_all_status_views).grid(row=1, column=0, sticky="w", pady=2)

        single_box = ttk.LabelFrame(frame, text="单月处理配置", padding=12)
        single_box.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        single_box.columnconfigure(1, weight=1)
        ttk.Label(single_box, text="原始数据目录", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(single_box, textvariable=self.single_input_dir_var).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(single_box, text="浏览", style="Secondary.TButton", command=lambda: self.choose_directory(self.single_input_dir_var)).grid(row=0, column=2)

        merge_box = ttk.LabelFrame(frame, text="多月合并配置", padding=12)
        merge_box.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        merge_box.columnconfigure(0, weight=1)
        self.merge_listbox = tk.Listbox(merge_box, height=5, relief="flat", bg="#fffdfb", fg=self.palette.text, selectbackground=self.palette.surface_alt)
        self.merge_listbox.grid(row=0, column=0, columnspan=3, sticky="ew")
        list_actions = ttk.Frame(merge_box, style="Surface.TFrame")
        list_actions.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(8, 8))
        ttk.Button(list_actions, text="添加目录", style="Secondary.TButton", command=self.add_merge_input_dir).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(list_actions, text="删除所选", style="Secondary.TButton", command=self.remove_selected_merge_dir).grid(row=0, column=1)
        ttk.Label(merge_box, text="合并输出目录", style="Body.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(merge_box, textvariable=self.merge_output_dir_var).grid(row=3, column=0, sticky="ew", pady=(4, 0))
        ttk.Button(merge_box, text="浏览", style="Secondary.TButton", command=lambda: self.choose_directory(self.merge_output_dir_var)).grid(row=3, column=1, padx=8, sticky="w")
        merge_button = ttk.Button(merge_box, text="执行合并", style="Primary.TButton", command=self.run_merge_task)
        merge_button.grid(row=3, column=2, sticky="e")
        self._register_start_button(merge_button)

        scan_box = ttk.LabelFrame(frame, text="当前输入目录与扫描结果", padding=12)
        scan_box.grid(row=3, column=0, sticky="nsew")
        scan_box.columnconfigure(0, weight=1)
        ttk.Label(scan_box, text="当前生效输入目录", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        self.effective_input_var = tk.StringVar()
        ttk.Label(scan_box, textvariable=self.effective_input_var, style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 8))
        ttk.Button(scan_box, text="刷新扫描结果", style="Secondary.TButton", command=self.refresh_all_status_views).grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.data_source_tree = ttk.Treeview(scan_box, columns=("path", "exists"), show="headings", height=8)
        self.data_source_tree.heading("path", text="路径")
        self.data_source_tree.heading("exists", text="状态")
        self.data_source_tree.column("path", width=620)
        self.data_source_tree.column("exists", width=120, anchor="center")
        self.data_source_tree.grid(row=3, column=0, sticky="nsew")
        return frame

    def _build_preprocess_page(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame")
        frame.columnconfigure(0, weight=1)

        phase_box = ttk.LabelFrame(frame, text="期次列修正", padding=12)
        phase_box.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(phase_box, text="说明：如果第三列存在“一期/二期”等期次值，则自动移动到最后一列。", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        phase_button = ttk.Button(phase_box, text="执行 phase_column_preprocess.py", style="Primary.TButton", command=self.run_phase_preprocess_task)
        phase_button.grid(row=1, column=0, sticky="w", pady=(10, 0))
        self._register_start_button(phase_button)

        fill_box = ttk.LabelFrame(frame, text="补写年份月份", padding=12)
        fill_box.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(fill_box, text="年份", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(fill_box, textvariable=self.year_value_var, width=12).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(fill_box, text="月份", style="Body.TLabel").grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Entry(fill_box, textvariable=self.month_value_var, width=12).grid(row=1, column=1, sticky="w", padx=(12, 0), pady=(4, 0))
        fill_button = ttk.Button(fill_box, text="执行 fill_year_month_columns.py", style="Secondary.TButton", command=self.run_fill_year_month_task)
        fill_button.grid(row=1, column=2, sticky="w", padx=(12, 0))
        self._register_start_button(fill_button)

        note_box = ttk.LabelFrame(frame, text="说明", padding=12)
        note_box.grid(row=2, column=0, sticky="ew")
        note_text = (
            "月份检查脚本文档已存在，但当前仓库中缺少 check_start_time_month.py。"
            " 第一版 GUI 先保留预处理主链：期次列修正、年份月份补写。"
        )
        ttk.Label(note_box, text=note_text, style="Muted.TLabel", wraplength=900, justify="left").grid(row=0, column=0, sticky="w")
        return frame

    def _build_stats_page(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame")
        frame.columnconfigure(0, weight=1)

        config_box = ttk.LabelFrame(frame, text="分项统计配置", padding=12)
        config_box.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        config_box.columnconfigure(1, weight=1)
        ttk.Label(config_box, text="输入目录", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        self.stats_effective_input_var = tk.StringVar()
        ttk.Label(config_box, textvariable=self.stats_effective_input_var, style="Muted.TLabel").grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(config_box, text="输出目录", style="Body.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(config_box, textvariable=self.stats_output_dir_var).grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(config_box, text="浏览", style="Secondary.TButton", command=lambda: self.choose_directory(self.stats_output_dir_var)).grid(row=1, column=2, sticky="w")
        ttk.Label(config_box, text="sheet", style="Body.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(config_box, textvariable=self.sheet_name_var, width=16).grid(row=2, column=1, sticky="w", padx=8)
        ttk.Label(config_box, text="计算口径", style="Body.TLabel").grid(row=2, column=2, sticky="w")
        ttk.Combobox(config_box, textvariable=self.calculation_mode_var, values=("template", "summary"), width=12, state="readonly").grid(row=2, column=3, sticky="w", padx=8)
        button_row = ttk.Frame(config_box, style="Surface.TFrame")
        button_row.grid(row=3, column=0, columnspan=4, sticky="w", pady=(10, 0))
        scan_button = ttk.Button(button_row, text="扫描可统计客群", style="Secondary.TButton", command=self.scan_stats_preview)
        scan_button.grid(row=0, column=0, padx=(0, 8))
        self._register_start_button(scan_button)
        dry_run_button = ttk.Button(button_row, text="dry-run 校验", style="Secondary.TButton", command=self.run_survey_stats_dry_run_task)
        dry_run_button.grid(row=0, column=1, padx=(0, 8))
        self._register_start_button(dry_run_button)
        stats_button = ttk.Button(button_row, text="开始生成分项统计", style="Primary.TButton", command=self.run_survey_stats_task)
        stats_button.grid(row=0, column=2, padx=(0, 8))
        self._register_start_button(stats_button)
        ttk.Button(button_row, text="保存当前批次", style="Secondary.TButton", command=self.save_current_batch).grid(row=0, column=3)

        summary_box = ttk.LabelFrame(frame, text="当前状态", padding=12)
        summary_box.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self.stats_summary_var = tk.StringVar(value="待扫描")
        ttk.Label(summary_box, textvariable=self.stats_summary_var, style="Muted.TLabel").grid(row=0, column=0, sticky="w")

        preview_box = ttk.LabelFrame(frame, text="客群扫描预览", padding=12)
        preview_box.grid(row=2, column=0, sticky="nsew", pady=(0, 12))
        preview_box.columnconfigure(0, weight=1)
        ttk.Label(preview_box, textvariable=self.stats_preview_summary_var, style="Muted.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Label(
            preview_box,
            text="双击某一行可切换勾选；正式执行和 dry-run 都只处理当前已勾选客群。",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 8))
        preview_actions = ttk.Frame(preview_box, style="Surface.TFrame")
        preview_actions.grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Button(preview_actions, text="全选", style="Secondary.TButton", command=self.select_all_customer_types).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(preview_actions, text="清空", style="Secondary.TButton", command=self.clear_selected_customer_types).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(preview_actions, text="仅选可生成", style="Secondary.TButton", command=self.select_ready_customer_types).grid(row=0, column=2)
        self.stats_preview_tree = ttk.Treeview(
            preview_box,
            columns=("selected", "customer_type", "source_file", "status", "output_name", "detail"),
            show="headings",
            height=10,
        )
        self.stats_preview_tree.heading("selected", text="选择")
        self.stats_preview_tree.heading("customer_type", text="客群")
        self.stats_preview_tree.heading("source_file", text="来源文件")
        self.stats_preview_tree.heading("status", text="状态")
        self.stats_preview_tree.heading("output_name", text="输出文件")
        self.stats_preview_tree.heading("detail", text="详情")
        self.stats_preview_tree.column("selected", width=70, anchor="center")
        self.stats_preview_tree.column("customer_type", width=180)
        self.stats_preview_tree.column("source_file", width=140)
        self.stats_preview_tree.column("status", width=130, anchor="center")
        self.stats_preview_tree.column("output_name", width=180)
        self.stats_preview_tree.column("detail", width=420)
        self.stats_preview_tree.grid(row=3, column=0, sticky="nsew")
        self.stats_preview_tree.bind("<Double-1>", self.on_stats_preview_row_double_click)

        result_box = ttk.LabelFrame(frame, text="输出目录文件预览", padding=12)
        result_box.grid(row=3, column=0, sticky="nsew")
        result_box.columnconfigure(0, weight=1)
        self.stats_tree = ttk.Treeview(result_box, columns=("name", "path"), show="headings", height=10)
        self.stats_tree.heading("name", text="文件名")
        self.stats_tree.heading("path", text="路径")
        self.stats_tree.column("name", width=220)
        self.stats_tree.column("path", width=650)
        self.stats_tree.grid(row=0, column=0, sticky="nsew")
        return frame

    def _build_summary_page(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame")
        frame.columnconfigure(0, weight=1)
        box = ttk.LabelFrame(frame, text="汇总统计配置", padding=12)
        box.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        box.columnconfigure(1, weight=1)
        ttk.Label(box, text="分项结果目录", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(box, textvariable=self.stats_output_dir_var, style="Muted.TLabel").grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(box, text="汇总输出目录", style="Body.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.summary_output_dir_var).grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(box, text="浏览", style="Secondary.TButton", command=lambda: self.choose_directory(self.summary_output_dir_var)).grid(row=1, column=2, sticky="w")
        ttk.Label(box, text="输出文件名", style="Body.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.summary_output_name_var).grid(row=2, column=1, sticky="ew", padx=8)
        summary_button = ttk.Button(box, text="生成汇总表", style="Primary.TButton", command=self.run_summary_task)
        summary_button.grid(row=3, column=0, sticky="w", pady=(10, 0))
        self._register_start_button(summary_button)

        result_box = ttk.LabelFrame(frame, text="结果状态", padding=12)
        result_box.grid(row=1, column=0, sticky="ew")
        self.summary_status_var = tk.StringVar(value="待执行")
        ttk.Label(result_box, textvariable=self.summary_status_var, style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        return frame

    def _build_ppt_page(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame")
        frame.columnconfigure(0, weight=1)
        box = ttk.LabelFrame(frame, text="PPT生成配置", padding=12)
        box.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        box.columnconfigure(1, weight=1)
        ttk.Label(box, text="统计结果目录", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(box, textvariable=self.stats_output_dir_var, style="Muted.TLabel").grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(box, text="模板PPT", style="Body.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.ppt_template_path_var).grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(box, text="浏览", style="Secondary.TButton", command=lambda: self.choose_file(self.ppt_template_path_var, [("PowerPoint", "*.pptx")])).grid(row=1, column=2)
        ttk.Label(box, text="输出PPT", style="Body.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.output_ppt_path_var).grid(row=2, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(box, text="浏览", style="Secondary.TButton", command=lambda: self.choose_save_file(self.output_ppt_path_var, ".pptx", [("PowerPoint", "*.pptx")])).grid(row=2, column=2)
        ttk.Label(box, text="section_mode", style="Body.TLabel").grid(row=3, column=0, sticky="w")
        ttk.Combobox(box, textvariable=self.ppt_section_mode_var, values=("auto", "template", "summary"), state="readonly", width=12).grid(row=3, column=1, sticky="w", padx=8)
        ttk.Label(box, text="空值显示", style="Body.TLabel").grid(row=4, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.ppt_blank_display_var, width=18).grid(row=4, column=1, sticky="w", padx=8, pady=6)
        button_row = ttk.Frame(box, style="Surface.TFrame")
        button_row.grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ppt_dry_run_button = ttk.Button(button_row, text="dry-run 校验", style="Secondary.TButton", command=self.run_ppt_dry_run_task)
        ppt_dry_run_button.grid(row=0, column=0, padx=(0, 8))
        self._register_start_button(ppt_dry_run_button)
        ppt_button = ttk.Button(button_row, text="生成PPT", style="Primary.TButton", command=self.run_ppt_task)
        ppt_button.grid(row=0, column=1)
        self._register_start_button(ppt_button)

        status_box = ttk.LabelFrame(frame, text="结果状态", padding=12)
        status_box.grid(row=1, column=0, sticky="ew")
        self.ppt_status_var = tk.StringVar(value="待执行")
        ttk.Label(status_box, textvariable=self.ppt_status_var, style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        return frame

    def _build_logs_page(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame")
        frame.columnconfigure(0, weight=1)
        info_box = ttk.LabelFrame(frame, text="说明", padding=12)
        info_box.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(
            info_box,
            text="右侧日志面板会实时显示脚本输出。运行 survey_stats.py 和 generate_ppt.py 时，GUI 会自动生成临时 TOML 配置写入 logs/gui_runtime。",
            style="Muted.TLabel",
            wraplength=900,
            justify="left",
        ).grid(row=0, column=0, sticky="w")
        runtime_box = ttk.LabelFrame(frame, text="运行目录", padding=12)
        runtime_box.grid(row=1, column=0, sticky="ew")
        ttk.Label(runtime_box, text=str(GUI_RUNTIME_DIR), style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        return frame

    def show_page(self, page_key: str) -> None:
        self._current_page_title_var.set(self.PAGE_TITLES[page_key])
        frame = self.pages[page_key]
        frame.tkraise()

    def _refresh_merge_listbox(self) -> None:
        self.merge_listbox.delete(0, "end")
        for item in self.merge_input_list:
            self.merge_listbox.insert("end", item)

    def _reset_stats_preview_state(self) -> None:
        self.stats_preview_summary = None
        self.stats_preview_rows = ()
        self.selected_customer_types = frozenset()
        self._refresh_stats_preview_tree()
        self._refresh_stats_preview_summary()

    def _refresh_saved_batch_profiles(self, *, selected_name: str | None = None) -> None:
        self._saved_batch_profiles = load_saved_batch_profiles()
        saved_batch_names = tuple(profile.batch_name for profile in self._saved_batch_profiles)
        self.saved_batch_combobox.configure(values=saved_batch_names)

        target_name = selected_name
        if target_name is None and self.active_saved_batch_name in saved_batch_names:
            target_name = self.active_saved_batch_name
        if target_name in saved_batch_names:
            self.saved_batch_var.set(target_name or "")
            return
        self.saved_batch_var.set("")

    def _persist_session_state(self) -> None:
        save_gui_session(
            self.current_config(),
            active_saved_batch_name=self.active_saved_batch_name,
        )

    def _apply_config_to_form(self, config: GuiBatchConfig) -> None:
        self.batch_name_var.set(config.batch_name)
        self.workflow_mode_var.set(config.workflow_mode.value)
        self.single_input_dir_var.set(str(config.single_input_dir))
        self.merge_input_list = [str(path) for path in config.merge_input_dirs]
        self._refresh_merge_listbox()
        self.merge_output_dir_var.set(str(config.merge_output_dir))
        self.sheet_name_var.set(config.sheet_name)
        self.year_value_var.set(config.year_value)
        self.month_value_var.set(config.month_value)
        self.stats_output_dir_var.set(str(config.stats_output_dir))
        self.calculation_mode_var.set(config.calculation_mode)
        self.output_format_var.set(config.output_format)
        self.summary_output_dir_var.set(str(config.summary_output_dir))
        self.summary_output_name_var.set(config.summary_output_name)
        self.ppt_template_path_var.set(str(config.ppt_template_path))
        self.output_ppt_path_var.set(str(config.output_ppt_path))
        self.ppt_section_mode_var.set(config.ppt_section_mode)
        self.ppt_blank_display_var.set(config.ppt_blank_display)
        self._reset_stats_preview_state()
        self.refresh_all_status_views()

    def _restore_persistent_state(self) -> None:
        self._refresh_saved_batch_profiles()
        config = DEFAULT_GUI_BATCH_CONFIG
        active_saved_batch_name: str | None = None
        if GUI_SESSION_PATH.exists():
            try:
                config, active_saved_batch_name = load_gui_session()
            except Exception as exc:
                self.append_log(f"[WARN] 最近会话恢复失败：{exc}")
        self._apply_config_to_form(config)
        saved_batch_names = {profile.batch_name for profile in self._saved_batch_profiles}
        if active_saved_batch_name in saved_batch_names:
            self.active_saved_batch_name = active_saved_batch_name
        else:
            self.active_saved_batch_name = None
        self._refresh_saved_batch_profiles(selected_name=self.active_saved_batch_name)

    def _build_new_batch_config(self) -> GuiBatchConfig:
        return replace(self.current_config(), batch_name="未命名批次")

    def _ensure_batch_switch_allowed(self) -> bool:
        if not self.workflow_controller.start_enabled or self.runner.is_running:
            messagebox.showwarning("任务执行中", "请等待当前任务完成后再切换或删除批次。")
            return False
        return True

    def load_selected_batch(self) -> None:
        if not self._ensure_batch_switch_allowed():
            return
        selected_batch_name = self.saved_batch_var.get().strip()
        if not selected_batch_name:
            messagebox.showinfo("未选择批次", "请先从下拉列表中选择一个已保存批次。")
            return
        profile = next(
            (item for item in self._saved_batch_profiles if item.batch_name == selected_batch_name),
            None,
        )
        if profile is None:
            messagebox.showerror("批次不存在", "所选批次配置不存在，请刷新后重试。")
            self._refresh_saved_batch_profiles()
            return
        self.active_saved_batch_name = profile.batch_name
        self._apply_config_to_form(profile.config)
        self._refresh_saved_batch_profiles(selected_name=profile.batch_name)
        self._persist_session_state()
        self.append_log(f"[INFO] 已加载批次：{profile.batch_name}")

    def create_new_batch(self) -> None:
        if not self._ensure_batch_switch_allowed():
            return
        self.active_saved_batch_name = None
        self._apply_config_to_form(self._build_new_batch_config())
        self._refresh_saved_batch_profiles()
        self._persist_session_state()
        self.append_log("[INFO] 已创建新的批次草稿。")

    def save_current_batch(self) -> None:
        config = self.current_config()
        if not config.batch_name.strip():
            messagebox.showerror("批次名称为空", "请先填写批次名称后再保存。")
            return
        profile_path = save_batch_profile(config)
        self.active_saved_batch_name = config.batch_name
        self._refresh_saved_batch_profiles(selected_name=config.batch_name)
        self._persist_session_state()
        messagebox.showinfo("已保存", f"批次配置已保存到：\n{profile_path}")
        self.append_log(f"[INFO] 已保存批次：{config.batch_name}")

    def delete_selected_batch(self) -> None:
        if not self._ensure_batch_switch_allowed():
            return
        selected_batch_name = self.saved_batch_var.get().strip() or self.active_saved_batch_name or ""
        if not selected_batch_name:
            messagebox.showinfo("无可删除批次", "当前没有选中的已保存批次。")
            return
        confirmed = messagebox.askyesno(
            "删除批次",
            f"确认删除批次“{selected_batch_name}”吗？\n该操作不会删除原始数据文件。",
        )
        if not confirmed:
            return
        deleted = delete_batch_profile(selected_batch_name)
        if not deleted:
            messagebox.showerror("删除失败", "未找到对应批次配置文件，可能已经被删除。")
            self._refresh_saved_batch_profiles()
            return
        if self.active_saved_batch_name == selected_batch_name:
            self.active_saved_batch_name = None
        self._refresh_saved_batch_profiles()
        self._persist_session_state()
        self.append_log(f"[INFO] 已删除批次：{selected_batch_name}")

    def current_config(self) -> GuiBatchConfig:
        merge_dirs = tuple(Path(item) for item in self.merge_input_list)
        mode = WorkflowMode(self.workflow_mode_var.get())
        return GuiBatchConfig(
            batch_name=self.batch_name_var.get().strip() or "未命名批次",
            workflow_mode=mode,
            single_input_dir=Path(self.single_input_dir_var.get().strip() or ".").resolve(),
            merge_input_dirs=merge_dirs,
            merge_output_dir=Path(self.merge_output_dir_var.get().strip() or ".").resolve(),
            sheet_name=self.sheet_name_var.get().strip() or DEFAULT_SHEET_NAME,
            year_value=self.year_value_var.get().strip(),
            month_value=self.month_value_var.get().strip(),
            stats_output_dir=Path(self.stats_output_dir_var.get().strip() or ".").resolve(),
            calculation_mode=self.calculation_mode_var.get().strip() or "template",
            output_format=self.output_format_var.get().strip() or "xlsx",
            summary_output_dir=Path(self.summary_output_dir_var.get().strip() or ".").resolve(),
            summary_output_name=self.summary_output_name_var.get().strip() or "客户类型满意度汇总表.xlsx",
            ppt_template_path=Path(self.ppt_template_path_var.get().strip() or ".").resolve(),
            output_ppt_path=Path(self.output_ppt_path_var.get().strip() or ".").resolve(),
            ppt_section_mode=self.ppt_section_mode_var.get().strip() or "auto",
            ppt_blank_display=self.ppt_blank_display_var.get(),
        )

    def choose_directory(self, variable: tk.StringVar) -> None:
        initial = variable.get().strip() or str(PROJECT_ROOT)
        selected = filedialog.askdirectory(initialdir=initial)
        if selected:
            variable.set(selected)
            self.refresh_all_status_views()

    def choose_file(self, variable: tk.StringVar, filetypes: list[tuple[str, str]]) -> None:
        initial = variable.get().strip() or str(PROJECT_ROOT)
        selected = filedialog.askopenfilename(initialdir=str(Path(initial).parent), filetypes=filetypes)
        if selected:
            variable.set(selected)
            self.refresh_all_status_views()

    def choose_save_file(self, variable: tk.StringVar, default_ext: str, filetypes: list[tuple[str, str]]) -> None:
        initial = variable.get().strip() or str(PROJECT_ROOT / f"output{default_ext}")
        selected = filedialog.asksaveasfilename(initialfile=Path(initial).name, initialdir=str(Path(initial).parent), defaultextension=default_ext, filetypes=filetypes)
        if selected:
            variable.set(selected)
            self.refresh_all_status_views()

    def add_merge_input_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=str(PROJECT_ROOT))
        if not selected:
            return
        if selected in self.merge_input_list:
            messagebox.showinfo("目录已存在", "该目录已经添加过了。")
            return
        self.merge_input_list.append(selected)
        self.merge_listbox.insert("end", selected)
        self.refresh_all_status_views()

    def remove_selected_merge_dir(self) -> None:
        selection = self.merge_listbox.curselection()
        if not selection:
            return
        for index in reversed(selection):
            del self.merge_input_list[index]
            self.merge_listbox.delete(index)
        self.refresh_all_status_views()

    def refresh_all_status_views(self) -> None:
        config = self.current_config()
        effective_input = config.effective_input_dir()
        stats_files = discover_excel_files(config.stats_output_dir)
        self.effective_input_var.set(str(effective_input))
        self.stats_effective_input_var.set(str(effective_input))
        self._raw_file_count_var.set(str(len(discover_excel_files(effective_input))))
        self._stats_file_count_var.set(str(len(stats_files)))
        summary_path = config.summary_output_dir / config.summary_output_name
        self._summary_file_var.set("已生成" if summary_path.exists() else "未生成")
        self._ppt_file_var.set("已生成" if config.output_ppt_path.exists() else "未生成")

        self._step_status_vars["data_source"].set("已完成" if effective_input.exists() else "待设置")
        if stats_files:
            self._step_status_vars["stats"].set("已完成")
        if summary_path.exists():
            self._step_status_vars["summary"].set("已完成")
        if config.output_ppt_path.exists():
            self._step_status_vars["ppt"].set("已完成")
        self.stats_summary_var.set(
            f"输入目录：{effective_input}；输出目录：{config.stats_output_dir}；当前分项文件数：{len(stats_files)}"
        )
        self.summary_status_var.set(
            f"目标汇总表：{summary_path}"
        )
        self.ppt_status_var.set(
            f"模板：{config.ppt_template_path}；输出：{config.output_ppt_path}"
        )

        self._refresh_treeviews(config)
        self._recompute_next_action(config)

    def _refresh_treeviews(self, config: GuiBatchConfig) -> None:
        for item in self.data_source_tree.get_children():
            self.data_source_tree.delete(item)
        data_source_paths: list[Path] = []
        if config.workflow_mode == WorkflowMode.SINGLE:
            data_source_paths.append(config.single_input_dir)
        else:
            data_source_paths.extend(config.merge_input_dirs)
            data_source_paths.append(config.merge_output_dir)
        for path in data_source_paths:
            status = "存在" if Path(path).exists() else "缺失"
            self.data_source_tree.insert("", "end", values=(str(path), status))

        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)
        for path in discover_excel_files(config.stats_output_dir):
            self.stats_tree.insert("", "end", values=(path.name, str(path)))

    def _recompute_next_action(self, config: GuiBatchConfig) -> None:
        if not config.effective_input_dir().exists():
            self._next_action_var.set("先配置并确认数据源")
            return
        if self._step_status_vars["preprocess"].get() == "待执行":
            self._next_action_var.set("建议先执行预处理")
            return
        if self._step_status_vars["stats"].get() == "待执行":
            self._next_action_var.set("建议先生成分项统计")
            return
        if self._step_status_vars["summary"].get() == "待执行":
            self._next_action_var.set("建议先生成汇总表")
            return
        if self._step_status_vars["ppt"].get() == "待执行":
            self._next_action_var.set("建议生成PPT")
            return
        self._next_action_var.set("本批次主流程已完成")

    def _update_task_status_text(self) -> None:
        self._task_status_label_var.set(
            build_workflow_status_text(self.workflow_controller, self._task_title_lookup)
        )

    def _set_stats_preview_selection(self, selected_customer_types: frozenset[str]) -> None:
        valid_customer_types = {
            row.customer_type_name
            for row in self.stats_preview_rows
        }
        self.selected_customer_types = frozenset(
            name for name in selected_customer_types if name in valid_customer_types
        )
        self._refresh_stats_preview_tree()
        self._refresh_stats_preview_summary()

    def _apply_stats_preview_summary(self, summary: StatsPreviewSummary) -> None:
        self.stats_preview_summary = summary
        self.stats_preview_rows = summary.rows
        self.selected_customer_types = default_selected_customer_types(summary)
        self._refresh_stats_preview_tree()
        self._refresh_stats_preview_summary()

    def _refresh_stats_preview_summary(self) -> None:
        if self.stats_preview_summary is None:
            self.stats_preview_summary_var.set("未扫描客群")
            return
        self.stats_preview_summary_var.set(
            build_stats_preview_summary_text(
                self.stats_preview_summary,
                self.selected_customer_types,
            )
        )

    def _refresh_stats_preview_tree(self) -> None:
        for item in self.stats_preview_tree.get_children():
            self.stats_preview_tree.delete(item)
        for row in self.stats_preview_rows:
            selected_marker = "✓" if row.customer_type_name in self.selected_customer_types else ""
            self.stats_preview_tree.insert(
                "",
                "end",
                iid=row.customer_type_name,
                values=(
                    selected_marker,
                    row.customer_type_name,
                    row.source_file_name,
                    customer_type_preview_status_text(row.status),
                    row.output_name,
                    row.detail,
                ),
            )

    def select_all_customer_types(self) -> None:
        self._set_stats_preview_selection(
            frozenset(row.customer_type_name for row in self.stats_preview_rows)
        )

    def clear_selected_customer_types(self) -> None:
        self._set_stats_preview_selection(frozenset())

    def select_ready_customer_types(self) -> None:
        if self.stats_preview_summary is None:
            return
        self._set_stats_preview_selection(default_selected_customer_types(self.stats_preview_summary))

    def on_stats_preview_row_double_click(self, event: tk.Event[tk.Misc]) -> None:
        item_id = self.stats_preview_tree.identify_row(event.y)
        if not item_id:
            return
        selected_customer_types = set(self.selected_customer_types)
        if item_id in selected_customer_types:
            selected_customer_types.remove(item_id)
        else:
            selected_customer_types.add(item_id)
        self._set_stats_preview_selection(frozenset(selected_customer_types))

    def _ensure_stats_preview_loaded(self, *, parent: tk.Misc | None = None) -> bool:
        if self.stats_preview_summary is not None:
            return True
        try:
            summary = build_stats_preview_summary(self.current_config())
        except Exception as exc:
            if parent is None:
                messagebox.showerror("扫描失败", str(exc))
            else:
                messagebox.showerror("扫描失败", str(exc), parent=parent)
            return False
        self._apply_stats_preview_summary(summary)
        return True

    def _selected_stats_job_names(self, *, parent: tk.Misc | None = None) -> tuple[str, ...] | None:
        if not self._ensure_stats_preview_loaded(parent=parent):
            return None
        selected_job_names = ordered_selected_customer_types(
            self.stats_preview_rows,
            self.selected_customer_types,
        )
        if selected_job_names:
            return selected_job_names
        if parent is None:
            messagebox.showinfo("未选择客群", "请至少勾选一个客群后再执行分项统计。")
        else:
            messagebox.showinfo("未选择客群", "请至少勾选一个客群后再执行分项统计。", parent=parent)
        return None

    def scan_stats_preview(self) -> None:
        try:
            summary = build_stats_preview_summary(self.current_config())
        except Exception as exc:
            messagebox.showerror("扫描失败", str(exc))
            return
        self._apply_stats_preview_summary(summary)

    def terminate_current_task(self) -> None:
        if not self.workflow_controller.terminate_enabled:
            messagebox.showinfo("无可终止任务", "当前没有正在执行的任务。")
            return
        self.workflow_controller.request_cancel()
        self._update_task_status_text()
        if self.workflow_controller.active_step_key is not None:
            self._set_workflow_progress_row(
                self.workflow_controller.active_step_key,
                "正在终止",
                "已发送终止请求，等待当前脚本退出",
            )
        self._refresh_workflow_progress_summary()
        self._sync_action_button_states()
        self.append_log("[INFO] 已请求终止当前任务。")
        self.runner.terminate()

    def append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _destroy_workflow_progress_dialog(self) -> None:
        dialog = self._workflow_progress_dialog
        self._workflow_progress_dialog = None
        self._workflow_progress_tree = None
        self._workflow_progress_item_ids = {}
        self._workflow_progress_last_page_key = None
        self._workflow_progress_summary_var.set("未启动主流程")
        self._workflow_progress_close_var.set("关闭")
        if dialog is not None and dialog.winfo_exists():
            dialog.destroy()

    def _hide_or_close_workflow_progress_dialog(self) -> None:
        dialog = self._workflow_progress_dialog
        if dialog is None or not dialog.winfo_exists():
            return
        if self.workflow_controller.status in {WorkflowRunStatus.RUNNING, WorkflowRunStatus.CANCELLING}:
            dialog.withdraw()
            return
        self._destroy_workflow_progress_dialog()

    def _open_workflow_progress_dialog(self, tasks: tuple[TaskCommand, ...]) -> None:
        self._destroy_workflow_progress_dialog()
        dialog = tk.Toplevel(self)
        dialog.title("主流程执行进度")
        dialog.transient(self)
        dialog.configure(bg=self.palette.background)
        dialog.geometry("980x420")
        dialog.minsize(860, 360)
        dialog.protocol("WM_DELETE_WINDOW", self._hide_or_close_workflow_progress_dialog)

        wrap = ttk.Frame(dialog, style="Surface.TFrame", padding=16)
        wrap.pack(fill="both", expand=True)
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(1, weight=1)

        ttk.Label(wrap, text="正在执行主流程", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(wrap, textvariable=self._workflow_progress_summary_var, style="Muted.TLabel").grid(row=0, column=1, sticky="e")

        tree = ttk.Treeview(
            wrap,
            columns=("step", "status", "detail"),
            show="headings",
            height=10,
        )
        tree.heading("step", text="步骤")
        tree.heading("status", text="状态")
        tree.heading("detail", text="进度明细")
        tree.column("step", width=220)
        tree.column("status", width=120, anchor="center")
        tree.column("detail", width=580)
        tree.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(12, 12))

        footer = ttk.Frame(wrap, style="Surface.TFrame")
        footer.grid(row=2, column=0, columnspan=2, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Button(footer, text="查看任务日志", style="Secondary.TButton", command=lambda: self.show_page("logs")).grid(row=0, column=0, sticky="w")
        close_button = ttk.Button(
            footer,
            textvariable=self._workflow_progress_close_var,
            style="Secondary.TButton",
            command=self._hide_or_close_workflow_progress_dialog,
        )
        close_button.grid(row=0, column=1, padx=(0, 8))
        terminate_button = ttk.Button(
            footer,
            text="终止当前任务",
            style="Primary.TButton",
            command=self.terminate_current_task,
        )
        terminate_button.grid(row=0, column=2)
        self._register_terminate_button(terminate_button)

        self._workflow_progress_dialog = dialog
        self._workflow_progress_tree = tree
        self._workflow_progress_item_ids = {}
        self._workflow_progress_last_page_key = page_key_for_task(tasks[-1].key) if tasks else None
        for index, task in enumerate(tasks, start=1):
            item_id = task.key
            self._workflow_progress_item_ids[task.key] = item_id
            tree.insert(
                "",
                "end",
                iid=item_id,
                values=(f"{index}. {task.title}", "待执行", "等待前置步骤"),
            )
        self._workflow_progress_summary_var.set(f"共 {len(tasks)} 步，等待启动")
        self._workflow_progress_close_var.set("后台运行")
        dialog.lift()
        self._sync_action_button_states()

    def _set_workflow_progress_row(
        self,
        task_key: str,
        status_text: str,
        detail_text: str,
    ) -> None:
        if self._workflow_progress_tree is None or not self._workflow_progress_tree.winfo_exists():
            return
        item_id = self._workflow_progress_item_ids.get(task_key)
        if item_id is None:
            return
        current_values = self._workflow_progress_tree.item(item_id, "values")
        if not current_values:
            return
        self._workflow_progress_tree.item(
            item_id,
            values=(current_values[0], status_text, detail_text),
        )

    def _refresh_workflow_progress_summary(self) -> None:
        summary = build_workflow_status_text(self.workflow_controller, self._task_title_lookup)
        total = len(self.workflow_controller.planned_step_keys)
        completed = len(self.workflow_controller.completed_step_keys)
        if total:
            summary = f"{summary}；进度 {completed}/{total}"
        self._workflow_progress_summary_var.set(summary)
        if self.workflow_controller.status in {WorkflowRunStatus.RUNNING, WorkflowRunStatus.CANCELLING}:
            self._workflow_progress_close_var.set("后台运行")
        else:
            self._workflow_progress_close_var.set("关闭")

    def _mark_pending_workflow_rows(self, status_text: str, detail_text: str) -> None:
        if self._workflow_progress_tree is None or not self._workflow_progress_tree.winfo_exists():
            return
        for task_key in self.workflow_controller.planned_step_keys:
            if task_key in self.workflow_controller.completed_step_keys:
                continue
            if task_key == self.workflow_controller.failed_step_key:
                continue
            item_id = self._workflow_progress_item_ids.get(task_key)
            if item_id is None:
                continue
            current_values = self._workflow_progress_tree.item(item_id, "values")
            if current_values and current_values[1] == "待执行":
                self._set_workflow_progress_row(task_key, status_text, detail_text)

    def run_task_list(
        self,
        tasks: tuple[TaskCommand, ...],
        *,
        show_workflow_progress: bool = False,
    ) -> None:
        if not tasks:
            messagebox.showinfo("无可执行任务", "当前没有需要执行的任务。")
            return
        if not self.workflow_controller.start_enabled or self.runner.is_running:
            messagebox.showwarning("任务执行中", "请等待当前任务完成后再继续。")
            return
        if show_workflow_progress:
            self._open_workflow_progress_dialog(tasks)
        else:
            self._destroy_workflow_progress_dialog()
        for task in tasks:
            self.append_log(f"[CMD] {task.title}: {shlex.join(task.command)}")
        self._task_title_lookup = {task.key: task.title for task in tasks}
        self.workflow_controller.begin(tuple(task.key for task in tasks))
        self._update_task_status_text()
        self._refresh_workflow_progress_summary()
        self._sync_action_button_states()
        self.runner.run_tasks(tasks)

    def run_merge_task(self) -> None:
        try:
            command = build_merge_command(self.current_config())
            if command is None:
                messagebox.showinfo("无需合并", "当前是单月模式，无需执行合并。")
                return
            self.run_task_list((TaskCommand("merge_workbooks", "合并多月问卷", tuple(command)),))
        except ValueError as exc:
            messagebox.showerror("参数不完整", str(exc))

    def run_phase_preprocess_task(self) -> None:
        self._run_single_builder_task("phase_preprocess", "修正期次列", build_phase_preprocess_command)

    def run_fill_year_month_task(self) -> None:
        self._run_single_builder_task("fill_year_month", "补写年份月份", build_fill_year_month_command)

    def run_survey_stats_task(self) -> None:
        selected_job_names = self._selected_stats_job_names()
        if selected_job_names is None:
            return
        self._run_single_command(
            TaskCommand(
                "survey_stats",
                "生成分项统计",
                tuple(
                    build_survey_stats_command(
                        self.current_config(),
                        selected_job_names=selected_job_names,
                    )
                ),
            )
        )

    def run_survey_stats_dry_run_task(self) -> None:
        selected_job_names = self._selected_stats_job_names()
        if selected_job_names is None:
            return
        self._run_single_command(
            TaskCommand(
                "survey_stats",
                "分项统计 dry-run",
                tuple(
                    build_survey_stats_command(
                        self.current_config(),
                        dry_run=True,
                        selected_job_names=selected_job_names,
                    )
                ),
            )
        )

    def run_summary_task(self) -> None:
        self._run_single_command(
            TaskCommand(
                "summary_table",
                "生成汇总表",
                tuple(build_summary_command(self.current_config())),
            )
        )

    def run_ppt_task(self) -> None:
        self._run_single_command(
            TaskCommand(
                "generate_ppt",
                "生成PPT",
                tuple(build_ppt_command(self.current_config())),
            )
        )

    def run_ppt_dry_run_task(self) -> None:
        self._run_single_command(
            TaskCommand(
                "generate_ppt",
                "PPT dry-run",
                tuple(build_ppt_command(self.current_config(), dry_run=True)),
            )
        )

    def _run_single_builder_task(
        self,
        key: str,
        title: str,
        builder,
    ) -> None:
        try:
            command = builder(self.current_config())
        except ValueError as exc:
            messagebox.showerror("参数不完整", str(exc))
            return
        self._run_single_command(TaskCommand(key, title, tuple(command)))

    def _run_single_command(self, task: TaskCommand) -> None:
        self.run_task_list((task,))

    def on_task_started(self, task: TaskCommand) -> None:
        self.workflow_controller.mark_task_started(task.key)
        self._update_task_status_text()
        self._set_workflow_progress_row(task.key, "执行中", shlex.join(task.command))
        self._refresh_workflow_progress_summary()
        self.append_log(f"[START] {task.title}")

    def on_task_finished(self, task: TaskCommand, success: bool) -> None:
        self.workflow_controller.mark_task_finished(task.key, success)
        if success:
            self.append_log(f"[DONE] {task.title}")
            self._set_workflow_progress_row(task.key, "已完成", "执行完成，可继续下一步")
            if task.key in {"merge_workbooks"}:
                self._step_status_vars["data_source"].set("已完成")
            elif task.key in {"phase_preprocess", "fill_year_month"}:
                self._step_status_vars["preprocess"].set("已完成")
            elif task.key == "survey_stats":
                self._step_status_vars["stats"].set("已完成")
            elif task.key == "summary_table":
                self._step_status_vars["summary"].set("已完成")
            elif task.key == "generate_ppt":
                self._step_status_vars["ppt"].set("已完成")
        else:
            if self.workflow_controller.cancel_requested:
                self.append_log(f"[CANCEL] {task.title}")
                self._set_workflow_progress_row(task.key, "已取消", "用户已终止当前步骤")
            else:
                self.append_log(f"[FAIL] {task.title}")
                self._set_workflow_progress_row(task.key, "失败", "执行失败，请查看日志")
        self._update_task_status_text()
        self._refresh_workflow_progress_summary()
        self.refresh_all_status_views()

    def on_all_tasks_finished(self, success: bool) -> None:
        self.workflow_controller.finish_run(success)
        self._update_task_status_text()
        if self.workflow_controller.status == WorkflowRunStatus.CANCELLED:
            self._mark_pending_workflow_rows("已取消", "主流程已终止，未再执行")
        elif self.workflow_controller.status == WorkflowRunStatus.FAILED:
            self._mark_pending_workflow_rows("未执行", "前序步骤失败，后续步骤未启动")
        self._refresh_workflow_progress_summary()
        self._sync_action_button_states()
        if success:
            self.append_log("[INFO] 所有任务执行完成。")
        elif self.workflow_controller.status == WorkflowRunStatus.CANCELLED:
            self.append_log("[INFO] 主流程已取消。")
        self.refresh_all_status_views()
        if self._workflow_progress_last_page_key is not None:
            self.show_page(self._workflow_progress_last_page_key)

    def save_profile_snapshot(self) -> None:
        self.save_current_batch()

    def on_app_close(self) -> None:
        try:
            self._persist_session_state()
        except Exception as exc:
            self.append_log(f"[WARN] 最近会话保存失败：{exc}")
        self.destroy()

    def open_main_workflow_dialog(self) -> None:
        if not self.workflow_controller.start_enabled:
            messagebox.showwarning("任务执行中", "请先等待当前任务完成或手动终止。")
            return
        config = self.current_config()
        dialog = tk.Toplevel(self)
        dialog.title("一键执行主流程")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=self.palette.background)

        include_merge_var = tk.BooleanVar(value=config.workflow_mode == WorkflowMode.MERGED)
        include_phase_var = tk.BooleanVar(value=True)
        include_fill_var = tk.BooleanVar(value=False)
        include_stats_var = tk.BooleanVar(value=True)
        include_summary_var = tk.BooleanVar(value=True)
        include_ppt_var = tk.BooleanVar(value=True)

        wrap = ttk.Frame(dialog, style="Surface.TFrame", padding=16)
        wrap.pack(fill="both", expand=True)
        ttk.Label(wrap, text="当前批次主流程", style="SubHeader.TLabel").pack(anchor="w")
        ttk.Label(wrap, text=f"批次：{config.batch_name}", style="Muted.TLabel").pack(anchor="w", pady=(4, 0))
        ttk.Label(wrap, text=f"输入目录：{config.effective_input_dir()}", style="Muted.TLabel").pack(anchor="w", pady=(0, 12))
        ttk.Label(
            wrap,
            text="若包含“生成分项统计”，将只运行分项统计页当前已勾选的客群。",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(0, 12))

        options = ttk.LabelFrame(wrap, text="执行步骤", padding=12)
        options.pack(fill="x")
        if config.workflow_mode == WorkflowMode.MERGED:
            ttk.Checkbutton(options, text="先合并多月问卷", variable=include_merge_var).pack(anchor="w")
        ttk.Checkbutton(options, text="预处理：修正期次列", variable=include_phase_var).pack(anchor="w")
        ttk.Checkbutton(options, text="预处理：补写年份月份", variable=include_fill_var).pack(anchor="w")
        ttk.Checkbutton(options, text="生成分项统计", variable=include_stats_var).pack(anchor="w")
        ttk.Checkbutton(options, text="生成汇总表", variable=include_summary_var).pack(anchor="w")
        ttk.Checkbutton(options, text="生成PPT", variable=include_ppt_var).pack(anchor="w")

        footer = ttk.Frame(wrap, style="Surface.TFrame")
        footer.pack(fill="x", pady=(16, 0))
        footer.columnconfigure(0, weight=1)

        def start() -> None:
            selection = MainWorkflowSelection(
                include_merge=include_merge_var.get(),
                include_phase_preprocess=include_phase_var.get(),
                include_fill_year_month=include_fill_var.get(),
                include_survey_stats=include_stats_var.get(),
                include_summary=include_summary_var.get(),
                include_ppt=include_ppt_var.get(),
            )
            selected_stats_job_names: tuple[str, ...] = ()
            if selection.include_survey_stats:
                selected_stats_job_names = self._selected_stats_job_names(parent=dialog)
                if selected_stats_job_names is None:
                    return
            try:
                tasks = build_task_commands(
                    self.current_config(),
                    selection,
                    selected_stats_job_names=selected_stats_job_names,
                )
            except ValueError as exc:
                messagebox.showerror("参数不完整", str(exc), parent=dialog)
                return
            dialog.destroy()
            self.run_task_list(tasks, show_workflow_progress=True)

        ttk.Button(footer, text="取消", style="Secondary.TButton", command=dialog.destroy).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(footer, text="开始执行", style="Primary.TButton", command=start).grid(row=0, column=2)


def main() -> None:
    app = SurveyPlatformApp()
    app.mainloop()


if __name__ == "__main__":
    main()
