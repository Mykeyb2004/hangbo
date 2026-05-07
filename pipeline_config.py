from __future__ import annotations

import tomllib
from dataclasses import dataclass
from pathlib import Path

from generate_ppt import normalize_section_mode
from survey_stats import normalize_calculation_mode

DEFAULT_SHEET_NAME = "问卷数据"
DEFAULT_CALCULATION_MODE = "template"
DEFAULT_SAMPLE_CONFIG_PATH = "sample_table.default.toml"
DEFAULT_PPT_TEMPLATE_PATH = "templates/template.pptx"
DEFAULT_PPT_SHEET_NAME_MODE = "first"
DEFAULT_PPT_SECTION_MODE = "auto"
DEFAULT_PPT_BLANK_DISPLAY = ""
DEFAULT_PPT_TITLE_SUFFIX = ""
DEFAULT_PPT_MAX_SINGLE_TABLE_ROWS = 18
DEFAULT_PPT_MAX_SPLIT_TABLE_ROWS = 19
DEFAULT_PPT_SORT_FILES = True
DEFAULT_PPT_BODY_FONT_SIZE_PT = 14.0
DEFAULT_PPT_HEADER_FONT_SIZE_PT = 14.0
DEFAULT_PPT_SUMMARY_FONT_SIZE_PT = 14.0
DEFAULT_PPT_TEMPLATE_SLIDE_INDEX = 0
DEFAULT_CHART_PAGE_ENABLED = True
DEFAULT_CHART_PAGE_PLACEHOLDER_TEXT = (
    "图表分析内容待补充。后续将在此处补充该客户分组二级指标的整体解读、优势项与待提升项。"
)
DEFAULT_CHART_PAGE_IMAGE_DPI = 220
DEFAULT_LLM_NOTES_ENABLED = True
DEFAULT_LLM_NOTES_ENV_PATH = ".env"
DEFAULT_LLM_NOTES_SYSTEM_ROLE_PATH = "system_role.md"
DEFAULT_LLM_NOTES_TARGET_CHARS = 120
DEFAULT_LLM_NOTES_TEMPERATURE = 0.6
DEFAULT_LLM_NOTES_MAX_TOKENS = 200
DEFAULT_LLM_NOTES_CHECKPOINT_CHARS = 80
DEFAULT_LLM_NOTES_HIGHLIGHT_THRESHOLD = 9.6


@dataclass(frozen=True)
class PipelineCategoryIntroSlideDefaults:
    ppt_path: Path
    slide_number: int


@dataclass(frozen=True)
class PipelineLlmNotesDefaults:
    enabled: bool
    env_path: Path
    system_role_path: Path
    target_chars: int
    temperature: float
    max_tokens: int
    checkpoint_chars: int
    highlight_threshold: float


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
    category_intro_slides: dict[str, PipelineCategoryIntroSlideDefaults]
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


def load_category_intro_slides_defaults(
    config_dir: Path,
    raw: object,
) -> dict[str, PipelineCategoryIntroSlideDefaults]:
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ValueError("ppt.category_intro_slides 必须是对象")

    intro_slides: dict[str, PipelineCategoryIntroSlideDefaults] = {}
    for category_label, item in raw.items():
        if not isinstance(category_label, str) or not category_label.strip():
            raise ValueError("ppt.category_intro_slides 的键必须是非空客户大类名称")
        if not isinstance(item, dict):
            raise ValueError(f"{category_label} 的章节页配置必须是对象")
        if "ppt_path" not in item:
            raise ValueError(f"{category_label} 缺少 ppt_path")
        if "slide_number" not in item:
            raise ValueError(f"{category_label} 缺少 slide_number")

        slide_number = int(item["slide_number"])
        if slide_number < 1:
            raise ValueError(f"{category_label} 的 slide_number 必须从 1 开始")

        intro_slides[category_label] = PipelineCategoryIntroSlideDefaults(
            ppt_path=resolve_config_path(config_dir, item["ppt_path"]),
            slide_number=slide_number,
        )
    return intro_slides


def load_pipeline_defaults(
    config_path: Path = Path("pipeline.defaults.toml"),
) -> PipelineDefaults:
    resolved_config_path = config_path.resolve()
    raw = tomllib.loads(resolved_config_path.read_text(encoding="utf-8"))
    config_dir = resolved_config_path.parent
    calculation_mode = normalize_calculation_mode(
        raw.get("calculation_mode", DEFAULT_CALCULATION_MODE)
    )
    ppt_raw = raw.get("ppt", {})
    chart_page_raw = ppt_raw.get("chart_page", {})
    llm_notes_raw = ppt_raw.get("llm_notes", {})
    category_intro_slides = load_category_intro_slides_defaults(
        config_dir,
        ppt_raw.get("category_intro_slides"),
    )

    return PipelineDefaults(
        sheet_name=str(raw.get("sheet_name", DEFAULT_SHEET_NAME)),
        calculation_mode=calculation_mode,
        sample_config_path=resolve_config_path(
            config_dir,
            raw.get("sample_config_path", DEFAULT_SAMPLE_CONFIG_PATH),
        ),
        ppt=PipelinePptDefaults(
            template_path=resolve_config_path(
                config_dir,
                ppt_raw.get("template_path", DEFAULT_PPT_TEMPLATE_PATH),
            ),
            sheet_name_mode=str(
                ppt_raw.get("sheet_name_mode", DEFAULT_PPT_SHEET_NAME_MODE)
            ),
            section_mode=normalize_section_mode(
                ppt_raw.get("section_mode", DEFAULT_PPT_SECTION_MODE)
            ),
            blank_display=str(ppt_raw.get("blank_display", DEFAULT_PPT_BLANK_DISPLAY)),
            title_suffix=str(ppt_raw.get("title_suffix", DEFAULT_PPT_TITLE_SUFFIX)),
            max_single_table_rows=int(
                ppt_raw.get(
                    "max_single_table_rows",
                    DEFAULT_PPT_MAX_SINGLE_TABLE_ROWS,
                )
            ),
            max_split_table_rows=int(
                ppt_raw.get(
                    "max_split_table_rows",
                    DEFAULT_PPT_MAX_SPLIT_TABLE_ROWS,
                )
            ),
            sort_files=bool(ppt_raw.get("sort_files", DEFAULT_PPT_SORT_FILES)),
            body_font_size_pt=float(
                ppt_raw.get("body_font_size_pt", DEFAULT_PPT_BODY_FONT_SIZE_PT)
            ),
            header_font_size_pt=float(
                ppt_raw.get("header_font_size_pt", DEFAULT_PPT_HEADER_FONT_SIZE_PT)
            ),
            summary_font_size_pt=float(
                ppt_raw.get("summary_font_size_pt", DEFAULT_PPT_SUMMARY_FONT_SIZE_PT)
            ),
            template_slide_index=int(
                ppt_raw.get(
                    "template_slide_index",
                    DEFAULT_PPT_TEMPLATE_SLIDE_INDEX,
                )
            ),
            category_intro_slides=category_intro_slides,
            chart_page=PipelineChartPageDefaults(
                enabled=bool(
                    chart_page_raw.get("enabled", DEFAULT_CHART_PAGE_ENABLED)
                ),
                placeholder_text=str(
                    chart_page_raw.get(
                        "placeholder_text",
                        DEFAULT_CHART_PAGE_PLACEHOLDER_TEXT,
                    )
                ),
                image_dpi=int(
                    chart_page_raw.get("image_dpi", DEFAULT_CHART_PAGE_IMAGE_DPI)
                ),
            ),
            llm_notes=PipelineLlmNotesDefaults(
                enabled=bool(llm_notes_raw.get("enabled", DEFAULT_LLM_NOTES_ENABLED)),
                env_path=resolve_config_path(
                    config_dir,
                    llm_notes_raw.get("env_path", DEFAULT_LLM_NOTES_ENV_PATH),
                ),
                system_role_path=resolve_config_path(
                    config_dir,
                    llm_notes_raw.get(
                        "system_role_path",
                        DEFAULT_LLM_NOTES_SYSTEM_ROLE_PATH,
                    ),
                ),
                target_chars=int(
                    llm_notes_raw.get("target_chars", DEFAULT_LLM_NOTES_TARGET_CHARS)
                ),
                temperature=float(
                    llm_notes_raw.get("temperature", DEFAULT_LLM_NOTES_TEMPERATURE)
                ),
                max_tokens=int(
                    llm_notes_raw.get("max_tokens", DEFAULT_LLM_NOTES_MAX_TOKENS)
                ),
                checkpoint_chars=int(
                    llm_notes_raw.get(
                        "checkpoint_chars",
                        DEFAULT_LLM_NOTES_CHECKPOINT_CHARS,
                    )
                ),
                highlight_threshold=float(
                    llm_notes_raw.get(
                        "highlight_threshold",
                        DEFAULT_LLM_NOTES_HIGHLIGHT_THRESHOLD,
                    )
                ),
            ),
        ),
    )
