from __future__ import annotations

import argparse
import math
import tomllib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from survey_stats import (
    RoleDefinition,
    TEMPLATE_DEFINITIONS,
    format_value,
    get_effective_role_definition,
)

VALID_SECTION_MODES = ("auto", "template", "summary")
DEFAULT_FILE_PATTERN = "*.xlsx"
DEFAULT_SHEET_NAME_MODE = "first"
DEFAULT_SECTION_MODE = "auto"

HEADER_FILL_COLOR = "D9E2F3"
OVERALL_FILL_COLOR = "F4B183"
SECTION_FILL_COLOR = "C6EFD1"
BODY_FILL_COLOR = "FFFFFF"
BORDER_COLOR = "BFBFBF"
HEADER_TEXT_COLOR = "1F1F1F"
BODY_TEXT_COLOR = "1F1F1F"


@dataclass(frozen=True)
class TableRegion:
    left: float
    top: float
    width: float
    height: float

    def emu(self) -> tuple[int, int, int, int]:
        return (
            Inches(self.left),
            Inches(self.top),
            Inches(self.width),
            Inches(self.height),
        )


@dataclass(frozen=True)
class PptLayoutConfig:
    summary_table: TableRegion = field(
        default_factory=lambda: TableRegion(0.73, 1.45, 11.87, 0.56)
    )
    detail_single_table: TableRegion = field(
        default_factory=lambda: TableRegion(0.73, 2.10, 11.87, 4.95)
    )
    detail_left_table: TableRegion = field(
        default_factory=lambda: TableRegion(0.73, 2.10, 5.78, 4.95)
    )
    detail_right_table: TableRegion = field(
        default_factory=lambda: TableRegion(6.82, 2.10, 5.78, 4.95)
    )


@dataclass(frozen=True)
class PptBatchConfig:
    template_path: Path
    input_dir: Path
    output_ppt: Path
    file_pattern: str = DEFAULT_FILE_PATTERN
    sheet_name_mode: str = DEFAULT_SHEET_NAME_MODE
    sheet_name: str | None = None
    blank_display: str = ""
    title_suffix: str = ""
    section_mode: str = DEFAULT_SECTION_MODE
    max_single_table_rows: int = 18
    max_split_table_rows: int = 19
    sort_files: bool = True
    layout: PptLayoutConfig = field(default_factory=PptLayoutConfig)
    body_font_size_pt: float = 10.5
    header_font_size_pt: float = 11.0
    summary_font_size_pt: float = 12.0
    template_slide_index: int = 0


@dataclass(frozen=True)
class SectionBlock:
    heading: str
    rows: tuple[tuple[str, float | None, float | None], ...]

    @property
    def row_count(self) -> int:
        return len(self.rows)


@dataclass(frozen=True)
class DetailLayout:
    is_split: bool
    single_rows: tuple[tuple[str, float | None, float | None], ...] = ()
    left_blocks: tuple[SectionBlock, ...] = ()
    right_blocks: tuple[SectionBlock, ...] = ()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="根据 Excel 批量生成 PPT")
    parser.add_argument("--config", type=Path, help="TOML 配置文件路径")
    parser.add_argument("--template-path", type=Path, help="PPT 模板路径")
    parser.add_argument("--input-dir", type=Path, help="Excel 输入目录")
    parser.add_argument("--output-ppt", type=Path, help="输出 PPT 路径")
    parser.add_argument(
        "--section-mode",
        choices=VALID_SECTION_MODES,
        help="二级标题识别口径：auto/template/summary",
    )
    parser.add_argument("--blank-display", help="空值显示文本，默认空字符串")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只校验输入和布局，不写出 PPT",
    )
    return parser


def load_batch_config(
    config_path: Path,
    *,
    template_path: Path | None = None,
    input_dir: Path | None = None,
    output_ppt: Path | None = None,
    section_mode: str | None = None,
    blank_display: str | None = None,
) -> PptBatchConfig:
    raw = tomllib.loads(config_path.read_text(encoding="utf-8"))
    layout = load_layout_config(raw.get("layout", {}))

    effective_section_mode = normalize_section_mode(section_mode or raw.get("section_mode"))
    return PptBatchConfig(
        template_path=Path(template_path or raw["template_path"]),
        input_dir=Path(input_dir or raw["input_dir"]),
        output_ppt=Path(output_ppt or raw["output_ppt"]),
        file_pattern=str(raw.get("file_pattern", DEFAULT_FILE_PATTERN)),
        sheet_name_mode=str(raw.get("sheet_name_mode", DEFAULT_SHEET_NAME_MODE)),
        sheet_name=raw.get("sheet_name"),
        blank_display=blank_display if blank_display is not None else str(raw.get("blank_display", "")),
        title_suffix=str(raw.get("title_suffix", "")),
        section_mode=effective_section_mode,
        max_single_table_rows=int(raw.get("max_single_table_rows", 18)),
        max_split_table_rows=int(raw.get("max_split_table_rows", 19)),
        sort_files=bool(raw.get("sort_files", True)),
        layout=layout,
        body_font_size_pt=float(raw.get("body_font_size_pt", 10.5)),
        header_font_size_pt=float(raw.get("header_font_size_pt", 11.0)),
        summary_font_size_pt=float(raw.get("summary_font_size_pt", 12.0)),
        template_slide_index=int(raw.get("template_slide_index", 0)),
    )


def load_layout_config(raw: dict[str, object]) -> PptLayoutConfig:
    return PptLayoutConfig(
        summary_table=load_table_region(
            raw.get("summary_table"),
            TableRegion(0.73, 1.45, 11.87, 0.56),
        ),
        detail_single_table=load_table_region(
            raw.get("detail_single_table"),
            TableRegion(0.73, 2.10, 11.87, 4.95),
        ),
        detail_left_table=load_table_region(
            raw.get("detail_left_table"),
            TableRegion(0.73, 2.10, 5.78, 4.95),
        ),
        detail_right_table=load_table_region(
            raw.get("detail_right_table"),
            TableRegion(6.82, 2.10, 5.78, 4.95),
        ),
    )


def load_table_region(raw: object, defaults: TableRegion) -> TableRegion:
    if raw is None:
        return defaults
    if not isinstance(raw, dict):
        raise ValueError("layout 下的表格区域必须是对象")
    return TableRegion(
        left=float(raw.get("left", defaults.left)),
        top=float(raw.get("top", defaults.top)),
        width=float(raw.get("width", defaults.width)),
        height=float(raw.get("height", defaults.height)),
    )


def normalize_section_mode(section_mode: str | None) -> str:
    normalized = str(section_mode or DEFAULT_SECTION_MODE).strip().lower()
    if normalized not in VALID_SECTION_MODES:
        raise ValueError(f"section_mode 仅支持: {', '.join(VALID_SECTION_MODES)}")
    return normalized


def discover_input_files(config: PptBatchConfig) -> list[Path]:
    if not config.input_dir.exists():
        raise FileNotFoundError(f"输入目录不存在: {config.input_dir}")
    files = list(config.input_dir.glob(config.file_pattern))
    files = [path for path in files if path.is_file()]
    if config.sort_files:
        files.sort(key=lambda path: path.name)
    if not files:
        raise FileNotFoundError(
            f"{config.input_dir} 下没有匹配 {config.file_pattern} 的 Excel 文件"
        )
    return files


def read_report_rows(
    workbook_path: Path,
    *,
    sheet_name_mode: str = DEFAULT_SHEET_NAME_MODE,
    sheet_name: str | None = None,
) -> list[tuple[str, float | None, float | None]]:
    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        if sheet_name_mode == "first":
            worksheet = workbook[workbook.sheetnames[0]]
        elif sheet_name_mode == "named":
            if not sheet_name:
                raise ValueError("sheet_name_mode=named 时必须提供 sheet_name")
            worksheet = workbook[sheet_name]
        else:
            raise ValueError("sheet_name_mode 仅支持 first 或 named")

        rows = list(worksheet.iter_rows(values_only=True))
    finally:
        workbook.close()

    if not rows:
        raise ValueError(f"{workbook_path} 为空工作簿")

    header = tuple("" if value is None else str(value).strip() for value in rows[0][:3])
    if header[:3] != ("指标", "满意度", "重要性"):
        raise ValueError(
            f"{workbook_path} 表头不符合预期，需为 指标/满意度/重要性，实际为: {header}"
        )

    report_rows: list[tuple[str, float | None, float | None]] = []
    for raw_row in rows[1:]:
        label = "" if raw_row[0] is None else str(raw_row[0]).strip()
        if not label:
            continue
        report_rows.append((label, to_optional_float(raw_row[1]), to_optional_float(raw_row[2])))

    if not report_rows:
        raise ValueError(f"{workbook_path} 没有可用数据行")
    return report_rows


def to_optional_float(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, float):
        if math.isnan(value):
            return None
        return value
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def resolve_section_definition(
    role_name: str,
    rows: Sequence[tuple[str, float | None, float | None]],
    *,
    section_mode: str = DEFAULT_SECTION_MODE,
) -> RoleDefinition | None:
    base_definition = next(
        (
            template
            for template in TEMPLATE_DEFINITIONS.values()
            if template.role_name == role_name
        ),
        None,
    )
    if base_definition is None:
        return None

    normalized_mode = normalize_section_mode(section_mode)
    if normalized_mode == "template":
        return get_effective_role_definition(base_definition, "template")
    if normalized_mode == "summary":
        return get_effective_role_definition(base_definition, "summary")

    candidates: list[RoleDefinition] = []
    for mode in ("template", "summary"):
        candidate = get_effective_role_definition(base_definition, mode)
        if candidate.sections not in [existing.sections for existing in candidates]:
            candidates.append(candidate)

    labels = {row[0] for row in rows}
    return max(
        candidates,
        key=lambda definition: sum(
            1 for section in definition.sections if section.name in labels
        ),
    )


def build_section_blocks(
    rows: Sequence[tuple[str, float | None, float | None]],
    role_definition: RoleDefinition | None,
) -> list[SectionBlock]:
    if not rows:
        return []

    start_index = 1 if role_definition and rows[0][0] == role_definition.role_name else 0
    detail_rows = rows[start_index:]
    if not detail_rows:
        return []

    if role_definition is None:
        return [SectionBlock(heading=detail_rows[0][0], rows=tuple(detail_rows))]

    section_names = {section.name for section in role_definition.sections}
    blocks: list[SectionBlock] = []
    current_rows: list[tuple[str, float | None, float | None]] = []
    current_heading: str | None = None

    for row in detail_rows:
        label = row[0]
        if label in section_names:
            if current_rows:
                blocks.append(SectionBlock(heading=current_heading or current_rows[0][0], rows=tuple(current_rows)))
            current_heading = label
            current_rows = [row]
            continue

        if not current_rows:
            current_heading = current_heading or label
            current_rows = [row]
            continue

        current_rows.append(row)

    if current_rows:
        blocks.append(SectionBlock(heading=current_heading or current_rows[0][0], rows=tuple(current_rows)))

    return blocks


def choose_detail_layout(
    *,
    detail_rows: Sequence[tuple[str, float | None, float | None]],
    role_definition: RoleDefinition | None,
    max_single_table_rows: int,
    max_split_table_rows: int,
) -> DetailLayout:
    if len(detail_rows) <= max_single_table_rows:
        return DetailLayout(is_split=False, single_rows=tuple(detail_rows))

    section_blocks = build_section_blocks(detail_rows, role_definition)
    if len(section_blocks) < 2:
        raise ValueError("数据超出单表容量，但无法按二级标题拆分为左右双表")

    candidates: list[tuple[int, int, int]] = []
    for split_index in range(1, len(section_blocks)):
        left_count = sum(block.row_count for block in section_blocks[:split_index])
        right_count = sum(block.row_count for block in section_blocks[split_index:])
        if left_count <= max_split_table_rows and right_count <= max_split_table_rows:
            candidates.append((abs(left_count - right_count), max(left_count, right_count), split_index))

    if not candidates:
        raise ValueError(
            "按二级标题拆分后仍超出左右双表容量，请调大 max_split_table_rows 或调整模板布局"
        )

    _, _, split_index = min(candidates)
    return DetailLayout(
        is_split=True,
        left_blocks=tuple(section_blocks[:split_index]),
        right_blocks=tuple(section_blocks[split_index:]),
    )


def flatten_blocks(blocks: Sequence[SectionBlock]) -> list[tuple[str, float | None, float | None]]:
    rows: list[tuple[str, float | None, float | None]] = []
    for block in blocks:
        rows.extend(block.rows)
    return rows


def format_report_value(value: float | None, *, blank_display: str = "") -> str:
    if value is None:
        return blank_display
    formatted = format_value(value)
    return formatted if formatted != "" else blank_display


def ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def find_title_shape(slide):
    if slide.shapes.title is not None:
        return slide.shapes.title
    for shape in slide.shapes:
        if getattr(shape, "is_placeholder", False):
            placeholder_type = str(shape.placeholder_format.type)
            if "TITLE" in placeholder_type:
                return shape
    raise ValueError("模板页缺少标题占位符")


def apply_title(slide, title: str) -> None:
    title_shape = find_title_shape(slide)
    title_shape.text = title
    if title_shape.has_text_frame:
        title_shape.text_frame.word_wrap = True


def generate_presentation(config: PptBatchConfig, *, dry_run: bool = False) -> Path:
    files = discover_input_files(config)
    presentation = Presentation(str(config.template_path))

    if config.template_slide_index >= len(presentation.slides):
        raise IndexError("template_slide_index 超出模板页数量")

    template_slide = presentation.slides[config.template_slide_index]
    template_layout = template_slide.slide_layout

    for index, workbook_path in enumerate(files):
        slide = template_slide if index == 0 else presentation.slides.add_slide(template_layout)
        render_workbook_slide(slide, workbook_path, config)

    if dry_run:
        return config.output_ppt

    ensure_parent_dir(config.output_ppt)
    presentation.save(str(config.output_ppt))
    return config.output_ppt


def render_workbook_slide(slide, workbook_path: Path, config: PptBatchConfig) -> None:
    title = workbook_path.stem + config.title_suffix
    apply_title(slide, title)

    report_rows = read_report_rows(
        workbook_path,
        sheet_name_mode=config.sheet_name_mode,
        sheet_name=config.sheet_name,
    )
    overall_row = report_rows[0]
    detail_rows = report_rows[1:]
    role_definition = resolve_section_definition(
        workbook_path.stem,
        report_rows,
        section_mode=config.section_mode,
    )

    render_table(
        slide,
        config.layout.summary_table,
        [overall_row],
        blank_display=config.blank_display,
        section_names=set(),
        overall_label=overall_row[0],
        header_font_size_pt=config.header_font_size_pt,
        body_font_size_pt=config.summary_font_size_pt,
    )

    if not detail_rows:
        return

    detail_layout = choose_detail_layout(
        detail_rows=detail_rows,
        role_definition=role_definition,
        max_single_table_rows=config.max_single_table_rows,
        max_split_table_rows=config.max_split_table_rows,
    )

    if detail_layout.is_split:
        left_section_names = {block.heading for block in detail_layout.left_blocks}
        right_section_names = {block.heading for block in detail_layout.right_blocks}
        render_table(
            slide,
            config.layout.detail_left_table,
            flatten_blocks(detail_layout.left_blocks),
            blank_display=config.blank_display,
            section_names=left_section_names,
            overall_label=None,
            header_font_size_pt=config.header_font_size_pt,
            body_font_size_pt=config.body_font_size_pt,
        )
        render_table(
            slide,
            config.layout.detail_right_table,
            flatten_blocks(detail_layout.right_blocks),
            blank_display=config.blank_display,
            section_names=right_section_names,
            overall_label=None,
            header_font_size_pt=config.header_font_size_pt,
            body_font_size_pt=config.body_font_size_pt,
        )
        return

    section_names = (
        {section.name for section in role_definition.sections} if role_definition else set()
    )
    render_table(
        slide,
        config.layout.detail_single_table,
        list(detail_layout.single_rows),
        blank_display=config.blank_display,
        section_names=section_names,
        overall_label=None,
        header_font_size_pt=config.header_font_size_pt,
        body_font_size_pt=config.body_font_size_pt,
    )


def render_table(
    slide,
    region: TableRegion,
    rows: Sequence[tuple[str, float | None, float | None]],
    *,
    blank_display: str,
    section_names: set[str],
    overall_label: str | None,
    header_font_size_pt: float,
    body_font_size_pt: float,
) -> None:
    left, top, width, height = region.emu()
    shape = slide.shapes.add_table(len(rows) + 1, 3, left, top, width, height)
    table = shape.table
    set_column_widths(table, width)

    headers = ("指标", "满意度", "重要性")
    for column_index, header in enumerate(headers):
        cell = table.cell(0, column_index)
        set_cell_text(
            cell,
            header,
            bold=True,
            align=PP_ALIGN.CENTER,
            font_size_pt=header_font_size_pt,
            fill_color=HEADER_FILL_COLOR,
            text_color=HEADER_TEXT_COLOR,
        )

    for row_index, row in enumerate(rows, start=1):
        label, satisfaction, importance = row
        if overall_label and label == overall_label:
            fill_color = OVERALL_FILL_COLOR
            bold = True
        elif label in section_names:
            fill_color = SECTION_FILL_COLOR
            bold = True
        else:
            fill_color = BODY_FILL_COLOR
            bold = False

        values = (
            label,
            format_report_value(satisfaction, blank_display=blank_display),
            format_report_value(importance, blank_display=blank_display),
        )
        for column_index, value in enumerate(values):
            cell = table.cell(row_index, column_index)
            set_cell_text(
                cell,
                value,
                bold=bold,
                align=PP_ALIGN.LEFT if column_index == 0 else PP_ALIGN.CENTER,
                font_size_pt=body_font_size_pt,
                fill_color=fill_color,
                text_color=BODY_TEXT_COLOR,
            )


def set_column_widths(table, total_width: int) -> None:
    first_col = int(total_width * 0.62)
    second_col = int(total_width * 0.19)
    third_col = total_width - first_col - second_col
    table.columns[0].width = first_col
    table.columns[1].width = second_col
    table.columns[2].width = third_col


def set_cell_text(
    cell,
    text: str,
    *,
    bold: bool,
    align,
    font_size_pt: float,
    fill_color: str,
    text_color: str,
) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor.from_string(fill_color)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Pt(4)
    cell.margin_right = Pt(4)
    cell.margin_top = Pt(2)
    cell.margin_bottom = Pt(2)

    text_frame = cell.text_frame
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.bold = bold
    run.font.size = Pt(font_size_pt)
    run.font.color.rgb = RGBColor.from_string(text_color)


def build_default_config_from_args(args: argparse.Namespace) -> PptBatchConfig:
    if not args.template_path or not args.input_dir or not args.output_ppt:
        raise ValueError("未提供 --config 时，必须同时提供 --template-path / --input-dir / --output-ppt")
    return PptBatchConfig(
        template_path=args.template_path,
        input_dir=args.input_dir,
        output_ppt=args.output_ppt,
        blank_display=args.blank_display or "",
        section_mode=normalize_section_mode(args.section_mode),
    )


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.config:
        config = load_batch_config(
            args.config,
            template_path=args.template_path,
            input_dir=args.input_dir,
            output_ppt=args.output_ppt,
            section_mode=args.section_mode,
            blank_display=args.blank_display,
        )
    else:
        config = build_default_config_from_args(args)

    output_path = generate_presentation(config, dry_run=args.dry_run)
    if args.dry_run:
        print(f"校验完成，未写出文件: {output_path}")
    else:
        print(f"PPT 已生成: {output_path}")


if __name__ == "__main__":
    main()
