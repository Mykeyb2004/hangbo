from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from pptx import Presentation

from generate_ppt import (
    PptBatchConfig,
    PptLayoutConfig,
    TableRegion,
    build_section_blocks,
    choose_detail_layout,
    format_report_value,
    generate_presentation,
    resolve_section_definition,
)


def create_report_workbook(path: Path, rows: list[tuple[object, object, object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = path.stem
    for row in rows:
        worksheet.append(list(row))
    workbook.save(path)


class GeneratePptTest(unittest.TestCase):
    def test_build_section_blocks_groups_rows_by_second_level_titles(self) -> None:
        rows = [
            ("专业观众", 9.93, 10.0),
            ("会展服务", 10.0, 10.0),
            ("工作人员仪容仪表", 10.0, 10.0),
            ("工作人员服务态度", 10.0, 10.0),
            ("硬件设施", 9.91, 10.0),
            ("展会路线安排", 9.93, 10.0),
        ]

        role_definition = resolve_section_definition("专业观众", rows)
        section_blocks = build_section_blocks(rows, role_definition)

        self.assertEqual([block.heading for block in section_blocks], ["会展服务", "硬件设施"])
        self.assertEqual([len(block.rows) for block in section_blocks], [3, 2])

    def test_choose_detail_layout_splits_into_two_tables_without_breaking_sections(self) -> None:
        rows = [
            ("专业观众", 9.93, 10.0),
            ("会展服务", 10.0, 10.0),
            ("工作人员仪容仪表", 10.0, 10.0),
            ("工作人员服务态度", 10.0, 10.0),
            ("工作人员业务技能", 10.0, 10.0),
            ("接待引导服务", 10.0, 10.0),
            ("硬件设施", 9.91, 10.0),
            ("展会路线安排", 9.93, 10.0),
            ("园区停车方便", 10.0, 10.0),
            ("交通便利，容易到达", 10.0, 10.0),
            ("标识标牌清晰", 9.73, 10.0),
            ("设施设备齐全", 9.89, 10.0),
            ("展厅使用情况", 9.84, 10.0),
            ("参展环境", 10.0, 10.0),
            ("配套服务", 9.8, 10.0),
            ("餐饮服务", 9.4, 10.0),
            ("客房服务", None, None),
            ("安保服务", 10.0, 10.0),
            ("保洁服务", 10.0, 10.0),
            ("智慧场馆", 10.0, 10.0),
            ("杭州国博APP", 10.0, 10.0),
            ("室内导航系统", None, None),
            ("寻车系统", None, None),
            ("云上看馆", None, None),
        ]

        role_definition = resolve_section_definition("专业观众", rows)
        detail_layout = choose_detail_layout(
            detail_rows=rows[1:],
            role_definition=role_definition,
            max_single_table_rows=18,
            max_split_table_rows=19,
        )

        self.assertTrue(detail_layout.is_split)
        self.assertEqual(
            [block.heading for block in detail_layout.left_blocks],
            ["会展服务", "硬件设施"],
        )
        self.assertEqual(
            [block.heading for block in detail_layout.right_blocks],
            ["配套服务", "智慧场馆"],
        )

    def test_format_report_value_hides_empty_values(self) -> None:
        self.assertEqual(format_report_value(None, blank_display=""), "")
        self.assertEqual(format_report_value(9.50, blank_display=""), "9.5")

    def test_generate_presentation_creates_single_and_double_table_slides(self) -> None:
        repo_root = Path(__file__).resolve().parents[1]
        template_path = repo_root / "templates" / "template.pptx"

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_dir = temp_path / "input"
            input_dir.mkdir()
            output_path = temp_path / "report.pptx"

            create_report_workbook(
                input_dir / "专业观众.xlsx",
                [
                    ("指标", "满意度", "重要性"),
                    ("专业观众", 9.93, 10.0),
                    ("会展服务", 10.0, 10.0),
                    ("工作人员仪容仪表", 10.0, 10.0),
                    ("工作人员服务态度", 10.0, 10.0),
                    ("工作人员业务技能", 10.0, 10.0),
                    ("接待引导服务", 10.0, 10.0),
                    ("硬件设施", 9.91, 10.0),
                    ("展会路线安排", 9.93, 10.0),
                    ("园区停车方便", 10.0, 10.0),
                    ("交通便利，容易到达", 10.0, 10.0),
                    ("标识标牌清晰", 9.73, 10.0),
                    ("设施设备齐全", 9.89, 10.0),
                    ("展厅使用情况", 9.84, 10.0),
                    ("参展环境", 10.0, 10.0),
                    ("配套服务", 9.8, 10.0),
                    ("餐饮服务", 9.4, 10.0),
                    ("客房服务", None, None),
                    ("安保服务", 10.0, 10.0),
                    ("保洁服务", 10.0, 10.0),
                    ("智慧场馆", 10.0, 10.0),
                    ("杭州国博APP", 10.0, 10.0),
                    ("室内导航系统", None, None),
                    ("寻车系统", None, None),
                    ("云上看馆", None, None),
                ],
            )
            create_report_workbook(
                input_dir / "自助餐.xlsx",
                [
                    ("指标", "满意度", "重要性"),
                    ("自助餐", 9.80, 9.60),
                    ("餐饮服务", 9.85, 9.72),
                    ("菜品口味", 9.90, 9.60),
                    ("菜品丰富度", 9.80, 9.60),
                    ("补菜及时性", None, None),
                    ("硬件设施", 9.70, 9.40),
                    ("环境卫生", 9.70, 9.40),
                    ("桌椅舒适度", 9.70, 9.40),
                ],
            )

            config = PptBatchConfig(
                template_path=template_path,
                input_dir=input_dir,
                output_ppt=output_path,
                blank_display="",
                max_single_table_rows=10,
                max_split_table_rows=19,
                layout=PptLayoutConfig(),
            )

            generate_presentation(config)

            self.assertTrue(output_path.exists())

            presentation = Presentation(output_path)
            self.assertEqual(len(presentation.slides), 2)

            slide_tables = {}
            for slide in presentation.slides:
                title = slide.shapes.title.text
                table_count = sum(1 for shape in slide.shapes if getattr(shape, "has_table", False))
                table_texts = [
                    "\n".join(
                        cell.text
                        for row in shape.table.rows
                        for cell in row.cells
                    )
                    for shape in slide.shapes
                    if getattr(shape, "has_table", False)
                ]
                slide_tables[title] = (table_count, table_texts)

            self.assertEqual(slide_tables["专业观众"][0], 3)
            self.assertEqual(slide_tables["自助餐"][0], 2)
            self.assertTrue(any("会展服务" in text for text in slide_tables["专业观众"][1]))
            self.assertTrue(any("智慧场馆" in text for text in slide_tables["专业观众"][1]))
            self.assertTrue(any("补菜及时性" in text for text in slide_tables["自助餐"][1]))
            self.assertTrue(any("\n\n" in text or text.endswith("\n") for text in slide_tables["专业观众"][1]))


if __name__ == "__main__":
    unittest.main()
