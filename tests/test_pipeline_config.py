from __future__ import annotations

import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from pipeline_config import load_pipeline_defaults


class PipelineConfigTest(unittest.TestCase):
    def test_project_pipeline_defaults_use_15pt_table_fonts(self) -> None:
        repo_root = Path(__file__).resolve().parents[1]
        defaults = load_pipeline_defaults(repo_root / "pipeline.defaults.toml")

        self.assertEqual(defaults.ppt.body_font_size_pt, 15.0)
        self.assertEqual(defaults.ppt.header_font_size_pt, 15.0)
        self.assertEqual(defaults.ppt.summary_font_size_pt, 15.0)

    def test_project_pipeline_defaults_use_balanced_llm_notes_settings(self) -> None:
        repo_root = Path(__file__).resolve().parents[1]
        defaults = load_pipeline_defaults(repo_root / "pipeline.defaults.toml")

        self.assertEqual(defaults.ppt.llm_notes.target_chars, 120)
        self.assertEqual(defaults.ppt.llm_notes.temperature, 0.4)
        self.assertEqual(defaults.ppt.llm_notes.max_tokens, 200)

    def test_project_pipeline_defaults_restore_category_intro_slides(self) -> None:
        repo_root = Path(__file__).resolve().parents[1]
        defaults = load_pipeline_defaults(repo_root / "pipeline.defaults.toml")

        self.assertIn("一、会展客户", defaults.ppt.category_intro_slides)
        self.assertIn("五、酒店客户", defaults.ppt.category_intro_slides)
        self.assertEqual(
            defaults.ppt.category_intro_slides["一、会展客户"].ppt_path,
            repo_root / "templates" / "chapter.pptx",
        )
        self.assertEqual(
            defaults.ppt.category_intro_slides["一、会展客户"].slide_number,
            3,
        )
        self.assertEqual(
            defaults.ppt.category_intro_slides["五、酒店客户"].slide_number,
            5,
        )

    def test_load_pipeline_defaults_reads_required_fields_and_resolves_relative_paths(self) -> None:
        with TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir)
            nested_dir = config_dir / "nested"
            nested_dir.mkdir()
            config_path = nested_dir / "pipeline.defaults.toml"
            config_path.write_text(
                "\n".join(
                    [
                        'sheet_name = "问卷数据"',
                        'calculation_mode = "template"',
                        'sample_config_path = "sample_table.default.toml"',
                        '',
                        '[ppt]',
                        'template_path = "templates/report.pptx"',
                        'sheet_name_mode = "first"',
                        'section_mode = "auto"',
                        'blank_display = "/"',
                        'title_suffix = "满意度"',
                        'max_single_table_rows = 18',
                        'max_split_table_rows = 19',
                        'sort_files = true',
                        'body_font_size_pt = 10.5',
                        'header_font_size_pt = 11.0',
                        'summary_font_size_pt = 12.0',
                        'template_slide_index = 0',
                        '',
                        '[ppt.chart_page]',
                        'enabled = true',
                        'placeholder_text = "图表分析内容待补充。\\n后续将在此处补充该客户分组二级指标的整体解读、优势项与待提升项。"',
                        'image_dpi = 220',
                        '',
                        '[ppt.llm_notes]',
                        'enabled = true',
                        'env_path = ".env"',
                        'system_role_path = "system_role.md"',
                        'target_chars = 300',
                        'temperature = 0.4',
                        'max_tokens = 500',
                        'checkpoint_chars = 80',
                        '',
                        '[ppt.category_intro_slides."一、会展客户"]',
                        'ppt_path = "templates/chapter.pptx"',
                        'slide_number = 3',
                    ]
                ),
                encoding="utf-8",
            )

            original_cwd = Path.cwd()
            try:
                os.chdir(config_dir)
                defaults = load_pipeline_defaults(Path("nested/pipeline.defaults.toml"))
            finally:
                os.chdir(original_cwd)

            resolved_nested_dir = nested_dir.resolve()
            self.assertEqual(defaults.sheet_name, "问卷数据")
            self.assertEqual(defaults.calculation_mode, "template")
            self.assertEqual(
                defaults.sample_config_path,
                resolved_nested_dir / "sample_table.default.toml",
            )
            self.assertEqual(
                defaults.ppt.template_path,
                resolved_nested_dir / "templates/report.pptx",
            )
            self.assertEqual(defaults.ppt.sheet_name_mode, "first")
            self.assertEqual(
                defaults.ppt.llm_notes.env_path,
                resolved_nested_dir / ".env",
            )
            self.assertEqual(
                defaults.ppt.llm_notes.system_role_path,
                resolved_nested_dir / "system_role.md",
            )
            self.assertEqual(
                defaults.ppt.category_intro_slides["一、会展客户"].ppt_path,
                resolved_nested_dir / "templates/chapter.pptx",
            )
            self.assertEqual(
                defaults.ppt.category_intro_slides["一、会展客户"].slide_number,
                3,
            )

    def test_load_pipeline_defaults_rejects_invalid_calculation_mode(self) -> None:
        with TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "pipeline.defaults.toml"
            config_path.write_text(
                "\n".join(
                    [
                        'sheet_name = "问卷数据"',
                        'calculation_mode = "invalid"',
                        'sample_config_path = "sample_table.default.toml"',
                        '',
                        '[ppt]',
                        'template_path = "templates/report.pptx"',
                        'sheet_name_mode = "first"',
                        'section_mode = "auto"',
                        'blank_display = ""',
                        'title_suffix = ""',
                        'max_single_table_rows = 18',
                        'max_split_table_rows = 19',
                        'sort_files = true',
                        'body_font_size_pt = 10.5',
                        'header_font_size_pt = 11.0',
                        'summary_font_size_pt = 12.0',
                        'template_slide_index = 0',
                        '',
                        '[ppt.chart_page]',
                        'enabled = true',
                        'placeholder_text = "图表分析内容待补充。后续将在此处补充该客户分组二级指标的整体解读、优势项与待提升项。"',
                        'image_dpi = 220',
                        '',
                        '[ppt.llm_notes]',
                        'enabled = true',
                        'env_path = ".env"',
                        'system_role_path = "system_role.md"',
                        'target_chars = 300',
                        'temperature = 0.6',
                        'max_tokens = 500',
                        'checkpoint_chars = 80',
                    ]
                ),
                encoding="utf-8",
            )

            with self.assertRaises(ValueError):
                load_pipeline_defaults(config_path)


if __name__ == "__main__":
    unittest.main()
