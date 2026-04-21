from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from pipeline_config import (
    PipelineChartPageDefaults,
    PipelineDefaults,
    PipelineLlmNotesDefaults,
    PipelinePptDefaults,
)
from generate_ppt import PptBatchConfig
from pipeline_models import PipelineIssue, PrecheckResult
from pipeline_paths import build_pipeline_paths
from pipeline_runtime import run_pipeline, wait_for_confirmation


def build_defaults() -> PipelineDefaults:
    return PipelineDefaults(
        sheet_name="问卷数据",
        calculation_mode="template",
        sample_config_path=Path("sample_table.default.toml"),
        ppt=PipelinePptDefaults(
            template_path=Path("templates/template.pptx"),
            sheet_name_mode="first",
            section_mode="auto",
            blank_display="",
            title_suffix="",
            max_single_table_rows=18,
            max_split_table_rows=19,
            sort_files=True,
            body_font_size_pt=10.5,
            header_font_size_pt=11.0,
            summary_font_size_pt=12.0,
            template_slide_index=0,
            chart_page=PipelineChartPageDefaults(
                True,
                "图表分析内容待补充。",
                220,
            ),
            llm_notes=PipelineLlmNotesDefaults(
                False,
                Path(".env"),
                Path("system_role.md"),
                300,
                0.4,
                500,
                80,
            ),
        ),
    )


class PipelineRuntimeTest(unittest.TestCase):
    def test_wait_for_confirmation_reprompts_until_y_yes_or_continue(self) -> None:
        prompts = iter(["不是", "y"])
        outputs: list[str] = []

        wait_for_confirmation(
            input_func=lambda prompt: next(prompts),
            output_func=outputs.append,
        )

        self.assertEqual(outputs, ["未识别的输入：不是"])

    @mock.patch("pipeline_runtime.generate_presentation")
    @mock.patch("pipeline_runtime.generate_sample_table_report")
    @mock.patch("pipeline_runtime.generate_summary_report")
    @mock.patch("pipeline_runtime.run_directory_batch")
    @mock.patch("pipeline_runtime.apply_year_month_to_directory")
    @mock.patch("pipeline_runtime.run_precheck")
    def test_run_pipeline_rechecks_after_confirmation_before_running_engines(
        self,
        mock_run_precheck: mock.Mock,
        mock_apply_year_month: mock.Mock,
        mock_run_directory_batch: mock.Mock,
        mock_generate_summary: mock.Mock,
        mock_generate_sample: mock.Mock,
        mock_generate_presentation: mock.Mock,
    ) -> None:
        blocking_issue = PipelineIssue("blocking", "unmapped", "存在未映射标签")
        mock_run_precheck.side_effect = [
            PrecheckResult((blocking_issue,), (), False),
            PrecheckResult((), (), True),
        ]
        engine_calls = mock.Mock()
        engine_calls.attach_mock(mock_run_directory_batch, "run_directory_batch")
        engine_calls.attach_mock(mock_generate_summary, "generate_summary_report")
        engine_calls.attach_mock(mock_generate_sample, "generate_sample_table_report")
        engine_calls.attach_mock(mock_generate_presentation, "generate_presentation")

        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=root / "data",
                logs_root=root / "logs/pipeline",
            )
            defaults = build_defaults()
            outputs: list[str] = []

            run_pipeline(
                paths=paths,
                defaults=defaults,
                single_month=3,
                input_func=lambda prompt: "继续",
                output_func=outputs.append,
            )

            self.assertTrue(paths.pipeline_log_path.exists())
            self.assertTrue(paths.precheck_log_path.exists())
            self.assertTrue(paths.satisfaction_detail_dir.exists())
            self.assertTrue(paths.satisfaction_summary_dir.exists())
            self.assertTrue(paths.sample_summary_dir.exists())
            self.assertTrue(paths.ppt_dir.exists())
            self.assertTrue(any("输出目录不存在，已创建" in message for message in outputs))
            self.assertEqual(
                engine_calls.mock_calls,
                [
                    mock.call.run_directory_batch(
                        input_dir=paths.raw_dir,
                        output_dir=paths.satisfaction_detail_dir,
                        sheet_name=defaults.sheet_name,
                        output_format="xlsx",
                        calculation_mode=defaults.calculation_mode,
                    ),
                    mock.call.generate_summary_report(
                        input_dir=paths.satisfaction_detail_dir,
                        output_dir=paths.satisfaction_summary_dir,
                        output_name=paths.summary_workbook_path.name,
                    ),
                    mock.call.generate_sample_table_report(
                        input_dir=paths.raw_dir,
                        output_dir=paths.sample_summary_dir,
                        output_name=paths.sample_workbook_path.name,
                        config_path=defaults.sample_config_path,
                        source_sheet_name=defaults.sheet_name,
                        default_year=paths.batch_ref.year,
                    ),
                    mock.call.generate_presentation(mock.ANY),
                ],
            )
            ppt_config = mock_generate_presentation.call_args.args[0]
            self.assertIsInstance(ppt_config, PptBatchConfig)
            self.assertEqual(ppt_config.template_path, defaults.ppt.template_path)
            self.assertEqual(ppt_config.input_dir, paths.satisfaction_detail_dir)
            self.assertEqual(ppt_config.output_ppt, paths.ppt_path)
            self.assertEqual(ppt_config.sheet_name_mode, defaults.ppt.sheet_name_mode)
            self.assertEqual(ppt_config.section_mode, defaults.ppt.section_mode)
            self.assertEqual(ppt_config.chart_page.enabled, defaults.ppt.chart_page.enabled)
            self.assertEqual(
                ppt_config.chart_page.placeholder_text,
                defaults.ppt.chart_page.placeholder_text,
            )
            self.assertEqual(ppt_config.chart_page.image_dpi, defaults.ppt.chart_page.image_dpi)
            self.assertEqual(ppt_config.llm_notes.enabled, defaults.ppt.llm_notes.enabled)
            self.assertEqual(ppt_config.llm_notes.env_path, defaults.ppt.llm_notes.env_path)
            self.assertEqual(
                ppt_config.llm_notes.system_role_path,
                defaults.ppt.llm_notes.system_role_path,
            )
            self.assertEqual(
                ppt_config.llm_notes.target_chars,
                defaults.ppt.llm_notes.target_chars,
            )
            self.assertEqual(
                ppt_config.llm_notes.temperature,
                defaults.ppt.llm_notes.temperature,
            )
            self.assertEqual(
                ppt_config.llm_notes.max_tokens,
                defaults.ppt.llm_notes.max_tokens,
            )
            self.assertEqual(
                ppt_config.llm_notes.checkpoint_chars,
                defaults.ppt.llm_notes.checkpoint_chars,
            )

        self.assertEqual(mock_run_precheck.call_count, 2)
        mock_apply_year_month.assert_called_once()
        mock_run_directory_batch.assert_called_once()
        mock_generate_summary.assert_called_once()
        mock_generate_sample.assert_called_once()
        mock_generate_presentation.assert_called_once()

    @mock.patch("pipeline_runtime.run_precheck")
    def test_run_pipeline_keeps_waiting_when_recheck_still_blocks(
        self,
        mock_run_precheck: mock.Mock,
    ) -> None:
        blocking_issue = PipelineIssue("blocking", "unmapped", "存在未映射标签")
        mock_run_precheck.side_effect = [
            PrecheckResult((blocking_issue,), (), False),
            PrecheckResult((blocking_issue,), (), False),
        ]

        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths(
                "2026",
                "Q1",
                data_root=Path(tmp_dir) / "data",
                logs_root=Path(tmp_dir) / "logs/pipeline",
            )
            prompts = iter(["yes", "stop"])

            with self.assertRaisesRegex(SystemExit, "用户取消主流程。"):
                run_pipeline(
                    paths=paths,
                    defaults=build_defaults(),
                    single_month=None,
                    input_func=lambda prompt: next(prompts),
                    output_func=lambda message: None,
                )

    @mock.patch("pipeline_runtime.generate_summary_report")
    @mock.patch("pipeline_runtime.run_directory_batch")
    @mock.patch("pipeline_runtime.run_precheck")
    def test_run_pipeline_stops_when_engine_step_fails(
        self,
        mock_run_precheck: mock.Mock,
        mock_run_directory_batch: mock.Mock,
        mock_generate_summary: mock.Mock,
    ) -> None:
        mock_run_precheck.return_value = PrecheckResult((), (), False)
        mock_run_directory_batch.side_effect = RuntimeError("满意度分项统计失败")

        with tempfile.TemporaryDirectory() as tmp_dir:
            paths = build_pipeline_paths(
                "2026",
                "3月",
                data_root=Path(tmp_dir) / "data",
                logs_root=Path(tmp_dir) / "logs/pipeline",
            )

            with self.assertRaisesRegex(RuntimeError, "满意度分项统计失败"):
                run_pipeline(
                    paths=paths,
                    defaults=build_defaults(),
                    single_month=3,
                    output_func=lambda message: None,
                )

        mock_generate_summary.assert_not_called()


if __name__ == "__main__":
    unittest.main()
