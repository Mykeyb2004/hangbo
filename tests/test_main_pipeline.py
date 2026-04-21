from __future__ import annotations

import unittest
from pathlib import Path
from unittest import mock

from main_pipeline import main, parse_args


class MainPipelineCliTest(unittest.TestCase):
    def test_parse_args_reads_year_batch_and_config(self) -> None:
        args = parse_args(
            ["--year", "2026", "--batch", "3月", "--config", "pipeline.defaults.toml"]
        )

        self.assertEqual(args.year, "2026")
        self.assertEqual(args.batch, "3月")
        self.assertEqual(str(args.config), "pipeline.defaults.toml")

    @mock.patch("main_pipeline.run_pipeline")
    @mock.patch("main_pipeline.load_pipeline_defaults")
    @mock.patch("main_pipeline.build_pipeline_paths")
    def test_main_builds_paths_and_invokes_runtime(
        self,
        mock_build_pipeline_paths: mock.Mock,
        mock_load_pipeline_defaults: mock.Mock,
        mock_run_pipeline: mock.Mock,
    ) -> None:
        mock_paths = mock.sentinel.paths
        mock_defaults = mock.sentinel.defaults
        mock_build_pipeline_paths.return_value = mock_paths
        mock_load_pipeline_defaults.return_value = mock_defaults

        main(["--year", "2026", "--batch", "3月"])

        mock_build_pipeline_paths.assert_called_once_with("2026", "3月")
        mock_load_pipeline_defaults.assert_called_once_with(
            Path("pipeline.defaults.toml")
        )
        mock_run_pipeline.assert_called_once_with(
            paths=mock_paths,
            defaults=mock_defaults,
            single_month=3,
        )

    @mock.patch("main_pipeline.run_pipeline")
    @mock.patch("main_pipeline.load_pipeline_defaults")
    @mock.patch("main_pipeline.build_pipeline_paths")
    def test_main_passes_none_for_combined_batch_single_month(
        self,
        mock_build_pipeline_paths: mock.Mock,
        mock_load_pipeline_defaults: mock.Mock,
        mock_run_pipeline: mock.Mock,
    ) -> None:
        mock_paths = mock.sentinel.paths
        mock_defaults = mock.sentinel.defaults
        mock_build_pipeline_paths.return_value = mock_paths
        mock_load_pipeline_defaults.return_value = mock_defaults

        main(["--year", "2026", "--batch", "Q1"])

        mock_build_pipeline_paths.assert_called_once_with("2026", "Q1")
        mock_load_pipeline_defaults.assert_called_once_with(
            Path("pipeline.defaults.toml")
        )
        mock_run_pipeline.assert_called_once_with(
            paths=mock_paths,
            defaults=mock_defaults,
            single_month=None,
        )


if __name__ == "__main__":
    unittest.main()
