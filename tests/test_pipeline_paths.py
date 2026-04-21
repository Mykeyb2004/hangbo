from __future__ import annotations

import unittest
from pathlib import Path

from pipeline_models import BatchRef, PipelinePaths
from pipeline_paths import (
    STANDARD_SOURCE_FILE_NAMES,
    build_pipeline_paths,
    parse_single_month_batch,
)


class PipelinePathsTest(unittest.TestCase):
    def test_build_pipeline_paths_uses_fixed_directory_contract(self) -> None:
        result = build_pipeline_paths(
            year=" 2026 ",
            batch=" 3月 ",
            data_root=Path("data"),
            logs_root=Path("logs/pipeline"),
        )

        self.assertEqual(
            result,
            PipelinePaths(
                batch_ref=BatchRef(year="2026", batch="3月"),
                data_root=Path("data"),
                logs_root=Path("logs/pipeline"),
                raw_dir=Path("data/raw/2026/3月"),
                satisfaction_detail_dir=Path("data/satisfaction_detail/2026/3月"),
                satisfaction_summary_dir=Path("data/satisfaction_summary/2026/3月"),
                sample_summary_dir=Path("data/sample_summary/2026/3月"),
                ppt_dir=Path("data/ppt/2026/3月"),
                logs_dir=Path("logs/pipeline/2026/3月"),
                summary_workbook_path=Path(
                    "data/satisfaction_summary/2026/3月/3月客户类型满意度汇总表.xlsx"
                ),
                sample_workbook_path=Path(
                    "data/sample_summary/2026/3月/3月客户类型样本统计表.xlsx"
                ),
                ppt_path=Path("data/ppt/2026/3月/3月满意度报告.pptx"),
                precheck_log_path=Path("logs/pipeline/2026/3月/precheck.log"),
                pipeline_log_path=Path("logs/pipeline/2026/3月/pipeline.log"),
                unmapped_log_path=Path(
                    "logs/pipeline/2026/3月/unmapped_customer_records.log"
                ),
                standard_source_paths=tuple(
                    Path("data/raw/2026/3月") / file_name
                    for file_name in STANDARD_SOURCE_FILE_NAMES
                ),
            ),
        )

    def test_build_pipeline_paths_exposes_standard_source_file_paths(self) -> None:
        result = build_pipeline_paths(year="2026", batch="Q1")

        self.assertEqual(
            result.standard_source_paths,
            (
                Path("data/raw/2026/Q1/展览.xlsx"),
                Path("data/raw/2026/Q1/会议.xlsx"),
                Path("data/raw/2026/Q1/酒店.xlsx"),
                Path("data/raw/2026/Q1/餐饮.xlsx"),
                Path("data/raw/2026/Q1/会展服务商.xlsx"),
                Path("data/raw/2026/Q1/旅游.xlsx"),
            ),
        )

    def test_parse_single_month_batch_returns_month_number(self) -> None:
        self.assertEqual(parse_single_month_batch("3月"), 3)
        self.assertEqual(parse_single_month_batch("03月"), 3)
        self.assertEqual(parse_single_month_batch(" 03月 "), 3)

    def test_parse_single_month_batch_returns_none_for_combined_batches(self) -> None:
        self.assertIsNone(parse_single_month_batch("1-2月"))
        self.assertIsNone(parse_single_month_batch("Q1"))
        self.assertIsNone(parse_single_month_batch("13月"))
        self.assertIsNone(parse_single_month_batch("00月"))


if __name__ == "__main__":
    unittest.main()
