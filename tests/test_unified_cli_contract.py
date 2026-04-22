from __future__ import annotations

import re
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

LEGACY_ENTRY_FILES = (
    "hangbo_gui.py",
    "tests/test_hangbo_gui.py",
    "job.toml",
    "job01-02.toml",
    "job03.toml",
    "job_Q1.toml",
    "report_jobs.1-2月.toml",
    "report_jobs.3月.toml",
    "report_jobs.Q1.toml",
    "report_jobs.example.toml",
    "report_jobs.directory.example.toml",
    "ppt_job.example.toml",
    "docs/UI适配新数据分析流程计划.md",
)

DOC_PATHS = (
    PROJECT_ROOT / "README.md",
    PROJECT_ROOT / "docs" / "README.md",
    PROJECT_ROOT / "docs" / "PPT生成说明.md",
    PROJECT_ROOT / "docs" / "用户运行用例故事.md",
    PROJECT_ROOT / "docs" / "统计口径与结果说明.md",
    PROJECT_ROOT / "docs" / "新数据分析流程说明.md",
    PROJECT_ROOT / "docs" / "数据准备与预查错.md",
)

STALE_DOC_PATTERNS = (
    "uv run python hangbo_gui.py",
    "--config pipeline.defaults.toml",
    "report_jobs.1-2月.toml",
    "report_jobs.3月.toml",
    "report_jobs.Q1.toml",
    "report_jobs.example.toml",
    "report_jobs.directory.example.toml",
    "ppt_job.example.toml",
    "job.toml",
    "job03.toml",
    "job_Q1.toml",
    "job01-02.toml",
    "GUI 入口",
    "GUI 工作台",
)


class UnifiedCliContractTest(unittest.TestCase):
    def test_legacy_gui_and_batch_config_entry_files_are_removed(self) -> None:
        existing_paths = [
            relative_path
            for relative_path in LEGACY_ENTRY_FILES
            if (PROJECT_ROOT / relative_path).exists()
        ]

        self.assertEqual(
            existing_paths,
            [],
            msg=f"Legacy entry/config files should be removed, but found: {existing_paths}",
        )

    def test_user_docs_do_not_recommend_legacy_entry_points(self) -> None:
        matches: list[str] = []
        for path in DOC_PATHS:
            content = path.read_text(encoding="utf-8")
            for pattern in STALE_DOC_PATTERNS:
                if pattern in content:
                    matches.append(f"{path.relative_to(PROJECT_ROOT)} contains {pattern!r}")

        self.assertEqual(
            matches,
            [],
            msg=f"Legacy docs patterns still present: {matches}",
        )

    def test_readme_recommends_default_cli_without_explicit_config(self) -> None:
        lines = (PROJECT_ROOT / "README.md").read_text(encoding="utf-8").splitlines()
        recommended_cli_pattern = re.compile(
            r"uv run python main_pipeline\.py --year \S+ --batch \S+"
        )
        matched_lines = [line.strip() for line in lines if recommended_cli_pattern.search(line)]

        self.assertTrue(
            matched_lines,
            msg=(
                "README should contain a recommended command like "
                "'uv run python main_pipeline.py --year <value> --batch <value>'."
            ),
        )
        self.assertTrue(
            all("--config pipeline.defaults.toml" not in line for line in matched_lines),
            msg=(
                "Recommended README command should not include explicit defaults config: "
                f"{matched_lines}"
            ),
        )


if __name__ == "__main__":
    unittest.main()
