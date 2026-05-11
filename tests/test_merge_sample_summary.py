from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from merge_sample_summary import (
    BatchNameError,
    build_merge_sample_paths,
    discover_source_directories,
    parse_number_selection,
    validate_batch_name,
)


class MergeSampleSummaryHelpersTest(unittest.TestCase):
    def test_discover_source_directories_lists_only_direct_folders_sorted_by_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            raw_year_dir = Path(temp_dir) / "data" / "raw" / "2026"
            raw_year_dir.mkdir(parents=True)
            (raw_year_dir / "3月").mkdir()
            (raw_year_dir / "1-2月").mkdir()
            (raw_year_dir / "说明.txt").write_text("ignore", encoding="utf-8")
            (raw_year_dir / "Q1").mkdir()
            (raw_year_dir / "Q1" / "nested").mkdir()

            result = discover_source_directories(raw_year_dir)

        self.assertEqual([item.name for item in result], ["1-2月", "3月", "Q1"])

    def test_discover_source_directories_rejects_missing_year_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaisesRegex(FileNotFoundError, "年份原始数据目录不存在"):
                discover_source_directories(Path(temp_dir) / "data" / "raw" / "2026")

    def test_discover_source_directories_rejects_empty_year_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            raw_year_dir = Path(temp_dir) / "data" / "raw" / "2026"
            raw_year_dir.mkdir(parents=True)

            with self.assertRaisesRegex(ValueError, "没有可选择的来源目录"):
                discover_source_directories(raw_year_dir)

    def test_parse_number_selection_supports_commas_and_ranges(self) -> None:
        result = parse_number_selection("1, 3-5, 2", item_count=5)

        self.assertEqual(result, (0, 2, 3, 4, 1))

    def test_parse_number_selection_rejects_out_of_range_and_empty_values(self) -> None:
        with self.assertRaisesRegex(ValueError, "至少选择一个来源目录"):
            parse_number_selection("", item_count=3)
        with self.assertRaisesRegex(ValueError, "超出范围"):
            parse_number_selection("4", item_count=3)
        with self.assertRaisesRegex(ValueError, "范围起点不能大于终点"):
            parse_number_selection("3-1", item_count=3)

    def test_validate_batch_name_rejects_empty_separator_and_source_conflict(self) -> None:
        selected_dirs = (Path("data/raw/2026/1-2月"), Path("data/raw/2026/3月"))

        self.assertEqual(validate_batch_name(" Q1 ", selected_dirs), "Q1")
        for raw_name in ("", "  ", "../Q1", "Q1/backup", "1-2月"):
            with self.assertRaises(BatchNameError):
                validate_batch_name(raw_name, selected_dirs)

    def test_build_merge_sample_paths_uses_existing_directory_contract(self) -> None:
        paths = build_merge_sample_paths(
            year="2026",
            batch_name="Q1",
            data_root=Path("data"),
        )

        self.assertEqual(paths.raw_year_dir, Path("data/raw/2026"))
        self.assertEqual(paths.merged_raw_dir, Path("data/raw/2026/Q1"))
        self.assertEqual(paths.sample_summary_dir, Path("data/sample_summary/2026/Q1"))
        self.assertEqual(
            paths.sample_summary_path,
            Path("data/sample_summary/2026/Q1/Q1客户类型样本统计表.xlsx"),
        )


if __name__ == "__main__":
    unittest.main()
