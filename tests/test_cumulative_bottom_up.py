"""Public workbook output must reconcile every satisfaction level."""

from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import load_workbook

from hangbo.pipeline.cumulative import PeriodResult, build_cumulative_result, run_cumulative_batch, write_cumulative_detail_reports
from hangbo.summary.table import SUMMARY_COLUMNS, SummaryRowResult, row_kind
from hangbo.survey.stats import MetricResult, SectionResult, SurveyStatistics


ROOT = Path(__file__).resolve().parents[1]


class CumulativeMissingMonthsTest(unittest.TestCase):
    def test_missing_months_and_decimal_ties_in_published_workbook(self):
        periods = []
        # Two valid customers per period give the existing weight 1/6.
        for label, scores in (("7月", (10, 8, 9.75)), ("8月", (6, None, 9.60)), ("9月", (None, None, None))):
            sections = (
                SectionResult("会展服务", 9 if label == "7月" else 6, None,
                              tuple(MetricResult(name, score, None) for name, score in zip(("态度", "效率"), scores[:2]))),
                SectionResult("硬件设施", scores[2], None, (MetricResult("设施", scores[2], None),)),
                SectionResult("配套服务", None, None, (MetricResult("餐饮", None, None),)),
            )
            stats = SurveyStatistics("展览主承办", 8, None, sections, 1)
            values = dict.fromkeys(SUMMARY_COLUMNS)
            values["总分"] = 8
            rows = tuple(SummaryRowResult("一、会展客户", name, values, frozenset(SUMMARY_COLUMNS))
                         for name in ("展览活动主（承）办", "参展商"))
            periods.append(PeriodResult(label, rows, {"展览主承办": stats}))
        with TemporaryDirectory() as directory:
            write_cumulative_detail_reports(build_cumulative_result(tuple(periods)), Path(directory))
            workbook = load_workbook(Path(directory) / "展览主承办.xlsx", data_only=True)
            try:
                scores = {row[0]: row[1] for row in workbook.active.iter_rows(values_only=True)}
                # Child means are 8 and 8, irrespective of differing month sets.
                self.assertEqual(scores["会展服务"], 8)
                # (9.75 + 9.60)/2 = 9.675, rounded half up, even at weight 1/6.
                self.assertEqual(scores["设施"], 9.68)
                self.assertEqual(scores["硬件设施"], 9.68)
                self.assertIsNone(scores["配套服务"])
                self.assertEqual(workbook.active["B2"].value, 8.84)
            finally:
                workbook.close()


def mean(values):
    defined = [Decimal(str(value)) for value in values if value is not None]
    if not defined:
        return None
    return float((sum(defined) / len(defined)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


@unittest.skipUnless((ROOT / "data/raw/2026/9月/会展服务商.xlsx").is_file(), "local 2026 survey data required")
class CumulativeBottomUpTest(unittest.TestCase):
    def test_quarter_sections_equal_child_means_without_display_rounding_differences(self):
        with TemporaryDirectory() as directory:
            output = Path(directory)
            run_cumulative_batch(
                year="2026", batch="Q3", data_root=ROOT / "data",
                detail_output_dir=output / "details", summary_output_path=output / "summary.xlsx",
                sheet_name="问卷数据", calculation_mode="template",
                january_february_mode="auto", log_path=output / "cumulative.json",
            )
            workbook = load_workbook(output / "details/参会人员.xlsx", data_only=True)
            try:
                sheet = workbook.active
                # Quarter children: dining 9.77, room 9.67, security/cleaning 10.
                self.assertEqual(sheet["B18"].value, 9.86)
                self.assertEqual(sheet["B2"].value, 9.82)
            finally:
                workbook.close()
            for path in sorted((output / "details").glob("*.xlsx")):
                workbook = load_workbook(path, data_only=True)
                try:
                    sheet = workbook.active
                    sections = []
                    for row_index in range(3, sheet.max_row + 1):
                        kind = row_kind(sheet, row_index, 1)
                        score = sheet.cell(row_index, 2).value
                        if kind == "section":
                            sections.append({"row": row_index, "score": score, "children": []})
                        elif kind == "metric":
                            sections[-1]["children"].append(score)
                        if score is not None:
                            with self.subTest(file=path.name, row=row_index):
                                self.assertEqual(Decimal(str(score)), Decimal(str(score)).quantize(Decimal("0.01")))
                                self.assertEqual(float(f"{score:.2f}"), score)
                    for section in sections:
                        with self.subTest(file=path.name, section=section["row"]):
                            self.assertEqual(section["score"], mean(section["children"]))
                    self.assertEqual(sheet["B2"].value, mean(section["score"] for section in sections))
                finally:
                    workbook.close()
            workbook = load_workbook(output / "summary.xlsx", data_only=True)
            try:
                self.assertEqual(workbook["汇总表"]["F8"].value, 9.86)
                self.assertEqual(workbook["汇总表"]["C8"].value, 9.82)
                self.assertEqual(workbook["汇总表"]["C22"].value, 9.73)
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
