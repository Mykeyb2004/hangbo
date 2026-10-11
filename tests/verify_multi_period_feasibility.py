"""Verify period-equal aggregation without modifying existing workbooks.

Run: uv run python tests/verify_multi_period_feasibility.py
"""

from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
PERIODS = ("1-2月", *(f"{month}月" for month in range(3, 10)))


def mean(values):
    return sum(values, Decimal(0)) / len(values)


def display(value):
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def verify():
    totals = []
    scores_by_period = []
    for period in PERIODS:
        path = ROOT / "data/satisfaction_summary/2026" / period / f"{period}客户类型满意度汇总表.xlsx"
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            rows = list(workbook.active.values)
            totals.append(Decimal(str(rows[-1][2])))
            scores_by_period.append({
                index: Decimal(str(row[2]))
                for index, row in enumerate(rows[2:-1])
                if isinstance(row[2], (int, float))
            })
        finally:
            workbook.close()

    target = mean(totals)
    row_ids = sorted(set().union(*(set(scores) for scores in scores_by_period)))
    simple_row_means = {
        row_id: mean([scores[row_id] for scores in scores_by_period if row_id in scores])
        for row_id in row_ids
    }
    period_count = Decimal(len(PERIODS))
    weights = {}
    contributions = {}
    weighted_row_scores = {}
    for row_id in row_ids:
        terms = [
            (Decimal(1) / period_count / len(scores), scores[row_id])
            for scores in scores_by_period if row_id in scores
        ]
        weights[row_id] = sum((weight for weight, _ in terms), Decimal(0))
        contributions[row_id] = sum((weight * score for weight, score in terms), Decimal(0))
        weighted_row_scores[row_id] = contributions[row_id] / weights[row_id]

    reconstructed = sum(contributions.values(), Decimal(0))
    precision_gap = target - reconstructed
    tolerance = Decimal("1e-20")
    assert abs(sum(weights.values()) - 1) < tolerance
    assert abs(reconstructed - mean([mean(list(scores.values())) for scores in scores_by_period])) < tolerance
    assert abs(reconstructed + precision_gap - target) < tolerance
    assert display(target) == Decimal("9.71")
    assert display(mean(list(simple_row_means.values()))) == Decimal("9.72")
    assert display(reconstructed) == Decimal("9.71")

    # Missing categories make plain averaging depend on the aggregation order.
    example_target = mean([Decimal(10), mean([Decimal(8), Decimal(6)])])
    example_naive = mean([mean([Decimal(10), Decimal(8)]), Decimal(6)])
    assert example_target == Decimal("8.5")
    assert example_naive == Decimal("7.5")

    print(f"period_count={len(PERIODS)}")
    print(f"mean_period_totals={target}; display={display(target)}")
    print(f"mean_type_period_means={mean(list(simple_row_means.values()))}; display={display(mean(list(simple_row_means.values())))}")
    print(f"weighted_type_rollup={reconstructed}; display={display(reconstructed)}")
    print(f"sum_type_weights={sum(weights.values())}")
    print(f"rounding_difference_to_saved_period_totals={precision_gap}")
    print("All feasibility checks passed; no workbooks or production files were written.")


if __name__ == "__main__":
    verify()
