from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from io import BytesIO
from math import pi
from typing import Literal, Sequence

import matplotlib

matplotlib.use("Agg")

from matplotlib import pyplot as plt
from matplotlib.font_manager import FontProperties, fontManager

ChartType = Literal["bar", "radar"]

PRIMARY_SERIES_COLOR = "#BF1E4B"
SECONDARY_SERIES_COLOR = "#E49AAF"
AXIS_TEXT_COLOR = "#4D5874"
GRID_COLOR = "#D9D9D9"


@dataclass(frozen=True)
class ChartPoint:
    label: str
    satisfaction: float
    importance: float


@dataclass(frozen=True)
class ChartRenderConfig:
    dpi: int = 220


def choose_chart_type(points: Sequence[ChartPoint]) -> ChartType | None:
    if len(points) < 2:
        return None
    if len(points) == 2:
        return "bar"
    return "radar"


def render_chart_image(
    points: Sequence[ChartPoint],
    *,
    config: ChartRenderConfig | None = None,
    width_inches: float = 5.6,
    height_inches: float = 5.1,
) -> bytes:
    chart_type = choose_chart_type(points)
    if chart_type is None:
        raise ValueError("至少需要 2 个二级指标才能生成图表")

    resolved_config = config or ChartRenderConfig()
    figure = plt.figure(figsize=(width_inches, height_inches), dpi=resolved_config.dpi)
    figure.patch.set_facecolor("white")

    if chart_type == "bar":
        _render_bar_chart(figure, points)
    else:
        _render_radar_chart(figure, points)

    output = BytesIO()
    figure.savefig(output, format="png", bbox_inches="tight", facecolor="white")
    plt.close(figure)
    return output.getvalue()


def _render_bar_chart(figure, points: Sequence[ChartPoint]) -> None:
    ax = figure.add_subplot(111)
    labels = [point.label for point in points]
    x_positions = list(range(len(points)))
    bar_width = 0.32

    satisfaction_values = [point.satisfaction for point in points]
    importance_values = [point.importance for point in points]

    satisfaction_bars = ax.bar(
        [position - bar_width / 2 for position in x_positions],
        satisfaction_values,
        width=bar_width,
        color=PRIMARY_SERIES_COLOR,
        alpha=0.92,
        label="满意度",
    )
    importance_bars = ax.bar(
        [position + bar_width / 2 for position in x_positions],
        importance_values,
        width=bar_width,
        color=SECONDARY_SERIES_COLOR,
        alpha=0.82,
        label="重要性",
    )

    _configure_font()
    ax.set_xticks(x_positions)
    ax.set_xticklabels(labels, color=AXIS_TEXT_COLOR, fontproperties=_font_properties())
    ax.set_ylim(0, 10)
    ax.set_yticks([0, 2, 4, 6, 8, 10])
    ax.tick_params(axis="y", colors=AXIS_TEXT_COLOR)
    ax.grid(axis="y", linestyle="--", color=GRID_COLOR, alpha=0.7)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID_COLOR)
    ax.spines["bottom"].set_color(GRID_COLOR)
    ax.legend(loc="upper center", ncol=2, frameon=False, prop=_font_properties())

    for bar in [*satisfaction_bars, *importance_bars]:
        height = bar.get_height()
        ax.annotate(
            f"{height:.2f}".rstrip("0").rstrip("."),
            xy=(bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 4),
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=9,
            color=AXIS_TEXT_COLOR,
            fontproperties=_font_properties(),
        )

    figure.tight_layout()


def _render_radar_chart(figure, points: Sequence[ChartPoint]) -> None:
    ax = figure.add_subplot(111, polar=True)
    labels = [point.label for point in points]
    angles = [index / len(points) * 2 * pi for index in range(len(points))]
    closed_angles = [*angles, angles[0]]

    satisfaction_values = [point.satisfaction for point in points]
    importance_values = [point.importance for point in points]
    closed_satisfaction = [*satisfaction_values, satisfaction_values[0]]
    closed_importance = [*importance_values, importance_values[0]]

    _configure_font()
    ax.set_theta_offset(pi / 2)
    ax.set_theta_direction(-1)
    ax.set_thetagrids(
        [angle * 180 / pi for angle in angles],
        labels=labels,
        fontproperties=_font_properties(),
        color=AXIS_TEXT_COLOR,
    )
    ax.set_ylim(0, 10)
    ax.set_yticks([2, 4, 6, 8, 10])
    ax.set_yticklabels(["2", "4", "6", "8", "10"], color=AXIS_TEXT_COLOR, fontproperties=_font_properties())
    ax.grid(color=GRID_COLOR, alpha=0.85)
    ax.spines["polar"].set_color(GRID_COLOR)

    ax.plot(
        closed_angles,
        closed_satisfaction,
        color=PRIMARY_SERIES_COLOR,
        linewidth=2.2,
        label="满意度",
    )
    ax.fill(closed_angles, closed_satisfaction, color=PRIMARY_SERIES_COLOR, alpha=0.28)

    ax.plot(
        closed_angles,
        closed_importance,
        color=SECONDARY_SERIES_COLOR,
        linewidth=2.0,
        linestyle="--",
        label="重要性",
    )
    ax.fill(closed_angles, closed_importance, color=SECONDARY_SERIES_COLOR, alpha=0.16)

    for angle, value in zip(angles, satisfaction_values):
        ax.annotate(
            f"{value:.2f}".rstrip("0").rstrip("."),
            xy=(angle, value),
            xytext=(0, 8),
            textcoords="offset points",
            ha="center",
            va="center",
            fontsize=9,
            color=AXIS_TEXT_COLOR,
            fontproperties=_font_properties(),
            bbox={"facecolor": "white", "edgecolor": "none", "alpha": 0.85, "pad": 0.8},
        )

    ax.legend(loc="upper center", bbox_to_anchor=(0.5, 1.18), ncol=2, frameon=False, prop=_font_properties())
    figure.tight_layout()


def _configure_font() -> None:
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["font.sans-serif"] = [_choose_font_family()]


def _font_properties() -> FontProperties:
    return FontProperties(family=_choose_font_family())


@lru_cache(maxsize=1)
def _choose_font_family() -> str:
    candidates = [
        "PingFang SC",
        "Hiragino Sans GB",
        "Microsoft YaHei",
        "SimHei",
        "Noto Sans CJK SC",
        "Arial Unicode MS",
        "DejaVu Sans",
    ]
    available = {font.name for font in fontManager.ttflist}
    for candidate in candidates:
        if candidate in available:
            return candidate
    return "DejaVu Sans"
