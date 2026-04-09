from __future__ import annotations

import unittest

from ppt_chart_renderer import ChartPoint, choose_chart_type, render_chart_image


class PptChartRendererTest(unittest.TestCase):
    def test_choose_chart_type_returns_bar_for_two_sections(self) -> None:
        points = [
            ChartPoint("服务", 9.6, 9.8),
            ChartPoint("设施", 9.2, 9.7),
        ]

        self.assertEqual(choose_chart_type(points), "bar")

    def test_choose_chart_type_returns_radar_for_three_or_more_sections(self) -> None:
        points = [
            ChartPoint("服务", 9.6, 9.8),
            ChartPoint("设施", 9.2, 9.7),
            ChartPoint("配套", 8.9, 9.4),
        ]

        self.assertEqual(choose_chart_type(points), "radar")

    def test_render_chart_image_returns_png_bytes(self) -> None:
        points = [
            ChartPoint("服务", 9.6, 9.8),
            ChartPoint("设施", 9.2, 9.7),
            ChartPoint("配套", 8.9, 9.4),
        ]

        image_bytes = render_chart_image(points)

        self.assertGreater(len(image_bytes), 1000)
        self.assertTrue(image_bytes.startswith(b"\x89PNG\r\n\x1a\n"))

    def test_render_chart_image_requires_at_least_two_sections(self) -> None:
        with self.assertRaises(ValueError):
            render_chart_image([ChartPoint("服务", 9.6, 9.8)])


if __name__ == "__main__":
    unittest.main()
