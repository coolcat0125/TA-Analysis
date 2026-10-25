# -*- coding: utf-8 -*-
"""Generate report charts from the normalized delivery payload."""

from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)
PAYLOAD = json.loads((ROOT / "data" / "delivery_payload.json").read_text(encoding="utf-8"))
FONT = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 28)
FONT_SMALL = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 22)
FONT_TITLE = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 36)


def draw_grouped_bars(
    output: Path,
    title: str,
    categories: list[str],
    series: list[tuple[str, list[float | None], str]],
    y_suffix: str = "",
    note: str | None = None,
) -> None:
    width, height = 1600, 920
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    left, top, right, bottom = 150, 130, 1520, 790

    plot_values = [value for _, values, _ in series for value in values if value is not None]
    max_value = max(plot_values) * 1.18 if plot_values else 1
    if max_value <= 10:
        max_value = 10

    for step in range(5):
        y = bottom - (bottom - top) * step / 4
        value = max_value * step / 4
        draw.line((left, y, right, y), fill="#E5E7EB", width=2)
        draw.text(
            (left - 20, y),
            f"{value:.0f}{y_suffix}",
            font=FONT_SMALL,
            fill="#6B7280",
            anchor="rm",
        )

    draw.text((left, 45), title, font=FONT_TITLE, fill="#111827")
    group_width = (right - left) / len(categories)
    bar_width = group_width * 0.18
    gap = group_width * 0.05
    for category_index, category in enumerate(categories):
        center = left + group_width * (category_index + 0.5)
        for series_index, (name, values, color) in enumerate(series):
            value = values[category_index]
            x_center = center + (series_index - (len(series) - 1) / 2) * (bar_width + gap)
            x0 = x_center - bar_width / 2
            x1 = x_center + bar_width / 2
            if value is None:
                draw.text(
                    ((x0 + x1) / 2, bottom - 12),
                    "n.a.",
                    font=FONT_SMALL,
                    fill="#9CA3AF",
                    anchor="ms",
                )
                continue
            y0 = bottom - (bottom - top) * value / max_value
            draw.rounded_rectangle((x0, y0, x1, bottom), radius=5, fill=color)
            draw.text(
                ((x0 + x1) / 2, y0 - 12),
                f"{value:.0f}",
                font=FONT_SMALL,
                fill="#374151",
                anchor="ms",
            )
        draw.text((center, bottom + 28), category, font=FONT, fill="#374151", anchor="ma")

    legend_x = left
    legend_y = 92
    for name, _, color in series:
        draw.rounded_rectangle((legend_x, legend_y, legend_x + 24, legend_y + 24), radius=3, fill=color)
        draw.text((legend_x + 34, legend_y + 12), name, font=FONT_SMALL, fill="#4B5563", anchor="lm")
        legend_x += 230
    if note:
        draw.text((left, 870), note, font=FONT_SMALL, fill="#6B7280", anchor="lm")
    image.save(output)


def main() -> None:
    summary = PAYLOAD["batch_summary"]
    batches = ["410", "411", "412"]
    bev = [next((row["数量"] for row in summary[b] if row["技术路线"] == "纯电(BEV)"), 0) for b in batches]
    phev = [next((row["数量"] for row in summary[b] if row["技术路线"] == "插电混动(PHEV)"), 0) for b in batches]
    erev = [next((row["数量"] for row in summary[b] if row["技术路线"] == "增程(EREV)"), 0) for b in batches]
    draw_grouped_bars(
        ASSETS / "chart_batch_structure.png",
        "各批次新能源乘用车结构",
        ["410正式", "411正式", "412公示"],
        [
            ("纯电(BEV)", bev, "#1F4E78"),
            ("插电混动(PHEV)", phev, "#2E75B6"),
            ("增程(EREV)", erev, "#D98C3F"),
        ],
        note="412为公示候选口径，正式公告发布后应以购置税目录附件复核。",
    )

    power_bev = [
        next((row["平均电机功率(kW)"] for row in summary[b] if row["技术路线"] == "纯电(BEV)"), None)
        for b in batches
    ]
    power_phev = [
        next((row["平均电机功率(kW)"] for row in summary[b] if row["技术路线"] == "插电混动(PHEV)"), None)
        for b in batches
    ]
    power_erev = [
        next((row["平均电机功率(kW)"] for row in summary[b] if row["技术路线"] == "增程(EREV)"), None)
        for b in batches
    ]
    draw_grouped_bars(
        ASSETS / "chart_power.png",
        "批次平均电机总功率",
        ["410正式", "411正式", "412公示"],
        [
            ("纯电(BEV)", power_bev, "#1F4E78"),
            ("插电混动(PHEV)", power_phev, "#2E75B6"),
            ("增程(EREV)", power_erev, "#D98C3F"),
        ],
        " kW",
        note="411 PHEV功率有效样本6款；均值只按公告已公开的有值样本计算。",
    )
    print(f"Wrote charts to {ASSETS}")


if __name__ == "__main__":
    main()
