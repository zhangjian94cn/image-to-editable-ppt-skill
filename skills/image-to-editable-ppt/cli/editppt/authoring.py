"""Stable source-pixel authoring components for editable slide reconstruction.

Page-specific scripts may compose these objects and then call ``editppt build``.
They must not package OOXML or reimplement font fitting themselves.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence


Box = Sequence[float]
Point = Sequence[float]


def _box(value: Box) -> list[float]:
    if len(value) != 4:
        raise ValueError("box_px must contain x, y, width, height")
    result = [float(item) for item in value]
    if result[2] <= 0 or result[3] <= 0:
        raise ValueError("box_px width and height must be positive")
    return result


def _point(value: Point) -> tuple[float, float]:
    if len(value) != 2:
        raise ValueError("point must contain x and y")
    return float(value[0]), float(value[1])


@dataclass
class SlideManifest:
    """Manifest builder that keeps every authored coordinate in source pixels."""

    width_px: int
    height_px: int
    slide_width: float = 13.333333
    slide_height: float = 7.5
    background: str = "#FFFFFF"
    shapes: list[dict[str, Any]] = field(default_factory=list)
    images: list[dict[str, Any]] = field(default_factory=list)
    tables: list[dict[str, Any]] = field(default_factory=list)
    text_boxes: list[dict[str, Any]] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.width_px <= 0 or self.height_px <= 0:
            raise ValueError("source dimensions must be positive")
        if self.slide_width <= 0 or self.slide_height <= 0:
            raise ValueError("slide dimensions must be positive")

    def add_shape(
        self,
        box_px: Box,
        kind: str = "rect",
        *,
        fill: str = "none",
        stroke: str = "none",
        stroke_width: float = 1,
        radius_px: float | None = None,
        preset: str = "",
        z_index: float = 100,
        **extra: Any,
    ) -> dict[str, Any]:
        item: dict[str, Any] = {
            "type": kind,
            "box_px": _box(box_px),
            "fill": fill,
            "stroke": stroke,
            "stroke_width": float(stroke_width),
            "z_index": float(z_index),
        }
        if radius_px is not None:
            item["source_corner_radius_px"] = float(radius_px)
        if preset:
            item["preset"] = preset
        item.update(extra)
        self.shapes.append(item)
        return item

    def add_connector(
        self,
        start_px: Point,
        end_px: Point,
        *,
        color: str = "#4472C4",
        width: float = 1.25,
        stroke: str | None = None,
        stroke_width: float | None = None,
        dash: str = "",
        start_arrow: str = "none",
        end_arrow: str = "triangle",
        preset: str = "line",
        z_index: float = 140,
    ) -> dict[str, Any]:
        if stroke is not None:
            color = stroke
        if stroke_width is not None:
            width = float(stroke_width)
        x1, y1 = _point(start_px)
        x2, y2 = _point(end_px)
        item = {
            "type": "line",
            "points_px": [x1, y1, x2, y2],
            "stroke": color,
            "stroke_width": float(width),
            "start_arrow": start_arrow,
            "end_arrow": end_arrow,
            "preset": preset,
            "z_index": float(z_index),
        }
        if dash:
            item["dash"] = dash
        self.shapes.append(item)
        return item

    def add_polygon(
        self,
        points_px: Sequence[Point],
        *,
        fill: str = "none",
        stroke: str = "none",
        stroke_width: float = 1,
        z_index: float = 100,
        **extra: Any,
    ) -> dict[str, Any]:
        points = [_point(value) for value in points_px]
        if len(points) < 3:
            raise ValueError("polygon requires at least three source-space points")
        xs = [value[0] for value in points]
        ys = [value[1] for value in points]
        item: dict[str, Any] = {
            "type": "polygon",
            "box_px": [min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)],
            "polygon_px": [[x, y] for x, y in points],
            "fill": fill,
            "stroke": stroke,
            "stroke_width": float(stroke_width),
            "z_index": float(z_index),
        }
        item.update(extra)
        self.shapes.append(item)
        return item

    def add_text(
        self,
        box_px: Box,
        text: str | None = None,
        *,
        runs: Iterable[dict[str, Any]] | None = None,
        paragraphs: Iterable[dict[str, Any] | str] | None = None,
        font: str = "PingFang SC",
        font_size: float | None = None,
        font_size_px: float | None = None,
        color: str = "#111111",
        bold: bool = False,
        align: str = "left",
        valign: str = "top",
        wrap: str = "none",
        fit_text: bool = True,
        z_index: float = 300,
        **extra: Any,
    ) -> dict[str, Any]:
        supplied = int(text is not None) + int(runs is not None) + int(paragraphs is not None)
        if supplied != 1:
            raise ValueError("provide exactly one of text, runs, or paragraphs")
        if font_size is not None and font_size_px is not None:
            raise ValueError("provide font_size in points or font_size_px, not both")
        if font_size_px is not None:
            font_size = float(font_size_px) * self.slide_height * 72.0 / self.height_px
        if font_size is None:
            font_size = 18.0
        item: dict[str, Any] = {
            "box_px": _box(box_px),
            "font": font,
            "font_size": float(font_size),
            "color": color,
            "bold": bool(bold),
            "align": align,
            "valign": valign,
            "wrap": wrap,
            "fit_text": bool(fit_text),
            "z_index": float(z_index),
        }
        if runs is not None:
            item["runs"] = [dict(run) for run in runs]
        elif paragraphs is not None:
            item["paragraphs"] = [dict(value) if isinstance(value, dict) else str(value) for value in paragraphs]
        else:
            item["text"] = str(text)
        item.update(extra)
        self.text_boxes.append(item)
        return item

    def add_image(
        self,
        box_px: Box,
        path: str | Path,
        *,
        alt: str,
        z_index: float = 200,
    ) -> dict[str, Any]:
        if not alt.strip():
            raise ValueError("image alt text is required")
        item = {
            "path": str(path),
            "box_px": _box(box_px),
            "alt": alt.strip(),
            "z_index": float(z_index),
        }
        self.images.append(item)
        return item

    def add_table(
        self,
        box_px: Box,
        rows: Sequence[Sequence[Any]],
        *,
        column_widths: Sequence[float] | None = None,
        row_heights: Sequence[float] | None = None,
        font: str = "PingFang SC",
        font_size: float | None = None,
        font_size_px: float | None = None,
        header_fill: str = "#0A65B7",
        header_color: str = "#FFFFFF",
        fill: str = "#FFFFFF",
        color: str = "#111111",
        border: str = "#B7C4D4",
        z_index: float = 250,
        **extra: Any,
    ) -> dict[str, Any]:
        if not rows or not max((len(row) for row in rows), default=0):
            raise ValueError("table rows must be non-empty")
        if font_size is not None and font_size_px is not None:
            raise ValueError("provide font_size in points or font_size_px, not both")
        if font_size_px is not None:
            font_size = float(font_size_px) * self.slide_height * 72.0 / self.height_px
        if font_size is None:
            font_size = 12.0
        item: dict[str, Any] = {
            "box_px": _box(box_px),
            "rows": [list(row) for row in rows],
            "font": font,
            "font_size": float(font_size),
            "header_fill": header_fill,
            "header_color": header_color,
            "fill": fill,
            "color": color,
            "border": border,
            "z_index": float(z_index),
        }
        if column_widths is not None:
            item["column_widths"] = [float(value) for value in column_widths]
        if row_heights is not None:
            item["row_heights"] = [float(value) for value in row_heights]
        item.update(extra)
        self.tables.append(item)
        return item

    def add_section_band(
        self,
        box_px: Box,
        label: str,
        *,
        fill: str,
        color: str = "#FFFFFF",
        font_size: float = 20,
        radius_px: float = 0,
    ) -> None:
        box = _box(box_px)
        self.add_shape(box, kind="roundRect" if radius_px else "rect", fill=fill, radius_px=radius_px)
        self.add_text(box, label, font_size=font_size, color=color, bold=True, align="center", valign="middle")

    def add_card(
        self,
        box_px: Box,
        *,
        title: str,
        body: str,
        fill: str = "#FFFFFF",
        stroke: str = "#D5E3F1",
        title_color: str = "#0A65B7",
        body_color: str = "#333333",
        radius_px: float = 12,
        padding_px: float = 18,
    ) -> None:
        x, y, width, height = _box(box_px)
        self.add_shape([x, y, width, height], kind="roundRect", fill=fill, stroke=stroke, radius_px=radius_px)
        self.add_text([x + padding_px, y + padding_px, width - 2 * padding_px, 34], title, font_size=18, color=title_color, bold=True)
        self.add_text(
            [x + padding_px, y + padding_px + 44, width - 2 * padding_px, height - 3 * padding_px - 34],
            body,
            font_size=14,
            color=body_color,
            wrap="square",
        )

    def add_timeline_stage(
        self,
        box_px: Box,
        *,
        number: str,
        period: str,
        heading: str = "",
        fill: str = "#4F8FEA",
        heading_color: str = "#1177C9",
    ) -> None:
        x, y, width, height = _box(box_px)
        node = min(height, 42)
        self.add_shape([x + node * 0.35, y, width - node * 0.35, height], kind="roundRect", fill=fill, stroke="none", radius_px=7)
        self.add_shape([x, y - (node - height) / 2, node, node], kind="ellipse", fill="#FFFFFF", stroke=fill, stroke_width=1.5)
        self.add_text([x, y - (node - height) / 2, node, node], number, font_size=16, color=fill, bold=True, align="center", valign="middle")
        self.add_text([x + node, y, width - node, height], period, font_size=15, color="#FFFFFF", align="center", valign="middle")
        if heading:
            self.add_text([x, y + height + 16, width, 34], heading, font_size=16, color=heading_color, bold=True, align="center")

    def as_dict(self) -> dict[str, Any]:
        return {
            "slide": {"width": self.slide_width, "height": self.slide_height, "background": self.background},
            "source": {"width_px": self.width_px, "height_px": self.height_px},
            "shapes": self.shapes,
            "images": self.images,
            "tables": self.tables,
            "text_boxes": self.text_boxes,
        }

    def write(self, path: str | Path) -> Path:
        output = Path(path)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(self.as_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return output
