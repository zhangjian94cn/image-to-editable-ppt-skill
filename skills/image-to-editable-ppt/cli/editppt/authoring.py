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
    typography: dict[str, Any] = field(
        default_factory=lambda: {"font_family": "Microsoft YaHei", "roles": {}}
    )
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
        box_or_kind: Box | str | None = None,
        kind: str | None = None,
        *,
        box_px: Box | None = None,
        points_px: Sequence[Point] | Sequence[float] | None = None,
        fill: str = "none",
        stroke: str = "none",
        stroke_width: float = 1,
        radius_px: float | None = None,
        preset: str = "",
        layer: str = "container",
        z_index: float | None = None,
        **extra: Any,
    ) -> dict[str, Any]:
        if isinstance(box_or_kind, str):
            if kind is not None and kind != box_or_kind:
                raise ValueError("shape kind was provided twice")
            resolved_kind = box_or_kind
            resolved_box = box_px
        else:
            resolved_kind = kind or "rect"
            if box_or_kind is not None and box_px is not None:
                raise ValueError("shape box was provided twice")
            resolved_box = box_or_kind if box_or_kind is not None else box_px

        if points_px is not None:
            if resolved_kind not in {"line", "polyline"}:
                raise ValueError("points_px is only valid for line or polyline shapes")
            values = list(points_px)
            if len(values) == 4 and not any(isinstance(value, (list, tuple)) for value in values):
                x1, y1, x2, y2 = [float(value) for value in values]
                item = {
                    "type": "line",
                    "points_px": [x1, y1, x2, y2],
                    "fill": "none",
                    "stroke": stroke,
                    "stroke_width": float(stroke_width),
                }
                item["layer"] = layer
                if z_index is not None:
                    item["z_index"] = float(z_index)
                item.update(extra)
                self.shapes.append(item)
                return item
            return self.add_polyline(
                values,
                stroke=stroke,
                stroke_width=stroke_width,
                layer=layer,
                z_index=z_index,
                **extra,
            )
        if resolved_box is None:
            raise ValueError("box_px is required unless line/polyline points_px is provided")
        item: dict[str, Any] = {
            "type": resolved_kind,
            "box_px": _box(resolved_box),
            "fill": fill,
            "stroke": stroke,
            "stroke_width": float(stroke_width),
            "layer": layer,
        }
        if z_index is not None:
            item["z_index"] = float(z_index)
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
        layer: str = "content",
        z_index: float | None = None,
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
            "layer": layer,
        }
        if z_index is not None:
            item["z_index"] = float(z_index)
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
        layer: str = "container",
        z_index: float | None = None,
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
            "layer": layer,
        }
        if z_index is not None:
            item["z_index"] = float(z_index)
        item.update(extra)
        self.shapes.append(item)
        return item

    def add_polyline(
        self,
        points_px: Sequence[Point],
        *,
        stroke: str = "#4472C4",
        stroke_width: float = 1.25,
        dash: str = "",
        start_arrow: str = "none",
        end_arrow: str = "none",
        layer: str = "content",
        z_index: float | None = None,
        **extra: Any,
    ) -> dict[str, Any]:
        """Add an open editable source-space path with two or more vertices."""

        points = [_point(value) for value in points_px]
        if len(points) < 2:
            raise ValueError("polyline requires at least two source-space points")
        xs = [value[0] for value in points]
        ys = [value[1] for value in points]
        item: dict[str, Any] = {
            "type": "polyline",
            "box_px": [min(xs), min(ys), max(1.0, max(xs) - min(xs)), max(1.0, max(ys) - min(ys))],
            "polyline_px": [[x, y] for x, y in points],
            "fill": "none",
            "stroke": stroke,
            "stroke_width": float(stroke_width),
            "start_arrow": start_arrow,
            "end_arrow": end_arrow,
            "layer": layer,
        }
        if z_index is not None:
            item["z_index"] = float(z_index)
        if dash:
            item["dash"] = dash
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
        font: str | None = None,
        font_size: float | None = None,
        font_size_px: float | None = None,
        color: str = "#111111",
        bold: bool = False,
        align: str = "left",
        valign: str = "top",
        wrap: str = "none",
        fit_text: bool = True,
        text_role: str = "body",
        layer: str = "text",
        z_index: float | None = None,
        **extra: Any,
    ) -> dict[str, Any]:
        supplied = int(text is not None) + int(runs is not None) + int(paragraphs is not None)
        if supplied != 1:
            raise ValueError("provide exactly one of text, runs, or paragraphs")
        if font_size is not None and font_size_px is not None:
            raise ValueError("provide font_size in points or font_size_px, not both")
        if font_size_px is not None:
            font_size = float(font_size_px) * self.slide_height * 72.0 / self.height_px
        item: dict[str, Any] = {
            "box_px": _box(box_px),
            "color": color,
            "bold": bool(bold),
            "align": align,
            "valign": valign,
            "wrap": wrap,
            "fit_text": bool(fit_text),
            "text_role": text_role,
            "layer": layer,
        }
        if font:
            item["font"] = font
        if font_size is not None:
            item["font_size"] = float(font_size)
        if z_index is not None:
            item["z_index"] = float(z_index)
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
        layer: str = "content",
        z_index: float | None = None,
    ) -> dict[str, Any]:
        if not alt.strip():
            raise ValueError("image alt text is required")
        item = {
            "path": str(path),
            "box_px": _box(box_px),
            "alt": alt.strip(),
            "layer": layer,
        }
        if z_index is not None:
            item["z_index"] = float(z_index)
        self.images.append(item)
        return item

    def add_table(
        self,
        box_px: Box,
        rows: Sequence[Sequence[Any]],
        *,
        column_widths: Sequence[float] | None = None,
        row_heights: Sequence[float] | None = None,
        font: str = "Microsoft YaHei",
        font_size: float | None = None,
        font_size_px: float | None = None,
        header_fill: str = "#0A65B7",
        header_color: str = "#FFFFFF",
        fill: str = "#FFFFFF",
        color: str = "#111111",
        border: str = "#B7C4D4",
        text_role: str = "body",
        layer: str = "content",
        z_index: float | None = None,
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
            "text_role": text_role,
            "layer": layer,
        }
        if z_index is not None:
            item["z_index"] = float(z_index)
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
        self.add_shape(box, kind="roundRect" if radius_px else "rect", fill=fill, radius_px=radius_px, layer="band")
        self.add_text(
            box,
            label,
            font_size=font_size,
            text_role="section_title",
            color=color,
            bold=True,
            align="center",
            valign="middle",
        )

    def add_layered_header(
        self,
        container_box_px: Box,
        band_box_px: Box,
        accent_box_px: Box,
        title: str,
        *,
        title_box_px: Box | None = None,
        container_fill: str = "#F5FAFD",
        container_stroke: str = "#D5E3F1",
        band_fill: str = "#B8E0F2",
        accent_fill: str = "#0795D2",
        title_color: str = "#0087CF",
        title_font_size_px: float | None = None,
        radius_px: float = 16,
    ) -> dict[str, dict[str, Any]]:
        """Add a card header whose accent is intentionally masked by the band."""

        container = self.add_shape(
            container_box_px,
            kind="roundRect",
            fill=container_fill,
            stroke=container_stroke,
            radius_px=radius_px,
            layer="container",
        )
        accent = self.add_shape(
            accent_box_px,
            kind="roundRect",
            fill=accent_fill,
            stroke="none",
            radius_px=radius_px / 2,
            layer="decoration_behind",
        )
        band = self.add_shape(
            band_box_px,
            kind="roundRect",
            fill=band_fill,
            stroke="none",
            radius_px=radius_px,
            layer="band",
        )
        text_box = title_box_px or band_box_px
        title_item = self.add_text(
            text_box,
            title,
            font_size_px=title_font_size_px,
            text_role="section_title",
            color=title_color,
            bold=True,
            align="center",
            valign="middle",
            wrap="none",
            layer="text",
        )
        return {"container": container, "accent": accent, "band": band, "title": title_item}

    def add_editable_bar_chart(
        self,
        box_px: Box,
        labels: Sequence[str],
        values: Sequence[float],
        *,
        title: str = "",
        maximum: float | None = None,
        grid_values: Sequence[float] | None = None,
        fill: str = "#FFFFFF",
        bar_fill: str = "#5C9BC2",
        grid_color: str = "#D9D9D9",
        text_color: str = "#111111",
        title_font_size_px: float = 24,
        label_font_size_px: float = 14,
        tick_font_size_px: float = 14,
        show_values: bool = False,
    ) -> dict[str, list[dict[str, Any]]]:
        """Build a deterministic native-shape bar chart in source coordinates."""

        if not labels or len(labels) != len(values):
            raise ValueError("bar chart labels and values must be non-empty and have equal length")
        numeric = [float(value) for value in values]
        if any(value < 0 for value in numeric):
            raise ValueError("bar chart values must be non-negative")
        upper = float(maximum) if maximum is not None else max(numeric, default=0.0)
        if upper <= 0 or any(value > upper for value in numeric):
            raise ValueError("bar chart maximum must be positive and cover every value")
        ticks = [float(value) for value in (grid_values or [0, upper / 2, upper])]
        if any(value < 0 or value > upper for value in ticks):
            raise ValueError("bar chart grid values must be within 0..maximum")

        x, y, width, height = _box(box_px)
        title_height = max(0.0, title_font_size_px * 1.7 if title else 0.0)
        label_height = max(label_font_size_px * 2.6, height * 0.16)
        tick_width = max(tick_font_size_px * 3.2, width * 0.08)
        plot_left = x + tick_width
        plot_top = y + title_height + tick_font_size_px * 0.4
        plot_width = max(1.0, width - tick_width - label_font_size_px * 0.8)
        plot_height = max(1.0, height - title_height - label_height - tick_font_size_px * 0.8)

        result: dict[str, list[dict[str, Any]]] = {"background": [], "grid": [], "bars": [], "labels": []}
        result["background"].append(
            self.add_shape([x, y, width, height], fill=fill, stroke="none", layer="container")
        )
        if title:
            result["labels"].append(
                self.add_text(
                    [x, y, width, title_height], title,
                    font_size_px=title_font_size_px, text_role="subheading",
                    bold=True, align="center", valign="middle", wrap="none",
                )
            )
        for tick in sorted(set(ticks)):
            py = plot_top + plot_height * (1.0 - tick / upper)
            result["grid"].append(
                self.add_connector(
                    [plot_left, py], [plot_left + plot_width, py], color=grid_color,
                    width=1, start_arrow="none", end_arrow="none", layer="decoration_behind",
                )
            )
            result["labels"].append(
                self.add_text(
                    [x, py - tick_font_size_px, tick_width * 0.82, tick_font_size_px * 2],
                    f"{tick:g}", font_size_px=tick_font_size_px, text_role="caption",
                    color=text_color, align="right", valign="middle", wrap="none",
                )
            )

        slot = plot_width / len(numeric)
        bar_width = slot * 0.46
        for index, (label, value) in enumerate(zip(labels, numeric)):
            center = plot_left + slot * (index + 0.5)
            bar_height = plot_height * value / upper
            result["bars"].append(
                self.add_shape(
                    [center - bar_width / 2, plot_top + plot_height - bar_height, bar_width, max(1.0, bar_height)],
                    fill=bar_fill, stroke="none", layer="content",
                )
            )
            result["labels"].append(
                self.add_text(
                    [center - slot * 0.48, plot_top + plot_height + 2, slot * 0.96, label_height - 2],
                    str(label), font_size_px=label_font_size_px, text_role="caption",
                    color=text_color, align="center", valign="top", wrap="square",
                )
            )
            if show_values:
                result["labels"].append(
                    self.add_text(
                        [center - slot * 0.48, plot_top + plot_height - bar_height - label_font_size_px * 1.8, slot * 0.96, label_font_size_px * 1.7],
                        f"{value:g}", font_size_px=label_font_size_px, text_role="caption",
                        color=text_color, align="center", valign="bottom", wrap="none",
                    )
                )
        return result

    def add_editable_line_chart(
        self,
        box_px: Box,
        labels: Sequence[str],
        values: Sequence[float],
        *,
        title: str = "",
        minimum: float = 0,
        maximum: float | None = None,
        grid_values: Sequence[float] | None = None,
        fill: str = "#FFFFFF",
        line_color: str = "#2387D0",
        marker_fill: str = "#F05A67",
        grid_color: str = "#D9D9D9",
        text_color: str = "#111111",
        title_font_size_px: float = 24,
        label_font_size_px: float = 14,
        tick_font_size_px: float = 14,
        show_values: bool = True,
    ) -> dict[str, list[dict[str, Any]]]:
        """Build a native editable line chart with deterministic typography."""

        if len(labels) < 2 or len(labels) != len(values):
            raise ValueError("line chart needs at least two labels and matching values")
        numeric = [float(value) for value in values]
        lower = float(minimum)
        upper = float(maximum) if maximum is not None else max(numeric)
        if upper <= lower or any(value < lower or value > upper for value in numeric):
            raise ValueError("line chart bounds must cover every value")
        ticks = [float(value) for value in (grid_values or [lower, (lower + upper) / 2, upper])]
        if any(value < lower or value > upper for value in ticks):
            raise ValueError("line chart grid values must be within the declared bounds")

        x, y, width, height = _box(box_px)
        title_height = max(0.0, title_font_size_px * 1.7 if title else 0.0)
        label_height = max(label_font_size_px * 1.9, height * 0.12)
        tick_width = max(tick_font_size_px * 3.2, width * 0.08)
        plot_left = x + tick_width
        plot_top = y + title_height + label_font_size_px * 1.8
        plot_width = max(1.0, width - tick_width - label_font_size_px)
        plot_height = max(1.0, height - title_height - label_height - label_font_size_px * 2.2)

        result: dict[str, list[dict[str, Any]]] = {"background": [], "grid": [], "line": [], "labels": []}
        result["background"].append(
            self.add_shape([x, y, width, height], fill=fill, stroke="none", layer="container")
        )
        if title:
            result["labels"].append(
                self.add_text(
                    [x, y, width, title_height], title,
                    font_size_px=title_font_size_px, text_role="subheading",
                    bold=True, align="center", valign="middle", wrap="none",
                )
            )
        for tick in sorted(set(ticks)):
            py = plot_top + plot_height * (1.0 - (tick - lower) / (upper - lower))
            result["grid"].append(
                self.add_connector(
                    [plot_left, py], [plot_left + plot_width, py], color=grid_color,
                    width=1, start_arrow="none", end_arrow="none", layer="decoration_behind",
                )
            )
            result["labels"].append(
                self.add_text(
                    [x, py - tick_font_size_px, tick_width * 0.82, tick_font_size_px * 2],
                    f"{tick:g}", font_size_px=tick_font_size_px, text_role="caption",
                    color=text_color, align="right", valign="middle", wrap="none",
                )
            )

        step = plot_width / (len(numeric) - 1)
        points = [
            [plot_left + step * index, plot_top + plot_height * (1.0 - (value - lower) / (upper - lower))]
            for index, value in enumerate(numeric)
        ]
        result["line"].append(self.add_polyline(points, stroke=line_color, stroke_width=3, layer="content"))
        marker = max(5.0, label_font_size_px * 0.55)
        for index, (label, value, point) in enumerate(zip(labels, numeric, points)):
            px, py = point
            result["line"].append(
                self.add_shape(
                    [px - marker / 2, py - marker / 2, marker, marker],
                    kind="ellipse", fill=marker_fill, stroke="none", layer="decoration_front",
                )
            )
            label_width = step if index not in {0, len(numeric) - 1} else step * 0.9
            label_left = px - label_width / 2
            if index == 0:
                label_left = px
            elif index == len(numeric) - 1:
                label_left = px - label_width
            result["labels"].append(
                self.add_text(
                    [label_left, plot_top + plot_height + 2, label_width, label_height - 2],
                    str(label), font_size_px=label_font_size_px, text_role="caption",
                    color=text_color, bold=True, align="center", valign="top", wrap="none",
                )
            )
            if show_values:
                result["labels"].append(
                    self.add_text(
                        [label_left, py - label_font_size_px * 2.1, label_width, label_font_size_px * 1.8],
                        f"{value:g}", font_size_px=label_font_size_px, text_role="caption",
                        color=text_color, bold=True, align="center", valign="bottom", wrap="none",
                    )
                )
        return result

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
            "typography": self.typography,
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
