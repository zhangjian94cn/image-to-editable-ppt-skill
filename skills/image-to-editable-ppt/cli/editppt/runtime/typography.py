#!/usr/bin/env python3
"""Semantic typography roles and named layer contracts."""

from __future__ import annotations

import statistics
from typing import Any


ROLE_SPECS: dict[str, dict[str, float]] = {
    "slide_title": {"min_px": 42, "default_px": 50, "max_px": 58, "shrink_warning": 0.90},
    "lead": {"min_px": 22, "default_px": 27, "max_px": 32, "shrink_warning": 0.90},
    "section_title": {"min_px": 26, "default_px": 32, "max_px": 38, "shrink_warning": 0.90},
    "subheading": {"min_px": 20, "default_px": 24, "max_px": 28, "shrink_warning": 0.90},
    "body": {"min_px": 17, "default_px": 21, "max_px": 25, "shrink_warning": 0.85},
    "metric": {"min_px": 28, "default_px": 36, "max_px": 48, "shrink_warning": 0.85},
    "caption": {"min_px": 13, "default_px": 16, "max_px": 19, "shrink_warning": 0.85},
    "footer": {"min_px": 9, "default_px": 12, "max_px": 15, "shrink_warning": 0.85},
}

LAYER_Z: dict[str, float] = {
    "background": 0,
    "container": 100,
    "decoration_behind": 120,
    "band": 140,
    "content": 200,
    "decoration_front": 240,
    "text": 300,
    "overlay": 400,
}


def scaled_role_specs(source_height: float) -> dict[str, dict[str, float]]:
    scale = max(float(source_height), 1.0) / 900.0
    return {
        role: {key: round(value * scale, 2) if key.endswith("_px") else value for key, value in spec.items()}
        for role, spec in ROLE_SPECS.items()
    }


def typography_hints(text_hints: dict[str, Any], source_height: float) -> dict[str, Any]:
    groups: dict[str, list[dict[str, Any]]] = {}
    for line in text_hints.get("lines", []):
        if not isinstance(line, dict):
            continue
        group = str(line.get("size_group") or "ungrouped")
        groups.setdefault(group, []).append(line)
    size_groups = []
    for name, lines in groups.items():
        glyphs = [float(line["glyph_height_px"]) for line in lines if line.get("glyph_height_px") is not None]
        font_points = [float(line["font_pt"]) for line in lines if line.get("font_pt") is not None]
        if not font_points:
            font_points = [float(line["font_pt_if_cjk"]) for line in lines if line.get("font_pt_if_cjk") is not None]
        size_groups.append(
            {
                "id": name,
                "line_count": len(lines),
                "median_glyph_height_px": round(statistics.median(glyphs), 2) if glyphs else None,
                "median_font_pt": round(statistics.median(font_points), 2) if font_points else None,
                "line_ids": [str(line.get("id") or "") for line in lines],
            }
        )
    size_groups.sort(key=lambda value: float(value.get("median_glyph_height_px") or 0), reverse=True)
    return {
        "schema_version": 1,
        "source_height_px": source_height,
        "fallback_roles": scaled_role_specs(source_height),
        "size_groups": size_groups,
        "instruction": (
            "Measured size groups are primary evidence. Codex assigns semantic text_role values; "
            "fallback role sizes apply only when measurement is unavailable."
        ),
    }


def resolve_layer(item: dict[str, Any], default_z: float) -> tuple[float, dict[str, Any] | None]:
    layer = str(item.get("layer") or "").strip()
    explicit = item.get("z_index")
    if layer and layer not in LAYER_Z:
        raise ValueError(f"unknown layer: {layer}")
    if explicit not in (None, ""):
        value = float(explicit)
        if layer and value != LAYER_Z[layer]:
            return value, {"layer": layer, "layer_z": LAYER_Z[layer], "z_index": value}
        return value, None
    return (LAYER_Z[layer] if layer else float(default_z)), None


def boxes_overlap(first: dict[str, Any], second: dict[str, Any]) -> bool:
    if not all(key in first for key in ("left", "top", "width", "height")):
        return False
    if not all(key in second for key in ("left", "top", "width", "height")):
        return False
    return (
        float(first["left"]) < float(second["left"]) + float(second["width"])
        and float(second["left"]) < float(first["left"]) + float(first["width"])
        and float(first["top"]) < float(second["top"]) + float(second["height"])
        and float(second["top"]) < float(first["top"]) + float(first["height"])
    )
