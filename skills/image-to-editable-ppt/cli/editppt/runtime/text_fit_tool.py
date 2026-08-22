#!/usr/bin/env python3
"""Portable advisory font measurement for slide authoring."""

from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any

from PIL import ImageFont

from font_registry import find_font, font_environment_fingerprint, resolve_font
from typography import ROLE_SPECS


@lru_cache(maxsize=2048)
def _font(path: Path | None, size: int, face_index: int = 0) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    if path:
        try:
            return ImageFont.truetype(str(path), size=size, index=face_index)
        except OSError:
            pass
    return ImageFont.load_default()


def measure(
    text: str,
    width_px: int,
    height_px: int,
    *,
    preferred_font: str = "",
    max_font_px: int = 120,
    min_font_px: int = 6,
    line_spacing: float = 1.2,
    single_line: bool = False,
    role: str = "",
    typography_profile: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if width_px <= 0 or height_px <= 0:
        raise ValueError("width and height must be positive")
    if role and role not in ROLE_SPECS:
        raise ValueError(f"unknown text role: {role}")
    role_profile = dict(ROLE_SPECS.get(role, {}))
    if typography_profile:
        configured = typography_profile.get("roles", {}).get(role, {}) if role else {}
        if isinstance(configured, dict):
            role_profile.update(configured)
    if role:
        max_font_px = min(max_font_px, int(round(float(role_profile.get("max_px", max_font_px)))))
        min_font_px = max(min_font_px, int(round(float(role_profile.get("min_px", min_font_px)))))
    face = resolve_font(preferred_font, require_cjk=any("\u4e00" <= char <= "\u9fff" for char in text))
    font_path = Path(face.path) if face else None
    font_name = face.family if face else (preferred_font or "default")
    face_index = face.face_index if face else 0
    text = str(text)
    chosen = min_font_px
    measured = (0, 0)
    for size in range(max_font_px, min_font_px - 1, -1):
        font = _font(font_path, size, face_index)
        lines = text.splitlines() or [""]
        widths = [font.getlength(line) for line in lines]
        bbox = font.getbbox("国Ag")
        line_height = max(1, bbox[3] - bbox[1]) * line_spacing
        candidate = (int(round(max(widths, default=0))), int(round(line_height * len(lines))))
        if candidate[0] <= width_px and candidate[1] <= height_px:
            chosen = size
            measured = candidate
            break
    font = _font(font_path, chosen, face_index)
    raw_width = int(round(font.getlength(text.replace("\n", ""))))
    result = {
        "schema_version": 1,
        "text": text,
        "box_px": [width_px, height_px],
        "font": font_name,
        "font_path": str(font_path) if font_path else "",
        "font_face_index": face_index,
        "font_sha256": face.sha256 if face else "",
        "font_provider": face.provider if face else "",
        "font_environment_fingerprint": font_environment_fingerprint(),
        "role": role,
        "role_profile": role_profile,
        "recommended_font_px": chosen,
        "measured_px": list(measured),
        "single_line": single_line,
        "single_line_width_px": raw_width,
        "fits_single_line": raw_width <= width_px,
        "overflow": measured[0] > width_px or measured[1] > height_px,
        "note": "Convert source pixels to slide points using the actual source-to-slide scale.",
    }
    if single_line and raw_width > width_px:
        result["warning"] = "single-line text does not fit at the recommended size"
    return result


def write_result(path: Path, result: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
