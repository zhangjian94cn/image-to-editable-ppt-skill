#!/usr/bin/env python3
"""Portable advisory font measurement for slide authoring."""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path
from typing import Any

from PIL import ImageFont


FONT_DIRS = (
    Path("/System/Library/Fonts"),
    Path("/Library/Fonts"),
    Path.home() / "Library/Fonts",
)

PORTABLE_CJK_FALLBACKS = {
    "microsoft yahei": "PingFang SC",
    "microsoft yahei ui": "PingFang SC",
    "微软雅黑": "PingFang SC",
    "dengxian": "PingFang SC",
    "等线": "PingFang SC",
    "simhei": "Heiti SC",
    "黑体": "Heiti SC",
}


def _fontconfig_match(query: str) -> tuple[Path | None, str]:
    executable = shutil.which("fc-match")
    if not executable:
        return None, ""
    completed = subprocess.run(
        [executable, "-f", "%{family[0]}\t%{file}\n", query or "sans-serif"],
        capture_output=True,
        text=True,
        check=False,
    )
    first = completed.stdout.splitlines()[0] if completed.stdout.strip() else ""
    family, separator, filename = first.partition("\t")
    candidate = Path(filename) if separator else None
    if candidate and candidate.is_file():
        return candidate.resolve(), family.strip() or candidate.stem
    return None, ""


def find_font(preferred: str) -> tuple[Path | None, str]:
    query = preferred.strip()
    if query:
        direct = Path(query).expanduser()
        if direct.is_file():
            return direct.resolve(), direct.stem
    portable = PORTABLE_CJK_FALLBACKS.get(query.casefold(), query)
    for name in filter(None, (portable, query)):
        matched_path, matched_family = _fontconfig_match(name)
        if matched_path and (name == portable or matched_family.casefold() == query.casefold()):
            return matched_path, matched_family
    names = [portable, query, "PingFang SC", "Heiti SC", "Noto Sans CJK SC", "Arial Unicode MS"]
    for name in filter(None, names):
        token = name.casefold().replace(" ", "")
        for directory in FONT_DIRS:
            if not directory.is_dir():
                continue
            for suffix in ("*.ttf", "*.ttc", "*.otf"):
                for candidate in directory.rglob(suffix):
                    if token in candidate.stem.casefold().replace(" ", ""):
                        return candidate.resolve(), name
    matched_path, matched_family = _fontconfig_match(portable or query or "sans-serif")
    if matched_path:
        return matched_path, matched_family
    return None, query or "default"


def _font(path: Path | None, size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    if path:
        try:
            return ImageFont.truetype(str(path), size=size)
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
) -> dict[str, Any]:
    if width_px <= 0 or height_px <= 0:
        raise ValueError("width and height must be positive")
    font_path, font_name = find_font(preferred_font)
    text = str(text)
    chosen = min_font_px
    measured = (0, 0)
    for size in range(max_font_px, min_font_px - 1, -1):
        font = _font(font_path, size)
        lines = text.splitlines() or [""]
        widths = [font.getlength(line) for line in lines]
        bbox = font.getbbox("国Ag")
        line_height = max(1, bbox[3] - bbox[1]) * line_spacing
        candidate = (int(round(max(widths, default=0))), int(round(line_height * len(lines))))
        if candidate[0] <= width_px and candidate[1] <= height_px:
            chosen = size
            measured = candidate
            break
    font = _font(font_path, chosen)
    raw_width = int(round(font.getlength(text.replace("\n", ""))))
    result = {
        "schema_version": 1,
        "text": text,
        "box_px": [width_px, height_px],
        "font": font_name,
        "font_path": str(font_path) if font_path else "",
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
