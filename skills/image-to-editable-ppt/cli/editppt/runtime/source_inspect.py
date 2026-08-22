#!/usr/bin/env python3
"""Advisory source-space layout and structure measurements."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import numpy as np
from PIL import Image


def _segments(mask: np.ndarray, minimum: int = 1) -> list[tuple[int, int]]:
    values = np.flatnonzero(mask)
    if not len(values):
        return []
    result: list[tuple[int, int]] = []
    start = previous = int(values[0])
    for value in values[1:]:
        value = int(value)
        if value > previous + 1:
            if previous - start + 1 >= minimum:
                result.append((start, previous + 1))
            start = value
        previous = value
    if previous - start + 1 >= minimum:
        result.append((start, previous + 1))
    return result


def _write(path: Path, payload: dict[str, Any]) -> dict[str, Any]:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return payload


def inspect_layout(source: Path, out: Path) -> dict[str, Any]:
    with Image.open(source).convert("RGB") as image:
        array = np.asarray(image)
        width, height = image.size
    # Ink relative to the locally dominant light background. The thresholds are
    # conservative and intentionally return broad regions rather than semantic
    # labels invented from pixels.
    darkness = 255.0 - np.mean(array.astype(np.float32), axis=2)
    chroma = np.max(array, axis=2).astype(np.int16) - np.min(array, axis=2).astype(np.int16)
    ink = (darkness > 18) | (chroma > 22)
    row_activity = ink.mean(axis=1)
    col_activity = ink.mean(axis=0)
    row_bands = _segments(row_activity > max(0.008, float(np.median(row_activity) * 1.3)), minimum=max(2, height // 600))
    col_bands = _segments(col_activity > max(0.008, float(np.median(col_activity) * 1.3)), minimum=max(2, width // 800))
    horizontal_whitespace = _segments(row_activity < 0.003, minimum=max(6, height // 80))
    vertical_whitespace = _segments(col_activity < 0.003, minimum=max(6, width // 80))
    payload = {
        "schema_version": 1,
        "provider": "local-pixel-layout",
        "source": str(source.resolve()),
        "size_px": [width, height],
        "row_bands": [[start, end] for start, end in row_bands],
        "column_bands": [[start, end] for start, end in col_bands],
        "horizontal_whitespace": [[start, end] for start, end in horizontal_whitespace],
        "vertical_whitespace": [[start, end] for start, end in vertical_whitespace],
        "note": "Coordinate evidence only; visually assign semantic regions.",
    }
    return _write(out, payload)


def _long_runs(mask: np.ndarray, horizontal: bool, minimum: int) -> list[list[int]]:
    result: list[list[int]] = []
    outer = mask.shape[0] if horizontal else mask.shape[1]
    for fixed in range(outer):
        values = mask[fixed, :] if horizontal else mask[:, fixed]
        for start, end in _segments(values, minimum=minimum):
            if horizontal:
                result.append([start, fixed, end, fixed])
            else:
                result.append([fixed, start, fixed, end])
    return result


def _dedupe_lines(lines: list[list[int]], horizontal: bool, tolerance: int = 3) -> list[list[int]]:
    if not lines:
        return []
    lines = sorted(lines, key=lambda line: (line[1] if horizontal else line[0], line[0] if horizontal else line[1]))
    groups: list[list[list[int]]] = []
    for line in lines:
        fixed = line[1] if horizontal else line[0]
        if not groups or abs(fixed - (groups[-1][-1][1] if horizontal else groups[-1][-1][0])) > tolerance:
            groups.append([line])
        else:
            groups[-1].append(line)
    merged: list[list[int]] = []
    for group in groups:
        if horizontal:
            merged.append([
                min(item[0] for item in group),
                round(sum(item[1] for item in group) / len(group)),
                max(item[2] for item in group),
                round(sum(item[3] for item in group) / len(group)),
            ])
        else:
            merged.append([
                round(sum(item[0] for item in group) / len(group)),
                min(item[1] for item in group),
                round(sum(item[2] for item in group) / len(group)),
                max(item[3] for item in group),
            ])
    return merged


def inspect_structure(source: Path, out: Path) -> dict[str, Any]:
    with Image.open(source).convert("RGB") as image:
        array = np.asarray(image).astype(np.int16)
        width, height = image.size
    gray = np.mean(array, axis=2)
    dx = np.abs(np.diff(gray, axis=1, prepend=gray[:, :1]))
    dy = np.abs(np.diff(gray, axis=0, prepend=gray[:1, :]))
    chroma = np.max(array, axis=2) - np.min(array, axis=2)
    horizontal_mask = (dy > 18) & ((chroma > 12) | (gray < 225))
    vertical_mask = (dx > 18) & ((chroma > 12) | (gray < 225))
    horizontal = _dedupe_lines(_long_runs(horizontal_mask, True, max(30, width // 18)), True)
    vertical = _dedupe_lines(_long_runs(vertical_mask, False, max(24, height // 15)), False)
    rectangles: list[list[int]] = []
    for top in horizontal:
        for bottom in horizontal:
            if bottom[1] - top[1] < max(16, height // 60):
                continue
            overlap_left = max(top[0], bottom[0])
            overlap_right = min(top[2], bottom[2])
            if overlap_right - overlap_left < max(40, width // 20):
                continue
            left_candidates = [line for line in vertical if abs(line[0] - overlap_left) <= 5 and line[1] <= top[1] + 5 and line[3] >= bottom[1] - 5]
            right_candidates = [line for line in vertical if abs(line[0] - overlap_right) <= 5 and line[1] <= top[1] + 5 and line[3] >= bottom[1] - 5]
            if left_candidates and right_candidates:
                rectangles.append([overlap_left, top[1], overlap_right, bottom[1]])
    # Keep largest unique candidates; nested candidates are useful evidence.
    unique = sorted({tuple(item) for item in rectangles}, key=lambda box: (-(box[2] - box[0]) * (box[3] - box[1]), box))[:120]
    payload = {
        "schema_version": 1,
        "provider": "local-source-space-edges",
        "source": str(source.resolve()),
        "size_px": [width, height],
        "horizontal_segments": horizontal[:300],
        "vertical_segments": vertical[:300],
        "rectangle_candidates": [list(value) for value in unique],
        "note": "Candidates are geometric evidence, not object IDs or containment decisions.",
    }
    return _write(out, payload)
