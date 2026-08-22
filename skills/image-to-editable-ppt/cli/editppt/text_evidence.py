"""Geometry-aware text evidence shared by PPTX inspection and benchmarks.

The source extractor may keep several visual paragraphs in one PowerPoint
shape, while an editable reconstruction may split that shape into smaller
text boxes.  Coverage therefore compares source paragraphs inside their
original region instead of requiring one long, z-order-dependent string.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Iterable, Sequence


_DECORATIVE = str.maketrans({
    "丨": "|", "｜": "|", "︱": "|",
    "➢": "", "➤": "", "▶": "", "►": "", "▸": "",
    "•": "", "●": "", "○": "", "▪": "", "■": "", "□": "",
})
_TOKEN = re.compile(r"[A-Za-z][A-Za-z0-9+./_-]*|\d+(?:\.\d+)?%?|[\u3400-\u9fff]+")


def normalize_text(value: str) -> str:
    """Normalize presentation-only variance without correcting source words."""

    normalized = unicodedata.normalize("NFKC", str(value or ""))
    normalized = re.sub(r"\\(?:geq?|ge)\b", "≥", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\\(?:leq?|le)\b", "≤", normalized, flags=re.IGNORECASE)
    normalized = normalized.replace("$", "").replace("\\(", "").replace("\\)", "")
    normalized = normalized.translate(_DECORATIVE)
    return re.sub(r"[\s\u000b]+", "", normalized).casefold()


def paragraph_segments(item: dict[str, Any]) -> list[str]:
    """Return semantic paragraphs from one extracted source text object."""

    paragraphs = ((item.get("text_style") or {}).get("paragraphs") or [])
    values: list[str] = []
    for paragraph in paragraphs:
        if not isinstance(paragraph, dict):
            continue
        runs = paragraph.get("runs") or []
        text = "".join(str(run.get("text") or "") for run in runs if isinstance(run, dict)).strip()
        if text:
            values.append(text)
    if values:
        return values
    return [part.strip() for part in str(item.get("text") or "").splitlines() if part.strip()]


def _intersection(first: Sequence[float], second: Sequence[float]) -> float:
    if len(first) != 4 or len(second) != 4:
        return 0.0
    left, top, width, height = [float(value) for value in first]
    other_left, other_top, other_width, other_height = [float(value) for value in second]
    overlap_width = max(0.0, min(left + width, other_left + other_width) - max(left, other_left))
    overlap_height = max(0.0, min(top + height, other_top + other_height) - max(top, other_top))
    return overlap_width * overlap_height


def _nearby(expected_box: Sequence[float], candidate_box: Sequence[float]) -> bool:
    if len(expected_box) != 4 or len(candidate_box) != 4:
        return True
    expected_area = max(1.0, float(expected_box[2]) * float(expected_box[3]))
    candidate_area = max(1.0, float(candidate_box[2]) * float(candidate_box[3]))
    overlap = _intersection(expected_box, candidate_box)
    if overlap / min(expected_area, candidate_area) >= 0.08:
        return True
    left, top, width, height = [float(value) for value in expected_box]
    cx = float(candidate_box[0]) + float(candidate_box[2]) / 2
    cy = float(candidate_box[1]) + float(candidate_box[3]) / 2
    margin_x = max(8.0, width * 0.08)
    margin_y = max(8.0, height * 0.12)
    return left - margin_x <= cx <= left + width + margin_x and top - margin_y <= cy <= top + height + margin_y


def _tokens(value: str) -> list[str]:
    return [normalize_text(match.group(0)) for match in _TOKEN.finditer(unicodedata.normalize("NFKC", value))]


def segment_is_covered(segment: str, candidates: Iterable[str]) -> bool:
    """Match exact text, then allow reordered labels/values within one region."""

    expected = normalize_text(segment)
    candidate_values = [normalize_text(value) for value in candidates if normalize_text(value)]
    if not expected:
        return True
    joined = "".join(candidate_values)
    if expected in joined or any(expected in value for value in candidate_values):
        return True
    atoms = [value for value in _tokens(segment) if value]
    # Reordered matching is only for genuinely grouped label/value regions.
    # A prose sentence usually yields one long CJK atom, so paraphrases and
    # misspellings remain visible instead of being silently accepted.
    return len(atoms) >= 2 and all(atom in joined for atom in atoms)


def region_text_coverage(
    expected_regions: Sequence[dict[str, Any]],
    candidate_regions: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    """Compare source text paragraphs with editable text in the same region."""

    expected_segments: list[dict[str, Any]] = []
    for region in expected_regions:
        values = region.get("segments") or [region.get("text") or ""]
        for value in values:
            if normalize_text(str(value)):
                expected_segments.append({"text": str(value), "box_px": region.get("box_px")})

    missing: list[str] = []
    all_candidate_texts = [str(region.get("text") or "") for region in candidate_regions]
    for expected in expected_segments:
        box = expected.get("box_px")
        nearby = [
            region for region in candidate_regions
            if not box or not region.get("box_px") or _nearby(box, region["box_px"])
        ]
        nearby.sort(key=lambda value: (float((value.get("box_px") or [0, 0])[1]), float((value.get("box_px") or [0, 0])[0])))
        exact = normalize_text(expected["text"])
        exact_anywhere = any(exact and exact in normalize_text(value) for value in all_candidate_texts)
        if not exact_anywhere and not segment_is_covered(
            expected["text"],
            [str(value.get("text") or "") for value in nearby],
        ):
            missing.append(expected["text"])

    total = len(expected_segments)
    matched = total - len(missing)
    return {
        "expected_text_count": total,
        "matched_text_count": matched,
        "missing_text_count": len(missing),
        "missing_texts": missing,
        "text_coverage": round(matched / total, 6) if total else None,
    }
