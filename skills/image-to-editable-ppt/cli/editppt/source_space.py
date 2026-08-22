"""Stable authoring coordinates for visual slide reconstruction.

One stable authoring view keeps geometry consistent across models and hosts.
Portable overlapping detail images preserve small-text readability, while
this module retains the original pixels for lossless asset extraction.
"""

from __future__ import annotations

import hashlib
import json
import math
import shutil
from pathlib import Path
from typing import Any

from PIL import Image


DEFAULT_AUTHORING_MAX_WIDTH = 2048
MAP_PATH = Path(".editppt/source-map.json")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def prepare_authoring_source(
    source: str | Path,
    page_dir: str | Path,
    *,
    max_width: int = DEFAULT_AUTHORING_MAX_WIDTH,
) -> dict[str, Any]:
    """Write ``source.png`` in a stable vision-sized coordinate space.

    The untouched input is copied under ``.editppt/source/``.  A sidecar map
    lets ``editppt assets crop`` convert authoring coordinates back to the
    original image and therefore preserve original pixels.
    """

    source = Path(source).expanduser().resolve()
    page_dir = Path(page_dir).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if max_width <= 0:
        raise ValueError("max_width must be positive")
    page_dir.mkdir(parents=True, exist_ok=True)
    view = page_dir / "source.png"
    if source == view.resolve():
        existing = read_source_map(view)
        if existing:
            return existing
    evidence_dir = page_dir / ".editppt/source"
    evidence_dir.mkdir(parents=True, exist_ok=True)
    suffix = source.suffix.lower() if source.suffix else ".bin"
    original = evidence_dir / f"original{suffix}"
    if source != original:
        shutil.copy2(source, original)

    with Image.open(original) as image:
        original_width, original_height = image.size
        if original_width <= 0 or original_height <= 0:
            raise ValueError("source dimensions must be positive")
        authoring_width = min(original_width, int(max_width))
        authoring_height = max(1, int(round(original_height * authoring_width / original_width)))
        rgb = image.convert("RGB")
        if (authoring_width, authoring_height) != image.size:
            rgb = rgb.resize((authoring_width, authoring_height), Image.Resampling.LANCZOS)
        rgb.save(view, format="PNG")

    payload = {
        "schema_version": 1,
        "coordinate_space": "authoring",
        "authoring_max_width": int(max_width),
        "authoring_size_px": [authoring_width, authoring_height],
        "original_size_px": [original_width, original_height],
        "scale_to_original": [
            original_width / authoring_width,
            original_height / authoring_height,
        ],
        "view_path": str(view.resolve()),
        "original_path": str(original.resolve()),
        "view_sha256": _sha256(view),
        "original_sha256": _sha256(original),
    }
    map_path = page_dir / MAP_PATH
    map_path.parent.mkdir(parents=True, exist_ok=True)
    map_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return payload


def read_source_map(view: str | Path) -> dict[str, Any]:
    """Return a trusted map only when it is bound to the supplied view."""

    view = Path(view).expanduser().resolve()
    path = view.parent / MAP_PATH
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict) or Path(str(payload.get("view_path") or "")).resolve() != view:
        return {}
    original = Path(str(payload.get("original_path") or "")).expanduser().resolve()
    allowed_root = (view.parent / ".editppt/source").resolve()
    if not original.is_file() or original.parent != allowed_root:
        return {}
    try:
        if payload.get("view_sha256") != _sha256(view) or payload.get("original_sha256") != _sha256(original):
            return {}
    except OSError:
        return {}
    return payload


def map_authoring_box_to_original(
    box: tuple[int, int, int, int],
    mapping: dict[str, Any],
    *,
    pad: int = 0,
) -> tuple[int, int, int, int]:
    """Map a half-open authoring box to an enclosing original-pixel box."""

    scale_x, scale_y = [float(value) for value in mapping["scale_to_original"]]
    left, top, right, bottom = box
    return (
        math.floor((left - pad) * scale_x),
        math.floor((top - pad) * scale_y),
        math.ceil((right + pad) * scale_x),
        math.ceil((bottom + pad) * scale_y),
    )
