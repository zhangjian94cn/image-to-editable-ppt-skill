"""Portable, deterministic image evidence for one multimodal page task.

The page agent always receives the whole authoring image. This module adds
overlapping, source-pixel detail views so small slide text remains readable
without depending on an operating-system OCR framework.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

from PIL import Image

from editppt.source_space import map_authoring_box_to_original, read_source_map


DEFAULT_OUTPUT = Path(".editppt/vision-inputs")
MANIFEST_NAME = "vision-inputs.json"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _save_detail(image: Image.Image, box: tuple[int, int, int, int], output: Path, target_edge: int) -> None:
    crop = image.crop(box).convert("RGB")
    width, height = crop.size
    longest = max(width, height)
    if longest and longest != target_edge:
        scale = target_edge / longest
        crop = crop.resize(
            (max(1, round(width * scale)), max(1, round(height * scale))),
            Image.Resampling.LANCZOS,
        )
    output.parent.mkdir(parents=True, exist_ok=True)
    crop.save(output, format="PNG")


def prepare_visual_inputs(
    page_dir: str | Path,
    *,
    source_name: str = "source.png",
    output_dir: str | Path = DEFAULT_OUTPUT,
    target_edge: int = 1792,
    overlap_ratio: float = 0.08,
) -> dict[str, Any]:
    """Create four overlapping quadrant views plus three footer detail views."""

    page_dir = Path(page_dir).expanduser().resolve()
    source = (page_dir / source_name).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if target_edge < 512 or target_edge > 4096:
        raise ValueError("target_edge must be between 512 and 4096")
    if not 0.0 <= overlap_ratio <= 0.25:
        raise ValueError("overlap_ratio must be between 0 and 0.25")

    mapping = read_source_map(source)
    original = Path(str(mapping.get("original_path") or source)).resolve()
    out = (page_dir / output_dir).resolve()
    allowed = (page_dir / ".editppt").resolve()
    if allowed != out and allowed not in out.parents:
        raise ValueError("visual input output must stay inside page_dir/.editppt")
    out.mkdir(parents=True, exist_ok=True)
    for stale in out.glob("*.png"):
        stale.unlink()

    with Image.open(source) as view_image, Image.open(original) as original_image:
        authoring_width, authoring_height = view_image.size
        overlap_x = round(authoring_width * overlap_ratio)
        overlap_y = round(authoring_height * overlap_ratio)
        half_width = authoring_width / 2
        half_height = authoring_height / 2
        authoring_boxes: list[tuple[str, tuple[int, int, int, int]]] = []
        for row in range(2):
            for column in range(2):
                left = max(0, round(column * half_width) - (overlap_x if column else 0))
                top = max(0, round(row * half_height) - (overlap_y if row else 0))
                right = min(
                    authoring_width,
                    round((column + 1) * half_width) + (overlap_x if column == 0 else 0),
                )
                bottom = min(
                    authoring_height,
                    round((row + 1) * half_height) + (overlap_y if row == 0 else 0),
                )
                authoring_boxes.append((f"detail-r{row + 1}c{column + 1}", (left, top, right, bottom)))
        footer_top = max(0, round(authoring_height * 0.94))
        footer_width = authoring_width / 3
        footer_overlap = round(authoring_width * 0.04)
        for column in range(3):
            left = max(0, round(column * footer_width) - (footer_overlap if column else 0))
            right = min(
                authoring_width,
                round((column + 1) * footer_width) + (footer_overlap if column < 2 else 0),
            )
            authoring_boxes.append(
                (f"detail-footer-c{column + 1}", (left, footer_top, right, authoring_height))
            )

        entries: list[dict[str, Any]] = []
        for name, authoring_box in authoring_boxes:
            original_box = map_authoring_box_to_original(authoring_box, mapping) if mapping else authoring_box
            original_box = (
                max(0, original_box[0]),
                max(0, original_box[1]),
                min(original_image.width, original_box[2]),
                min(original_image.height, original_box[3]),
            )
            path = out / f"{name}.png"
            _save_detail(original_image, original_box, path, target_edge)
            entries.append({
                "id": name,
                "path": str(path.relative_to(page_dir)),
                "source_box_px": [
                    authoring_box[0],
                    authoring_box[1],
                    authoring_box[2] - authoring_box[0],
                    authoring_box[3] - authoring_box[1],
                ],
                "sha256": _sha256(path),
            })

    payload = {
        "schema_version": 1,
        "provider": "deterministic-source-crops",
        "source": source_name,
        "source_size_px": [authoring_width, authoring_height],
        "instruction": (
            "These attached detail views are source evidence, not separate slides. "
            "Use source_box_px to map their text back into source.png coordinates."
        ),
        "images": entries,
    }
    manifest = out / MANIFEST_NAME
    manifest.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    payload["manifest"] = str(manifest)
    return payload


def visual_input_paths(page_dir: str | Path) -> list[Path]:
    page_dir = Path(page_dir).expanduser().resolve()
    manifest = page_dir / DEFAULT_OUTPUT / MANIFEST_NAME
    try:
        payload = json.loads(manifest.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    paths: list[Path] = []
    for entry in payload.get("images") or []:
        if not isinstance(entry, dict):
            continue
        path = (page_dir / str(entry.get("path") or "")).resolve()
        if path.is_file() and page_dir in path.parents:
            paths.append(path)
    return paths
