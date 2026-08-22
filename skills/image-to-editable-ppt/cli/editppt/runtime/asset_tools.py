#!/usr/bin/env python3
"""Source-bound asset extraction helpers with auditable coordinates."""

from __future__ import annotations

import json
import shutil
from collections import deque
from pathlib import Path
from typing import Any

import numpy as np
from PIL import Image

from editppt.source_space import map_authoring_box_to_original, read_source_map


def _write(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def crop(source: Path, out: Path, box: tuple[int, int, int, int], pad: int = 0) -> dict[str, Any]:
    source = source.resolve()
    requested_source = source
    authoring_box = list(box)
    mapping = read_source_map(source)
    if mapping:
        source = Path(mapping["original_path"]).resolve()
        box = map_authoring_box_to_original(box, mapping, pad=pad)
        pad = 0
    with Image.open(source) as image:
        width, height = image.size
        left, top, right, bottom = box
        left = max(0, left - pad)
        top = max(0, top - pad)
        right = min(width, right + pad)
        bottom = min(height, bottom + pad)
        if right <= left or bottom <= top:
            raise ValueError("crop box is empty after clipping")
        out.parent.mkdir(parents=True, exist_ok=True)
        image.crop((left, top, right, bottom)).save(out)
    coverage = ((right - left) * (bottom - top)) / max(1, width * height)
    result = {
        "status": "ready",
        "operation": "original-source-pixel-crop" if mapping else "source-pixel-crop",
        "source": str(source.resolve()),
        "output": str(out.resolve()),
        "source_size_px": [width, height],
        "box_px": [left, top, right, bottom],
        "coverage": round(coverage, 6),
        "near_full_page_risk": coverage >= 0.8,
    }
    if mapping:
        result.update({
            "authoring_source": str(requested_source),
            "authoring_size_px": mapping["authoring_size_px"],
            "authoring_box_px": authoring_box,
            "scale_to_original": mapping["scale_to_original"],
        })
    _write(out.with_suffix(out.suffix + ".json"), result)
    return result


def _edge_background(array: np.ndarray) -> np.ndarray:
    border = np.concatenate((array[0], array[-1], array[:, 0], array[:, -1]), axis=0)
    # Median is robust to a logo touching a small part of one edge.
    return np.median(border[:, :3].astype(np.float32), axis=0)


def separate_flat_background(
    source: Path,
    out: Path,
    *,
    tolerance: float = 24.0,
    feather: float = 14.0,
) -> dict[str, Any]:
    """Remove a flat edge-connected background without repainting foreground pixels."""

    with Image.open(source).convert("RGBA") as image:
        array = np.asarray(image).copy()
    rgb = array[:, :, :3].astype(np.float32)
    background = _edge_background(array)
    distance = np.sqrt(np.sum((rgb - background[None, None, :]) ** 2, axis=2))
    # Only clear background connected to the image edge. This keeps white holes
    # inside a logo or diagram instead of punching through them.
    candidate = distance <= tolerance + feather
    height, width = candidate.shape
    connected = np.zeros_like(candidate, dtype=bool)
    queue: deque[tuple[int, int]] = deque()
    queued = np.zeros_like(candidate, dtype=bool)

    def enqueue(y: int, x: int) -> None:
        if candidate[y, x] and not queued[y, x]:
            queued[y, x] = True
            queue.append((y, x))

    for x in range(width):
        enqueue(0, x)
        enqueue(height - 1, x)
    for y in range(height):
        enqueue(y, 0)
        enqueue(y, width - 1)
    while queue:
        y, x = queue.popleft()
        connected[y, x] = True
        if y:
            enqueue(y - 1, x)
        if y + 1 < height:
            enqueue(y + 1, x)
        if x:
            enqueue(y, x - 1)
        if x + 1 < width:
            enqueue(y, x + 1)
    alpha = np.full((height, width), 255.0, dtype=np.float32)
    soft = np.clip((distance - tolerance) / max(1.0, feather), 0.0, 1.0) * 255.0
    alpha[connected] = soft[connected]
    output = array.copy()
    output[:, :, 3] = np.uint8(alpha)
    out.parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(output, "RGBA").save(out)
    # Put-back fidelity measures whether compositing on the estimated original
    # background reconstructs the source pixels.
    a = alpha[:, :, None] / 255.0
    put_back = rgb * a + background[None, None, :] * (1.0 - a)
    mae = float(np.abs(put_back - rgb).mean() / 255.0)
    result = {
        "status": "ready",
        "operation": "edge-connected-flat-background-separation",
        "source": str(source.resolve()),
        "output": str(out.resolve()),
        "background_rgb": [round(float(value), 2) for value in background],
        "tolerance": tolerance,
        "feather": feather,
        "transparent_fraction": round(float(np.mean(alpha < 1.0)), 6),
        "put_back_mae": round(mae, 8),
    }
    _write(out.with_suffix(out.suffix + ".json"), result)
    return result


def brand_asset(skill_root: Path, asset_id: str, out: Path | None = None) -> dict[str, Any]:
    catalog_path = skill_root / "assets/brand-catalog.json"
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    for pack_id, pack in (catalog.get("brand_packs") or {}).items():
        for asset in pack.get("assets") or []:
            if asset.get("id") != asset_id:
                continue
            source = (skill_root / "assets" / str(asset["path"])).resolve()
            result = {
                "status": "found",
                "pack": pack_id,
                "asset": asset,
                "source": str(source),
            }
            if out:
                out.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, out)
                result["output"] = str(out.resolve())
            return result
    raise KeyError(f"brand asset not found: {asset_id}")


def list_brand_assets(skill_root: Path) -> dict[str, Any]:
    catalog = json.loads((skill_root / "assets/brand-catalog.json").read_text(encoding="utf-8"))
    assets: list[dict[str, Any]] = []
    for pack_id, pack in (catalog.get("brand_packs") or {}).items():
        for asset in pack.get("assets") or []:
            assets.append({"pack": pack_id, **asset})
    return {"schema_version": 1, "assets": assets}
