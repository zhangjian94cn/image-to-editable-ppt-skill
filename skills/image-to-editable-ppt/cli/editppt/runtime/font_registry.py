#!/usr/bin/env python3
"""Portable, auditable font discovery for editable slide reconstruction.

PowerPoint ships fonts in a private ``DFonts`` directory on macOS.  They are
valid authoring/measurement fonts even though fontconfig and the public system
font directories do not expose them.  This module inventories every configured
font source, reads internal family names with fontTools, and records a stable
environment fingerprint so callers never need to silently guess a substitute.
"""

from __future__ import annotations

import hashlib
import json
import os
import platform
from dataclasses import asdict, dataclass
from functools import lru_cache
from pathlib import Path
from typing import Iterable

from fontTools.ttLib import TTCollection, TTFont


FONT_SUFFIXES = {".ttf", ".ttc", ".otf", ".otc"}
CACHE_VERSION = 1


@dataclass(frozen=True)
class FontFace:
    path: str
    family: str
    aliases: tuple[str, ...]
    weight: int
    face_index: int
    sha256: str
    provider: str
    cjk_coverage: bool

    def as_dict(self) -> dict[str, object]:
        return asdict(self)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _font_roots() -> list[tuple[Path, str]]:
    values: list[tuple[Path, str]] = []
    explicit = os.environ.get("EDITPPT_FONT_ROOTS", "")
    for raw in explicit.split(os.pathsep):
        if raw.strip():
            values.append((Path(raw).expanduser(), "explicit"))
    if platform.system() == "Darwin":
        values.extend(
            [
                (Path.home() / "Library/Fonts", "user"),
                (Path("/Library/Fonts"), "system"),
                (Path("/System/Library/Fonts"), "system"),
                (Path("/System/Library/AssetsV2/com_apple_MobileAsset_Font7"), "system-asset"),
                (
                    Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"),
                    "powerpoint-dfonts",
                ),
            ]
        )
    elif platform.system() == "Windows":
        values.append((Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts", "system"))
    else:
        values.extend(
            [
                (Path.home() / ".local/share/fonts", "user"),
                (Path.home() / ".fonts", "user"),
                (Path("/usr/local/share/fonts"), "system"),
                (Path("/usr/share/fonts"), "system"),
            ]
        )
    seen: set[str] = set()
    result: list[tuple[Path, str]] = []
    for root, provider in values:
        key = str(root.resolve()) if root.exists() else str(root)
        if key not in seen:
            result.append((root, provider))
            seen.add(key)
    return result


def _cache_key(roots: Iterable[tuple[Path, str]]) -> str:
    records = []
    for root, provider in roots:
        try:
            stat = root.stat()
            marker: list[object] = [str(root.resolve()), provider, stat.st_mtime_ns, stat.st_size]
            if root.is_dir():
                marker.append([
                    [str(path.relative_to(root)), path.stat().st_mtime_ns, path.stat().st_size]
                    for path in sorted(root.rglob("*"))
                    if path.is_file() and path.suffix.lower() in FONT_SUFFIXES
                ])
        except OSError:
            marker = [str(root), provider, 0, 0]
        records.append(marker)
    raw = json.dumps({"version": CACHE_VERSION, "roots": records}, sort_keys=True).encode()
    return hashlib.sha256(raw).hexdigest()


def _name_values(font: TTFont) -> list[str]:
    values: list[str] = []
    name_table = font.get("name")
    if name_table is None:
        return values
    for record in name_table.names:
        if record.nameID not in {1, 2, 4, 6, 16, 17}:
            continue
        try:
            value = record.toUnicode().strip()
        except Exception:
            continue
        if value and value not in values:
            values.append(value)
    return values


def _font_faces(path: Path, provider: str) -> list[FontFace]:
    digest = _sha256(path)
    fonts: list[TTFont] = []
    collection: TTCollection | None = None
    try:
        if path.suffix.lower() in {".ttc", ".otc"}:
            collection = TTCollection(str(path), lazy=True)
            fonts = list(collection.fonts)
        else:
            fonts = [TTFont(str(path), lazy=True)]
        result: list[FontFace] = []
        for index, font in enumerate(fonts):
            names = _name_values(font)
            family_candidates = []
            name_table = font.get("name")
            if name_table is not None:
                for name_id in (16, 1):
                    for record in name_table.names:
                        if record.nameID == name_id:
                            try:
                                value = record.toUnicode().strip()
                            except Exception:
                                continue
                            if value and value not in family_candidates:
                                family_candidates.append(value)
            family = family_candidates[0] if family_candidates else path.stem
            os2 = font.get("OS/2")
            weight = int(getattr(os2, "usWeightClass", 400) or 400)
            cmap = font.getBestCmap() or {}
            cjk = any(codepoint in cmap for codepoint in (0x4E2D, 0x56FD, 0x6587))
            result.append(
                FontFace(
                    path=str(path.resolve()),
                    family=family,
                    aliases=tuple(names),
                    weight=weight,
                    face_index=index,
                    sha256=digest,
                    provider=provider,
                    cjk_coverage=cjk,
                )
            )
        return result
    finally:
        for font in fonts:
            try:
                font.close()
            except Exception:
                pass
        if collection is not None:
            try:
                collection.close()
            except Exception:
                pass


def _cache_path() -> Path:
    configured = os.environ.get("EDITPPT_FONT_CACHE", "").strip()
    if configured:
        return Path(configured).expanduser()
    return Path.home() / ".cache/editppt/font-inventory-v1.json"


@lru_cache(maxsize=1)
def font_inventory() -> tuple[FontFace, ...]:
    roots = _font_roots()
    key = _cache_key(roots)
    cache = _cache_path()
    try:
        payload = json.loads(cache.read_text(encoding="utf-8"))
        if payload.get("cache_key") == key:
            return tuple(
                FontFace(
                    path=str(value["path"]),
                    family=str(value["family"]),
                    aliases=tuple(value.get("aliases") or ()),
                    weight=int(value.get("weight") or 400),
                    face_index=int(value.get("face_index") or 0),
                    sha256=str(value["sha256"]),
                    provider=str(value.get("provider") or "unknown"),
                    cjk_coverage=bool(value.get("cjk_coverage")),
                )
                for value in payload.get("faces", [])
                if Path(str(value.get("path") or "")).is_file()
            )
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
        pass

    faces: list[FontFace] = []
    seen_files: set[str] = set()
    for root, provider in roots:
        if not root.is_dir():
            continue
        for path in sorted(root.rglob("*")):
            if path.suffix.lower() not in FONT_SUFFIXES or not path.is_file():
                continue
            resolved = str(path.resolve())
            if resolved in seen_files:
                continue
            seen_files.add(resolved)
            try:
                faces.extend(_font_faces(path, provider))
            except Exception:
                # A broken font file must not make slide authoring unavailable.
                continue
    try:
        cache.parent.mkdir(parents=True, exist_ok=True)
        cache.write_text(
            json.dumps({"cache_key": key, "faces": [face.as_dict() for face in faces]}, ensure_ascii=False),
            encoding="utf-8",
        )
    except OSError:
        pass
    return tuple(faces)


def font_environment_fingerprint() -> str:
    records = [
        [face.path, face.family, face.face_index, face.weight, face.sha256, face.provider]
        for face in font_inventory()
    ]
    return hashlib.sha256(json.dumps(records, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def _normalized(value: str) -> str:
    return "".join(character for character in value.casefold() if character.isalnum())


FALLBACK_FAMILIES = (
    "Microsoft YaHei",
    "Noto Sans CJK SC",
    "Source Han Sans SC",
    "PingFang SC",
    "DengXian",
    "SimHei",
    "Arial Unicode MS",
)


def resolve_font(preferred: str, *, require_cjk: bool = False) -> FontFace | None:
    query = preferred.strip()
    if query:
        direct = Path(query).expanduser()
        if direct.is_file():
            faces = _font_faces(direct, "explicit-file")
            return faces[0] if faces else None
    inventory = font_inventory()
    token = _normalized(query)
    if token:
        exact = [
            face for face in inventory
            if token in {_normalized(face.family), *(_normalized(alias) for alias in face.aliases)}
        ]
        if exact:
            # Prefer PowerPoint's own copy: measurement then matches the
            # authoritative rendering environment.
            exact.sort(key=lambda face: (face.provider != "powerpoint-dfonts", abs(face.weight - 400), face.path))
            return exact[0]

    candidates = list(FALLBACK_FAMILIES)
    if query and query not in candidates:
        candidates.insert(0, query)
    for family in candidates:
        candidate_token = _normalized(family)
        matches = [
            face for face in inventory
            if candidate_token in {_normalized(face.family), *(_normalized(alias) for alias in face.aliases)}
            and (not require_cjk or face.cjk_coverage)
        ]
        if matches:
            matches.sort(key=lambda face: (face.provider != "powerpoint-dfonts", abs(face.weight - 400), face.path))
            return matches[0]
    cjk = [face for face in inventory if face.cjk_coverage]
    return sorted(cjk, key=lambda face: (face.provider != "powerpoint-dfonts", abs(face.weight - 400), face.path))[0] if cjk else None


def find_font(preferred: str) -> tuple[Path | None, str]:
    """Compatibility wrapper for callers that only need path and family."""

    face = resolve_font(preferred, require_cjk=any("\u4e00" <= char <= "\u9fff" for char in preferred))
    if face is None:
        return None, preferred or "default"
    return Path(face.path), face.family
