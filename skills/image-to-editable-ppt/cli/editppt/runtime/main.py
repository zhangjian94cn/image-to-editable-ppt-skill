#!/usr/bin/env python3
"""Curated deterministic tools for one-Codex-task-per-page PPT rebuilding."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from PIL import Image, ImageOps

# ``main.py`` remains directly executable for diagnostics and tests.  Make the
# package root discoverable before importing runtime modules that share public
# helpers from ``editppt``.
CLI_ROOT = Path(__file__).resolve().parents[2]
if str(CLI_ROOT) not in sys.path:
    sys.path.insert(0, str(CLI_ROOT))

from _input_normalization import normalize_inputs
from asset_tools import brand_asset, crop, list_brand_assets, separate_flat_background
from formula_renderer import FormulaRenderError, render_latex_asset
from pptx_assemble import assemble_pptx_packages
from quality_tools import compare_images, find_extractor, inspect_pptx, powerpoint_version, render_powerpoint
from runtime_env import config_path, read_config_file
from source_inspect import inspect_layout, inspect_structure
from text_fit_tool import measure as measure_text
from text_fit_tool import write_result as write_text_fit
from font_registry import font_environment_fingerprint, resolve_font
from typography import ROLE_SPECS, typography_hints
from editppt.text_evidence import normalize_text, region_text_coverage
from editppt.visual_inputs import prepare_visual_inputs


RUNTIME_DIR = Path(__file__).resolve().parent
SKILL_ROOT = RUNTIME_DIR.parents[2]
HELP = argparse.RawDescriptionHelpFormatter


def _script(
    name: str,
    *argv: object,
    capture: bool = False,
    env: dict[str, str] | None = None,
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(RUNTIME_DIR / name), *[str(value) for value in argv]],
        text=True,
        capture_output=capture,
        env=env,
        check=False,
    )


def _json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return value if isinstance(value, dict) else {}


def _print_json(value: Any) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2))


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _configured_paddle_token() -> str:
    if os.environ.get("EDITPPT_DISABLE_PADDLE_OCR", "").strip().lower() in {"1", "true", "yes"}:
        return ""
    token = os.environ.get("PADDLE_OCR_TOKEN", "").strip()
    if token:
        return token
    try:
        return str(read_config_file(config_path()).get("PADDLE_OCR_TOKEN", "")).strip()
    except (OSError, SystemExit):
        return ""


def _write_footer_detail(source: Path, page_dir: Path) -> tuple[Path, list[dict[str, float]]]:
    """Write a high-contrast, enlarged footer strip for exact small-text review."""

    detail_dir = page_dir / ".editppt/inspect"
    detail_dir.mkdir(parents=True, exist_ok=True)
    output = detail_dir / "footer-detail.png"
    with Image.open(source).convert("RGB") as image:
        top = int(round(image.height * 0.94))
        strip_height = image.height - top
        tile_count = 3
        core_width = int(round(image.width / tile_count))
        overlap = max(24, int(round(image.width * 0.035)))
        # Legal footers are often lighter than #E0E0E0.  A high white-point
        # threshold reveals them without fabricating glyphs, while the narrow
        # band keeps main-page copy out of the detail OCR pass.
        scale = 4.0
        separator = 18
        tiles = []
        mappings: list[dict[str, float]] = []
        for index in range(tile_count):
            left = max(0, index * core_width - (overlap if index else 0))
            right = min(image.width, (index + 1) * core_width + (overlap if index < tile_count - 1 else 0))
            tile = image.crop((left, top, right, image.height)).convert("L")
            tile = ImageOps.autocontrast(tile, cutoff=1).point(lambda value: 0 if value < 247 else 255)
            tile = tile.resize((int((right - left) * scale), int(strip_height * scale)), Image.Resampling.LANCZOS)
            tiles.append(tile)
            mappings.append({
                "source_left": float(left),
                "source_top": float(top),
                "scale": scale,
                "output_top": float(index * (int(strip_height * scale) + separator)),
                "output_height": float(tile.height),
            })
        canvas = Image.new("L", (max(tile.width for tile in tiles), sum(tile.height for tile in tiles) + separator * (tile_count - 1)), 255)
        for mapping, tile in zip(mappings, tiles):
            canvas.paste(tile, (0, int(mapping["output_top"])))
        canvas.convert("RGB").save(output)
    return output, mappings


def _merge_footer_hints(
    page_dir: Path,
    hints: dict[str, Any],
    footer_path: Path,
    mappings: list[dict[str, float]],
) -> dict[str, Any]:
    detail = _json(page_dir / ".editppt/inspect/footer-hints.json")
    existing = {normalize_text(str(value.get("text") or "")) for value in hints.get("lines", []) if isinstance(value, dict)}
    added = 0
    for line in detail.get("lines", []):
        if not isinstance(line, dict) or not normalize_text(str(line.get("text") or "")):
            continue
        text_key = normalize_text(str(line.get("text") or ""))
        if text_key in existing or any(
            len(text_key) >= 4 and (text_key in known or known in text_key)
            for known in existing if len(known) >= 4
        ):
            continue
        value = dict(line)
        box = value.get("box_px") or []
        mapping = None
        if len(box) == 4:
            center_y = float(box[1]) + float(box[3]) / 2
            mapping = next(
                (
                    candidate for candidate in mappings
                    if candidate["output_top"] <= center_y <= candidate["output_top"] + candidate["output_height"]
                ),
                None,
            )
        scale = mapping["scale"] if mapping else 1.0
        if len(box) == 4 and mapping:
            value["box_px"] = [
                mapping["source_left"] + float(box[0]) / scale,
                mapping["source_top"] + (float(box[1]) - mapping["output_top"]) / scale,
                float(box[2]) / scale,
                float(box[3]) / scale,
            ]
        elif len(box) == 4:
            continue
        for key in ("glyph_height_px", "group_glyph_px"):
            if value.get(key) is not None:
                value[key] = round(float(value[key]) / scale, 1)
        for key in ("font_pt", "font_pt_if_cjk", "font_pt_if_latin"):
            if value.get(key) is not None:
                value[key] = round(float(value[key]) / scale, 1)
        value["evidence_pass"] = "footer-detail"
        hints.setdefault("lines", []).append(value)
        existing.add(text_key)
        added += 1
    hints["lines"].sort(key=lambda line: (float((line.get("box_px") or [0, 0])[1]), float((line.get("box_px") or [0, 0])[0])))
    for index, line in enumerate(hints["lines"], start=1):
        line["id"] = f"P{index:02d}"
    hints["detail_images"] = [str(footer_path.relative_to(page_dir))]
    hints["footer_detail_added_lines"] = added
    return hints


def cmd_prepare(args: argparse.Namespace) -> int:
    deck = normalize_inputs(args.inputs, out_root=args.out_root, job_dir=args.job_dir, dpi=args.dpi)
    payload = _json(deck)
    _print_json({"status": "ready", "deck_manifest": str(deck), "page_count": int(payload.get("page_count") or 0)})
    return 0


def _collect_text_hints(
    page_dir: Path,
    *,
    source_name: str,
    out_name: str,
    overlay_name: str,
    detail_ocr: bool,
) -> dict[str, Any]:
    source = page_dir / source_name
    if not source.is_file():
        raise FileNotFoundError(f"source image not found: {source}")
    command: list[object] = [page_dir, "--source", source_name, "--out", out_name]
    command += ["--overlay", overlay_name if overlay_name else ""]
    footer_path, footer_mappings = _write_footer_detail(source, page_dir)
    token = _configured_paddle_token()
    ocr_status = "not_configured"
    ocr_reason = "PADDLE_OCR_TOKEN is not configured"
    paddle_succeeded = False
    if token:
        child_env = os.environ.copy()
        child_env["PADDLE_OCR_TOKEN"] = token
        completed = _script("paddle_text_hints.py", *command, capture=True, env=child_env)
        if completed.returncode != 0:
            reason = (completed.stderr or completed.stdout or "unknown OCR error").strip()
            print(f"content-aware OCR unavailable ({reason[:240]}); using geometric hints", file=sys.stderr)
            ocr_status = "degraded"
            ocr_reason = reason[:500]
            completed = _script("text_hints.py", *command, capture=True)
        else:
            ocr_status = "ready"
            ocr_reason = ""
            paddle_succeeded = True
    else:
        completed = _script("text_hints.py", *command, capture=True)
    if completed.returncode != 0:
        raise RuntimeError((completed.stderr or completed.stdout or "text evidence failed").strip())
    hints = _json(page_dir / out_name)
    if paddle_succeeded and detail_ocr:
        footer_command: list[object] = [
            page_dir,
            "--source", str(footer_path.relative_to(page_dir)),
            "--out", ".editppt/inspect/footer-hints.json",
            "--overlay", ".editppt/inspect/footer-hints.png",
            "--min-glyph", 4,
        ]
        child_env = os.environ.copy()
        child_env["PADDLE_OCR_TOKEN"] = token
        footer_completed = _script("paddle_text_hints.py", *footer_command, capture=True, env=child_env)
        if footer_completed.returncode != 0:
            hints["footer_detail_warning"] = (footer_completed.stderr or footer_completed.stdout or "footer OCR failed").strip()[:240]
    hints = _merge_footer_hints(page_dir, hints, footer_path, footer_mappings)
    hints["ocr"] = {
        "status": ocr_status,
        "configured": bool(token),
        "provider": "paddleocr-vl" if paddle_succeeded else "local-geometric",
        "model": os.environ.get("PADDLE_OCR_MODEL", "PaddleOCR-VL-1.6") if token else "",
        "degraded_reason": ocr_reason,
    }
    (page_dir / out_name).write_text(json.dumps(hints, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    lines = hints.get("lines") if isinstance(hints.get("lines"), list) else []
    backend = str(hints.get("backend") or hints.get("source") or "local-geometric")
    real_ocr = backend not in {"local", "local-geometric", "geometric", "projection-components"}
    return {
        "status": "ready",
        "provider": backend,
        "real_ocr": real_ocr,
        "ocr": hints["ocr"],
        "text_hints": str(page_dir / out_name),
        "overlay": str(page_dir / overlay_name) if overlay_name else "",
        "detail_images": hints.get("detail_images", []),
        "footer_detail_added_lines": int(hints.get("footer_detail_added_lines") or 0),
        "line_count": len(lines),
    }


def cmd_inspect_text(args: argparse.Namespace) -> int:
    page_dir = Path(args.page_dir).expanduser().resolve()
    payload = _collect_text_hints(
        page_dir,
        source_name=args.source,
        out_name=args.out,
        overlay_name=args.overlay,
        detail_ocr=getattr(args, "detail_ocr", True),
    )
    _print_json(payload)
    return 0


def cmd_inspect_evidence(args: argparse.Namespace) -> int:
    """Build one cached evidence index for the page-owning Codex task."""

    page = Path(args.page_dir).expanduser().resolve()
    source = page / args.source
    if not source.is_file():
        print(f"source image not found: {source}", file=sys.stderr)
        return 2
    editppt_dir = page / ".editppt"
    editppt_dir.mkdir(parents=True, exist_ok=True)
    output = editppt_dir / "evidence.json"
    font_fingerprint = font_environment_fingerprint()
    contract = _json(SKILL_ROOT / "skill.json")
    cache_inputs = {
        "source_sha256": _sha256(source),
        "paddle_configured": bool(_configured_paddle_token()),
        "paddle_model": os.environ.get("PADDLE_OCR_MODEL", "PaddleOCR-VL-1.6"),
        "skill_contract_version": str(contract.get("contract_version") or "unknown"),
        "font_environment_fingerprint": font_fingerprint,
    }
    cache_key = hashlib.sha256(json.dumps(cache_inputs, sort_keys=True).encode()).hexdigest()
    cached = _json(output)
    if cached.get("cache_key") == cache_key:
        required = [
            cached.get("visual_inputs", {}).get("manifest"),
            cached.get("text", {}).get("text_hints"),
            cached.get("layout", {}).get("path"),
            cached.get("structure", {}).get("path"),
            cached.get("typography", {}).get("path"),
        ]
        if all(value and Path(str(value)).is_file() for value in required):
            cached["cache_hit"] = True
            _print_json(cached)
            return 0

    visual = prepare_visual_inputs(
        page,
        source_name=args.source,
        output_dir=args.vision_out_dir,
        target_edge=args.target_edge,
        overlap_ratio=args.overlap_ratio,
    )
    text = _collect_text_hints(
        page,
        source_name=args.source,
        out_name=args.text_out,
        overlay_name=args.text_overlay,
        detail_ocr=args.detail_ocr,
    )
    layout_path = editppt_dir / "layout.json"
    structure_path = editppt_dir / "structure.json"
    layout = inspect_layout(source, layout_path)
    structure = inspect_structure(source, structure_path)
    with Image.open(source) as image:
        source_size = [image.width, image.height]
    text_payload = _json(page / args.text_out)
    typography = typography_hints(text_payload, source_size[1])
    typography_path = editppt_dir / "typography-hints.json"
    typography_path.write_text(json.dumps(typography, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    payload = {
        "schema_version": 1,
        "status": "ready",
        "cache_key": cache_key,
        "cache_hit": False,
        "cache_inputs": cache_inputs,
        "source": {"path": str(source), "size_px": source_size, "sha256": cache_inputs["source_sha256"]},
        "visual_inputs": {
            "provider": visual["provider"],
            "manifest": visual["manifest"],
            "images": [value["path"] for value in visual["images"]],
        },
        "text": text,
        "layout": {"path": str(layout_path), "summary": layout},
        "structure": {"path": str(structure_path), "summary": structure},
        "typography": {"path": str(typography_path), "size_groups": typography["size_groups"]},
        "font_environment_fingerprint": font_fingerprint,
        "instruction": (
            "Read this index before authoring. OCR is positional evidence; Codex remains responsible "
            "for semantic roles, exact source reading, and conflict resolution."
        ),
    }
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    _print_json(payload)
    return 0


def cmd_inspect_vision(args: argparse.Namespace) -> int:
    try:
        payload = prepare_visual_inputs(
            args.page_dir,
            source_name=args.source,
            output_dir=args.out_dir,
            target_edge=args.target_edge,
            overlap_ratio=args.overlap_ratio,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 2
    _print_json({
        "status": "ready",
        "provider": payload["provider"],
        "manifest": payload["manifest"],
        "images": [value["path"] for value in payload["images"]],
    })
    return 0


def cmd_inspect_transcript(args: argparse.Namespace) -> int:
    page = Path(args.page_dir).expanduser().resolve()
    transcript_path = page / args.input
    reference_path = page / args.against
    transcript = _json(transcript_path)
    reference = _json(reference_path)
    if not transcript_path.is_file() or not isinstance(transcript.get("lines"), list):
        print(f"invalid transcript evidence: {transcript_path}", file=sys.stderr)
        return 2
    if not reference_path.is_file() or not isinstance(reference.get("lines"), list):
        print(f"invalid reference text evidence: {reference_path}", file=sys.stderr)
        return 2
    expected = [
        {"text": str(value.get("text") or ""), "box_px": value.get("box_px")}
        for value in reference["lines"] if isinstance(value, dict)
    ]
    candidate = [
        {"text": str(value.get("text") or ""), "box_px": value.get("box_px")}
        for value in transcript["lines"] if isinstance(value, dict)
    ]
    evidence = region_text_coverage(expected, candidate)
    payload = {
        "status": "ready",
        "advisory": True,
        "transcript": str(transcript_path),
        "reference": str(reference_path),
        "reference_coverage": evidence["text_coverage"],
        "missing_reference_count": evidence["missing_text_count"],
        "missing_reference_texts": evidence["missing_texts"],
        "instruction": (
            "Re-view the attached source/detail image for every disagreement. "
            "Never automatically replace the transcript with OCR text."
        ),
    }
    out = page / args.out
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    _print_json(payload)
    return 0


def cmd_inspect_layout(args: argparse.Namespace) -> int:
    page = Path(args.page_dir).expanduser().resolve()
    payload = inspect_layout(page / args.source, page / args.out)
    _print_json(payload)
    return 0


def cmd_inspect_structure(args: argparse.Namespace) -> int:
    page = Path(args.page_dir).expanduser().resolve()
    payload = inspect_structure(page / args.source, page / args.out)
    _print_json(payload)
    return 0


def cmd_inspect_pptx(args: argparse.Namespace) -> int:
    page = Path(args.page_dir).expanduser().resolve()
    path = page / args.input
    hints = page / args.text_hints if args.text_hints else None
    payload = inspect_pptx(path, hints)
    out = page / args.out
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    _print_json(payload)
    return 0


def cmd_assets_crop(args: argparse.Namespace) -> int:
    payload = crop(
        Path(args.input).expanduser().resolve(),
        Path(args.out).expanduser().resolve(),
        (args.left, args.top, args.right, args.bottom),
        args.pad,
    )
    _print_json(payload)
    return 0


def cmd_assets_separate(args: argparse.Namespace) -> int:
    payload = separate_flat_background(
        Path(args.input).expanduser().resolve(),
        Path(args.out).expanduser().resolve(),
        tolerance=args.tolerance,
        feather=args.feather,
    )
    _print_json(payload)
    return 0


def cmd_assets_split_alpha(args: argparse.Namespace) -> int:
    argv: list[object] = ["--input", args.input, "--out-dir", args.out_dir, "--sort", args.sort]
    for name in ("names", "threshold", "min_area", "pad", "limit", "contact_sheet", "manifest"):
        value = getattr(args, name)
        if value not in (None, ""):
            argv += ["--" + name.replace("_", "-"), value]
    if args.square:
        argv.append("--square")
    return _script("split_alpha_components.py", *argv).returncode


def cmd_assets_remove_chroma(args: argparse.Namespace) -> int:
    argv: list[object] = ["--input", args.input, "--out", args.out, "--key-color", args.key_color]
    argv += ["--tolerance", args.tolerance, "--auto-key", args.auto_key]
    if args.soft_matte:
        argv += ["--soft-matte", "--transparent-threshold", args.transparent_threshold, "--opaque-threshold", args.opaque_threshold]
    if args.force:
        argv.append("--force")
    return _script("remove_chroma_key.py", *argv).returncode


def cmd_assets_brand(args: argparse.Namespace) -> int:
    if not args.id:
        _print_json(list_brand_assets(SKILL_ROOT))
        return 0
    output = Path(args.out).expanduser().resolve() if args.out else None
    _print_json(brand_asset(SKILL_ROOT, args.id, output))
    return 0


def cmd_build(args: argparse.Namespace) -> int:
    page = Path(args.page_dir).expanduser().resolve()
    manifest = page / args.manifest
    if not manifest.is_file():
        print(f"manifest not found: {manifest}", file=sys.stderr)
        return 2
    build_report = page / ".editppt/build.json"
    argv: list[object] = [manifest, "--out", page / args.out, "--report", build_report]
    draft_preview = args.draft_preview or args.preview
    if draft_preview:
        argv += ["--preview", page / draft_preview]
    completed = _script("build_pptx_from_manifest.py", *argv, capture=True)
    if completed.returncode != 0:
        print(completed.stderr or completed.stdout, file=sys.stderr)
        return completed.returncode
    report = _json(build_report)
    layer_report = page / ".editppt/layer-report.json"
    layer_report.parent.mkdir(parents=True, exist_ok=True)
    layer_report.write_text(
        json.dumps(report.get("layer_report", {}), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    _print_json({
        "status": "ready",
        "output_pptx": str(page / args.out),
        "build_report": str(build_report),
        "layer_report": str(layer_report),
        "font_substitutions": report.get("font_substitutions", []),
        "font_resolution": report.get("font_resolution", []),
        "font_environment_fingerprint": report.get("font_environment_fingerprint", ""),
        "text_adjustments": report.get("text_adjustments", []),
        "severe_text_adjustments": report.get("severe_text_adjustments", []),
        "role_shrink_warnings": report.get("role_shrink_warnings", []),
        "typography_adjustments": report.get("typography_adjustments", []),
        "role_size_deviations": report.get("role_size_deviations", []),
        "warnings": report.get("warnings", []),
        "draft_preview": str(page / draft_preview) if draft_preview else "",
        "warning": "draft preview is not a Microsoft PowerPoint render" if draft_preview else "",
    })
    return 0


def cmd_text_fit(args: argparse.Namespace) -> int:
    text = Path(args.text_file).read_text(encoding="utf-8") if args.text_file else args.text
    profile = _json(Path(args.typography_profile).expanduser().resolve()) if args.typography_profile else None
    payload = measure_text(
        text,
        args.width_px,
        args.height_px,
        preferred_font=args.font,
        max_font_px=args.max_font_px,
        min_font_px=args.min_font_px,
        line_spacing=args.line_spacing,
        single_line=args.single_line,
        role=args.role,
        typography_profile=profile,
    )
    if args.out:
        write_text_fit(Path(args.out).expanduser().resolve(), payload)
    _print_json(payload)
    return 0


def cmd_render(args: argparse.Namespace) -> int:
    page = Path(args.page_dir).expanduser().resolve()
    payload = render_powerpoint(
        page / args.input,
        page / args.out,
        evidence_dir=page / args.evidence_dir,
        extractor=args.extractor,
        dpi=args.dpi,
    )
    _print_json(payload)
    return 0


def cmd_compare(args: argparse.Namespace) -> int:
    page = Path(args.page_dir).expanduser().resolve()
    payload = compare_images(page / args.source, page / args.candidate, page / args.out_dir)
    _print_json(payload)
    return 0


def cmd_assemble(args: argparse.Namespace) -> int:
    page_dirs = [Path(value).expanduser().resolve() for value in args.page_dirs]
    inputs = [page / args.page_pptx for page in page_dirs]
    payload = assemble_pptx_packages(
        inputs,
        Path(args.out).expanduser().resolve(),
        Path(args.evidence_dir).expanduser().resolve(),
    )
    _print_json(payload)
    return 0


def cmd_doctor(args: argparse.Namespace) -> int:
    completed = _script("runtime_env.py", "doctor", "--json", capture=True)
    try:
        payload = json.loads(completed.stdout or "{}")
    except json.JSONDecodeError:
        payload = {"ok": False, "runtime_stdout": completed.stdout, "runtime_stderr": completed.stderr}
    contract = _json(SKILL_ROOT / "skill.json")
    renderer = find_extractor()
    checks = {
        "powerpoint": {
            "available": Path("/Applications/Microsoft PowerPoint.app").is_dir(),
            "version": powerpoint_version(),
        },
        "powerpoint_renderer": {
            "available": bool(renderer),
            "path": str(renderer or ""),
            "target_only": True,
            "canary_created": False,
        },
        "pdf_to_png": {"available": bool(shutil.which("pdftoppm")), "path": shutil.which("pdftoppm") or ""},
        "rsvg_convert": {"available": bool(shutil.which("rsvg-convert")), "path": shutil.which("rsvg-convert") or ""},
        "font_registry": {
            "available": bool(resolve_font("Microsoft YaHei", require_cjk=True)),
            "fingerprint": font_environment_fingerprint(),
            "microsoft_yahei": (
                resolve_font("Microsoft YaHei", require_cjk=True).as_dict()
                if resolve_font("Microsoft YaHei", require_cjk=True)
                else {}
            ),
        },
    }
    payload["skill"] = {
        "root": str(SKILL_ROOT),
        "contract_version": str(contract.get("contract_version") or "unknown"),
        "upstream_base": str(contract.get("upstream_base") or ""),
    }
    payload["editppt_checks"] = checks
    payload["ok"] = bool(payload.get("ok")) and all(check["available"] for check in checks.values() if check is not checks["rsvg_convert"])
    if args.json:
        _print_json(payload)
    else:
        print("ok=" + ("yes" if payload.get("ok") else "no"))
        print(f"contract={payload['skill']['contract_version']}")
        for name, check in checks.items():
            print(f"{name}={'yes' if check['available'] else 'no'}")
    return 0 if payload.get("ok") else 1


def cmd_formula(args: argparse.Namespace) -> int:
    tex = Path(args.tex_file).read_text(encoding="utf-8") if args.tex_file else (args.tex or "")
    if not tex:
        print("formula render-latex requires --tex or --tex-file", file=sys.stderr)
        return 2
    try:
        result = render_latex_asset(
            tex=tex,
            out=args.out,
            page_dir=args.page_dir,
            output_format=args.format,
            engine=args.engine,
            preamble="",
            full_document=args.full_document,
            display=not args.inline,
            dpi=args.dpi,
            timeout=args.timeout,
            shell_escape=False,
            keep_workdir=None,
        )
    except FormulaRenderError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    _print_json(result)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=os.environ.get("IMAGE_TO_EDITABLE_PPT_CLI_PROG", "editppt"),
        description="Deterministic evidence and authoring tools for editable slide reconstruction.",
        formatter_class=HELP,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    prepare = sub.add_parser("prepare", help="Normalize images, PDFs, or PPTX inputs into ordered source pages.")
    prepare.add_argument("inputs", nargs="+")
    prepare.add_argument("--out-root", default="output/image-to-editable-ppt")
    prepare.add_argument("--job-dir")
    prepare.add_argument("--dpi", type=int, default=180)
    prepare.set_defaults(func=cmd_prepare)

    inspect = sub.add_parser("inspect", help="Inspect source evidence or the authored PPTX.")
    inspect_sub = inspect.add_subparsers(dest="inspect_command", required=True)
    inspect_evidence = inspect_sub.add_parser(
        "evidence",
        help="Create the cached multimodal, OCR, layout, structure, typography, and font evidence index.",
    )
    inspect_evidence.add_argument("page_dir")
    inspect_evidence.add_argument("--source", default="source.png")
    inspect_evidence.add_argument("--text-out", default="text_hints.json")
    inspect_evidence.add_argument("--text-overlay", default="text_hints.png")
    inspect_evidence.add_argument("--vision-out-dir", default=str(Path(".editppt/vision-inputs")))
    inspect_evidence.add_argument("--target-edge", type=int, default=1792)
    inspect_evidence.add_argument("--overlap-ratio", type=float, default=0.08)
    inspect_evidence.add_argument(
        "--detail-ocr",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Run a second OCR pass on the enlarged footer when PaddleOCR-VL succeeds.",
    )
    inspect_evidence.set_defaults(func=cmd_inspect_evidence)
    inspect_text = inspect_sub.add_parser("text", help="Write OCR/text coordinate evidence.")
    inspect_text.add_argument("page_dir")
    inspect_text.add_argument("--source", default="source.png")
    inspect_text.add_argument("--out", default="text_hints.json")
    inspect_text.add_argument("--overlay", default="text_hints.png")
    inspect_text.add_argument(
        "--detail-ocr",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Run a second OCR pass on an enlarged high-contrast footer strip when content-aware OCR is available.",
    )
    inspect_text.set_defaults(func=cmd_inspect_text)
    inspect_vision = inspect_sub.add_parser(
        "vision",
        help="Create portable overlapping detail images for the page's multimodal Codex task.",
    )
    inspect_vision.add_argument("page_dir")
    inspect_vision.add_argument("--source", default="source.png")
    inspect_vision.add_argument("--out-dir", default=str(Path(".editppt/vision-inputs")))
    inspect_vision.add_argument("--target-edge", type=int, default=1792)
    inspect_vision.add_argument("--overlap-ratio", type=float, default=0.08)
    inspect_vision.set_defaults(func=cmd_inspect_vision)
    inspect_transcript = inspect_sub.add_parser(
        "transcript",
        help="Compare Codex multimodal transcription with independent OCR evidence.",
    )
    inspect_transcript.add_argument("page_dir")
    inspect_transcript.add_argument("--input", default="source-transcript.json")
    inspect_transcript.add_argument("--against", default="text_hints.json")
    inspect_transcript.add_argument("--out", default=".editppt/inspect-transcript.json")
    inspect_transcript.set_defaults(func=cmd_inspect_transcript)
    inspect_layout_parser = inspect_sub.add_parser("layout", help="Measure broad content and whitespace bands.")
    inspect_layout_parser.add_argument("page_dir")
    inspect_layout_parser.add_argument("--source", default="source.png")
    inspect_layout_parser.add_argument("--out", default="layout.json")
    inspect_layout_parser.set_defaults(func=cmd_inspect_layout)
    inspect_structure_parser = inspect_sub.add_parser("structure", help="Measure source-space line and rectangle candidates.")
    inspect_structure_parser.add_argument("page_dir")
    inspect_structure_parser.add_argument("--source", default="source.png")
    inspect_structure_parser.add_argument("--out", default="structure.json")
    inspect_structure_parser.set_defaults(func=cmd_inspect_structure)
    inspect_pptx_parser = inspect_sub.add_parser("pptx", help="Read back objects, text, tables, pictures, and boundary risks.")
    inspect_pptx_parser.add_argument("page_dir")
    inspect_pptx_parser.add_argument("--input", default="page.pptx")
    inspect_pptx_parser.add_argument("--out", default="pptx-inspect.json")
    inspect_pptx_parser.add_argument(
        "--text-hints",
        default="text_hints.json",
        help="Advisory OCR evidence to compare with editable PPTX text; pass an empty string to skip.",
    )
    inspect_pptx_parser.set_defaults(func=cmd_inspect_pptx)

    assets = sub.add_parser("assets", help="Extract or resolve independent source-bound assets.")
    assets_sub = assets.add_subparsers(dest="assets_command", required=True)
    crop_parser = assets_sub.add_parser("crop", help="Make a lossless source-pixel crop with coordinate evidence.")
    crop_parser.add_argument("--input", required=True)
    crop_parser.add_argument("--out", required=True)
    crop_parser.add_argument("--left", type=int, required=True)
    crop_parser.add_argument("--top", type=int, required=True)
    crop_parser.add_argument("--right", type=int, required=True)
    crop_parser.add_argument("--bottom", type=int, required=True)
    crop_parser.add_argument("--pad", type=int, default=0)
    crop_parser.set_defaults(func=cmd_assets_crop)
    separate = assets_sub.add_parser("separate", help="Remove an edge-connected flat background and report put-back fidelity.")
    separate.add_argument("--input", required=True)
    separate.add_argument("--out", required=True)
    separate.add_argument("--tolerance", type=float, default=24.0)
    separate.add_argument("--feather", type=float, default=14.0)
    separate.set_defaults(func=cmd_assets_separate)
    split_alpha = assets_sub.add_parser("split-alpha", help="Split a transparent sheet into component PNG assets.")
    split_alpha.add_argument("--input", required=True)
    split_alpha.add_argument("--out-dir", required=True)
    split_alpha.add_argument("--names", default="")
    split_alpha.add_argument("--sort", choices=("x", "y", "area"), default="x")
    split_alpha.add_argument("--threshold", type=int)
    split_alpha.add_argument("--min-area", type=int)
    split_alpha.add_argument("--pad", type=int)
    split_alpha.add_argument("--limit", type=int)
    split_alpha.add_argument("--square", action="store_true")
    split_alpha.add_argument("--manifest", default="")
    split_alpha.add_argument("--contact-sheet", default="")
    split_alpha.set_defaults(func=cmd_assets_split_alpha)
    chroma = assets_sub.add_parser("remove-chroma", help="Remove a chroma-key background with explicit parameters.")
    chroma.add_argument("--input", required=True)
    chroma.add_argument("--out", required=True)
    chroma.add_argument("--key-color", default="#00ff00")
    chroma.add_argument("--tolerance", type=int, default=12)
    chroma.add_argument("--auto-key", choices=("none", "corners", "border"), default="none")
    chroma.add_argument("--soft-matte", action="store_true")
    chroma.add_argument("--transparent-threshold", type=float, default=12.0)
    chroma.add_argument("--opaque-threshold", type=float, default=96.0)
    chroma.add_argument("--force", action="store_true")
    chroma.set_defaults(func=cmd_assets_remove_chroma)
    brand = assets_sub.add_parser("brand", help="List or export an audited brand asset.")
    brand.add_argument("--id", default="")
    brand.add_argument("--out", default="")
    brand.set_defaults(func=cmd_assets_brand)

    build = sub.add_parser("build", help="Build page.pptx from the shared manifest Builder.")
    build.add_argument("page_dir")
    build.add_argument("--manifest", default="manifest.json")
    build.add_argument("--out", default="page.pptx")
    build.add_argument("--draft-preview", default="")
    build.add_argument("--preview", default="", help=argparse.SUPPRESS)
    build.set_defaults(func=cmd_build)

    text_fit = sub.add_parser("text-fit", help="Measure text and recommend a source-pixel font size.")
    text_fit.add_argument("--text", default="")
    text_fit.add_argument("--text-file", default="")
    text_fit.add_argument("--width-px", type=int, required=True)
    text_fit.add_argument("--height-px", type=int, required=True)
    text_fit.add_argument("--font", default="PingFang SC")
    text_fit.add_argument("--max-font-px", type=int, default=120)
    text_fit.add_argument("--min-font-px", type=int, default=6)
    text_fit.add_argument("--line-spacing", type=float, default=1.2)
    text_fit.add_argument("--single-line", action="store_true")
    text_fit.add_argument("--role", choices=tuple(ROLE_SPECS), default="")
    text_fit.add_argument("--typography-profile", default="")
    text_fit.add_argument("--out", default="")
    text_fit.set_defaults(func=cmd_text_fit)

    render = sub.add_parser("render", help="Render page.pptx through Microsoft PowerPoint; fail closed.")
    render.add_argument("page_dir")
    render.add_argument("--input", default="page.pptx")
    render.add_argument("--out", default="preview.png")
    render.add_argument("--evidence-dir", default=".editppt/render")
    render.add_argument("--extractor", default="")
    render.add_argument("--dpi", type=int, default=200)
    render.set_defaults(func=cmd_render)

    compare = sub.add_parser("compare", help="Create source-versus-PowerPoint diagnostic images and metrics.")
    compare.add_argument("page_dir")
    compare.add_argument("--source", default="source.png")
    compare.add_argument("--candidate", default="preview.png")
    compare.add_argument("--out-dir", default=".editppt/compare")
    compare.set_defaults(func=cmd_compare)

    assemble = sub.add_parser("assemble", help="Merge independent page.pptx files in supplied order through PowerPoint.")
    assemble.add_argument("page_dirs", nargs="+")
    assemble.add_argument("--out", required=True)
    assemble.add_argument("--page-pptx", default="page.pptx")
    assemble.add_argument("--evidence-dir", default=".editppt/assemble")
    assemble.set_defaults(func=cmd_assemble)

    formula = sub.add_parser("formula", help="Render a complex formula as an independent local asset.")
    formula_sub = formula.add_subparsers(dest="formula_command", required=True)
    render_formula = formula_sub.add_parser("render-latex")
    render_formula.add_argument("page_dir", nargs="?")
    render_formula.add_argument("--tex")
    render_formula.add_argument("--tex-file")
    render_formula.add_argument("--out", required=True)
    render_formula.add_argument("--format", choices=("svg", "png", "pdf"))
    render_formula.add_argument("--engine", default="auto")
    render_formula.add_argument("--inline", action="store_true")
    render_formula.add_argument("--full-document", action="store_true")
    render_formula.add_argument("--dpi", type=int, default=300)
    render_formula.add_argument("--timeout", type=int, default=120)
    render_formula.set_defaults(func=cmd_formula)

    doctor = sub.add_parser("doctor", help="Report Codex/Skill/OCR/font/PowerPoint/Open XML dependencies.")
    doctor.add_argument("--json", action="store_true")
    doctor.set_defaults(func=cmd_doctor)
    return parser


def _trace(args: argparse.Namespace, started: float, exit_code: int, error: str = "") -> None:
    trace_file = os.environ.get("EDITPPT_TRACE_FILE", "").strip()
    if not trace_file:
        return
    payload = {
        "schema_version": 1,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "command": getattr(args, "command", ""),
        "subcommand": getattr(args, "inspect_command", "") or getattr(args, "assets_command", "") or getattr(args, "formula_command", ""),
        "argv": sys.argv[1:],
        "exit_code": exit_code,
        "elapsed_sec": round(time.monotonic() - started, 6),
        "error": error,
    }
    path = Path(trace_file).expanduser()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(payload, ensure_ascii=False) + "\n")


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    started = time.monotonic()
    exit_code = 1
    error = ""
    try:
        exit_code = int(args.func(args) or 0)
    except (FileNotFoundError, ValueError, KeyError, RuntimeError) as exc:
        error = str(exc)
        print(error, file=sys.stderr)
        exit_code = 1
    finally:
        _trace(args, started, exit_code, error)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
