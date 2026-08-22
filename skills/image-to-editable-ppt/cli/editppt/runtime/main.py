#!/usr/bin/env python3
"""Curated deterministic tools for one-Codex-task-per-page PPT rebuilding."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

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


def _configured_paddle_token() -> str:
    token = os.environ.get("PADDLE_OCR_TOKEN", "").strip()
    if token:
        return token
    try:
        return str(read_config_file(config_path()).get("PADDLE_OCR_TOKEN", "")).strip()
    except (OSError, SystemExit):
        return ""


def cmd_prepare(args: argparse.Namespace) -> int:
    deck = normalize_inputs(args.inputs, out_root=args.out_root, job_dir=args.job_dir, dpi=args.dpi)
    payload = _json(deck)
    _print_json({"status": "ready", "deck_manifest": str(deck), "page_count": int(payload.get("page_count") or 0)})
    return 0


def cmd_inspect_text(args: argparse.Namespace) -> int:
    page_dir = Path(args.page_dir).expanduser().resolve()
    source = page_dir / args.source
    if not source.is_file():
        print(f"source image not found: {source}", file=sys.stderr)
        return 2
    command: list[object] = [page_dir, "--source", args.source, "--out", args.out]
    command += ["--overlay", args.overlay if args.overlay else ""]
    token = _configured_paddle_token()
    if token:
        child_env = os.environ.copy()
        child_env["PADDLE_OCR_TOKEN"] = token
        completed = _script("paddle_text_hints.py", *command, capture=True, env=child_env)
        if completed.returncode != 0:
            reason = (completed.stderr or completed.stdout or "unknown OCR error").strip()
            print(f"content-aware OCR unavailable ({reason[:240]}); using geometric hints", file=sys.stderr)
            completed = _script("text_hints.py", *command, capture=True)
    else:
        completed = _script("text_hints.py", *command, capture=True)
    if completed.returncode != 0:
        print(completed.stderr or completed.stdout, file=sys.stderr)
        return completed.returncode
    hints = _json(page_dir / args.out)
    lines = hints.get("lines") if isinstance(hints.get("lines"), list) else []
    backend = str(hints.get("backend") or hints.get("source") or "local-geometric")
    real_ocr = backend not in {"local", "local-geometric", "geometric", "projection-components"}
    _print_json({
        "status": "ready",
        "provider": backend,
        "real_ocr": real_ocr,
        "text_hints": str(page_dir / args.out),
        "overlay": str(page_dir / args.overlay) if args.overlay else "",
        "line_count": len(lines),
    })
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
    payload = inspect_pptx(path)
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
    _print_json({
        "status": "ready",
        "output_pptx": str(page / args.out),
        "build_report": str(build_report),
        "font_substitutions": report.get("font_substitutions", []),
        "text_adjustments": report.get("text_adjustments", []),
        "severe_text_adjustments": report.get("severe_text_adjustments", []),
        "warnings": report.get("warnings", []),
        "draft_preview": str(page / draft_preview) if draft_preview else "",
        "warning": "draft preview is not a Microsoft PowerPoint render" if draft_preview else "",
    })
    return 0


def cmd_text_fit(args: argparse.Namespace) -> int:
    text = Path(args.text_file).read_text(encoding="utf-8") if args.text_file else args.text
    payload = measure_text(
        text,
        args.width_px,
        args.height_px,
        preferred_font=args.font,
        max_font_px=args.max_font_px,
        min_font_px=args.min_font_px,
        line_spacing=args.line_spacing,
        single_line=args.single_line,
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
    inspect_text = inspect_sub.add_parser("text", help="Write OCR/text coordinate evidence.")
    inspect_text.add_argument("page_dir")
    inspect_text.add_argument("--source", default="source.png")
    inspect_text.add_argument("--out", default="text_hints.json")
    inspect_text.add_argument("--overlay", default="text_hints.png")
    inspect_text.set_defaults(func=cmd_inspect_text)
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
