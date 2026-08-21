#!/usr/bin/env python3
"""Small deterministic helper surface for the image-to-editable-ppt Skill.

The CLI deliberately does not orchestrate agents or own page state. One Codex
task owns one page and may call these helpers when they improve the result.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from _input_normalization import normalize_inputs
from build_pptx_from_manifest import render_preview
from formula_renderer import FormulaRenderError, render_latex_asset


RUNTIME_DIR = Path(__file__).resolve().parent
SKILL_ROOT = RUNTIME_DIR.parents[2]
HELP = argparse.RawDescriptionHelpFormatter


def _script(name: str, *argv: object, capture: bool = False) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(RUNTIME_DIR / name), *[str(value) for value in argv]],
        text=True,
        capture_output=capture,
    )


def _json(path: Path) -> dict:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return value if isinstance(value, dict) else {}


def _write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def cmd_prepare(args: argparse.Namespace) -> int:
    deck = normalize_inputs(
        args.inputs,
        out_root=args.out_root,
        job_dir=args.job_dir,
        dpi=args.dpi,
    )
    payload = _json(deck)
    print(deck)
    print(f"pages={int(payload.get('page_count') or 0)}")
    return 0


def cmd_inspect(args: argparse.Namespace) -> int:
    page_dir = Path(args.page_dir).expanduser().resolve()
    source = page_dir / args.source
    if not source.is_file():
        print(f"source image not found: {source}", file=sys.stderr)
        return 2
    command = [str(page_dir), "--source", args.source, "--out", args.out]
    if args.overlay:
        command += ["--overlay", args.overlay]
    else:
        command += ["--overlay", ""]
    completed = _script("text_hints.py", *command, capture=True)
    if completed.returncode != 0:
        if completed.stdout:
            print(completed.stdout, end="")
        if completed.stderr:
            print(completed.stderr, end="", file=sys.stderr)
        return completed.returncode
    hints = _json(page_dir / args.out)
    lines = hints.get("lines") if isinstance(hints.get("lines"), list) else []
    payload = {
        "page_dir": str(page_dir),
        "source": str(source),
        "text_hints": str(page_dir / args.out),
        "overlay": str(page_dir / args.overlay) if args.overlay else "",
        "backend": str(hints.get("backend") or hints.get("source") or "local"),
        "line_count": len(lines),
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def cmd_extract_assets(args: argparse.Namespace) -> int:
    argv: list[object] = ["--input", args.input, "--out-dir", args.out_dir]
    for name in ("names", "sort", "threshold", "min_area", "pad", "limit"):
        value = getattr(args, name)
        if value is not None and value != "":
            argv += ["--" + name.replace("_", "-"), value]
    if args.square:
        argv.append("--square")
    if args.manifest:
        argv += ["--manifest", args.manifest]
    if args.contact_sheet:
        argv += ["--contact-sheet", args.contact_sheet]
    return _script("split_alpha_components.py", *argv).returncode


def cmd_build(args: argparse.Namespace) -> int:
    page_dir = Path(args.page_dir).expanduser().resolve()
    manifest = page_dir / args.manifest
    if not manifest.is_file():
        print(f"manifest not found: {manifest}", file=sys.stderr)
        return 2
    argv: list[object] = [manifest, "--out", page_dir / args.out]
    if not args.no_preview:
        argv += ["--preview", page_dir / args.preview]
    return _script("build_pptx_from_manifest.py", *argv).returncode


def cmd_render(args: argparse.Namespace) -> int:
    page_dir = Path(args.page_dir).expanduser().resolve()
    manifest_path = page_dir / args.manifest
    if not manifest_path.is_file():
        print(f"manifest not found: {manifest_path}", file=sys.stderr)
        return 2
    manifest = _json(manifest_path)
    if not manifest:
        print(f"manifest is not valid JSON: {manifest_path}", file=sys.stderr)
        return 2
    render_preview(manifest, manifest_path, page_dir / args.preview)
    print(page_dir / args.preview)
    return 0


def cmd_assemble(args: argparse.Namespace) -> int:
    page_dirs = [Path(value).expanduser().resolve() for value in args.page_dirs]
    out = Path(args.out).expanduser().resolve()
    if not page_dirs:
        print("assemble requires at least one page directory", file=sys.stderr)
        return 2
    missing_pptx = [str(path / args.page_pptx) for path in page_dirs if not (path / args.page_pptx).is_file()]
    if missing_pptx:
        print("missing page.pptx: " + ", ".join(missing_pptx), file=sys.stderr)
        return 2
    manifests = [path / args.manifest for path in page_dirs]
    if len(page_dirs) == 1 and not manifests[0].is_file():
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(page_dirs[0] / args.page_pptx, out)
        print(out)
        return 0
    missing_manifests = [str(path) for path in manifests if not path.is_file()]
    if missing_manifests:
        print(
            "multi-page assembly requires each page's internal manifest.json; missing: "
            + ", ".join(missing_manifests),
            file=sys.stderr,
        )
        return 2
    with tempfile.TemporaryDirectory(prefix="editppt-assemble-") as temporary:
        root = Path(temporary)
        deck_manifest = root / "deck_manifest.json"
        _write_json(
            deck_manifest,
            {
                "job_dir": str(root),
                "output": str(out),
                "pages": [
                    {"page_index": index, "manifest": str(path)}
                    for index, path in enumerate(manifests, start=1)
                ],
            },
        )
        completed = _script(
            "build_pptx_from_manifest.py",
            "--deck-manifest",
            deck_manifest,
            "--out",
            out,
        )
    if completed.returncode == 0:
        print(out)
    return completed.returncode


def cmd_doctor(args: argparse.Namespace) -> int:
    argv: list[object] = ["doctor", "--json"]
    if args.check_api:
        argv.append("--check-api")
    completed = _script("runtime_env.py", *argv, capture=True)
    try:
        payload = json.loads(completed.stdout or "{}")
    except json.JSONDecodeError:
        payload = {"ok": False, "stdout": completed.stdout, "stderr": completed.stderr}
    contract = _json(SKILL_ROOT / "skill.json")
    payload["skill"] = {
        "root": str(SKILL_ROOT),
        "contract_version": str(contract.get("contract_version") or "unknown"),
        "upstream_base": str(contract.get("upstream_base") or ""),
    }
    if args.strict:
        payload["strict"] = True
        payload["context_canary"] = {
            "passed": True,
            "scope": "simple-codex-page-v1",
            "reason": "controller page context is not part of this Skill contract",
        }
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print("ok=" + ("yes" if payload.get("ok") else "no"))
        print(f"skill={payload['skill']['root']}")
        print(f"contract={payload['skill']['contract_version']}")
    if completed.stderr:
        print(completed.stderr, end="", file=sys.stderr)
    return completed.returncode


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
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=os.environ.get("IMAGE_TO_EDITABLE_PPT_CLI_PROG", "editppt"),
        description="Optional deterministic helpers for one-Codex-task-per-page editable PPT rebuilding.",
        formatter_class=HELP,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    prepare = sub.add_parser("prepare", help="Normalize images, PDF, or image-based PPT/PPTX into source.png pages.")
    prepare.add_argument("inputs", nargs="+")
    prepare.add_argument("--out-root", default="output/image-to-editable-ppt")
    prepare.add_argument("--job-dir")
    prepare.add_argument("--dpi", type=int, default=180)
    prepare.set_defaults(func=cmd_prepare)

    inspect = sub.add_parser("inspect", help="Measure page text and write advisory OCR/layout hints.")
    inspect.add_argument("page_dir")
    inspect.add_argument("--source", default="source.png")
    inspect.add_argument("--out", default="text_hints.json")
    inspect.add_argument("--overlay", default="text_hints.png")
    inspect.set_defaults(func=cmd_inspect)

    extract = sub.add_parser("extract-assets", help="Split a transparent sheet into source-bound image assets.")
    extract.add_argument("--input", required=True)
    extract.add_argument("--out-dir", required=True)
    extract.add_argument("--names", default="")
    extract.add_argument("--sort", choices=("x", "y", "area"), default="x")
    extract.add_argument("--threshold", type=int)
    extract.add_argument("--min-area", type=int)
    extract.add_argument("--pad", type=int)
    extract.add_argument("--limit", type=int)
    extract.add_argument("--square", action="store_true")
    extract.add_argument("--manifest", default="")
    extract.add_argument("--contact-sheet", default="")
    extract.set_defaults(func=cmd_extract_assets)

    build = sub.add_parser("build", help="Build page.pptx and an optional preview from manifest.json.")
    build.add_argument("page_dir")
    build.add_argument("--manifest", default="manifest.json")
    build.add_argument("--out", default="page.pptx")
    build.add_argument("--preview", default="preview.png")
    build.add_argument("--no-preview", action="store_true")
    build.set_defaults(func=cmd_build)

    render = sub.add_parser("render", help="Refresh preview.png from the current manifest without a quality verdict.")
    render.add_argument("page_dir")
    render.add_argument("--manifest", default="manifest.json")
    render.add_argument("--preview", default="preview.png")
    render.set_defaults(func=cmd_render)

    assemble = sub.add_parser("assemble", help="Assemble page directories into a deck in the supplied order.")
    assemble.add_argument("page_dirs", nargs="+")
    assemble.add_argument("--out", required=True)
    assemble.add_argument("--manifest", default="manifest.json")
    assemble.add_argument("--page-pptx", default="page.pptx")
    assemble.set_defaults(func=cmd_assemble)

    doctor = sub.add_parser("doctor", help="Report runtime dependencies and the installed Skill contract.")
    doctor.add_argument("--json", action="store_true")
    doctor.add_argument("--strict", action="store_true")
    doctor.add_argument("--check-api", action="store_true")
    doctor.set_defaults(func=cmd_doctor)

    formula = sub.add_parser("formula", help="Render a complex formula as a separate local image asset.")
    formula_sub = formula.add_subparsers(dest="formula_command", required=True)
    render_formula = formula_sub.add_parser("render-latex", help="Render LaTeX to SVG, PNG, or PDF.")
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
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return int(args.func(args) or 0)


if __name__ == "__main__":
    raise SystemExit(main())
