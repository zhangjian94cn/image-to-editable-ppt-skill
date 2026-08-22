#!/usr/bin/env python3
"""Local dependency and optional OCR configuration diagnostics."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import shutil
import sys
from pathlib import Path


DEFAULT_CONFIG_HOME = "~/.editppt"
PADDLE_TOKEN_APPLY_URL = "https://aistudio.baidu.com/account/accessToken"


def runtime_home() -> Path:
    return Path(os.getenv("EDITPPT_CONFIG_HOME", DEFAULT_CONFIG_HOME)).expanduser()


def config_path(home: Path | None = None) -> Path:
    return (home or runtime_home()) / "config.yaml"


def read_config_file(path: Path) -> dict:
    path = Path(path).expanduser()
    if not path.exists():
        return {}
    try:
        import yaml
    except ImportError as exc:
        raise SystemExit("PyYAML is required to read ~/.editppt/config.yaml") from exc
    with path.open("r", encoding="utf-8") as handle:
        value = yaml.safe_load(handle) or {}
    if not isinstance(value, dict):
        raise SystemExit(f"invalid config file: {path}")
    return value


def collect_status() -> dict:
    configured = read_config_file(config_path())
    token = os.environ.get("PADDLE_OCR_TOKEN", "").strip() or str(configured.get("PADDLE_OCR_TOKEN") or "").strip()
    dependencies = {
        module: importlib.util.find_spec(module) is not None
        for module in ("fitz", "PIL", "pptx", "yaml", "numpy", "requests", "fontTools")
    }
    commands = {name: shutil.which(name) or "" for name in ("codex", "pdftoppm", "rsvg-convert", "osascript")}
    return {
        "ok": all(dependencies.values()) and bool(commands["codex"]),
        "cli_python": sys.executable,
        "dependencies": dependencies,
        "commands": commands,
        "text_hints": {
            "selection": "paddleocr-vl" if token else "local-geometric",
            "paddle_token": "set" if token else "unset",
            "apply_url": PADDLE_TOKEN_APPLY_URL,
            "configuration": "PADDLE_OCR_TOKEN env or ~/.editppt/config.yaml",
        },
        "note": "PowerPoint and the authoritative extractor are checked by the top-level editppt doctor.",
    }


def doctor(args: argparse.Namespace) -> int:
    payload = collect_status()
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print("ok=" + ("yes" if payload["ok"] else "no"))
        for module, ready in payload["dependencies"].items():
            print(f"python:{module}={'yes' if ready else 'no'}")
        for command, path in payload["commands"].items():
            print(f"command:{command}={path or 'missing'}")
        print(f"text_hints={payload['text_hints']['selection']}")
    return 0 if payload["ok"] else 1


def main() -> None:
    parser = argparse.ArgumentParser(prog="editppt-runtime")
    sub = parser.add_subparsers(required=True)
    doctor_parser = sub.add_parser("doctor")
    doctor_parser.add_argument("--json", action="store_true")
    doctor_parser.set_defaults(func=doctor)
    args = parser.parse_args()
    raise SystemExit(args.func(args))


if __name__ == "__main__":
    main()
