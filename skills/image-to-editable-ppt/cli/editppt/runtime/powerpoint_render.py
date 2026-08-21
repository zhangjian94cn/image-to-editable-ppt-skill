#!/usr/bin/env python3
"""Fail-closed Microsoft PowerPoint renderer for one-page candidate PPTX files.

Only the collision-proof staged target is opened, exported, and closed. The
renderer never creates an unrelated canary and never closes user presentations.
"""

from __future__ import annotations

import fcntl
import json
import os
import plistlib
import shutil
import subprocess
import tempfile
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

import fitz


POWERPOINT_APP = Path("/Applications/Microsoft PowerPoint.app")
POWERPOINT_BUNDLE_ID = "com.microsoft.Powerpoint"
LOCK_PATH = Path(tempfile.gettempdir()) / "editppt-powerpoint-render.lock"
REPAIR_PATTERNS = (
    'Data.Repair":true',
    "OnFileCorrupt",
    "ReopenDocumentForRepair",
    "DocumentXmlReader::DocumentXmlReadStart: Failed",
    "XmlReader::ReadXmlHelper Failed",
    "modal window blocking open event",
)
LOG_ROOT = (
    Path.home()
    / "Library/Containers/com.microsoft.Powerpoint/Data/Library/Logs/Diagnostics/POWERPOINT"
)


class PowerPointRenderError(RuntimeError):
    pass


def _powerpoint_version() -> str:
    try:
        payload = plistlib.loads((POWERPOINT_APP / "Contents/Info.plist").read_bytes())
    except (OSError, plistlib.InvalidFileException):
        return ""
    short = str(payload.get("CFBundleShortVersionString") or "")
    build = str(payload.get("CFBundleVersion") or "")
    return f"{short} ({build})" if short and build else short or build


def _escape(value: str | Path) -> str:
    return str(value).replace("\\", "\\\\").replace('"', '\\"')


def _run_osascript(script: str, *, timeout: int) -> str:
    try:
        completed = subprocess.run(
            ["/usr/bin/osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise PowerPointRenderError("PowerPoint AppleScript timed out; a modal dialog may be open") from exc
    if completed.returncode != 0:
        message = completed.stderr.strip() or completed.stdout.strip() or "AppleScript failed"
        raise PowerPointRenderError(message[:1200])
    return completed.stdout.strip()


def presentation_snapshot() -> list[dict[str, str]]:
    script = r'''
tell application "Microsoft PowerPoint"
  set rows to {}
  repeat with deck in presentations
    try
      set end of rows to ((name of deck as text) & tab & (full name of deck as text))
    on error
      set end of rows to ((name of deck as text) & tab)
    end try
  end repeat
end tell
set AppleScript's text item delimiters to linefeed
return rows as text
'''.strip()
    raw = _run_osascript(script, timeout=10)
    rows: list[dict[str, str]] = []
    for line in raw.splitlines():
        name, separator, full_name = line.partition("\t")
        if name:
            rows.append({"name": name, "full_name": full_name if separator else ""})
    return rows


def _readback(expected_name: str) -> dict[str, Any] | None:
    script = f'''
set expectedName to "{_escape(expected_name)}"
tell application "Microsoft PowerPoint"
  try
    set deck to presentation expectedName
    if (name of deck as text) is not expectedName then return ""
    return (name of deck as text) & tab & (count of slides of deck as text) & tab & (full name of deck as text)
  on error
    return ""
  end try
end tell
'''.strip()
    raw = _run_osascript(script, timeout=10)
    values = raw.split("\t")
    if len(values) != 3 or not values[1].isdigit():
        return None
    return {"name": values[0], "slide_count": int(values[1]), "full_name": values[2]}


def _close_exact(expected_name: str) -> None:
    script = f'''
tell application "Microsoft PowerPoint"
  if exists presentation "{_escape(expected_name)}" then
    close presentation "{_escape(expected_name)}" saving no
  end if
end tell
'''.strip()
    _run_osascript(script, timeout=20)


def _export_exact(expected_name: str, output_pdf: Path) -> None:
    script = f'''
set outputFile to POSIX file "{_escape(output_pdf.resolve())}"
set expectedName to "{_escape(expected_name)}"
tell application "Microsoft PowerPoint"
  set deck to presentation expectedName
  if (name of deck as text) is not expectedName then error "staged presentation identity changed"
  save deck in outputFile as save as PDF
end tell
'''.strip()
    _run_osascript(script, timeout=180)


def _log_cursor() -> dict[Path, int]:
    if not LOG_ROOT.is_dir():
        return {}
    cursor: dict[Path, int] = {}
    for path in LOG_ROOT.glob("*.log"):
        try:
            cursor[path] = path.stat().st_size
        except OSError:
            continue
    return cursor


def _repair_evidence(cursor: dict[Path, int]) -> list[str]:
    if not LOG_ROOT.is_dir():
        return []
    evidence: list[str] = []
    for path in LOG_ROOT.glob("*.log"):
        try:
            start = cursor.get(path, 0)
            with path.open("rb") as handle:
                handle.seek(min(start, path.stat().st_size))
                text = handle.read().decode("utf-8", errors="ignore")
        except OSError:
            continue
        for line in text.splitlines():
            if any(pattern.casefold() in line.casefold() for pattern in REPAIR_PATTERNS):
                evidence.append(line[-800:])
    return evidence[-20:]


@contextmanager
def _render_lock(timeout: float = 180.0) -> Iterator[None]:
    LOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOCK_PATH.open("a+") as handle:
        deadline = time.monotonic() + timeout
        while True:
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except BlockingIOError:
                if time.monotonic() >= deadline:
                    raise PowerPointRenderError("another PowerPoint render owns the host lock")
                time.sleep(0.25)
        try:
            yield
        finally:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def _valid_pdf(path: Path, expected_slides: int) -> bool:
    if not path.is_file() or path.stat().st_size < 8 or path.read_bytes()[:5] != b"%PDF-":
        return False
    try:
        with fitz.open(path) as document:
            return document.page_count == expected_slides
    except (OSError, RuntimeError, ValueError):
        return False


def _render_pdf_page(pdf: Path, output_png: Path, dpi: int) -> None:
    with fitz.open(pdf) as document:
        if document.page_count != 1:
            raise PowerPointRenderError(f"page renderer requires exactly one slide, found {document.page_count}")
        pixmap = document[0].get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0), alpha=False)
        output_png.parent.mkdir(parents=True, exist_ok=True)
        pixmap.save(output_png)


def render_one_page(
    pptx: Path,
    output_png: Path,
    *,
    evidence_dir: Path,
    dpi: int = 200,
) -> dict[str, Any]:
    if os.uname().sysname != "Darwin" or not POWERPOINT_APP.is_dir():
        raise PowerPointRenderError("Microsoft PowerPoint for macOS is unavailable")
    if shutil.which("osascript") is None or shutil.which("open") is None:
        raise PowerPointRenderError("required macOS automation commands are unavailable")
    pptx = pptx.resolve()
    output_png = output_png.resolve()
    evidence_dir = evidence_dir.resolve()
    evidence_dir.mkdir(parents=True, exist_ok=True)
    if not pptx.is_file() or pptx.stat().st_size == 0:
        raise PowerPointRenderError(f"non-empty PPTX not found: {pptx}")

    report_path = evidence_dir / "powerpoint-render.json"
    with _render_lock():
        before = presentation_snapshot()
        cursor = _log_cursor()
        configured_root = os.environ.get("EDITPPT_POWERPOINT_CONTAINER_ROOT", "").strip()
        container_root = (
            Path(configured_root).expanduser()
            if configured_root
            else Path.home() / "Library/Containers" / POWERPOINT_BUNDLE_ID / "Data/tmp"
        )
        container_root.mkdir(parents=True, exist_ok=True)
        staging = Path(tempfile.mkdtemp(prefix="editppt-render-", dir=container_root))
        expected_name = f"editppt-render-{uuid.uuid4().hex}.pptx"
        staged_pptx = staging / expected_name
        staged_pdf = staging / "rendered.pdf"
        shutil.copy2(pptx, staged_pptx)
        readback: dict[str, Any] | None = None
        closed = False
        try:
            launched = subprocess.run(
                ["/usr/bin/open", "-gj", "-a", str(POWERPOINT_APP), str(staged_pptx)],
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
            if launched.returncode != 0:
                raise PowerPointRenderError(
                    (launched.stderr.strip() or launched.stdout.strip() or "PowerPoint launch failed")[:1200]
                )
            deadline = time.monotonic() + 90
            while time.monotonic() < deadline:
                repairs = _repair_evidence(cursor)
                if repairs:
                    raise PowerPointRenderError("PowerPoint reported repair/corruption: " + repairs[-1])
                readback = _readback(expected_name)
                if readback:
                    break
                time.sleep(0.5)
            if not readback:
                raise PowerPointRenderError("PowerPoint did not open the exact staged presentation")
            if readback["slide_count"] != 1:
                raise PowerPointRenderError(
                    f"page renderer requires one slide; PowerPoint read back {readback['slide_count']}"
                )
            _export_exact(expected_name, staged_pdf)
            repairs = _repair_evidence(cursor)
            if repairs:
                raise PowerPointRenderError("PowerPoint reported repair/corruption: " + repairs[-1])
            if not _valid_pdf(staged_pdf, readback["slide_count"]):
                raise PowerPointRenderError("PowerPoint did not export a validated one-page PDF")
            _close_exact(expected_name)
            closed = True
            after = presentation_snapshot()
            if after != before:
                raise PowerPointRenderError("user presentation state changed during rendering")
            evidence_pdf = evidence_dir / "powerpoint.pdf"
            shutil.copy2(staged_pdf, evidence_pdf)
            _render_pdf_page(evidence_pdf, output_png, dpi)
            report = {
                "schema_version": 1,
                "success": True,
                "authoritative": True,
                "renderer": "microsoft-powerpoint",
                "renderer_version": _powerpoint_version(),
                "target_only": True,
                "canary_created": False,
                "readback": readback,
                "before_presentations": before,
                "after_presentations": after,
                "repair_evidence": [],
                "pdf": str(evidence_pdf),
                "png": str(output_png),
                "dpi": dpi,
            }
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            shutil.rmtree(staging, ignore_errors=True)
            return report
        except Exception as exc:
            if readback and not closed:
                try:
                    _close_exact(expected_name)
                    closed = True
                except PowerPointRenderError:
                    pass
            report = {
                "schema_version": 1,
                "success": False,
                "authoritative": False,
                "renderer": "microsoft-powerpoint",
                "renderer_version": _powerpoint_version(),
                "target_only": True,
                "canary_created": False,
                "error": str(exc),
                "readback": readback or {},
                "before_presentations": before,
                "repair_evidence": _repair_evidence(cursor),
                "staging": str(staging),
                "closed": closed,
            }
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            if isinstance(exc, PowerPointRenderError):
                raise
            raise PowerPointRenderError(str(exc)) from exc
