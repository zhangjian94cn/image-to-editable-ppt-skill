from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import fitz


ROOT = Path(__file__).resolve().parents[1]
RUNTIME_DIR = ROOT / "skills/image-to-editable-ppt/cli/editppt/runtime"
sys.path.insert(0, str(RUNTIME_DIR))
import powerpoint_render as renderer  # noqa: E402


def _one_page_pdf(path: Path) -> None:
    document = fitz.open()
    document.new_page(width=960, height=540)
    document.save(path)
    document.close()


def test_presentation_snapshot_parses_exact_names_and_paths(monkeypatch):
    monkeypatch.setattr(
        renderer,
        "_run_osascript",
        lambda script, timeout: "existing.pptx\t/Users/test/existing.pptx\nunsaved\t",
    )
    assert renderer.presentation_snapshot() == [
        {"name": "existing.pptx", "full_name": "/Users/test/existing.pptx"},
        {"name": "unsaved", "full_name": ""},
    ]


def test_target_only_render_writes_authoritative_png_without_canary(tmp_path, monkeypatch):
    app = tmp_path / "Microsoft PowerPoint.app"
    app.mkdir()
    monkeypatch.setattr(renderer, "POWERPOINT_APP", app)
    monkeypatch.setattr(renderer.shutil, "which", lambda name: f"/usr/bin/{name}")
    monkeypatch.setenv("EDITPPT_POWERPOINT_CONTAINER_ROOT", str(tmp_path / "office"))
    before = [{"name": "user.pptx", "full_name": "/Users/test/user.pptx"}]
    snapshots = iter((before, before))
    monkeypatch.setattr(renderer, "presentation_snapshot", lambda: next(snapshots))
    monkeypatch.setattr(renderer, "_log_cursor", lambda: {})
    monkeypatch.setattr(renderer, "_repair_evidence", lambda cursor: [])
    monkeypatch.setattr(
        renderer,
        "_readback",
        lambda expected: {"name": expected, "slide_count": 1, "full_name": f"/staged/{expected}"},
    )
    closed: list[str] = []
    monkeypatch.setattr(renderer, "_close_exact", closed.append)
    monkeypatch.setattr(renderer, "_export_exact", lambda expected, output: _one_page_pdf(output))
    monkeypatch.setattr(
        renderer.subprocess,
        "run",
        lambda *args, **kwargs: subprocess.CompletedProcess(args[0], 0, "", ""),
    )
    monkeypatch.setattr(renderer, "_powerpoint_version", lambda: "test-version")

    pptx = tmp_path / "page.pptx"
    pptx.write_bytes(b"not-opened-by-test")
    output = tmp_path / "preview.png"
    evidence = tmp_path / "evidence"
    report = renderer.render_one_page(pptx, output, evidence_dir=evidence, dpi=96)

    assert output.is_file()
    assert report["authoritative"] is True
    assert report["target_only"] is True
    assert report["canary_created"] is False
    assert report["before_presentations"] == report["after_presentations"]
    assert closed == [report["readback"]["name"]]
    assert (evidence / "powerpoint-render.json").is_file()
