from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pymupdf
from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
RUNTIME_DIR = ROOT / "skills/image-to-editable-ppt/cli/editppt/runtime"
sys.path.insert(0, str(ROOT / "skills/image-to-editable-ppt/cli"))
sys.path.insert(0, str(RUNTIME_DIR))
import powerpoint_render as renderer  # noqa: E402
import quality_tools  # noqa: E402


def _one_page_pdf(path: Path) -> None:
    document = pymupdf.open()
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


def test_presentation_snapshot_handles_zero_open_presentations(monkeypatch):
    scripts: list[str] = []

    def run(script: str, timeout: int) -> str:
        scripts.append(script)
        return ""

    monkeypatch.setattr(renderer, "_run_osascript", run)
    assert renderer.presentation_snapshot() == []
    assert "count of presentations" in scripts[0]
    assert "repeat with deck in presentations" not in scripts[0]


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


def test_quality_render_reuses_sha_version_and_dpi_bound_output(tmp_path, monkeypatch):
    app = tmp_path / "Microsoft PowerPoint.app"
    app.mkdir()
    monkeypatch.setattr(quality_tools, "POWERPOINT_APP", app)
    monkeypatch.setattr(quality_tools.sys, "platform", "darwin")
    monkeypatch.setattr(quality_tools, "powerpoint_version", lambda: "test-version")
    calls = []

    def fake_render(pptx, output, *, evidence_dir, dpi):
        calls.append((pptx, dpi))
        Image.new("RGB", (320, 180), "white").save(output)
        quality_tools.write_json(evidence_dir / "powerpoint-render.json", {"success": True})
        return {"success": True}

    monkeypatch.setattr(quality_tools, "render_one_page", fake_render)
    pptx = tmp_path / "page.pptx"
    pptx.write_bytes(b"stable-pptx")
    output = tmp_path / "preview.png"
    evidence = tmp_path / "evidence"

    first = quality_tools.render_powerpoint(pptx, output, evidence_dir=evidence, dpi=200)
    second = quality_tools.render_powerpoint(pptx, output, evidence_dir=evidence, dpi=200)

    assert first["cache_hit"] is False
    assert second["cache_hit"] is True
    assert len(calls) == 1


def test_compare_emits_clustered_local_source_candidate_pairs(tmp_path):
    source = Image.new("RGB", (320, 180), "white")
    candidate = source.copy()
    ImageDraw.Draw(candidate).rectangle([40, 30, 110, 90], fill="black")
    source_path = tmp_path / "source.png"
    candidate_path = tmp_path / "candidate.png"
    source.save(source_path)
    candidate.save(candidate_path)

    report = quality_tools.compare_images(source_path, candidate_path, tmp_path / "compare")

    assert report["regions"]
    region = report["regions"][0]
    assert Path(region["source"]).is_file()
    assert Path(region["candidate"]).is_file()
    assert Path(region["diff"]).is_file()
