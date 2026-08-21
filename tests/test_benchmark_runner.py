from __future__ import annotations

import json
from argparse import Namespace
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches

from benchmark_runner.runner import ROOT_FILES, _issues, _pptx_metrics, _resolve_cases


def _corpus(tmp_path: Path) -> Path:
    root = tmp_path / "corpus"
    page = root / "decks/case/slides/001"
    page.mkdir(parents=True)
    Image.new("RGB", (1600, 900), "white").save(page / "input.png")
    (page / "expected.json").write_text(json.dumps({"objects": [{"kind": "text", "text": "标题", "visible": True}]}))
    (root / "decks/case/deck.json").write_text(json.dumps({"slides": [{"number": 1, "path": "slides/001"}]}))
    (root / "suites.json").write_text(json.dumps({"page_suites": {"round-0": [{"deck_id": "case", "slide": 1}]}}))
    return root


def test_page_suite_resolves_one_frozen_page(tmp_path: Path):
    cases = _resolve_cases(_corpus(tmp_path), "round-0")
    assert [value.page_id for value in cases] == ["case-p001"]


def test_pptx_readback_flags_full_page_picture_and_text(tmp_path: Path):
    pptx = tmp_path / "candidate.pptx"
    picture = tmp_path / "picture.png"
    Image.new("RGB", (1600, 900), "white").save(picture)
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_picture(str(picture), 0, 0, presentation.slide_width, presentation.slide_height)
    slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1)).text = "标题"
    presentation.save(pptx)
    expected = tmp_path / "expected.json"
    expected.write_text(json.dumps({"objects": [{"kind": "text", "text": "标题", "visible": True}]}))
    metrics = _pptx_metrics(pptx, expected)
    assert metrics["text_coverage"] == 1.0
    assert metrics["max_picture_coverage"] > 0.99
    verdict, issues = _issues({**metrics, "coarse_rgb_loss": 0.0, "content_ink_loss": 0.0})
    assert verdict == "reject"
    assert any(value["category"] == "editability" for value in issues)


def test_snapshot_root_contract_is_review_first():
    assert ROOT_FILES == {"source.png", "candidate.pptx", "candidate.png", "report.md", "artifacts"}
