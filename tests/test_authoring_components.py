from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image
from pptx import Presentation


ROOT = Path(__file__).resolve().parents[1]
CLI_ROOT = ROOT / "skills/image-to-editable-ppt/cli"
RUNTIME = CLI_ROOT / "editppt/runtime/main.py"
sys.path.insert(0, str(CLI_ROOT))
from editppt.authoring import SlideManifest  # noqa: E402


def run_cli(*args: object) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(RUNTIME), *[str(value) for value in args]],
        text=True,
        capture_output=True,
    )


def test_components_build_native_connector_and_rich_table():
    with tempfile.TemporaryDirectory() as temporary:
        page = Path(temporary)
        Image.new("RGB", (1600, 900), "white").save(page / "source.png")
        manifest = SlideManifest(1600, 900)
        manifest.add_section_band([80, 60, 1440, 64], "原生组件", fill="#0A65B7")
        manifest.add_shape(
            [80, 140, 1440, 40],
            gradient={
                "angle": 0,
                "stops": [
                    {"position": 0, "color": "#EAF3FF"},
                    {"position": 1, "color": "#FFFFFF"},
                ],
            },
        )
        manifest.add_timeline_stage([120, 220, 280, 48], number="1", period="第1-2周", heading="完成闭环")
        manifest.add_connector([400, 244], [620, 244], end_arrow="triangle")
        manifest.add_table(
            [120, 420, 1360, 260],
            [
                ["阶段", "说明"],
                ["准备", {"runs": [{"text": "可编辑", "bold": True, "color": "#D92D20"}, {"text": "单元格"}]}],
            ],
            column_widths=[1, 3],
        )
        manifest.write(page / "manifest.json")

        built = run_cli("build", page)
        assert built.returncode == 0, built.stderr
        build_payload = json.loads(built.stdout)
        assert Path(build_payload["build_report"]).is_file()
        inspected = run_cli("inspect", "pptx", page)
        assert inspected.returncode == 0, inspected.stderr
        payload = json.loads(inspected.stdout)
        assert payload["summary"]["connector_count"] == 1
        assert payload["summary"]["table_count"] == 1
        with zipfile.ZipFile(page / "page.pptx") as package:
            slide_xml = package.read("ppt/slides/slide1.xml").decode("utf-8")
        assert "<p:cxnSp>" in slide_xml
        assert '<a:tailEnd type="triangle"/>' in slide_xml
        assert '<a:gradFill rotWithShape="1">' in slide_xml
        assert "可编辑" in slide_xml and "单元格" in slide_xml


def test_builder_resolves_unavailable_font_and_protects_single_line_title():
    with tempfile.TemporaryDirectory() as temporary:
        page = Path(temporary)
        Image.new("RGB", (1600, 900), "white").save(page / "source.png")
        value = SlideManifest(1600, 900)
        value.add_text(
            [53, 34, 1040, 53],
            "各阶段里程碑 | 看闭环、看资产，而不是只看功能数量",
            font="Microsoft YaHei",
            font_size=30.5,
            font_size_source="measured",
            bold=True,
        )
        value.write(page / "manifest.json")
        built = run_cli("build", page)
        assert built.returncode == 0, built.stderr
        build_payload = json.loads(built.stdout)
        assert build_payload["font_substitutions"]
        assert build_payload["text_adjustments"]
        deck = Presentation(str(page / "page.pptx"))
        shape = next(shape for shape in deck.slides[0].shapes if getattr(shape, "text", ""))
        run = shape.text_frame.paragraphs[0].runs[0]
        assert run.font.name != "Microsoft YaHei"
        assert run.font.size.pt < 30.5


def test_builder_accepts_common_aliases_and_source_pixel_font_sizes():
    with tempfile.TemporaryDirectory() as temporary:
        page = Path(temporary)
        Image.new("RGB", (1600, 900), "white").save(page / "source.png")
        value = SlideManifest(1600, 900)
        value.add_shape([100, 100, 300, 80], "roundRect", fill="#EAF3FF", radius=12)
        value.add_connector([400, 140], [600, 140], stroke="#0A65B7", stroke_width=2)
        value.add_text([120, 110, 250, 50], "", font_size_px=30, valign="mid")
        value.add_table([100, 250, 600, 160], [["表头"], [{"text": "内容", "font_size_px": 20}]], font_size_px=24)
        value.write(page / "manifest.json")

        built = run_cli("build", page)
        assert built.returncode == 0, built.stderr
        with zipfile.ZipFile(page / "page.pptx") as package:
            slide_xml = package.read("ppt/slides/slide1.xml").decode("utf-8")
        assert 'anchor="ctr"' in slide_xml
        assert '<a:gd name="adj" fmla="val 15000"/>' in slide_xml
        assert 'w="25400"' in slide_xml
        assert 'sz="1800"' in slide_xml
        assert 'sz="1440"' in slide_xml
        assert 'sz="1200"' in slide_xml


def test_builder_surfaces_severe_automatic_text_shrink():
    with tempfile.TemporaryDirectory() as temporary:
        page = Path(temporary)
        Image.new("RGB", (1600, 900), "white").save(page / "source.png")
        value = SlideManifest(1600, 900)
        value.add_text([10, 10, 120, 20], "这是一个明显放不下的很长标题", font_size=32)
        value.write(page / "manifest.json")

        built = run_cli("build", page)
        assert built.returncode == 0, built.stderr
        payload = json.loads(built.stdout)
        assert payload["severe_text_adjustments"]
        assert payload["warnings"]
