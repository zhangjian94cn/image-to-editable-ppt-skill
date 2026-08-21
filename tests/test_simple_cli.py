import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
RUNTIME = ROOT / "skills/image-to-editable-ppt/cli/editppt/runtime/main.py"
sys.path.insert(0, str(RUNTIME.parent))
import main as runtime_main  # noqa: E402
sys.path.pop(0)


def run_cli(*args: object) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(RUNTIME), *[str(value) for value in args]],
        text=True,
        capture_output=True,
    )


def manifest(text: str) -> dict:
    return {
        "slide": {"width": 13.333, "height": 7.5, "background": "#FFFFFF"},
        "source": {"width_px": 1600, "height_px": 900},
        "shapes": [
            {
                "type": "rect",
                "box_px": [80, 180, 1440, 500],
                "fill": "#EAF3FF",
                "stroke": "#2375D8",
            }
        ],
        "images": [],
        "text_boxes": [
            {
                "text": text,
                "box_px": [120, 80, 1360, 90],
                "font_size": 28,
                "color": "#0A65B7",
            }
        ],
    }


class SimpleCliTest(unittest.TestCase):
    def test_help_exposes_only_simple_page_helpers(self):
        completed = run_cli("--help")
        self.assertEqual(0, completed.returncode, completed.stderr)
        for command in ("prepare", "inspect", "extract-assets", "build", "render", "assemble", "doctor"):
            self.assertIn(command, completed.stdout)
        self.assertNotIn("dispatch", completed.stdout)
        self.assertNotIn("finalize", completed.stdout)

    def test_prepare_writes_ordered_source_pages_without_agent_state(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            first = root / "one.png"
            second = root / "two.png"
            Image.new("RGB", (160, 90), "white").save(first)
            Image.new("RGB", (160, 90), "blue").save(second)
            run_dir = root / "run"
            completed = run_cli("prepare", first, second, "--job-dir", run_dir)
            self.assertEqual(0, completed.returncode, completed.stderr)
            self.assertTrue((run_dir / "pages/page_001/source.png").is_file())
            self.assertTrue((run_dir / "pages/page_002/source.png").is_file())
            self.assertFalse((run_dir / "page_jobs.json").exists())

    def test_build_render_and_assemble_keep_page_order(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            pages = []
            for index, title in enumerate(("第一页", "第二页"), start=1):
                page = root / f"page-{index}"
                page.mkdir()
                Image.new("RGB", (1600, 900), "white").save(page / "source.png")
                (page / "manifest.json").write_text(
                    json.dumps(manifest(title), ensure_ascii=False), encoding="utf-8"
                )
                built = run_cli("build", page)
                self.assertEqual(0, built.returncode, built.stderr)
                self.assertTrue((page / "page.pptx").is_file())
                self.assertTrue((page / "preview.png").is_file())
                pages.append(page)

            out = root / "deck.pptx"
            assembled = run_cli("assemble", *pages, "--out", out)
            self.assertEqual(0, assembled.returncode, assembled.stderr)
            with zipfile.ZipFile(out) as package:
                presentation = package.read("ppt/presentation.xml").decode("utf-8")
                self.assertIn("<Slides>2</Slides>", package.read("docProps/app.xml").decode("utf-8"))
                self.assertEqual(2, presentation.count("<p:sldId "))
                first_xml = package.read("ppt/slides/slide1.xml").decode("utf-8")
                second_xml = package.read("ppt/slides/slide2.xml").decode("utf-8")
                self.assertIn("第一页", first_xml)
                self.assertIn("第二页", second_xml)

    def test_doctor_reports_contract_version(self):
        completed = run_cli("doctor", "--json")
        self.assertEqual(0, completed.returncode, completed.stderr)
        payload = json.loads(completed.stdout)
        self.assertEqual("simple-codex-page-v1", payload["skill"]["contract_version"])

    def test_inspect_uses_configured_content_aware_ocr_without_exposing_token(self):
        with tempfile.TemporaryDirectory() as temporary:
            page = Path(temporary)
            Image.new("RGB", (160, 90), "white").save(page / "source.png")
            calls: list[tuple[str, dict[str, str] | None]] = []

            def fake_script(name: str, *argv: object, capture: bool = False, env=None):
                calls.append((name, env))
                (page / "text_hints.json").write_text(
                    json.dumps({"backend": "paddleocr-vl", "lines": []}),
                    encoding="utf-8",
                )
                return subprocess.CompletedProcess([], 0, "", "")

            args = SimpleNamespace(
                page_dir=page,
                source="source.png",
                out="text_hints.json",
                overlay="text_hints.png",
            )
            with mock.patch.object(runtime_main, "_configured_paddle_token", return_value="secret-value"), mock.patch.object(
                runtime_main, "_script", side_effect=fake_script
            ):
                self.assertEqual(0, runtime_main.cmd_inspect(args))

            self.assertEqual("paddle_text_hints.py", calls[0][0])
            self.assertEqual("secret-value", calls[0][1]["PADDLE_OCR_TOKEN"])


if __name__ == "__main__":
    unittest.main()
