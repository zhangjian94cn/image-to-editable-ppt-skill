import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RUNTIME_DIR = ROOT / "skills/image-to-editable-ppt/cli/editppt/runtime"
sys.path.insert(0, str(RUNTIME_DIR))

from pptx_integrity import (  # noqa: E402
    classify_authored_validator_errors,
    is_allowed_authored_extension_warning,
    sanitize_orphaned_comment_authors,
)


class AuthoredIntegrityClassificationTest(unittest.TestCase):
    def test_known_office_chart_extension_is_an_allowlisted_warning(self):
        error = {
            "error_type": "Schema",
            "part": "/ppt/charts/chart1.xml",
            "xpath": "/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]",
            "description": (
                "The element has unexpected child element "
                "'http://schemas.microsoft.com/office/drawing/2012/chart:showLeaderLines'."
            ),
        }
        self.assertTrue(is_allowed_authored_extension_warning(error))
        blocking, warnings = classify_authored_validator_errors([error])
        self.assertEqual([], blocking)
        self.assertEqual([error], warnings)

    def test_duplicate_comment_author_id_blocks_powerpoint(self):
        error = {
            "error_type": "Semantic",
            "part": "/ppt/commentAuthors.xml",
            "xpath": "/p:cmAuthorLst[1]/p:cmAuthor[2979]",
            "description": (
                "Attribute 'id' should have unique value. Its current value "
                "'2979' duplicates with others."
            ),
        }
        self.assertFalse(is_allowed_authored_extension_warning(error))
        blocking, warnings = classify_authored_validator_errors([error])
        self.assertEqual([error], blocking)
        self.assertEqual([], warnings)

    def _write_comment_author_package(self, path: Path, *, with_comments: bool = False):
        content_types = b'''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>
</Types>'''
        relationships = b'''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>
</Relationships>'''
        authors = b'''<?xml version="1.0" encoding="UTF-8"?>
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="7" name="A" initials="A" lastIdx="0" clrIdx="0"/>
  <p:cmAuthor id="7" name="B" initials="B" lastIdx="0" clrIdx="1"/>
</p:cmAuthorLst>'''
        with zipfile.ZipFile(path, "w") as package:
            package.writestr("[Content_Types].xml", content_types)
            package.writestr("ppt/_rels/presentation.xml.rels", relationships)
            package.writestr("ppt/commentAuthors.xml", authors)
            package.writestr("ppt/presentation.xml", b"<p:presentation xmlns:p='urn:p'/>")
            if with_comments:
                package.writestr("ppt/comments/comment1.xml", b"<p:cmLst xmlns:p='urn:p'/>")

    def _duplicate_report(self):
        return {
            "validator_classification": {
                "blocking_errors": [{
                    "error_type": "Semantic",
                    "part": "/ppt/commentAuthors.xml",
                    "xpath": "/p:cmAuthorLst[1]/p:cmAuthor[2]",
                    "description": "duplicate id",
                }]
            }
        }

    def test_orphaned_comment_authors_are_removed_from_a_derived_copy(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.pptx"
            derived = Path(temporary) / "derived.pptx"
            self._write_comment_author_package(source)
            original = source.read_bytes()
            result = sanitize_orphaned_comment_authors(
                source, derived, integrity_report=self._duplicate_report()
            )
            self.assertEqual("applied", result["status"])
            self.assertFalse(result["visual_content_changed"])
            self.assertEqual(original, source.read_bytes())
            with zipfile.ZipFile(derived) as package:
                self.assertNotIn("ppt/commentAuthors.xml", package.namelist())
                self.assertNotIn(b"commentAuthors", package.read("[Content_Types].xml"))
                self.assertNotIn(
                    b"commentAuthors", package.read("ppt/_rels/presentation.xml.rels")
                )

    def test_real_comment_parts_make_sanitization_inapplicable(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.pptx"
            derived = Path(temporary) / "derived.pptx"
            self._write_comment_author_package(source, with_comments=True)
            result = sanitize_orphaned_comment_authors(
                source, derived, integrity_report=self._duplicate_report()
            )
            self.assertEqual("not_applicable", result["status"])
            self.assertFalse(derived.exists())

    def test_core_schema_error_is_not_hidden_by_extension_allowlist(self):
        error = {
            "error_type": "Schema",
            "part": "/ppt/presentation.xml",
            "xpath": "/p:presentation[1]/p:sldSz[1]",
            "description": "The attribute 'type' has invalid value 'wide'.",
        }
        self.assertFalse(is_allowed_authored_extension_warning(error))


if __name__ == "__main__":
    unittest.main()
