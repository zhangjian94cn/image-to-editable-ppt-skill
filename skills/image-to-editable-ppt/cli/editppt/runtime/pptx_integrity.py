#!/usr/bin/env python3
"""Open XML integrity gate for generated PowerPoint packages."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def validator_project() -> Path:
    return Path(__file__).resolve().parents[3] / "tools/openxml-validator/openxml-validator.csproj"


def validator_dll() -> Path:
    override = os.environ.get("EDITPPT_OPENXML_VALIDATOR")
    if override:
        return Path(override).expanduser().resolve()
    return validator_project().parent / "bin/Release/net6.0/openxml-validator.dll"


def ensure_validator() -> tuple[Path | None, str]:
    dll = validator_dll()
    if dll.is_file():
        return dll, ""
    return None, (
        "Open XML validator is not installed; run the locked deployment build before jobs "
        f"(see {validator_project().parent / 'README.md'})"
    )


def is_allowed_authored_extension_warning(error: dict[str, Any]) -> bool:
    """Return whether an SDK error is a known Office chart-extension warning.

    Authored Office files can contain extension elements that the pinned SDK
    does not know yet.  Those are not package corruption.  Keep this allowlist
    deliberately narrow: semantic errors, core PresentationML errors and
    extension errors outside chart parts remain blocking so they never reach
    PowerPoint's repair dialog.
    """

    error_type = str(error.get("error_type") or "").casefold()
    part = str(error.get("part") or "")
    description = str(error.get("description") or "").casefold()
    if error_type != "schema" or not part.startswith("/ppt/charts/"):
        return False
    if "unexpected child element" not in description:
        return False
    return (
        "schemas.microsoft.com/office/drawing/" in description
        or "schemas.openxmlformats.org/drawingml/2006/chart:ext" in description
    )


def classify_authored_validator_errors(
    errors: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Split authored-file SDK findings into blocking errors and allowlisted warnings."""

    allowed_warnings = [error for error in errors if is_allowed_authored_extension_warning(error)]
    blocking_errors = [error for error in errors if not is_allowed_authored_extension_warning(error)]
    return blocking_errors, allowed_warnings


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _orphan_comment_author_duplicate_count(package: zipfile.ZipFile) -> int:
    root = ET.fromstring(package.read("ppt/commentAuthors.xml"))
    counts: dict[str, int] = {}
    for author in root:
        if _local_name(author.tag) != "cmAuthor":
            continue
        author_id = author.get("id")
        if author_id is not None:
            counts[author_id] = counts.get(author_id, 0) + 1
    return sum(count - 1 for count in counts.values() if count > 1)


def sanitize_orphaned_comment_authors(
    pptx_path: Path,
    output_path: Path,
    *,
    integrity_report: dict[str, Any],
) -> dict[str, Any]:
    """Remove corrupt *orphaned* legacy comment-author metadata from a copy.

    This deliberately does not repair comments.  It is applicable only when
    the package has no comment parts or comment relationships and every
    blocking SDK finding points to duplicate IDs in commentAuthors.xml.  The
    caller must validate the derived package again and still require a clean
    PowerPoint render.
    """

    pptx_path = Path(pptx_path).resolve()
    output_path = Path(output_path).resolve()
    blocking = list(
        (integrity_report.get("validator_classification") or {}).get("blocking_errors") or []
    )
    if not blocking:
        return {"status": "not_applicable", "reason": "no blocking validator errors"}
    if any(
        str(error.get("error_type") or "").casefold() != "semantic"
        or str(error.get("part") or "") != "/ppt/commentAuthors.xml"
        for error in blocking
    ):
        return {
            "status": "not_applicable",
            "reason": "blocking errors are not confined to legacy commentAuthors.xml",
        }

    try:
        with zipfile.ZipFile(pptx_path) as package:
            names = set(package.namelist())
            if "ppt/commentAuthors.xml" not in names:
                return {"status": "not_applicable", "reason": "no commentAuthors part"}
            comment_parts = sorted(
                name for name in names
                if name.startswith("ppt/comments/") and name.endswith(".xml")
            )
            comment_relationships: list[tuple[str, str]] = []
            relationship_roots: dict[str, ET.Element] = {}
            for name in sorted(item for item in names if item.endswith(".rels")):
                root = ET.fromstring(package.read(name))
                relationship_roots[name] = root
                for relationship in root:
                    rel_type = str(relationship.get("Type") or "")
                    if rel_type.endswith("/comments"):
                        comment_relationships.append((name, relationship.get("Id") or ""))
            if comment_parts or comment_relationships:
                return {
                    "status": "not_applicable",
                    "reason": "the package contains real comments; automatic metadata removal is unsafe",
                    "comment_parts": comment_parts,
                    "comment_relationships": comment_relationships,
                }

            duplicate_count = _orphan_comment_author_duplicate_count(package)
            if duplicate_count != len(blocking):
                return {
                    "status": "not_applicable",
                    "reason": "duplicate author count does not match blocking SDK evidence",
                    "duplicate_count": duplicate_count,
                    "blocking_error_count": len(blocking),
                }

            content_types = ET.fromstring(package.read("[Content_Types].xml"))
            removed_overrides = 0
            for child in list(content_types):
                if child.get("PartName") == "/ppt/commentAuthors.xml":
                    content_types.remove(child)
                    removed_overrides += 1

            modified_relationships: dict[str, bytes] = {}
            removed_relationships = 0
            for name, root in relationship_roots.items():
                changed = False
                for relationship in list(root):
                    rel_type = str(relationship.get("Type") or "")
                    if rel_type.endswith("/commentAuthors"):
                        root.remove(relationship)
                        removed_relationships += 1
                        changed = True
                if changed:
                    ET.register_namespace("", RELATIONSHIPS_NS)
                    modified_relationships[name] = ET.tostring(
                        root, encoding="utf-8", xml_declaration=True
                    )

            ET.register_namespace("", CONTENT_TYPES_NS)
            content_types_bytes = ET.tostring(
                content_types, encoding="utf-8", xml_declaration=True
            )
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(output_path, "w") as derived:
                for info in package.infolist():
                    if info.filename == "ppt/commentAuthors.xml":
                        continue
                    if info.filename == "[Content_Types].xml":
                        data = content_types_bytes
                    elif info.filename in modified_relationships:
                        data = modified_relationships[info.filename]
                    else:
                        data = package.read(info.filename)
                    derived.writestr(info, data)
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        return {"status": "failed", "reason": str(exc)}

    return {
        "status": "applied",
        "strategy": "remove-orphaned-legacy-comment-authors",
        "original_pptx_sha256": _sha256(pptx_path),
        "derived_pptx_sha256": _sha256(output_path),
        "removed_parts": ["/ppt/commentAuthors.xml"],
        "removed_relationship_count": removed_relationships,
        "removed_content_type_override_count": removed_overrides,
        "duplicate_author_id_count": duplicate_count,
        "visual_content_changed": False,
    }


def validate_pptx_integrity(
    pptx_path: Path,
    *,
    report_path: Path | None = None,
    source_kind: str = "generated",
    build_id: str | None = None,
    raise_on_failure: bool = True,
    template_path: Path | None = None,
) -> dict[str, Any]:
    pptx_path = Path(pptx_path).resolve()
    build_id = build_id or str(uuid.uuid4())
    report: dict[str, Any] = {
        "schema_version": 1,
        "created_at": _now_iso(),
        "build_id": build_id,
        "source_kind": source_kind,
        "pptx_path": str(pptx_path),
        "pptx_sha256": _sha256(pptx_path) if pptx_path.is_file() else "",
        "template": {
            "path": str(Path(template_path).resolve()) if template_path else "",
            "sha256": _sha256(Path(template_path)) if template_path and Path(template_path).is_file() else "",
        },
        "generator": {
            "name": "editppt-build-pptx-from-manifest",
            "version": "powerpoint-template-v1",
        },
        "platform": platform.platform(),
        "status": "manual_pending",
        "schema_status": "unavailable",
        "powerpoint_preflight_status": "blocked",
        "validator_classification": {
            "blocking_error_count": 0,
            "allowed_warning_count": 0,
            "blocking_errors": [],
            "allowed_warnings": [],
        },
        "validator": {},
        "powerpoint": {"status": "pending", "repair_evidence": []},
    }
    dll, error = ensure_validator()
    if dll is not None:
        completed = subprocess.run(
            [shutil.which("dotnet") or "dotnet", str(dll), str(pptx_path)],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        try:
            validator = json.loads(completed.stdout)
        except json.JSONDecodeError:
            validator = {
                "passed": False,
                "open_error": completed.stderr.strip() or "validator emitted invalid JSON",
                "error_count": 0,
                "errors": [],
            }
        report["validator"] = validator
        if validator.get("passed") is True:
            report["schema_status"] = "passed"
            report["status"] = "schema_passed"
            report["powerpoint_preflight_status"] = "passed"
        elif source_kind == "authored" and not validator.get("open_error"):
            errors = validator.get("errors") or []
            blocking_errors, allowed_warnings = classify_authored_validator_errors(errors)
            report["validator_classification"] = {
                "blocking_error_count": len(blocking_errors),
                "allowed_warning_count": len(allowed_warnings),
                "blocking_errors": blocking_errors,
                "allowed_warnings": allowed_warnings,
            }
            if blocking_errors:
                report["schema_status"] = "failed"
                report["status"] = "manual_pending"
                report["powerpoint_preflight_status"] = "blocked"
                error = (
                    "authored PPTX has non-allowlisted Open XML errors; "
                    "PowerPoint was not started"
                )
            else:
                report["schema_status"] = "warnings"
                report["status"] = "powerpoint_validation_required"
                report["powerpoint_preflight_status"] = "passed_with_warnings"
        else:
            report["schema_status"] = "failed"
            report["status"] = "failed"
            report["powerpoint_preflight_status"] = "blocked"
            error = "generated PPTX failed Open XML validation"
    else:
        report["validator"] = {"passed": False, "error_count": 0, "errors": [], "open_error": error}
    if report_path is not None:
        report_path = Path(report_path)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if raise_on_failure and report["schema_status"] != "passed":
        raise RuntimeError(error or "PPTX integrity could not be proven")
    return report


def prepare_authored_reference(
    pptx_path: Path,
    output_path: Path,
    *,
    report_path: Path | None = None,
    build_id: str | None = None,
) -> dict[str, Any]:
    """Create the exact authored reference that PowerPoint may receive.

    Clean inputs are copied byte-for-byte.  The only automatic transformation
    is the narrow, evidence-backed removal implemented by
    :func:`sanitize_orphaned_comment_authors`.  The derived package is validated
    again before it is returned.  This keeps sanitisation policy in one owner
    and gives callers a stable JSON/exit-code contract.
    """

    source = Path(pptx_path).expanduser().resolve()
    target = Path(output_path).expanduser().resolve()
    build_id = build_id or str(uuid.uuid4())
    result: dict[str, Any] = {
        "schema_version": 1,
        "build_id": build_id,
        "status": "manual_pending",
        "original_pptx_sha256": _sha256(source) if source.is_file() else "",
        "reference_pptx_sha256": "",
        "sanitization": {"status": "not_applicable"},
        "original_validation": {},
        "reference_validation": {},
        "error": "",
    }
    if not source.is_file():
        result["error"] = f"input PPTX does not exist: {source}"
    else:
        original = validate_pptx_integrity(
            source,
            source_kind="authored",
            build_id=build_id,
            raise_on_failure=False,
        )
        result["original_validation"] = original
        preflight = str(original.get("powerpoint_preflight_status") or "")
        target.parent.mkdir(parents=True, exist_ok=True)
        if preflight in {"passed", "passed_with_warnings"}:
            shutil.copy2(source, target)
            result["reference_validation"] = original
            result["status"] = "passed"
        else:
            sanitization = sanitize_orphaned_comment_authors(
                source, target, integrity_report=original
            )
            result["sanitization"] = sanitization
            if sanitization.get("status") == "applied":
                derived = validate_pptx_integrity(
                    target,
                    source_kind="authored",
                    build_id=build_id,
                    raise_on_failure=False,
                )
                result["reference_validation"] = derived
                derived_preflight = str(derived.get("powerpoint_preflight_status") or "")
                if derived_preflight in {"passed", "passed_with_warnings"}:
                    result["status"] = "passed"
                else:
                    result["error"] = "sanitized reference still fails Open XML preflight"
                    target.unlink(missing_ok=True)
            else:
                result["error"] = str(
                    sanitization.get("reason")
                    or "authored PPTX has blocking Open XML errors"
                )
        if target.is_file():
            result["reference_pptx_sha256"] = _sha256(target)

    if report_path is not None:
        destination = Path(report_path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate or prepare PPTX packages")
    subparsers = parser.add_subparsers(dest="command")

    validate = subparsers.add_parser("validate", help="run the pinned Open XML validator")
    validate.add_argument("pptx")
    validate.add_argument("--report")
    validate.add_argument("--source-kind", choices=("generated", "authored"), default="generated")

    prepare = subparsers.add_parser(
        "prepare", help="create a PowerPoint-safe authored reference copy"
    )
    prepare.add_argument("pptx")
    prepare.add_argument("--output", required=True)
    prepare.add_argument("--report")
    prepare.add_argument("--build-id")

    argv = sys.argv[1:]
    # Backward compatibility for the existing font-source endpoint and local
    # operators that used ``pptx_integrity.py FILE --source-kind authored``.
    if argv and argv[0] not in {"validate", "prepare", "-h", "--help"}:
        argv.insert(0, "validate")
    args = parser.parse_args(argv)
    if args.command == "prepare":
        report = prepare_authored_reference(
            Path(args.pptx),
            Path(args.output),
            report_path=Path(args.report) if args.report else None,
            build_id=args.build_id,
        )
        print(json.dumps(report, ensure_ascii=False, indent=2))
        raise SystemExit(0 if report["status"] == "passed" else 2)
    if args.command == "validate":
        report = validate_pptx_integrity(
            Path(args.pptx),
            report_path=Path(args.report) if args.report else None,
            source_kind=args.source_kind,
            raise_on_failure=False,
        )
        print(json.dumps(report, ensure_ascii=False, indent=2))
        success = (
            report.get("powerpoint_preflight_status")
            in {"passed", "passed_with_warnings"}
        )
        raise SystemExit(0 if success else 2)
    parser.error("choose 'validate' or 'prepare'")


if __name__ == "__main__":
    main()
