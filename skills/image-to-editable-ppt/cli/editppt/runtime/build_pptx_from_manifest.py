#!/usr/bin/env python3
import argparse
import hashlib
import html
import importlib.util
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from copy import deepcopy
from pathlib import Path

from pptx_integrity import validate_pptx_integrity


EMU_PER_INCH = 914400
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
ASPECT_16_9 = 16 / 9
ASPECT_16_10 = 16 / 10
ASPECT_4_3 = 4 / 3
ASPECT_TOLERANCE = 0.03
DEFAULT_TEXT_FIT_SAFETY = 0.9
DEFAULT_TEXT_LINE_HEIGHT = 1.22
DEFAULT_MIN_FONT_SIZE = 4.0
TEXT_ALIGNMENTS = {
    "left": ("l", "left"),
    "center": ("ctr", "center"),
    "right": ("r", "right"),
    "l": ("l", "left"),
    "ctr": ("ctr", "center"),
    "r": ("r", "right"),
}
TEXT_VERTICAL_ALIGNMENTS = {
    "top": ("t", "top"),
    "middle": ("ctr", "middle"),
    "center": ("ctr", "middle"),
    "bottom": ("b", "bottom"),
    "t": ("t", "top"),
    "ctr": ("ctr", "middle"),
    "b": ("b", "bottom"),
}
AUTOFIT_VALUES = {
    "none": "<a:noAutofit/>",
    "shrink_text": "<a:normAutofit/>",
    "resize_shape": "<a:spAutoFit/>",
    # Backward compatibility for manifests produced before schema v2.
    "shape": "<a:spAutoFit/>",
}
WRAP_VALUES = {"none", "square"}
DASH_VALUES = {
    "solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot",
    "lgDashDotDot", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot",
}
PRESET_GEOMETRIES = {
    "rect", "roundRect", "ellipse", "line", "triangle", "rtTriangle",
    "parallelogram", "trapezoid", "diamond", "pentagon", "hexagon",
    "heptagon", "octagon", "decagon", "dodecagon", "pie", "chord",
    "teardrop", "frame", "halfFrame", "corner", "diagStripe", "plus",
    "plaque", "can", "cube", "bevel", "donut", "noSmoking", "blockArc",
    "foldedCorner", "smileyFace", "heart", "lightningBolt", "sun", "moon",
    "cloud", "arc", "bracketPair", "bracePair", "leftBracket", "rightBracket",
    "leftBrace", "rightBrace", "rightArrow", "leftArrow", "upArrow",
    "downArrow", "leftRightArrow", "upDownArrow", "quadArrow", "leftRightUpArrow",
    "bentArrow", "uturnArrow", "leftUpArrow", "bentUpArrow", "curvedRightArrow",
    "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow", "stripedRightArrow",
    "notchedRightArrow", "homePlate", "chevron", "rightArrowCallout",
    "downArrowCallout", "leftArrowCallout", "upArrowCallout", "leftRightArrowCallout",
    "quadArrowCallout", "circularArrow", "flowChartProcess", "flowChartDecision",
    "flowChartInputOutput", "flowChartPredefinedProcess", "flowChartInternalStorage",
    "flowChartDocument", "flowChartMultidocument", "flowChartTerminator",
    "flowChartPreparation", "flowChartManualInput", "flowChartManualOperation",
    "flowChartConnector", "flowChartOffpageConnector", "flowChartPunchedCard",
    "flowChartPunchedTape", "flowChartSummingJunction", "flowChartOr",
    "flowChartCollate", "flowChartSort", "flowChartExtract", "flowChartMerge",
    "flowChartOnlineStorage", "flowChartDelay", "flowChartMagneticTape",
    "flowChartMagneticDisk", "flowChartMagneticDrum", "flowChartDisplay",
}

ET.register_namespace("p", P_NS)
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("r", R_NS)
ET.register_namespace("", REL_NS)


def powerpoint_base_template():
    """Return the Microsoft PowerPoint-authored blank package used as the base.

    python-pptx ships a zero-slide template whose package properties identify
    Microsoft Macintosh PowerPoint as its creator.  Reusing that package keeps
    the real master, blank layout, theme, view/presentation properties and
    relationship topology instead of attempting another incomplete skeleton.
    Deployments may pin an independently accepted template with the environment
    override; no generated package is allowed to fall back to handwritten OOXML.
    """

    override = os.environ.get("EDITPPT_POWERPOINT_TEMPLATE", "")
    if override:
        candidate = Path(override).expanduser().resolve()
        if candidate.is_file():
            return candidate
        raise ValueError(f"PowerPoint base template does not exist: {candidate}")
    spec = importlib.util.find_spec("pptx")
    if spec and spec.submodule_search_locations:
        candidate = Path(next(iter(spec.submodule_search_locations))) / "templates/default.pptx"
        if candidate.is_file():
            return candidate.resolve()
    raise ValueError(
        "PowerPoint-authored base template is unavailable; install python-pptx "
        "or set EDITPPT_POWERPOINT_TEMPLATE"
    )


def emu(value):
    return int(round(float(value) * EMU_PER_INCH))


def hex_color(value, default="000000"):
    if not value:
        return default
    return str(value).strip().lstrip("#").upper()


def content_type_for(path):
    suffix = Path(path).suffix.lower()
    if suffix == ".png":
        return "image/png"
    if suffix in (".jpg", ".jpeg"):
        return "image/jpeg"
    if suffix == ".gif":
        return "image/gif"
    if suffix == ".svg":
        return "image/svg+xml"
    raise ValueError(f"Unsupported image type: {path}")


def image_ext(path):
    suffix = Path(path).suffix.lower()
    return ".jpg" if suffix == ".jpeg" else suffix


def xml_text(value):
    return html.escape(str(value), quote=True)


def text_alignment(value="left"):
    key = str(value or "left").strip().lower()
    if key not in TEXT_ALIGNMENTS:
        raise ValueError(f"Unsupported text align value: {value}")
    return TEXT_ALIGNMENTS[key]


def text_vertical_alignment(value="top"):
    key = str(value or "top").strip().lower()
    if key not in TEXT_VERTICAL_ALIGNMENTS:
        raise ValueError(f"Unsupported text valign value: {value}")
    return TEXT_VERTICAL_ALIGNMENTS[key]


def source_size_px(manifest):
    source = manifest.get("source", {})
    width = source.get("width_px")
    height = source.get("height_px")
    if width and height:
        return float(width), float(height)
    return None


def slide_size(manifest):
    slide = manifest.get("slide", {})
    return float(slide.get("width", 13.333)), float(slide.get("height", 7.5))


def fit_content_box(source_width, source_height, slide_width, slide_height):
    source_aspect = source_width / source_height
    slide_aspect = slide_width / slide_height
    if source_aspect >= slide_aspect:
        width = slide_width
        height = width / source_aspect
        left = 0
        top = (slide_height - height) / 2
    else:
        height = slide_height
        width = height * source_aspect
        left = (slide_width - width) / 2
        top = 0
    return {"left": left, "top": top, "width": width, "height": height}


def content_box_for_manifest(manifest):
    content_box = manifest.get("content_box")
    if content_box:
        return {
            "left": float(content_box.get("left", 0)),
            "top": float(content_box.get("top", 0)),
            "width": float(content_box.get("width", 1)),
            "height": float(content_box.get("height", 1)),
        }
    source_size = source_size_px(manifest)
    slide_width, slide_height = slide_size(manifest)
    if not source_size:
        return {"left": 0, "top": 0, "width": slide_width, "height": slide_height}
    source_width, source_height = source_size
    return fit_content_box(source_width, source_height, slide_width, slide_height)


def px_to_inches(manifest, x, y, width, height):
    source_size = source_size_px(manifest)
    if not source_size:
        raise ValueError("Manifest uses pixel coordinates but lacks source.width_px/source.height_px")
    source_width, source_height = source_size
    content_box = content_box_for_manifest(manifest)
    return {
        "left": content_box["left"] + float(x) / source_width * content_box["width"],
        "top": content_box["top"] + float(y) / source_height * content_box["height"],
        "width": float(width) / source_width * content_box["width"],
        "height": float(height) / source_height * content_box["height"],
    }


def normalize_position_item(manifest, item):
    item = dict(item)
    if "polygon_px" in item:
        points = [(float(point[0]), float(point[1])) for point in item["polygon_px"]]
        if points and "box_px" not in item:
            xs = [point[0] for point in points]
            ys = [point[1] for point in points]
            item["box_px"] = [min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)]
        item["polygon"] = [
            [px_to_inches(manifest, point[0], point[1], 0, 0)["left"], px_to_inches(manifest, point[0], point[1], 0, 0)["top"]]
            for point in points
        ]
    if "box_px" in item:
        x, y, width, height = item["box_px"]
        item.update(px_to_inches(manifest, x, y, width, height))
    if "points_px" in item:
        x1, y1, x2, y2 = item["points_px"]
        left = min(float(x1), float(x2))
        top = min(float(y1), float(y2))
        width = abs(float(x2) - float(x1))
        height = abs(float(y2) - float(y1))
        item.update(px_to_inches(manifest, left, top, width, height))
        start = px_to_inches(manifest, x1, y1, 0, 0)
        end = px_to_inches(manifest, x2, y2, 0, 0)
        item["points"] = [start["left"], start["top"], end["left"], end["top"]]
        if float(x2) < float(x1):
            item["flip_h"] = True
        if float(y2) < float(y1):
            item["flip_v"] = True
    if item.get("source_corner_radius_px") is not None and "radius" not in item:
        radius = float(item.get("source_corner_radius_px") or 0)
        item["radius"] = px_to_inches(manifest, 0, 0, radius, radius)["width"]
    return item


def iter_text_lines(item):
    if item.get("paragraphs"):
        lines = []
        for paragraph in item["paragraphs"]:
            if isinstance(paragraph, str):
                lines.append(paragraph)
            else:
                runs = paragraph.get("runs")
                if runs:
                    lines.append("".join(str(run.get("text", "")) for run in runs))
                else:
                    lines.append(str(paragraph.get("text", "")))
        return lines or [""]
    if item.get("runs"):
        return ["".join(str(run.get("text", "")) for run in item["runs"])]
    return str(item.get("text", "")).splitlines() or [""]


def text_width_units(text):
    units = 0.0
    for char in str(text):
        codepoint = ord(char)
        if char.isspace():
            units += 0.32
        elif codepoint <= 0x7F:
            units += 0.55
        elif 0xFF00 <= codepoint <= 0xFFEF:
            units += 1.0
        elif 0x4E00 <= codepoint <= 0x9FFF:
            units += 1.0
        else:
            units += 0.85
    return max(units, 1.0)


def longest_unbreakable_units(text):
    tokens = [token for token in re.split(r"\s+", str(text)) if token]
    if not tokens:
        return text_width_units(text)
    return max(text_width_units(token) for token in tokens)


def is_measured_text(item):
    """True when the author marked this box as sized from `page hints` measurement."""
    return str(item.get("font_size_source", "")).strip().lower() in {"measured", "hints"}


def fitted_font_size(item, manifest):
    if item.get("fit_text") is False or manifest.get("fit_text") is False:
        return None
    if "width" not in item or "height" not in item:
        return None
    lines = iter_text_lines(item)
    requested = float(item.get("font_size", 18))
    width_pt = max(1.0, float(item.get("width", 1)) * 72)
    height_pt = max(1.0, float(item.get("height", 0.4)) * 72)
    if is_measured_text(item):
        # Box and font size were both measured from source ink; the safety
        # discount exists to absorb estimation error in hand-written boxes
        # and would systematically shrink correct text here. Clamp only at
        # the geometric limit.
        safety = 1.0
    else:
        safety = float(item.get("text_fit_safety", manifest.get("text_fit_safety", DEFAULT_TEXT_FIT_SAFETY)))
    line_height = float(item.get("line_height", manifest.get("text_line_height", DEFAULT_TEXT_LINE_HEIGHT)))
    wrap_enabled = item.get("wrap") not in (None, "", "none")
    if wrap_enabled:
        line_count = sum(max(1, math.ceil(text_width_units(line) * requested / width_pt)) for line in lines)
        width_limit = width_pt / max(longest_unbreakable_units(line) for line in lines)
    else:
        line_count = max(1, len(lines))
        width_limit = width_pt / max(text_width_units(line) for line in lines)
    height_limit = height_pt / (line_count * max(line_height, 1.0))
    max_font_size = min(width_limit, height_limit) * safety
    explicit_max = item.get("max_font_size")
    if explicit_max not in (None, ""):
        max_font_size = min(max_font_size, float(explicit_max))
    min_font_size = float(item.get("min_font_size", manifest.get("min_font_size", DEFAULT_MIN_FONT_SIZE)))
    return max(min_font_size, max_font_size)


def scale_run_font_sizes(item, ratio):
    def scale_run(run):
        if run.get("font_size") not in (None, ""):
            run["font_size"] = round(float(run["font_size"]) * ratio, 1)

    for run in item.get("runs", []):
        scale_run(run)
    for paragraph in item.get("paragraphs", []):
        if isinstance(paragraph, dict):
            for run in paragraph.get("runs", []):
                scale_run(run)


def fit_text_item(item, manifest):
    fitted = fitted_font_size(item, manifest)
    if fitted is None:
        return item
    requested = float(item.get("font_size", fitted))
    effective = min(requested, fitted)
    if effective < requested:
        item["_requested_font_size"] = requested
        item["font_size"] = round(effective, 1)
        scale_run_font_sizes(item, effective / requested)
    elif "font_size" not in item:
        item["font_size"] = round(effective, 1)
    return item


def _require_finite_number(value, field, *, minimum=None, maximum=None):
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field} must be numeric") from exc
    if not math.isfinite(number):
        raise ValueError(f"{field} must be finite")
    if minimum is not None and number < minimum:
        raise ValueError(f"{field} must be >= {minimum}")
    if maximum is not None and number > maximum:
        raise ValueError(f"{field} must be <= {maximum}")
    return number


def _validate_hex_color(value, field, *, allow_none=False):
    if value in (None, ""):
        return
    if allow_none and str(value).lower() == "none":
        return
    if not re.fullmatch(r"#?[0-9A-Fa-f]{6}", str(value).strip()):
        raise ValueError(f"{field} must be a six-digit RGB color")


def validate_manifest_values(manifest):
    """Reject values that would serialize into invalid DrawingML."""

    slide = manifest.get("slide", {})
    _require_finite_number(slide.get("width", 13.333), "slide.width", minimum=0.01)
    _require_finite_number(slide.get("height", 7.5), "slide.height", minimum=0.01)
    _validate_hex_color(slide.get("background"), "slide.background")
    for index, item in enumerate(manifest.get("text_boxes", [])):
        prefix = f"text_boxes[{index}]"
        text_alignment(item.get("align", "left"))
        text_vertical_alignment(item.get("valign", "top"))
        wrap = str(item.get("wrap") or "none")
        if wrap not in WRAP_VALUES:
            raise ValueError(f"{prefix}.wrap must be one of {sorted(WRAP_VALUES)}")
        autofit = str(item.get("autofit") or "none")
        if autofit not in AUTOFIT_VALUES:
            raise ValueError(f"{prefix}.autofit must be one of {sorted(AUTOFIT_VALUES)}")
        _require_finite_number(item.get("font_size", 18), f"{prefix}.font_size", minimum=0.01)
        _validate_hex_color(item.get("color", "#111111"), f"{prefix}.color")
        for field in ("margin_left", "margin_right", "margin_top", "margin_bottom",
                      "paragraph_spacing_before", "paragraph_spacing_after"):
            _require_finite_number(item.get(field, 0), f"{prefix}.{field}", minimum=0)
        _require_finite_number(item.get("line_spacing", 1.0), f"{prefix}.line_spacing", minimum=0.01)
        _require_finite_number(item.get("character_spacing", 0), f"{prefix}.character_spacing",
                               minimum=-4000, maximum=4000)
    for index, item in enumerate(manifest.get("shapes", [])):
        prefix = f"shapes[{index}]"
        preset = item.get("preset")
        if preset and preset not in PRESET_GEOMETRIES:
            raise ValueError(f"{prefix}.preset is not an allowed DrawingML geometry: {preset}")
        dash = item.get("dash")
        if dash and dash not in DASH_VALUES:
            raise ValueError(f"{prefix}.dash is not an allowed DrawingML dash: {dash}")
        _validate_hex_color(item.get("fill"), f"{prefix}.fill", allow_none=True)
        _validate_hex_color(item.get("stroke"), f"{prefix}.stroke", allow_none=True)
        _require_finite_number(item.get("stroke_width", 1), f"{prefix}.stroke_width", minimum=0)
    for table_index, table in enumerate(manifest.get("tables", [])):
        prefix = f"tables[{table_index}]"
        rows = table.get("cells") or []
        if not rows or not all(isinstance(row, list) and row for row in rows):
            raise ValueError(f"{prefix}.cells must be a non-empty rectangular matrix")
        columns = len(rows[0])
        if any(len(row) != columns for row in rows):
            raise ValueError(f"{prefix}.cells must be rectangular")
        if table.get("columns_px") and len(table["columns_px"]) != columns:
            raise ValueError(f"{prefix}.columns_px must match the cell column count")
        if table.get("rows_px") and len(table["rows_px"]) != len(rows):
            raise ValueError(f"{prefix}.rows_px must match the cell row count")
        text_alignment(table.get("align", "left"))
        text_vertical_alignment(table.get("valign", "middle"))
        _validate_hex_color(table.get("color", "#111111"), f"{prefix}.color")
        _validate_hex_color(table.get("cell_fill"), f"{prefix}.cell_fill", allow_none=True)
        _validate_hex_color(table.get("stroke", "#BFBFBF"), f"{prefix}.stroke", allow_none=True)
        _require_finite_number(table.get("font_size", 12), f"{prefix}.font_size", minimum=0.01)
        for row_index, row in enumerate(rows):
            for cell_index, raw in enumerate(row):
                if not isinstance(raw, dict):
                    continue
                cell_prefix = f"{prefix}.cells[{row_index}][{cell_index}]"
                text_alignment(raw.get("align", table.get("align", "left")))
                text_vertical_alignment(raw.get("valign", table.get("valign", "middle")))
                _validate_hex_color(raw.get("color", table.get("color", "#111111")), f"{cell_prefix}.color")
                _validate_hex_color(raw.get("fill"), f"{cell_prefix}.fill", allow_none=True)
                _validate_hex_color(raw.get("stroke"), f"{cell_prefix}.stroke", allow_none=True)
                for field in ("col_span", "row_span"):
                    _require_finite_number(raw.get(field, 1), f"{cell_prefix}.{field}", minimum=1)


def _validate_strict_font(item, label, required):
    spec = item.get("font_spec")
    if not isinstance(spec, dict) or required - set(spec):
        raise ValueError(f"{label} requires an exact manifest-v2 font_spec")
    if not item.get("font"):
        raise ValueError(f"{label} requires a builder-facing font name")
    font_file = Path(str(spec["font_file"])).expanduser()
    if not font_file.is_absolute() or not font_file.is_file():
        raise ValueError(f"{label} font_file must be an installed absolute path")
    digest = hashlib.sha256(font_file.read_bytes()).hexdigest()
    if digest != spec["file_sha256"]:
        raise ValueError(f"{label} font file SHA-256 mismatch")
    item["_strict_font"] = True


def normalize_manifest(manifest):
    """Return a manifest copy with pixel authoring fields resolved to inches."""
    normalized = deepcopy(manifest)
    validate_manifest_values(normalized)
    strict_fonts = normalized.get("schema_version") == 2
    if strict_fonts:
        required = {
            "east_asian_font", "latin_font", "file_sha256", "internal_name",
            "font_source", "font_file", "face_index",
        }
        for index, item in enumerate(normalized.get("text_boxes", [])):
            _validate_strict_font(item, f"text_boxes[{index}]", required)
        for table_index, table in enumerate(normalized.get("tables", [])):
            _validate_strict_font(table, f"tables[{table_index}]", required)
            for row_index, row in enumerate(table.get("cells") or []):
                for cell_index, raw in enumerate(row):
                    if not isinstance(raw, dict):
                        raise ValueError(
                            f"tables[{table_index}].cells[{row_index}][{cell_index}] "
                            "must be an object in manifest v2"
                        )
                    _validate_strict_font(
                        raw,
                        f"tables[{table_index}].cells[{row_index}][{cell_index}]",
                        required,
                    )
    normalized["text_boxes"] = [
        fit_text_item(normalize_position_item(normalized, item), normalized) for item in normalized.get("text_boxes", [])
    ]
    for key in ("images", "shapes"):
        normalized[key] = [normalize_position_item(normalized, item) for item in normalized.get(key, [])]
    normalized["tables"] = [
        normalize_table_item(normalized, item) for item in normalized.get("tables", [])
    ]
    return normalized


def normalize_table_item(manifest, item):
    item = normalize_position_item(manifest, item)
    if item.get("columns_px"):
        item["columns"] = [
            px_to_inches(manifest, 0, 0, float(value), 0)["width"]
            for value in item["columns_px"]
        ]
    if item.get("rows_px"):
        item["rows"] = [
            px_to_inches(manifest, 0, 0, 0, float(value))["height"]
            for value in item["rows_px"]
        ]
    return item


def preview_color(value):
    if not value or value == "none":
        return value
    value = str(value).strip()
    if value.startswith("#"):
        return value
    if len(value) == 6 and all(ch in "0123456789abcdefABCDEF" for ch in value):
        return f"#{value}"
    return value


def shape_fill(fill):
    if not fill or fill == "none":
        return '<a:noFill/>'
    return f'<a:solidFill><a:srgbClr val="{hex_color(fill)}"/></a:solidFill>'


def shape_line_xml(stroke, width, dash=None):
    if not stroke or stroke == "none":
        return '<a:ln><a:noFill/></a:ln>'
    dash_xml = f'<a:prstDash val="{xml_text(dash)}"/>' if dash else ""
    return (
        f'<a:ln w="{int(float(width or 1) * 12700)}">'
        f'<a:solidFill><a:srgbClr val="{hex_color(stroke)}"/></a:solidFill>'
        f"{dash_xml}"
        "</a:ln>"
    )


def slide_background_xml(slide):
    background = slide.get("background")
    if not background:
        return ""
    return (
        "<p:bg><p:bgPr>"
        f'<a:solidFill><a:srgbClr val="{hex_color(background, "FFFFFF")}"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
    )


def text_box_xml(idx, item):
    left = emu(item.get("left", 0))
    top = emu(item.get("top", 0))
    width = emu(item.get("width", 1))
    height = emu(item.get("height", 0.4))
    rotation = item.get("rotation")
    rotation_attr = f' rot="{int(float(rotation) * 60000)}"' if rotation not in (None, "") else ""
    font_size = int(float(item.get("font_size", 18)) * 100)
    if item.get("_strict_font") and not item.get("font"):
        raise ValueError(f"strict text box {item.get('id', idx)!r} has no font")
    align = text_alignment(item.get("align", "left"))[0]
    anchor = text_vertical_alignment(item.get("valign", "top"))[0]
    wrap = str(item.get("wrap") or "none")
    autofit = str(item.get("autofit") or "none")
    autofit_xml = AUTOFIT_VALUES[autofit]
    margins = {
        "lIns": int(round(float(item.get("margin_left", 0)) * 12700)),
        "tIns": int(round(float(item.get("margin_top", 0)) * 12700)),
        "rIns": int(round(float(item.get("margin_right", 0)) * 12700)),
        "bIns": int(round(float(item.get("margin_bottom", 0)) * 12700)),
    }
    paragraphs = item.get("paragraphs")
    runs = item.get("runs")

    def run_xml(run):
        run_font_size = int(float(run.get("font_size", item.get("font_size", 18))) * 100)
        spec = run.get("font_spec") or item.get("font_spec") or {}
        fallback_font = run.get("font") or item.get("font") or (
            None if item.get("_strict_font") else "PingFang SC"
        )
        latin_font_value = spec.get("latin_font") or fallback_font
        east_asian_font_value = spec.get("east_asian_font") or fallback_font
        complex_font_value = spec.get("complex_script_font") or latin_font_value
        if not latin_font_value or not east_asian_font_value:
            raise ValueError(f"strict text run in {item.get('id', idx)!r} has no font")
        latin_font = xml_text(latin_font_value)
        east_asian_font = xml_text(east_asian_font_value)
        complex_font = xml_text(complex_font_value)
        run_color = hex_color(run.get("color", item.get("color", "#111111")))
        run_bold = ' b="1"' if run.get("bold", item.get("bold")) else ""
        run_italic = ' i="1"' if run.get("italic", item.get("italic")) else ""
        spacing = run.get("character_spacing", item.get("character_spacing", 0))
        run_spacing = f' spc="{int(round(float(spacing) * 100))}"' if float(spacing or 0) else ""
        run_baseline = run.get("baseline")
        run_baseline_attr = f' baseline="{int(float(run_baseline))}"' if run_baseline not in (None, "") else ""
        run_text = xml_text(run.get("text", ""))
        return (
            f'<a:r><a:rPr lang="zh-CN" sz="{run_font_size}"{run_bold}{run_italic}{run_spacing}{run_baseline_attr}>'
            f'<a:solidFill><a:srgbClr val="{run_color}"/></a:solidFill>'
            f'<a:latin typeface="{latin_font}"/><a:ea typeface="{east_asian_font}"/><a:cs typeface="{complex_font}"/>'
            f'</a:rPr><a:t>{run_text}</a:t></a:r>'
        )

    def paragraph_xml(paragraph):
        if isinstance(paragraph, str):
            paragraph_runs = [{"text": paragraph}]
            paragraph_values = {}
        else:
            paragraph_runs = paragraph.get("runs", [{"text": paragraph.get("text", "")}])
            paragraph_values = paragraph
        line_spacing = float(paragraph_values.get("line_spacing", item.get("line_spacing", 1.0)))
        spacing_before = float(paragraph_values.get(
            "paragraph_spacing_before", item.get("paragraph_spacing_before", 0)
        ))
        spacing_after = float(paragraph_values.get(
            "paragraph_spacing_after", item.get("paragraph_spacing_after", 0)
        ))
        spacing_xml = (
            f'<a:lnSpc><a:spcPct val="{int(round(line_spacing * 100000))}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="{int(round(spacing_before * 100))}"/></a:spcBef>'
            f'<a:spcAft><a:spcPts val="{int(round(spacing_after * 100))}"/></a:spcAft>'
        )
        return (
            f'<a:p><a:pPr algn="{align}">{spacing_xml}</a:pPr>'
            + "".join(run_xml(run) for run in paragraph_runs)
            + f'<a:endParaRPr lang="zh-CN" sz="{font_size}"/></a:p>'
        )

    if paragraphs:
        text_body = "".join(paragraph_xml(paragraph) for paragraph in paragraphs)
    elif runs:
        text_body = paragraph_xml({"runs": runs})
    else:
        text_body = "".join(paragraph_xml(part) for part in str(item.get("text", "")).split("\n"))
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{idx}" name="TextBox {idx}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm{rotation_attr}><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
        <p:txBody>
          <a:bodyPr wrap="{xml_text(wrap)}" anchor="{anchor}" lIns="{margins['lIns']}" tIns="{margins['tIns']}" rIns="{margins['rIns']}" bIns="{margins['bIns']}">{autofit_xml}</a:bodyPr><a:lstStyle/>
          {text_body}
        </p:txBody>
      </p:sp>"""


def image_xml(idx, rel_id, item):
    left = emu(item.get("left", 0))
    top = emu(item.get("top", 0))
    width = emu(item.get("width", 1))
    height = emu(item.get("height", 1))
    name = xml_text(item.get("alt") or Path(item.get("path", "")).stem or f"Image {idx}")
    return f"""
      <p:pic>
        <p:nvPicPr><p:cNvPr id="{idx}" name="{name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      </p:pic>"""


def table_cell_xml(cell, table):
    """Serialize one editable native table cell without a font fallback."""

    if cell.get("h_merge") or cell.get("v_merge"):
        merge = ' hMerge="1"' if cell.get("h_merge") else ' vMerge="1"'
        return (
            f'<a:tc{merge}><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>'
            '<a:tcPr/></a:tc>'
        )
    span = ""
    if int(cell.get("col_span", 1) or 1) > 1:
        span += f' gridSpan="{int(cell["col_span"])}"'
    if int(cell.get("row_span", 1) or 1) > 1:
        span += f' rowSpan="{int(cell["row_span"])}"'

    size = int(float(cell.get("font_size", table.get("font_size", 12))) * 100)
    spec = cell.get("font_spec") or table.get("font_spec") or {}
    fallback = cell.get("font") or table.get("font")
    if not fallback and not (cell.get("_strict_font") or table.get("_strict_font")):
        fallback = "PingFang SC"
    latin = spec.get("latin_font") or fallback
    east_asian = spec.get("east_asian_font") or fallback
    complex_script = spec.get("complex_script_font") or latin
    if not latin or not east_asian:
        raise ValueError("strict table cell is missing exact Latin or East Asian font names")
    color = hex_color(cell.get("color", table.get("color", "#111111")))
    bold = ' b="1"' if cell.get("bold", table.get("bold")) else ""
    italic = ' i="1"' if cell.get("italic", table.get("italic")) else ""
    spacing = float(cell.get("character_spacing", table.get("character_spacing", 0)) or 0)
    spacing_attr = f' spc="{int(round(spacing * 100))}"' if spacing else ""
    align = text_alignment(cell.get("align", table.get("align", "left")))[0]
    anchor = text_vertical_alignment(cell.get("valign", table.get("valign", "middle")))[0]
    paragraphs = "".join(
        f'<a:p><a:pPr algn="{align}"/><a:r>'
        f'<a:rPr lang="zh-CN" sz="{size}"{bold}{italic}{spacing_attr}>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{xml_text(latin)}"/>'
        f'<a:ea typeface="{xml_text(east_asian)}"/>'
        f'<a:cs typeface="{xml_text(complex_script)}"/>'
        f'</a:rPr><a:t>{xml_text(line)}</a:t></a:r></a:p>'
        for line in str(cell.get("text", "")).split("\n")
    ) or "<a:p/>"
    fill = cell.get("fill", table.get("cell_fill"))
    fill_xml = shape_fill(fill) if fill else "<a:noFill/>"
    stroke = cell.get("stroke", table.get("stroke", "#BFBFBF"))
    stroke_width = float(cell.get("stroke_width", table.get("stroke_width", 0.75)) or 0.75)
    if not stroke or stroke == "none":
        line_xml = ""
    else:
        pen = f'<a:solidFill><a:srgbClr val="{hex_color(stroke)}"/></a:solidFill>'
        edge = f'w="{int(stroke_width * 12700)}" cap="flat" cmpd="sng"'
        line_xml = "".join(
            f'<a:ln{side} {edge}>{pen}</a:ln{side}>' for side in ("L", "R", "T", "B")
        )
    margins = {
        "marL": int(float(cell.get("margin_left", table.get("margin_left", 3.6))) * 12700),
        "marR": int(float(cell.get("margin_right", table.get("margin_right", 3.6))) * 12700),
        "marT": int(float(cell.get("margin_top", table.get("margin_top", 2.16))) * 12700),
        "marB": int(float(cell.get("margin_bottom", table.get("margin_bottom", 2.16))) * 12700),
    }
    return (
        f'<a:tc{span}><a:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</a:txBody>'
        f'<a:tcPr marL="{margins["marL"]}" marR="{margins["marR"]}" '
        f'marT="{margins["marT"]}" marB="{margins["marB"]}" anchor="{anchor}">'
        f'{line_xml}{fill_xml}</a:tcPr></a:tc>'
    )


def table_xml(idx, item):
    left, top = emu(item.get("left", 0)), emu(item.get("top", 0))
    width, height = emu(item.get("width", 1)), emu(item.get("height", 1))
    rows = item.get("cells") or []
    column_count = max((len(row) for row in rows), default=0)
    if not rows or not column_count:
        raise ValueError(f"table {item.get('id', idx)!r} has no cells")
    columns = item.get("columns") or [item.get("width", 1) / column_count] * column_count
    heights = item.get("rows") or [item.get("height", 1) / len(rows)] * len(rows)
    grid = "".join(f'<a:gridCol w="{emu(value)}"/>' for value in columns)
    body = "".join(
        f'<a:tr h="{emu(row_height)}">'
        + "".join(
            table_cell_xml(cell if isinstance(cell, dict) else {"text": cell}, item)
            for cell in row
        )
        + "</a:tr>"
        for row, row_height in zip(rows, heights)
    )
    first_row = ' firstRow="1"' if item.get("header", True) else ""
    return f"""
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="{idx}" name="Table {idx}"/>
          <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl><a:tblPr{first_row}/><a:tblGrid>{grid}</a:tblGrid>{body}</a:tbl>
        </a:graphicData></a:graphic>
      </p:graphicFrame>"""


def shape_xml(idx, item):
    kind = item.get("type", "rect")
    left = emu(item.get("left", 0))
    top = emu(item.get("top", 0))
    width = emu(item.get("width", 1))
    height = emu(item.get("height", 1))
    stroke_width = item.get("stroke_width", 1)
    flip_h = ' flipH="1"' if item.get("flip_h") else ""
    flip_v = ' flipV="1"' if item.get("flip_v") else ""
    fill = shape_fill(item.get("fill"))
    line = shape_line_xml(item.get("stroke", "#000000"), stroke_width, item.get("dash"))
    preset = item.get("preset")
    if item.get("polygon_px"):
        geometry = custom_polygon_geometry_xml(item)
    else:
        if not preset:
            preset = "line" if kind == "line" else "ellipse" if kind == "ellipse" else "roundRect" if kind == "roundRect" else "rect"
        geometry = preset_geometry_xml(preset, item)
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{idx}" name="{xml_text(kind.title())} {idx}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm{flip_h}{flip_v}><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>{geometry}{fill}{line}</p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>"""


def custom_polygon_geometry_xml(item):
    points = [(float(point[0]), float(point[1])) for point in item.get("polygon_px", [])]
    if len(points) < 3:
        return '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    box = item.get("box_px")
    if box and len(box) == 4:
        left, top, width, height = [float(value) for value in box]
    else:
        xs = [point[0] for point in points]
        ys = [point[1] for point in points]
        left, top = min(xs), min(ys)
        width, height = max(xs) - left, max(ys) - top
    width = max(width, 1.0)
    height = max(height, 1.0)

    def rel_coord(point):
        x, y = point
        return int(round((x - left) / width * 21600)), int(round((y - top) / height * 21600))

    first_x, first_y = rel_coord(points[0])
    segments = [f'<a:moveTo><a:pt x="{first_x}" y="{first_y}"/></a:moveTo>']
    for point in points[1:]:
        x, y = rel_coord(point)
        segments.append(f'<a:lnTo><a:pt x="{x}" y="{y}"/></a:lnTo>')
    segments.append("<a:close/>")
    return (
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        '<a:rect l="l" t="t" r="r" b="b"/>'
        '<a:pathLst><a:path w="21600" h="21600">'
        + "".join(segments)
        + "</a:path></a:pathLst></a:custGeom>"
    )


def round_rect_adjustment(item):
    box_px = item.get("box_px")
    if item.get("source_corner_radius_px") is not None and box_px and len(box_px) == 4:
        radius = float(item.get("source_corner_radius_px") or 0)
        min_dim = max(1.0, min(float(box_px[2]), float(box_px[3])))
    elif item.get("radius") is not None:
        radius = float(item.get("radius") or 0)
        min_dim = max(0.01, min(float(item.get("width", 1)), float(item.get("height", 1))))
    else:
        return None
    return max(0, min(50000, int(round(radius / min_dim * 100000))))


def preset_geometry_xml(preset, item):
    if preset != "roundRect":
        return f'<a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom>'
    adjustment = round_rect_adjustment(item)
    if adjustment is None:
        return '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
    return (
        '<a:prstGeom prst="roundRect"><a:avLst>'
        f'<a:gd name="adj" fmla="val {adjustment}"/>'
        '</a:avLst></a:prstGeom>'
    )


def slide_xml(manifest):
    slide = manifest.get("slide", {})
    next_id = 2
    parts = []
    layered = []
    for index, item in enumerate(manifest.get("shapes", [])):
        layered.append((float(item.get("z_index", 100)), index, "shape", item, None))
    for rel_index, item in enumerate(manifest.get("images", []), start=1):
        layered.append((float(item.get("z_index", 200)), rel_index, "image", item, f"rId{rel_index + 1}"))
    for index, item in enumerate(manifest.get("tables", [])):
        layered.append((float(item.get("z_index", 250)), index, "table", item, None))
    for index, item in enumerate(manifest.get("text_boxes", [])):
        layered.append((float(item.get("z_index", 300)), index, "text", item, None))

    for _z_index, _order, kind, item, rel_id in sorted(layered, key=lambda entry: (entry[0], entry[1])):
        if kind == "shape":
            parts.append(shape_xml(next_id, item))
        elif kind == "image":
            parts.append(image_xml(next_id, rel_id, item))
        elif kind == "table":
            parts.append(table_xml(next_id, item))
        else:
            parts.append(text_box_xml(next_id, item))
        next_id += 1
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    {slide_background_xml(slide)}
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {''.join(parts)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def rels_xml(manifest, media_start=1, notes_index=None, layout_target="slideLayout7.xml"):
    rels = [f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/{layout_target}"/>']
    for i, item in enumerate(manifest.get("images", []), start=1):
        target = f"../media/image{media_start + i - 1}{image_ext(item['path'])}"
        rels.append(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{target}"/>')
    if notes_index is not None:
        rels.append(
            f'<Relationship Id="rId{len(rels) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{notes_index}.xml"/>'
        )
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(rels) + "</Relationships>"


def notes_slide_xml(text):
    paras = "".join(
        f'<a:p><a:r><a:rPr lang="zh-CN" sz="1200"/><a:t>{xml_text(line)}</a:t></a:r><a:endParaRPr lang="zh-CN" sz="1200"/></a:p>'
        for line in str(text).splitlines()
    ) or '<a:p><a:endParaRPr lang="zh-CN" sz="1200"/></a:p>'
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="914400"/><a:ext cx="5486400" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>"""


def notes_rels_xml(slide_index):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{slide_index}.xml"/>
</Relationships>"""


def notes_master_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
</p:notesMaster>"""


def notes_master_rels_xml():
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""


def content_types_xml(manifests, notes_indices=None):
    notes_indices = notes_indices or []
    defaults = {
        "rels": "application/vnd.openxmlformats-package.relationships+xml",
        "xml": "application/xml",
    }
    for manifest in manifests:
        for item in manifest.get("images", []):
            ext = image_ext(item["path"]).lstrip(".")
            defaults[ext] = content_type_for(item["path"])
    default_xml = "".join(f'<Default Extension="{ext}" ContentType="{ctype}"/>' for ext, ctype in defaults.items())
    overrides = [
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for i in range(1, len(manifests) + 1):
        overrides.append(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
    if notes_indices:
        overrides.append('<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>')
        for i in notes_indices:
            overrides.append(f'<Override PartName="/ppt/notesSlides/notesSlide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{default_xml}{"".join(overrides)}</Types>'


def is_wide_slide(width, height):
    return abs((float(width) / float(height)) / ASPECT_16_9 - 1) <= ASPECT_TOLERANCE


def slide_size_type(width, height):
    aspect = float(width) / float(height)
    for expected, value in (
        (ASPECT_16_9, "screen16x9"),
        (ASPECT_16_10, "screen16x10"),
        (ASPECT_4_3, "screen4x3"),
    ):
        if abs(aspect / expected - 1) <= ASPECT_TOLERANCE:
            return value
    return None


def presentation_xml(slide_count, width, height):
    slide_ids = "".join(f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, slide_count + 1))
    size_type = slide_size_type(width, height)
    size_type_attr = f' type="{size_type}"' if size_type else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:sldSz cx="{width}" cy="{height}"{size_type_attr}/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""


def presentation_rels_xml(slide_count):
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    for i in range(1, slide_count + 1):
        rels.append(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{REL_NS}">{"".join(rels)}</Relationships>'


def theme_xml():
    """Return a minimal theme that PowerPoint accepts without repairing.

    DrawingML requires the format scheme's fill, line, effect, and background
    style lists to contain three entries.  The former one-entry lists were
    readable by zip/XML tooling but PowerPoint repaired every generated deck.
    Runs still carry explicit typefaces and colors; these entries are complete
    structural defaults, not a font-substitution path.
    """
    fills = "".join(
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        for _ in range(3)
    )
    lines = "".join(
        f'<a:ln w="{width}" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
        for width in (6350, 12700, 19050)
    )
    effects = "".join('<a:effectStyle><a:effectLst/></a:effectStyle>' for _ in range(3))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ImageToEditablePPT">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F1F1F"/></a:dk2>
      <a:lt2><a:srgbClr val="F8F8F8"/></a:lt2>
      <a:accent1><a:srgbClr val="0F766E"/></a:accent1>
      <a:accent2><a:srgbClr val="E66B00"/></a:accent2>
      <a:accent3><a:srgbClr val="F6D365"/></a:accent3>
      <a:accent4><a:srgbClr val="57C4B8"/></a:accent4>
      <a:accent5><a:srgbClr val="666666"/></a:accent5>
      <a:accent6><a:srgbClr val="111111"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="PingFang">
      <a:majorFont><a:latin typeface="PingFang SC"/><a:ea typeface="PingFang SC"/><a:cs typeface="PingFang SC"/></a:majorFont>
      <a:minorFont><a:latin typeface="PingFang SC"/><a:ea typeface="PingFang SC"/><a:cs typeface="PingFang SC"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>{fills}</a:fillStyleLst>
      <a:lnStyleLst>{lines}</a:lnStyleLst>
      <a:effectStyleLst>{effects}</a:effectStyleLst>
      <a:bgFillStyleLst>{fills}</a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>'''


def write_common_parts(z, slide_count, width, height, notes_count):
    size_type = slide_size_type(width, height)
    presentation_format = {
        "screen16x9": "Widescreen",
        "screen16x10": "Screen 16:10",
        "screen4x3": "On-screen Show (4:3)",
    }.get(size_type, "Custom")
    z.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>""")
    z.writestr("docProps/core.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Image to editable PPT</dc:title></cp:coreProperties>""")
    z.writestr("docProps/app.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application><PresentationFormat>{presentation_format}</PresentationFormat><Slides>{slide_count}</Slides></Properties>""")
    z.writestr("ppt/presentation.xml", presentation_xml(slide_count, width, height))
    z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(slide_count))
    z.writestr("ppt/slideMasters/slideMaster1.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>""")
    z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>""")
    z.writestr("ppt/slideLayouts/slideLayout1.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>""")
    z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>""")
    z.writestr("ppt/theme/theme1.xml", theme_xml())
    if notes_count:
        z.writestr("ppt/notesMasters/notesMaster1.xml", notes_master_xml())
        z.writestr("ppt/notesMasters/_rels/notesMaster1.xml.rels", notes_master_rels_xml())


def _template_content_types(template_bytes, manifests, notes_indices):
    root = ET.fromstring(template_bytes)
    existing_defaults = {
        node.get("Extension") for node in root.findall(f"{{{CT_NS}}}Default")
    }
    existing_overrides = {
        node.get("PartName") for node in root.findall(f"{{{CT_NS}}}Override")
    }
    for manifest in manifests:
        for item in manifest.get("images", []):
            extension = image_ext(item["path"]).lstrip(".")
            if extension not in existing_defaults:
                ET.SubElement(root, f"{{{CT_NS}}}Default", {
                    "Extension": extension,
                    "ContentType": content_type_for(item["path"]),
                })
                existing_defaults.add(extension)
    for index in range(1, len(manifests) + 1):
        part = f"/ppt/slides/slide{index}.xml"
        if part not in existing_overrides:
            ET.SubElement(root, f"{{{CT_NS}}}Override", {
                "PartName": part,
                "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            })
    if notes_indices:
        master = "/ppt/notesMasters/notesMaster1.xml"
        if master not in existing_overrides:
            ET.SubElement(root, f"{{{CT_NS}}}Override", {
                "PartName": master,
                "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
            })
        for index in notes_indices:
            part = f"/ppt/notesSlides/notesSlide{index}.xml"
            if part not in existing_overrides:
                ET.SubElement(root, f"{{{CT_NS}}}Override", {
                    "PartName": part,
                    "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
                })
    # OPC readers require `Types` itself to carry the default content-types
    # namespace.  A generic `ns0:Types` serialization is XML-equivalent but is
    # rejected by Packaging APIs before the Open XML schema is even reached.
    ET.register_namespace("", CT_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _template_presentation_parts(presentation_bytes, rels_bytes, slide_count, width, height, notes_count):
    rels_root = ET.fromstring(rels_bytes)
    used_ids = {
        int(match.group(1))
        for node in list(rels_root)
        if (match := re.fullmatch(r"rId(\d+)", str(node.get("Id") or "")))
    }
    next_rid = max(used_ids or {0}) + 1
    slide_rids = []
    for index in range(1, slide_count + 1):
        rid = f"rId{next_rid}"
        next_rid += 1
        slide_rids.append(rid)
        ET.SubElement(rels_root, f"{{{REL_NS}}}Relationship", {
            "Id": rid,
            "Type": f"{R_NS}/slide",
            "Target": f"slides/slide{index}.xml",
        })
    if notes_count:
        ET.SubElement(rels_root, f"{{{REL_NS}}}Relationship", {
            "Id": f"rId{next_rid}",
            "Type": f"{R_NS}/notesMaster",
            "Target": "notesMasters/notesMaster1.xml",
        })

    root = ET.fromstring(presentation_bytes)
    old_list = root.find(f"{{{P_NS}}}sldIdLst")
    if old_list is not None:
        root.remove(old_list)
    slide_list = ET.Element(f"{{{P_NS}}}sldIdLst")
    for index, rid in enumerate(slide_rids, start=256):
        ET.SubElement(slide_list, f"{{{P_NS}}}sldId", {
            "id": str(index), f"{{{R_NS}}}id": rid,
        })
    master_list = root.find(f"{{{P_NS}}}sldMasterIdLst")
    root.insert(list(root).index(master_list) + 1 if master_list is not None else 0, slide_list)
    size = root.find(f"{{{P_NS}}}sldSz")
    if size is None:
        size = ET.SubElement(root, f"{{{P_NS}}}sldSz")
    size.set("cx", str(width))
    size.set("cy", str(height))
    size_type = slide_size_type(width, height)
    if size_type:
        size.set("type", size_type)
    else:
        size.attrib.pop("type", None)
    presentation_xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    ET.register_namespace("", REL_NS)
    rels_xml_bytes = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    return presentation_xml_bytes, rels_xml_bytes


def _template_app_properties(app_bytes, slide_count, notes_count, width, height):
    root = ET.fromstring(app_bytes)
    values = {
        "Application": "Microsoft Macintosh PowerPoint",
        "PresentationFormat": {
            "screen16x9": "Widescreen",
            "screen16x10": "Screen 16:10",
            "screen4x3": "On-screen Show (4:3)",
        }.get(slide_size_type(width, height), "Custom"),
        "Slides": str(slide_count),
        "Notes": str(notes_count),
    }
    for name, value in values.items():
        node = root.find(f"{{{EP_NS}}}{name}")
        if node is not None:
            node.text = value
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def write_template_parts(z, manifests, width, height, notes_indices):
    template = powerpoint_base_template()
    skip = {
        "[Content_Types].xml", "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels", "docProps/app.xml",
    }
    with zipfile.ZipFile(template) as source:
        names = set(source.namelist())
        required = {
            "[Content_Types].xml", "ppt/presentation.xml",
            "ppt/_rels/presentation.xml.rels", "docProps/app.xml",
            "ppt/slideLayouts/slideLayout7.xml",
        }
        missing = sorted(required - names)
        if missing:
            raise ValueError(f"PowerPoint base template is missing required parts: {missing}")
        for name in source.namelist():
            if name not in skip:
                z.writestr(name, source.read(name))
        presentation, rels = _template_presentation_parts(
            source.read("ppt/presentation.xml"),
            source.read("ppt/_rels/presentation.xml.rels"),
            len(manifests), width, height, len(notes_indices),
        )
        z.writestr("[Content_Types].xml", _template_content_types(
            source.read("[Content_Types].xml"), manifests, notes_indices
        ))
        z.writestr("ppt/presentation.xml", presentation)
        z.writestr("ppt/_rels/presentation.xml.rels", rels)
        z.writestr("docProps/app.xml", _template_app_properties(
            source.read("docProps/app.xml"), len(manifests), len(notes_indices), width, height
        ))
    if notes_indices:
        z.writestr("ppt/notesMasters/notesMaster1.xml", notes_master_xml())
        z.writestr("ppt/notesMasters/_rels/notesMaster1.xml.rels", notes_master_rels_xml())
    return template


def write_pptx(manifest, out_path, manifest_path, *, integrity_report_path=None, build_id=None):
    width = emu(manifest.get("slide", {}).get("width", 13.333))
    height = emu(manifest.get("slide", {}).get("height", 7.5))
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    normalized = normalize_manifest(manifest)
    media_index = 1
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        template = write_template_parts(z, [normalized], width, height, [])
        z.writestr("ppt/slides/slide1.xml", slide_xml(normalized))
        z.writestr("ppt/slides/_rels/slide1.xml.rels", rels_xml(normalized, media_index, None))
        base = Path(manifest_path).resolve().parent
        for item in normalized.get("images", []):
            src = Path(item["path"])
            if not src.is_absolute():
                src = base / src
            z.write(src, f"ppt/media/image{media_index}{image_ext(src)}")
            media_index += 1
    validate_pptx_integrity(
        out,
        report_path=Path(integrity_report_path) if integrity_report_path else out.parent / "pptx-integrity-report.json",
        source_kind="generated",
        build_id=build_id,
        template_path=template,
    )


def deck_slide_size(deck, page_entries):
    slide = deck.get("slide") or {}
    if not slide and page_entries:
        slide = page_entries[0]["manifest"].get("slide", {})
    return emu(slide.get("width", 13.333)), emu(slide.get("height", 7.5))


def write_deck(deck, page_entries, out_path, notes_entries, *, integrity_report_path=None, build_id=None):
    if not page_entries:
        raise ValueError("Deck has no pages")
    width, height = deck_slide_size(deck, page_entries)
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    notes_by_page = {int(entry.get("page_index", 0)): entry for entry in notes_entries if entry.get("text")}
    notes_indices = sorted(notes_by_page)
    normalized_entries = [{**entry, "manifest": normalize_manifest(entry["manifest"])} for entry in page_entries]
    manifests = [entry["manifest"] for entry in normalized_entries]
    media_index = 1
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        template = write_template_parts(z, manifests, width, height, notes_indices)
        for slide_index, entry in enumerate(normalized_entries, start=1):
            manifest = entry["manifest"]
            notes_index = slide_index if slide_index in notes_by_page else None
            z.writestr(f"ppt/slides/slide{slide_index}.xml", slide_xml(manifest))
            z.writestr(f"ppt/slides/_rels/slide{slide_index}.xml.rels", rels_xml(manifest, media_index, notes_index))
            base = Path(entry["manifest_path"]).resolve().parent
            for item in manifest.get("images", []):
                src = Path(item["path"])
                if not src.is_absolute():
                    src = base / src
                z.write(src, f"ppt/media/image{media_index}{image_ext(src)}")
                media_index += 1
            if notes_index is not None:
                note = notes_by_page[slide_index]
                notes_xml = note.get("notes_xml")
                if notes_xml and Path(notes_xml).exists():
                    z.writestr(f"ppt/notesSlides/notesSlide{notes_index}.xml", Path(notes_xml).read_bytes())
                else:
                    z.writestr(f"ppt/notesSlides/notesSlide{notes_index}.xml", notes_slide_xml(note.get("text", "")))
                z.writestr(f"ppt/notesSlides/_rels/notesSlide{notes_index}.xml.rels", notes_rels_xml(slide_index))
    validate_pptx_integrity(
        out,
        report_path=Path(integrity_report_path) if integrity_report_path else out.parent / "pptx-integrity-report.json",
        source_kind="generated",
        build_id=build_id,
        template_path=template,
    )


def page_entries_from_deck_manifest(deck_manifest_path):
    deck_path = Path(deck_manifest_path).resolve()
    deck = json.loads(deck_path.read_text(encoding="utf-8"))
    root = Path(deck.get("job_dir", deck_path.parent)).resolve()
    entries = []
    for page in deck.get("pages", []):
        manifest_path = Path(page.get("manifest", ""))
        if not manifest_path.is_absolute():
            manifest_path = root / manifest_path
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        entries.append({"manifest": manifest, "manifest_path": manifest_path})
    notes_path = deck.get("notes_manifest")
    notes_entries = []
    if notes_path:
        notes_file = Path(notes_path)
        if not notes_file.is_absolute():
            notes_file = root / notes_file
        if notes_file.exists():
            notes_entries = json.loads(notes_file.read_text(encoding="utf-8")).get("notes", [])
            for note in notes_entries:
                notes_xml = note.get("notes_xml")
                if notes_xml:
                    notes_xml_path = Path(notes_xml)
                    if not notes_xml_path.is_absolute():
                        notes_xml_path = root / notes_xml_path
                    note["notes_xml"] = str(notes_xml_path)
    return deck, entries, notes_entries


def output_path_from_deck_manifest(deck_manifest_path):
    deck_path = Path(deck_manifest_path).resolve()
    deck = json.loads(deck_path.read_text(encoding="utf-8"))
    root = Path(deck.get("job_dir", deck_path.parent)).resolve()
    output = Path(deck.get("output", "final/deck_edited.pptx"))
    if not output.is_absolute():
        output = root / output
    return output


def render_preview(manifest, manifest_path, out_path):
    from PIL import Image, ImageColor, ImageDraw, ImageFont

    manifest = normalize_manifest(manifest)
    slide = manifest.get("slide", {})
    width_in = float(slide.get("width", 13.333))
    height_in = float(slide.get("height", 7.5))
    scale = int(manifest.get("preview_scale", 120))
    canvas = Image.new("RGB", (int(width_in * scale), int(height_in * scale)), ImageColor.getrgb(slide.get("background", "#ffffff")))
    base = Path(manifest_path).resolve().parent
    draw = ImageDraw.Draw(canvas)

    def open_preview_image(src):
        if src.suffix.lower() != ".svg":
            return Image.open(src).convert("RGBA")
        convert = "/opt/homebrew/bin/magick"
        if not Path(convert).exists():
            convert = "/opt/homebrew/bin/convert"
        if not Path(convert).exists():
            print(f"Warning: cannot preview SVG without ImageMagick: {src}", file=sys.stderr)
            return None
        with tempfile.NamedTemporaryFile(suffix=".png") as handle:
            subprocess.run([convert, str(src), handle.name], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return Image.open(handle.name).convert("RGBA")

    def render_shape(item):
        box = [item.get("left", 0) * scale, item.get("top", 0) * scale, (item.get("left", 0) + item.get("width", 1)) * scale, (item.get("top", 0) + item.get("height", 1)) * scale]
        fill = preview_color(item.get("fill"))
        outline = preview_color(item.get("stroke", "#000000"))
        width = max(1, int(float(item.get("stroke_width", 1))))
        if item.get("polygon"):
            points = [(point[0] * scale, point[1] * scale) for point in item["polygon"]]
            draw.polygon(points, fill=None if fill in (None, "none") else fill, outline=None if outline == "none" else outline)
        elif item.get("type") == "line":
            if "points" in item:
                points = [value * scale for value in item["points"]]
                draw.line(points, fill=outline, width=width)
                return
            if item.get("dash"):
                draw_dashed_line(draw, box, outline, width)
            else:
                draw.line(box, fill=outline, width=width)
        elif item.get("type") == "ellipse":
            draw.ellipse(box, fill=None if fill in (None, "none") else fill, outline=None if outline == "none" else outline, width=width)
        elif item.get("type") == "roundRect" or item.get("preset") == "roundRect":
            radius = int(float(item.get("radius", 0.12)) * scale)
            draw.rounded_rectangle(box, radius=radius, fill=None if fill in (None, "none") else fill, outline=None if outline == "none" else outline, width=width)
        elif item.get("preset") == "diamond":
            left, top, right, bottom = box
            center_x = (left + right) / 2
            center_y = (top + bottom) / 2
            points = [(center_x, top), (right, center_y), (center_x, bottom), (left, center_y)]
            draw.polygon(points, fill=None if fill in (None, "none") else fill, outline=None if outline == "none" else outline)
        else:
            draw.rectangle(box, fill=None if fill in (None, "none") else fill, outline=None if outline == "none" else outline, width=width)

    def render_image(item):
        src = Path(item["path"])
        if not src.is_absolute():
            src = base / src
        img = open_preview_image(src)
        if img is None:
            return
        img = img.resize((max(1, int(item.get("width", 1) * scale)), max(1, int(item.get("height", 1) * scale))))
        canvas.paste(img, (int(item.get("left", 0) * scale), int(item.get("top", 0) * scale)), img)

    def render_text(item):
        preview_font_scale = float(item.get("preview_font_scale", manifest.get("preview_font_scale", 1.0)))
        size = max(1, int(float(item.get("font_size", 18)) * scale / 72 * preview_font_scale))
        font_path = choose_preview_font(item.get("preview_font"))
        try:
            font = ImageFont.truetype(font_path, size=size) if font_path else ImageFont.load_default()
        except Exception:
            font = ImageFont.load_default()
        if item.get("paragraphs"):
            lines = []
            for paragraph in item["paragraphs"]:
                if isinstance(paragraph, str):
                    lines.append(paragraph)
                else:
                    lines.append("".join(str(run.get("text", "")) for run in paragraph.get("runs", [])))
            preview_text = "\n".join(lines)
        elif item.get("runs"):
            preview_text = "".join(str(run.get("text", "")) for run in item["runs"])
        else:
            preview_text = item.get("text", "")
        fill = preview_color(item.get("color", "#111111"))
        align = text_alignment(item.get("align", "left"))[1]
        valign = text_vertical_alignment(item.get("valign", "top"))[1]
        box_x = int(item.get("left", 0) * scale)
        box_y = int(item.get("top", 0) * scale)
        box_width = max(1, int(item.get("width", 1) * scale))
        box_height = max(1, int(item.get("height", 0.4) * scale))
        rotation = float(item.get("rotation", 0) or 0)

        def aligned_origin(bounds, origin_x, origin_y):
            text_width = bounds[2] - bounds[0]
            text_height = bounds[3] - bounds[1]
            if align == "center":
                text_x = origin_x + (box_width - text_width) // 2 - bounds[0]
            elif align == "right":
                text_x = origin_x + box_width - text_width - bounds[0]
            else:
                text_x = origin_x
            if valign == "middle":
                text_y = origin_y + (box_height - text_height) // 2 - bounds[1]
            elif valign == "bottom":
                text_y = origin_y + box_height - text_height - bounds[1]
            else:
                text_y = origin_y
            return text_x, text_y

        run_specs = []
        if item.get("runs"):
            cursor_x = 0
            base_size = size
            for run in item["runs"]:
                run_size = max(1, int(float(run.get("font_size", item.get("font_size", 18))) * scale / 72 * preview_font_scale))
                run_font_path = choose_preview_font(run.get("preview_font") or item.get("preview_font"))
                try:
                    run_font = ImageFont.truetype(run_font_path, size=run_size) if run_font_path else ImageFont.load_default()
                except Exception:
                    run_font = font
                run_fill = preview_color(run.get("color", item.get("color", "#111111")))
                baseline = float(run.get("baseline", 0) or 0)
                run_y = int(-baseline / 100000 * base_size)
                run_text = str(run.get("text", ""))
                run_specs.append((cursor_x, run_y, run_text, run_fill, run_font, run_font.getbbox(run_text)))
                cursor_x += int(round(draw.textlength(run_text, font=run_font)))

        def draw_content(target_draw, origin_x, origin_y):
            if run_specs:
                bounds = (
                    min(spec[0] + spec[5][0] for spec in run_specs),
                    min(spec[1] + spec[5][1] for spec in run_specs),
                    max(spec[0] + spec[5][2] for spec in run_specs),
                    max(spec[1] + spec[5][3] for spec in run_specs),
                )
                text_x, text_y = aligned_origin(bounds, origin_x, origin_y)
                for run_x, run_y, run_text, run_fill, run_font, _run_bounds in run_specs:
                    target_draw.text((text_x + run_x, text_y + run_y), run_text, fill=run_fill, font=run_font)
                return
            bounds = target_draw.multiline_textbbox((0, 0), preview_text, font=font, spacing=4, align=align)
            text_x, text_y = aligned_origin(bounds, origin_x, origin_y)
            target_draw.multiline_text((text_x, text_y), preview_text, fill=fill, font=font, spacing=4, align=align)

        if rotation:
            layer = Image.new("RGBA", (box_width, box_height), (0, 0, 0, 0))
            draw_content(ImageDraw.Draw(layer), 0, 0)
            rotated = layer.rotate(-rotation, expand=True)
            paste_x = box_x + (box_width - rotated.width) // 2
            paste_y = box_y + (box_height - rotated.height) // 2
            canvas.paste(rotated, (paste_x, paste_y), rotated)
            return
        draw_content(draw, box_x, box_y)

    def render_table(item):
        rows = item.get("cells") or []
        if not rows:
            return
        column_count = len(rows[0])
        columns = item.get("columns") or [item.get("width", 1) / column_count] * column_count
        heights = item.get("rows") or [item.get("height", 1) / len(rows)] * len(rows)
        lefts, tops = [], []
        running = float(item.get("left", 0))
        for value in columns:
            lefts.append(running)
            running += value
        running = float(item.get("top", 0))
        for value in heights:
            tops.append(running)
            running += value
        for row_index, row in enumerate(rows):
            for column_index, raw in enumerate(row):
                cell = raw if isinstance(raw, dict) else {"text": raw}
                if cell.get("h_merge") or cell.get("v_merge"):
                    continue
                col_span = max(1, int(cell.get("col_span", 1) or 1))
                row_span = max(1, int(cell.get("row_span", 1) or 1))
                x0, y0 = lefts[column_index] * scale, tops[row_index] * scale
                x1 = (lefts[column_index] + sum(columns[column_index:column_index + col_span])) * scale
                y1 = (tops[row_index] + sum(heights[row_index:row_index + row_span])) * scale
                fill = preview_color(cell.get("fill", item.get("cell_fill")))
                stroke = preview_color(cell.get("stroke", item.get("stroke", "#BFBFBF")))
                draw.rectangle(
                    [x0, y0, x1, y1],
                    fill=None if fill in (None, "none") else fill,
                    outline=None if stroke in (None, "none") else stroke,
                    width=max(1, int(float(cell.get("stroke_width", item.get("stroke_width", 0.75))))),
                )
                value = str(cell.get("text", ""))
                if not value:
                    continue
                size = max(1, int(float(cell.get("font_size", item.get("font_size", 12))) * scale / 72))
                font_path = choose_preview_font(
                    cell.get("preview_font") or item.get("preview_font")
                    or (cell.get("font_spec") or {}).get("font_file")
                    or (item.get("font_spec") or {}).get("font_file")
                )
                try:
                    font = ImageFont.truetype(font_path, size) if font_path else ImageFont.load_default()
                except OSError:
                    font = ImageFont.load_default()
                color = preview_color(cell.get("color", item.get("color", "#111111")))
                padding = max(2, int(size * 0.3))
                draw.multiline_text((x0 + padding, y0 + padding), value, font=font, fill=color, spacing=2)

    layered = []
    for index, item in enumerate(manifest.get("shapes", [])):
        layered.append((float(item.get("z_index", 100)), index, render_shape, item))
    for index, item in enumerate(manifest.get("images", [])):
        layered.append((float(item.get("z_index", 200)), index, render_image, item))
    for index, item in enumerate(manifest.get("tables", [])):
        layered.append((float(item.get("z_index", 250)), index, render_table, item))
    for index, item in enumerate(manifest.get("text_boxes", [])):
        layered.append((float(item.get("z_index", 300)), index, render_text, item))
    for _z_index, _order, renderer, item in sorted(layered, key=lambda entry: (entry[0], entry[1])):
        renderer(item)
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out_path)
    output = Path(out_path)
    if output.name == "preview.png":
        shutil.copy2(output, output.with_name("quick-preview.png"))


def choose_preview_font(preferred):
    candidates = [
        preferred,
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return candidate
    return None


def draw_dashed_line(draw, box, fill, width):
    x1, y1, x2, y2 = box
    dash = 8
    gap = 6
    if abs(y2 - y1) <= abs(x2 - x1):
        step = dash + gap
        x = min(x1, x2)
        end = max(x1, x2)
        y = y1
        while x < end:
            draw.line((x, y, min(x + dash, end), y), fill=fill, width=width)
            x += step
    else:
        step = dash + gap
        y = min(y1, y2)
        end = max(y1, y2)
        x = x1
        while y < end:
            draw.line((x, y, x, min(y + dash, end)), fill=fill, width=width)
            y += step


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("manifest", nargs="?")
    parser.add_argument("--deck-manifest")
    parser.add_argument("--out")
    parser.add_argument("--preview")
    parser.add_argument("--integrity-report")
    parser.add_argument("--build-id")
    args = parser.parse_args()
    if args.deck_manifest:
        deck, entries, notes_entries = page_entries_from_deck_manifest(args.deck_manifest)
        out = Path(args.out) if args.out else output_path_from_deck_manifest(args.deck_manifest)
        write_deck(
            deck,
            entries,
            out,
            notes_entries,
            integrity_report_path=args.integrity_report,
            build_id=args.build_id,
        )
        print(f"Wrote {out}")
        return
    if not args.manifest:
        parser.error("manifest is required unless --deck-manifest is used")
    if not args.out:
        parser.error("--out is required unless --deck-manifest provides an output")
    manifest = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
    write_pptx(
        manifest,
        args.out,
        args.manifest,
        integrity_report_path=args.integrity_report,
        build_id=args.build_id,
    )
    if args.preview:
        render_preview(manifest, args.manifest, args.preview)
    print(f"Wrote {args.out}")
    if args.preview:
        print(f"Wrote {args.preview}")


if __name__ == "__main__":
    main()
