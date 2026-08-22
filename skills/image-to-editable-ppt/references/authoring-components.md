# Shared authoring components

Use `editppt.authoring.SlideManifest` when a page benefits from a short Python
authoring script. The script writes `manifest.json`; `editppt build` remains the
only OOXML packager and automatically resolves installed fonts and fits text.

## Minimal page

```python
from editppt.authoring import SlideManifest

page = SlideManifest(1600, 900)
page.add_text(
    [55, 34, 1040, 54],
    "源图中的单行标题",
    font="Microsoft YaHei",  # Builder resolves an installed equivalent.
    font_size_px=48,
    bold=True,
    color="#0588D4",
)
page.add_image([1340, 28, 190, 72], "assets/logo.png", alt="品牌标识")
page.write("manifest.json")
```

Then run:

```bash
editppt build .
editppt render . --input page.pptx --out preview.png
```

## Timeline

```python
page.add_timeline_stage(
    [90, 250, 260, 44],
    number="1",
    period="第1-2周",
    heading="完成最小可用协作",
)
page.add_timeline_stage(
    [410, 250, 260, 44],
    number="2",
    period="第3-6周",
    heading="跑通真实闭环",
)
page.add_connector([350, 272], [410, 272], end_arrow="triangle")
```

The connector is a native PowerPoint connector, not a decorative rectangle or
raster arrow. Use `preset="bentConnector3"` for orthogonal routing.

## Cards and section bands

```python
page.add_section_band([60, 740, 1480, 48], "三个阶段", fill="#2D80D2")
page.add_card(
    [80, 810, 430, 170],
    title="阶段一",
    body="原生文本、容器和分组说明。",
)
```

These components do not redesign a page. Match their boxes, fills, strokes,
text, and rounding to the source; use lower-level `add_shape` and `add_text`
when the source differs.

## Native table with rich text

```python
page.add_table(
    [100, 220, 1400, 460],
    [
        ["序号", "阶段", "说明"],
        ["1", "准备", {"runs": [
            {"text": "必须保留", "bold": True, "color": "#D92D20"},
            {"text": " 的字段"},
        ]}],
    ],
    column_widths=[1, 1.5, 4],
    row_heights=[0.8, 1.2],
)
```

Do not emulate a real grid with dozens of unrelated shapes. Use cell-level
fill, color, bold, border, alignment, and rich-text runs where needed.

## Low-level methods

- `add_shape`: rectangle, rounded rectangle, ellipse, arbitrary supported
  preset, or source-space polygon.
- `add_connector`: straight or bent native connector with start/end arrows.
- `add_text`: one text box with plain text, runs, or paragraphs.
- `add_image`: compact independent asset with required alt text.
- `add_table`: native editable table.
- `add_section_band`, `add_card`, `add_timeline_stage`: reusable composites.

All coordinates are source pixels. Do not hard-code a benchmark filename or
case-specific coordinate rule into the shared library.
The page `width_px/height_px` are the prepared authoring dimensions from
`editppt inspect layout`, not an inferred screenshot display size. Use
`font_size_px` for source-derived typography; use `font_size` only when you
intentionally know the PowerPoint point size.
