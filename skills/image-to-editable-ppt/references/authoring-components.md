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
    text_role="slide_title",
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
For a bracket or open path with more than one bend, use
`page.add_polyline([[x1, y1], [x2, y2], ...])`. The equivalent
`page.add_shape("line", points_px=[[...], ...])` form is supported. Use
`add_polygon` only for a closed filled shape.

## Cards and section bands

```python
page.add_section_band([60, 740, 1480, 48], "三个阶段", fill="#2D80D2")
page.add_card(
    [80, 810, 430, 170],
    title="阶段一",
    body="原生文本、容器和分组说明。",
)
```

For a title band that masks part of a decorative accent, do not rely on write
order. Use the explicit layered component:

```python
page.add_layered_header(
    [10, 45, 420, 300],
    [10, 45, 420, 58],
    [90, 91, 260, 18],
    "整体使用分析",
    title_font_size_px=32,
)
```

It produces `container → decoration_behind → band → text`, so the accent is
behind the title band even if the manifest is later reorganized.

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
- `add_polygon`: three or more source-space points, with native fill and stroke.
- `add_connector`: straight or bent native connector with start/end arrows.
- `add_text`: one text box with plain text, runs, or paragraphs; pass `fill`,
  `stroke`, and `stroke_width` when the rectangular text box itself is the card.
- `add_image`: compact independent asset with required alt text.
- `add_table`: native editable table.
- `add_section_band`, `add_layered_header`, `add_card`, `add_timeline_stage`:
  reusable composites.
- `add_editable_bar_chart`, `add_editable_line_chart`: deterministic native
  shape charts with editable grids, axes, marks, labels, consistent caption
  typography, and named layers.

## Editable charts

```python
page.add_editable_bar_chart(
    [610, 410, 410, 290],
    ["泰州", "镇江", "徐州"],
    [19, 17, 13],
    title="各地市TPD",
    maximum=20,
    grid_values=[0, 5, 10, 15, 20],
)
page.add_editable_line_chart(
    [120, 470, 400, 280],
    ["02月", "03月", "04月", "05月"],
    [19.3, 23.2, 44.2, 46.6],
    title="Token 整体使用情况",
    minimum=10,
    maximum=55,
    grid_values=[10, 25, 40, 55],
)
```

Chart geometry is expressed in source pixels. All chart output is native and
editable; the helpers never rasterize a graph. Use them for conventional flat
charts and keep unusual infographics in lower-level source-space objects.

All coordinates are source pixels. Do not hard-code a benchmark filename or
case-specific coordinate rule into the shared library.
The page `width_px/height_px` are the prepared authoring dimensions from
`editppt inspect layout`, not an inferred screenshot display size. Use
`font_size_px` for source-derived typography; use `font_size` only when you
intentionally know the PowerPoint point size.
Declare `text_role` for text and `layer` for overlapping objects. Use
`z_index` only when the named layers cannot express a genuine source stack.
