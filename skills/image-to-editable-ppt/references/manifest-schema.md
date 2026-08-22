# Shared editable-page manifest

Use the shared Builder instead of ad-hoc OOXML or `python-pptx` when the page is
primarily text, shapes, tables, connectors, and compact images.

## Minimum document

```json
{
  "slide": {"width": 13.333, "height": 7.5, "background": "#FFFFFF"},
  "source": {"width_px": 1600, "height_px": 900},
  "typography": {
    "font_family": "Microsoft YaHei",
    "roles": {"slide_title": {"font_size_px": 50}}
  },
  "shapes": [],
  "images": [],
  "tables": [],
  "text_boxes": []
}
```

Objects may use inches (`left`, `top`, `width`, `height`) or source pixels
(`box_px: [x, y, width, height]`). Lines use `points_px: [x1, y1, x2, y2]`.
Polygons use `polygon_px: [[x1, y1], [x2, y2], [x3, y3], ...]`; nested
`points_px` is accepted as the same polygon spelling, but the explicit field is
preferred.
Use named `layer` values for normal stacking: `background` (0), `container`
(100), `decoration_behind` (120), `band` (140), `content` (200),
`decoration_front` (240), `text` (300), and `overlay` (400). Legacy `z_index`
remains supported. If both are present, `z_index` wins and the Builder writes a
conflict warning to `.editppt/layer-report.json`. Same-z overlapping objects
without named layers are also reported. Shape effects default to none.
For task directories prepared by the Skill, `source.width_px/height_px` must
equal `editppt inspect layout`'s `size_px`; never substitute the original
pre-normalization dimensions or the chat preview dimensions.

`font_size` is always PowerPoint points. `font_size_px` is the preferred input
when reading type size from the source image and is converted to points by the
Builder. Do not provide both fields. `valign` accepts `top`, `middle`/`mid`, or
`bottom`.

Every text object should set `text_role` to one of `slide_title`, `lead`,
`section_title`, `subheading`, `body`, `metric`, `caption`, or `footer`.
Measured object sizes win; then explicit object size; then the role profile;
only then the resolution-scaled fallback. At a 900px page height the fallback
ranges are respectively 42–58, 22–32, 26–38, 20–28, 17–25, 28–48, 13–19,
and 9–15 source pixels. Other heights scale proportionally.

## Rich text in one box

```json
{
  "box_px": [80, 50, 1300, 70],
  "font_size_px": 42,
  "font": "Microsoft YaHei",
  "text_role": "slide_title",
  "layer": "text",
  "wrap": "none",
  "runs": [
    {"text": "3个月看：", "bold": true, "color": "#0085D0"},
    {"text": "是否形成完整闭环", "bold": true, "color": "#111111"}
  ]
}
```

Use runs instead of fragmenting one sentence into several independently placed
text boxes. A source single-line title should use `wrap: "none"` and a measured
box wide enough to keep one line.

## Native table

```json
{
  "box_px": [100, 220, 1400, 420],
  "column_widths": [1.2, 1, 2.4],
  "row_heights": [0.8, 1, 1],
  "font_size": 12,
  "header_fill": "#0A65B7",
  "header_color": "#FFFFFF",
  "fill": "#FFFFFF",
  "border": "#B7C4D4",
  "rows": [
    ["序号", "阶段", "说明"],
    ["1", "准备", "可编辑单元格"],
    ["2", "执行", {"text": "重点", "bold": true, "color": "#D92D20"}]
  ]
}
```

Use a native table for real grids. Do not simulate a large data table with
dozens of unrelated rectangles and text boxes.

## Shapes, connectors, and images

- `shapes`: `rect`, `roundRect`, `ellipse`, `line`, or a supported DrawingML
  `preset`; use `fill: "none"` and `stroke: "none"` explicitly where needed.
  A `line` becomes a native connector and supports `start_arrow`, `end_arrow`,
  `dash`, and straight or bent connector presets. Shapes may use a source-backed
  `gradient` with `angle` and two or more `{position, color}` stops; do not add
  a gradient when the source is flat.
- `images`: local source-bound paths for logos, photos, maps, screenshots, or
  complex illustrations. Keep their boxes tight.
- `text_boxes`: editable text with runs, paragraphs, alignment, vertical
  alignment, explicit wrapping, and optional native `fill`, `stroke`, and
  `stroke_width`. Use these properties for a label/card whose background is
  rectangular; do not add a second shape unless the source needs different
  rounding, geometry, or stacking.

After `editppt build`, inspect with `editppt render` and `editppt inspect pptx`.
The manifest draft is not final visual evidence.
The build JSON reports font resolution, role sizes, shrink adjustments, role
deviations, and layer conflicts. Core text roles shrinking below 90% and body
text below 85% mean the font file, box geometry, or line grouping needs repair;
do not accept automatic tiny text.

For page-specific Python authoring, use `editppt.authoring.SlideManifest` and
the reusable patterns in [authoring-components.md](authoring-components.md).
The Builder searches explicit roots, user/system fonts, and Microsoft
PowerPoint's private `DFonts`, reads internal family names and glyph coverage,
and measures with the resolved file before fitting. Substitutions and the font
environment fingerprint are always recorded.
