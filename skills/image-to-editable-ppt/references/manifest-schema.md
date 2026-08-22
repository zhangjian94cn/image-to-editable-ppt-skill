# Shared editable-page manifest

Use the shared Builder instead of ad-hoc OOXML or `python-pptx` when the page is
primarily text, shapes, tables, connectors, and compact images.

## Minimum document

```json
{
  "slide": {"width": 13.333, "height": 7.5, "background": "#FFFFFF"},
  "source": {"width_px": 1600, "height_px": 900},
  "shapes": [],
  "images": [],
  "tables": [],
  "text_boxes": []
}
```

Objects may use inches (`left`, `top`, `width`, `height`) or source pixels
(`box_px: [x, y, width, height]`). Lines use `points_px: [x1, y1, x2, y2]`.
`z_index` controls the shared object stack. Shape effects default to none.
For task directories prepared by the Skill, `source.width_px/height_px` must
equal `editppt inspect layout`'s `size_px`; never substitute the original
pre-normalization dimensions or the chat preview dimensions.

`font_size` is always PowerPoint points. `font_size_px` is the preferred input
when reading type size from the source image and is converted to points by the
Builder. Do not provide both fields. `valign` accepts `top`, `middle`/`mid`, or
`bottom`.

## Rich text in one box

```json
{
  "box_px": [80, 50, 1300, 70],
  "font_size_px": 42,
  "font": "PingFang SC",
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
  alignment, and explicit wrapping.

After `editppt build`, inspect with `editppt render` and `editppt inspect pptx`.
The manifest draft is not final visual evidence.
The build JSON also reports font shrink adjustments. Any ratio below 0.75
usually means the coordinate frame, box geometry, or line grouping is wrong;
repair that input rather than accepting tiny text.

For page-specific Python authoring, use `editppt.authoring.SlideManifest` and
the reusable patterns in [authoring-components.md](authoring-components.md).
The Builder resolves requested fonts to installed families before fitting,
which prevents PowerPoint-only substitution from changing line breaks.
