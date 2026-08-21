# Optional editable-page manifest

Use this contract only when the deterministic Builder is the best authoring
route. A Codex task may instead author `page.pptx` with another suitable local
library.

## Minimum document

```json
{
  "slide": {"width": 13.333, "height": 7.5},
  "source": {"width_px": 1600, "height_px": 900},
  "shapes": [],
  "images": [],
  "text_boxes": []
}
```

Positioned objects may use inches (`left`, `top`, `width`, `height`) or source
pixel coordinates (`box_px: [x, y, width, height]`). Lines use `points_px` or
their inch equivalents. The Builder maps source pixels to the declared slide.

## Native objects

- `shapes`: rectangles, rounded rectangles, ellipses, lines, arrows and other
  supported DrawingML shapes. Use them for containers, bands, timelines and
  ordinary diagrams.
- `text_boxes`: editable text with font, size, color, alignment and paragraph
  properties.
- `images`: local image assets for logos, photos, maps, screenshots and complex
  illustrations. Paths resolve relative to the manifest.

The existing Builder accepts additional fields for tables, notes, typography,
z-order and provenance. Inspect its help or existing examples only when the
page needs those features.

## Result file

After the PPTX is ready, write:

```json
{
  "status": "ready",
  "output_pptx": "page.pptx",
  "warnings": []
}
```

Warnings describe real limitations; they do not activate another workflow.
