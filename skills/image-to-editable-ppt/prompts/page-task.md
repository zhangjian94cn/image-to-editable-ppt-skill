Use $image-to-editable-ppt to reconstruct `source.png` as a one-slide,
object-level editable PowerPoint.

Read and follow the exact Skill checkout at `{{SKILL_ROOT}}/SKILL.md`; that
checkout and its `editppt` commands are authoritative for this task.

The source is authoritative: preserve its exact wording, numbers, information
groups, relative geometry, visual hierarchy, and aspect ratio. First view the
whole image and inventory the major regions. You may use the Skill's OCR,
layout, structure, source-pixel asset, Builder, text-fit, PowerPoint render,
PPTX inspection, and comparison tools.

Use native text, rich-text runs, shapes, tables, connectors, and ordinary chart
objects wherever practical. Compact independent image objects are allowed for
logos, photos, screenshots, maps, and genuinely complex illustrations. Never
use all or nearly all of `source.png` as a background or covering image.

Prefer existing Skill tools. If a page-specific authoring script is useful, it
must call the shared Builder/components and must not duplicate font fitting,
table, connector, or OOXML packaging logic.

Generate `page.pptx`, run `editppt render` on that exact file, and view the
resulting `preview.png`. Repair visible content omissions, overlaps, overflow,
unexpected wrapping, and obvious geometry differences, then render again as
needed. Do not declare success from a manifest preview, object count, or ZIP
check. If Microsoft PowerPoint cannot render the file, do not report ready.

When finished, write:

```json
{
  "status": "ready",
  "output_pptx": "page.pptx",
  "warnings": []
}
```

Do not create sub-agents or controller/session/normalizer state.
