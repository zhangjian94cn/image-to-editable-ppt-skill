Use $image-to-editable-ppt to reconstruct `source.png` as a one-slide,
object-level editable PowerPoint.

Read and follow the exact Skill checkout at `{{SKILL_ROOT}}/SKILL.md`; that
checkout and its `editppt` commands are authoritative for this task.

The source is authoritative: preserve its exact wording, numbers, information
groups, relative geometry, visual hierarchy, and aspect ratio. First view the
whole image and inventory the major regions, including the footer and page
number. Preserve small legal lines, captions, units, punctuation, and numbered
steps verbatim; do not shorten or paraphrase text that is difficult to read.
`source.png` plus the overlapping detail images listed in
`.editppt/vision-inputs/vision-inputs.json` are attached to this same task.
They are different views of the same slide, not additional slides. Use the
detail views to read small and dense text exactly and map it back through each
entry's `source_box_px`.

Before authoring, write `source-transcript.json` with this minimal public
evidence contract:

```json
{
  "schema_version": 1,
  "backend": "codex-multimodal",
  "source": {"width_px": 1600, "height_px": 900},
  "lines": [
    {"id": "T001", "text": "visible text verbatim", "box_px": [0, 0, 100, 20], "certainty": "high"}
  ],
  "uncertain": []
}
```

Record every visible line you can read, character-for-character. Never use
plausible completion or paraphrasing. Put genuinely unreadable fragments in
`uncertain` instead of inventing text. Treat OCR output only as a second
opinion against this multimodal transcript.

You may use the Skill's OCR,
layout, structure, source-pixel asset, Builder, text-fit, PowerPoint render,
PPTX inspection, and comparison tools.
After `editppt inspect text .`, compare its hints with
`source-transcript.json`. Before finalizing, run
`editppt inspect pptx . --text-hints source-transcript.json`; read
`text_evidence.missing_texts` and either restore
each visible source line or explicitly verify that the OCR hint is wrong.

`source.png` has already been normalized to the canonical authoring coordinate
space used by Codex vision. Run `editppt inspect layout .`, use the returned
`size_px` directly for the manifest, and keep every geometry/crop coordinate in
that space. Never bulk-rescale a completed manifest based on the chat preview
or inferred display size. `editppt assets crop` automatically maps those
coordinates back to the retained original pixels when a source map is present.

Use native text, rich-text runs, shapes, tables, connectors, and ordinary chart
objects wherever practical. Compact independent image objects are allowed for
logos, photos, screenshots, maps, and genuinely complex illustrations. Never
use all or nearly all of `source.png` as a background or covering image.
Any source-pixel crop must be tight around one allowed asset and must not carry
text, timeline segments, container borders, or other editable structure that
you also rebuild. Check the true render for duplicated content around every
crop.

Prefer existing Skill tools. If a page-specific authoring script is useful, it
must use `editppt.authoring.SlideManifest` plus the shared Builder and must not
duplicate font fitting, table, connector, or OOXML packaging logic. For a
source single-line title, use one text box with `wrap: none`; let the Builder
resolve the installed font and fit it, and use `editppt text-fit` when the box
remains tight.
Use `font_size_px` when estimating a size from the source image; plain
`font_size` is expressed in PowerPoint points. After each build, inspect the
reported text adjustments. A shrink ratio below 0.75 is a layout defect to fix
by correcting the coordinate space, box size, line grouping, or requested
font—not a successful automatic fit.

Generate `page.pptx`, run `editppt render` on that exact file, and view the
resulting `preview.png`. Repair visible content omissions, overlaps, overflow,
unexpected wrapping, and obvious geometry differences, then render again as
needed. Do not declare success from a manifest preview, object count, or ZIP
check. If Microsoft PowerPoint cannot render the file, do not report ready.
Keep product names, English identifiers, capitalization, uncommon spellings,
and version strings exactly as shown. Do not rewrite a rare source term into a
more familiar one.

Treat the Skill checkout as read-only. Do not inspect, patch, or replace its
runtime implementation from a page task. `editppt render` exclusively owns the
Microsoft PowerPoint process lifecycle: never call `open`, `osascript`,
`cua-driver`, `pgrep`, `kill`, or application activation/quit commands to
control or diagnose PowerPoint. If the curated renderer fails, preserve its
evidence, explain the failure in your final message, and do not write a ready
result.

Work only in the current page directory plus the exact Skill checkout. Do not
search or read parent directories, sibling benchmark pages, earlier manifests,
another task's output, Codex memory, or prior task history.

When finished, write:

```json
{
  "status": "ready",
  "output_pptx": "page.pptx",
  "warnings": []
}
```

Do not create sub-agents or controller/session/normalizer state.
