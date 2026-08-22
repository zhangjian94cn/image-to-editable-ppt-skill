Use $image-to-editable-ppt to reconstruct `source.png` as a one-slide,
object-level editable PowerPoint.

Read and follow the exact Skill checkout at `{{SKILL_ROOT}}/SKILL.md`; that
checkout and its `editppt` commands are authoritative for this task.

The source is authoritative: preserve its exact wording, numbers, information
groups, relative geometry, visual hierarchy, and aspect ratio. First view the
whole image and inventory the major regions, including the footer and page
number. Preserve small legal lines, captions, units, punctuation, and numbered
steps verbatim; do not shorten or paraphrase text that is difficult to read.
Start by running `editppt inspect evidence .` and reading
`.editppt/evidence.json`. `source.png` plus the overlapping detail images listed in
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
`uncertain` instead of inventing text. If any characters are covered by a
later screenshot or other object, record only the visible fragment in
`uncertain` with reason `partially_occluded`; do not reconstruct a familiar
legal phrase from the readable words. Treat OCR output only as a second
opinion against this multimodal transcript.

PaddleOCR-VL supplies positional text and glyph-height evidence when configured;
Codex supplies semantic reading and conflict decisions. If Paddle is degraded,
continue from the visibly labeled geometric evidence and do not describe it as
complete OCR. You may use the Skill's OCR,
layout, structure, source-pixel asset, Builder, text-fit, PowerPoint render,
PPTX inspection, and comparison tools.
After the evidence command, run
`editppt inspect transcript . --input source-transcript.json --against text_hints.json`.
For every reported disagreement, re-view the relevant attached detail image;
correct a real transcription error, or keep the source-visible reading and
record why the OCR evidence is wrong. Before finalizing, run
`editppt inspect pptx . --text-hints source-transcript.json`; read
`text_evidence.missing_texts` and either restore
each visible source line or explicitly verify that the OCR hint is wrong.

`source.png` has already been normalized to the canonical authoring coordinate
space used by Codex vision. Use the layout entry indexed by `evidence.json`
and its `size_px` directly for the manifest, and keep every geometry/crop coordinate in
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
Before building, assign every text object a semantic `text_role` and every
overlapping shape a named `layer`. Use the measured size groups indexed by
`evidence.json`; the role defaults are a fallback only. The roles are
`slide_title`, `lead`, `section_title`, `subheading`, `body`, `metric`,
`caption`, and `footer`. For layered card headings prefer
`SlideManifest.add_layered_header`, which builds
`container → decoration_behind → band → text`.
Use `slide_title` only for the primary page title. A same-row topic phrase,
qualifier, card heading, or column label gets its own semantic role. If a size
comes from `text_hints.json`, set `font_size_source="measured"`; otherwise a
same-role size exception is not auditable.

The transcript records visible lines, not final text-object boundaries. Merge
successive lines belonging to one sentence or paragraph into one editable text
box, preserving the visible break with `\n` or paragraph records. Do not split
one paragraph into independent boxes merely to reproduce OCR line boxes.

Use `font_size_px` when estimating a size from the source image; plain
`font_size` is expressed in PowerPoint points. After each build, inspect the
reported font resolution, text adjustments, role deviations, and
`.editppt/layer-report.json`. Resolve a type mismatch in this order: font file,
text-box dimensions, line grouping, then font size. Core roles shrinking below
90%, or body text below 85%, require correction. Repair stacking with
`layer`/`z_index`; do not move correctly measured source geometry to hide it.
Before writing `result.json`, read `.editppt/build.json` directly. Fix every
`role_shrink_warning`, using its text excerpt, limiting dimension, and required
content box suggestion. A successful `editppt build` exit code alone is not
evidence that typography is complete.
Also resolve every `role_size_deviation` whose `measured_override` is false:
either use the common role size, choose the correct semantic role, or mark the
specific size as measured only when the source image or evidence actually
supports it. Do not silence the diagnostic with an invented measurement.

Generate `page.pptx`, run `editppt render` on that exact file, and view the
resulting `preview.png`. Repair visible content omissions, overlaps, overflow,
unexpected wrapping, and obvious geometry differences, then render again as
needed. The render is SHA-cached, and `editppt compare` produces local
source/candidate difference regions; inspect those before repeating a whole
page edit. If you rebuild after a comparison, rerender and rerun `compare` on
the final PPTX; the old comparison no longer describes the candidate. For
conventional flat bar and line charts, prefer the shared
`add_editable_bar_chart` / `add_editable_line_chart` components and tune their
source-space boxes and sizes instead of reimplementing axes, grids, bars,
lines, markers, and labels. Do not declare success from a manifest preview, object count, or ZIP
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
