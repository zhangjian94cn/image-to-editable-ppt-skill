# Page object decisions

Use this reference when a visual element could reasonably be represented in
more than one way.

## Prefer native PowerPoint objects

Use native text, shapes, tables and connectors for information the user is
likely to edit: wording, numbers, stages, labels, lists, cards, table cells,
background regions, timelines and ordinary diagrams.

Preserve source geometry and hierarchy before decorative precision. Correct
grouping, alignment, reading order, and emphasis matter more than a one-pixel
border difference. Do not add a shadow when the source has none.

## Keep a local image object

Use a separate image object when the source is inherently raster or cannot be
honestly reconstructed at reasonable cost: a photograph, brand mark, map,
product screenshot, textured illustration or complex chart artwork.

Extract from the original pixels when that preserves fidelity. Keep the crop
tight, bind it to the intended object, and leave surrounding labels and
containers editable.

## Never use a raster substitute

Do not place the complete slide or an almost-complete slide region behind or in
front of editable text. Do not rasterize a whole table, card group, text panel,
dashboard or simple diagram merely to finish faster.

## Text and layout

- Preserve source wording, numbers, punctuation and Chinese list numbering.
- Include visible footers, legal lines, page numbers, units, and small captions
  in the inventory. Reinspect the source instead of inventing or shortening
  low-contrast text.
- Use measured OCR boxes as evidence; merge fragmented lines when they form one
  paragraph.
- Assign each text object a semantic role and use measured size groups before
  the resolution-scaled role fallback. Diagnose a mismatch in font-file,
  text-box, line-break, then font-size order.
- Build a small region inventory before authoring dense pages. Count the major
  columns, panels, stages, tables and footer bands visible in the source, then
  make sure each region is represented in the editable result. OCR may miss an
  entire low-contrast or visually dense region.
- Match font family where available and use a visually compatible local font
  when it is not.
- Keep source line breaks when they communicate grouping. Source single-line
  titles remain single-line; measure instead of letting PowerPoint wrap them.
- Keep rich phrases in one box and use runs for emphasis instead of forcing OCR
  fragments into separate objects.

## Overlap and stacking

Overlapping objects must use named layers. Prefer `background`, `container`,
`decoration_behind`, `band`, `content`, `decoration_front`, `text`, then
`overlay`. A card accent intended to peek below its title band belongs in
`decoration_behind`; moving it away from the measured source position only
hides the actual stack bug. Use explicit `z_index` only for a source stack that
the named layers cannot express, and review the resulting conflict report.

## Final judgment

Before writing `result.json`, render the exact `page.pptx` through Microsoft
PowerPoint and view the result as a whole. Correct missing text, overlap,
overflow, wrong wrapping, broken grouping, or misplaced local assets. A ZIP
check, manifest draft, or object count is not visual inspection. This remains
Codex's own authoring judgment, not a backend acceptance state.
