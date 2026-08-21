# Page object decisions

Use this reference when a visual element could reasonably be represented in
more than one way.

## Prefer native PowerPoint objects

Use native text, shapes, tables and connectors for information the user is
likely to edit: wording, numbers, stages, labels, lists, cards, table cells,
background regions, timelines and ordinary diagrams.

Preserve visual hierarchy before decorative precision. Correct grouping,
alignment, reading order and emphasis matter more than an exact shadow blur or
one-pixel border difference.

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
- Use measured OCR boxes as evidence; merge fragmented lines when they form one
  paragraph.
- Match font family where available and use a visually compatible local font
  when it is not.
- Prefer readable line breaks and stable text boxes over forcing exact OCR
  fragments into separate objects.

## Final judgment

Before writing `result.json`, inspect the page as a whole. Correct obvious
missing text, overlap, overflow, broken grouping or misplaced local assets when
you can. This is Codex's own authoring judgment; it does not create a separate
acceptance state or hand control back to a backend controller.
