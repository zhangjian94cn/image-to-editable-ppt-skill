# `editppt` command reference

Every command has one responsibility, writes only to caller-selected paths,
prints JSON on success, and exits non-zero on failure. Setting
`EDITPPT_TRACE_FILE=/path/events.jsonl` records command, arguments, duration,
exit code, and error without recording secrets.

## Prepare

```bash
editppt prepare input.pdf --job-dir /absolute/run
```

Normalizes images, PDFs, or visual PPTX files into ordered page directories.

## Inspect

```bash
editppt inspect text /page
editppt inspect layout /page
editppt inspect structure /page
editppt inspect pptx /page --input page.pptx
```

- `text`: content-aware OCR when configured, otherwise explicitly labeled
  geometric hints.
- `layout`: content and whitespace bands in source pixels.
- `structure`: source-space line and rectangle candidates; no semantic object
  IDs or containment decisions.
- `pptx`: native object, text, table, connector, picture-coverage, and boundary
  readback.

## Assets

```bash
editppt assets crop --input source.png --out assets/logo.png \
  --left 1300 --top 30 --right 1540 --bottom 120
editppt assets separate --input assets/crop.png --out assets/crop-rgba.png
editppt assets split-alpha --input assets/sheet.png --out-dir assets/items
editppt assets remove-chroma --input green.png --out clean.png --auto-key border
editppt assets brand
editppt assets brand --id cmcc-logo-horizontal-blue-white --out assets/cmcc.svg
```

`crop` reports source coverage and warns at 80% page coverage. `separate`
removes only an edge-connected flat background, preserves foreground pixels,
and reports put-back error. Use these for compact independent assets, never as
a page-raster shortcut.

## Build and text fitting

```bash
editppt build /page
editppt build /page --draft-preview manifest-draft.png
editppt text-fit --text '单行标题' --width-px 900 --height-px 80 --single-line
```

`build` uses the shared manifest Builder. A draft preview is optional and is
explicitly not a PowerPoint render. `text-fit` measures installed fonts and
returns source-pixel measurements; convert pixels to points using the actual
source-to-slide scale.

## True render and compare

```bash
editppt render /page --input page.pptx --out preview.png
editppt compare /page --source source.png --candidate preview.png
```

`render` requires Microsoft PowerPoint, rejects non-authoritative or repaired
opens, and binds `preview.png` to the exact `page.pptx` SHA. The Skill-owned
renderer opens only a collision-proof copy of the target; it does not create a
separate canary or close unrelated presentations. It never falls back to
LibreOffice or the manifest draft. `compare` writes overlay, diff, heatmap, and
coarse diagnostic metrics; Codex must still view the images.

## Assemble and diagnose

```bash
editppt assemble /page-001 /page-002 --out /absolute/deck.pptx
editppt doctor --json
```

Assembly consumes each independent `page.pptx` in supplied order and copies its
slide relationship graph (media, tables, charts, layouts, themes, and notes)
into one Open XML package. It does not require a manifest. Render the assembled
deck through PowerPoint in the caller's final smoke test.
