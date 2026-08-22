# Image to Editable PPT Skill

[简体中文](README.md) · **English** · [한국어](README_ko.md)

Rebuild slide images, scanned PDFs, or image-based PPT/PPTX pages into object-level editable PowerPoint files. This fork is based on `ningzimu/main@fb869763` and intentionally keeps one open Codex + Skill workflow—without page subagents, a controller state machine, or mandatory acceptance gates.

## Usage

Give Codex one slide image and explicitly select the Skill:

```text
Use $image-to-editable-ppt to rebuild source.png as one object-level editable PowerPoint slide.
Use the Skill's OCR, asset extraction, builder, and preview helpers as needed.
Do not use the whole source.png as a full-slide background or overlay.
Write page.pptx and result.json in the current page directory.
```

One Codex task owns one page from inspection through optional preview repairs.

Required outputs:

```text
source.png
page.pptx
result.json
preview.png   # true render of page.pptx from Microsoft PowerPoint
```

```json
{"status":"ready","output_pptx":"page.pptx","warnings":[]}
```

## CLI

```bash
editppt prepare <input...>
editppt inspect text|layout|structure|pptx ...
editppt assets crop|separate|split-alpha|remove-chroma|brand ...
editppt build <page-dir>
editppt text-fit ...
editppt render <page-dir>
editppt compare <page-dir>
editppt assemble <page-dir...> --out <deck.pptx>
editppt formula render-latex ...
editppt doctor --json
```

These are optional deterministic helpers selected by Codex, not a mandatory pipeline. A caller handles multiple pages by starting one Codex task per page and assembling the page files in source order.

## Install and verify

```bash
pipx install --force --python python3.12 ./skills/image-to-editable-ppt/cli
ln -sfn "$PWD/skills/image-to-editable-ppt" ~/.codex/skills/image-to-editable-ppt
editppt doctor --json
python -m unittest discover -s tests -v
```

The shared Builder resolves installed fonts, authors native connectors and rich
tables, and can be composed from `editppt.authoring.SlideManifest`. Never hide
missing editable structure behind a whole-slide raster. Local image objects
remain appropriate for logos, photos, maps, screenshots, and genuinely complex
illustrations. `result.json` may be written only after the exact `page.pptx`
has been opened and rendered by Microsoft PowerPoint.
