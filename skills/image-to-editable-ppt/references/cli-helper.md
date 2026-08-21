# `editppt` command reference

The CLI contains optional deterministic helpers. It does not own page state and
does not dispatch agents.

## Prepare inputs

```bash
editppt prepare input.pdf --out-root output/image-to-editable-ppt
editppt prepare page-1.png page-2.png --job-dir /absolute/run-dir
```

The command normalizes visual inputs into ordered page directories containing
`source.png`. Existing run metadata is an input convenience only; Codex does
not need to advance it.

## Inspect one page

```bash
editppt inspect /absolute/page-dir
```

This writes or refreshes `text_hints.json` and `text_hints.png`, then prints a
small JSON summary. Treat the measurements as evidence, not as a page plan.

## Extract transparent components

```bash
editppt extract-assets \
  --input /absolute/sheet.png \
  --out-dir /absolute/page-dir/assets \
  --manifest /absolute/page-dir/assets/components.json
```

Use this for an alpha asset sheet or a source-bound transparent asset. It does
not infer semantics.

## Build and render

```bash
editppt build /absolute/page-dir
editppt build /absolute/page-dir --manifest manifest.json --out page.pptx
editppt render /absolute/page-dir
```

`build` creates `page.pptx` and, unless disabled, `preview.png` from the
optional manifest contract. `render` refreshes only the deterministic preview
from the current manifest; it is a visual aid, not an acceptance gate.

## Assemble pages

```bash
editppt assemble /absolute/page-001 /absolute/page-002 --out /absolute/deck.pptx
```

Assembly consumes each page directory in command order. The deterministic
Builder uses `manifest.json` when present. A single page without a manifest may
be copied directly; a multi-page deck requires build manifests so slide media
and relationships stay valid.

## Diagnose installation

```bash
editppt doctor --json
```

The JSON report includes the Skill contract version and local dependencies.
