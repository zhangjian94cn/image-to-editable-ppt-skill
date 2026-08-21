# FAQ

## Why is an object count insufficient?

Existing objects can still wrap, overlap, substitute fonts, or sit in the wrong
place. Always view the true Microsoft PowerPoint render.

## What if PowerPoint rendering fails?

Do not write `ready` and do not fall back to LibreOffice or a screenshot deck.
Preserve `.editppt/render/` evidence and resolve the permission, modal, or file
problem. The renderer opens and closes only the uniquely staged target and does
not create an extra canary.

## Are image objects allowed?

Yes, for tight source-pixel crops of logos, photos, maps, screenshots, and
complex illustrations. Do not rasterize titles, body copy, tables, timelines,
containers, or ordinary charts.

## How are multiple pages handled?

Run one independent Codex task per page, then call `editppt assemble` in source
order. No page-worker or controller state is required.
