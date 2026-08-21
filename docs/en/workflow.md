# Standard workflow

The Skill keeps one simple page path:

1. The caller runs `editppt prepare` to create ordered `source.png` pages.
2. One independent local Codex task owns each page and explicitly uses
   `$image-to-editable-ppt`.
3. Codex inventories the whole page, then selects OCR, inspection, asset,
   Builder, font, and comparison tools as needed.
4. Text, data, tables, timelines, containers, ordinary charts, and connections
   are native editable objects. Only compact logos, photos, maps, screenshots,
   or complex illustrations remain images.
5. Codex builds `page.pptx`, runs `editppt render`, views the true PowerPoint
   `preview.png`, and repairs concrete visible problems.
6. It writes `result.json` only after the exact final PPTX renders. The caller
   uses `editppt assemble` to merge independent pages in source order.

There are no page workers, controllers, resumable sessions, dispatch/record
state, coverage/containment graphs, or Hybrid fallback. A failed PowerPoint
open or render must fail clearly and must never become a screenshot delivery.
