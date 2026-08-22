# Image to Editable PPT Skill

Rebuild a slide image, scanned PDF, or visual PPTX page as an object-level
editable PowerPoint slide.

One local Codex task owns one page and uses `$image-to-editable-ppt` plus the
Skill's deterministic tools. It writes `page.pptx`, opens and renders the exact
file through Microsoft PowerPoint, views the result, and only then writes
`result.json`. The caller assembles independent pages in source order.

## Documentation

- [Quick Start](/en/quickstart.md)
- [Design and Boundaries](/en/design.md)
- [Installation](/en/installation.md)
- [Standard Workflow](/en/workflow.md)
- [FAQ](/en/faq.md)
- [Example Prompts](/en/prompts.md)

The Skill includes OCR/layout evidence, source-pixel assets, installed-font
fitting, rich text, native tables and connectors, reusable authoring
components, target-only PowerPoint rendering, and relationship-aware assembly.
It has no page subagents, controller state, containment workflow, or screenshot
fallback.
