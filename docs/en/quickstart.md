# Quick start

Install the Skill and CLI and run `editppt doctor --json`. Give Codex one slide
image and ask:

```text
Use $image-to-editable-ppt to reconstruct source.png as one object-level
editable PowerPoint slide. Never use the whole source as a background or
overlay. Render and view the exact page.pptx through Microsoft PowerPoint
before writing result.json.
```

A successful page contains `source.png`, `page.pptx`, `preview.png`, and
`result.json`. For multiple pages, run one Codex task per prepared page and use
`editppt assemble` in source order. Start with one page and manually review the
source, PowerPoint render, and editable objects.
