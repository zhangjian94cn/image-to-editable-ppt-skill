# Design and boundaries

- The caller owns input splitting, one Codex process per page, cancellation,
  and ordered assembly.
- Codex owns visual judgment, object choices, tool selection, true-render
  inspection, and targeted repairs.
- The Skill owns deterministic OCR hints, source-pixel assets, editable object
  construction, installed-font fitting, PowerPoint rendering, comparison,
  PPTX readback, and assembly.

Reconstruction is not redesign. Source wording, data, groups, relative
geometry, and hierarchy are authoritative. The shared Builder supports rich
text, native tables, native connectors and arrows, shapes, and compact images.
Page-specific scripts use `editppt.authoring.SlideManifest`; they do not
reimplement OOXML, font fitting, tables, or connectors.

PowerPoint rendering is the Codex page task's final self-check, not a second
product acceptance state machine.
