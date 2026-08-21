# Lean benchmark runner

This development-only harness evaluates the same one-local-Codex-task-per-page
workflow used by the installable Skill. It does not dispatch agents, resume
sessions, normalize page plans, or decide product downloads.

Freeze a new image case with independently reviewed OCR hints:

```bash
uv run --project skills/image-to-editable-ppt/cli python -m benchmark_runner ingest-image \
  --corpus /absolute/me-miaobi-ppt/benchmarks/image-to-ppt \
  --source /absolute/slide.png --name "session timeline" \
  --text-hints /absolute/text_hints.json
```

```bash
uv run --project skills/image-to-editable-ppt/cli python -m benchmark_runner verify \
  --corpus /absolute/me-miaobi-ppt/benchmarks/image-to-ppt --suite round-0

uv run --project skills/image-to-editable-ppt/cli python -m benchmark_runner run \
  --corpus /absolute/me-miaobi-ppt/benchmarks/image-to-ppt \
  --suite round-0 --out /absolute/benchmark-results/image-to-ppt/runs \
  --label round-0-baseline \
  --skill-root /absolute/image-to-editable-ppt-skill/skills/image-to-editable-ppt
```

Successful page snapshots contain only `source.png`, `candidate.pptx`, the
Microsoft PowerPoint-rendered `candidate.png`, `report.md`, and `artifacts/`.
Failed pages never receive placeholder candidate files.

The final benchmark render calls the same Skill-owned `editppt render` command
used by the page task; the runner does not maintain a second Office adapter.

Because an authoritative render controls the native Microsoft PowerPoint app,
benchmark Codex tasks run with `danger-full-access`. Use this development runner
only with frozen, trusted benchmark inputs on a controlled workstation; it is
not imported by the Miaobi product runtime.
