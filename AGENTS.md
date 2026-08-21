# Repository guidance

This repository is the canonical source of the installable
`image-to-editable-ppt` Skill. Generated conversions belong in `output/` and
must not be committed.

## Architecture

- One Codex task owns one slide page and writes `page.pptx` plus `result.json`.
- `editppt` exposes optional deterministic helpers only. It must not dispatch
  agents, persist controller sessions, or authorize product downloads.
- Keep the Skill entry concise. Object decisions live in
  `references/page-decision-tree.md`, optional manifest fields in
  `references/manifest-schema.md`, and command syntax in
  `references/cli-helper.md`.
- Preserve OCR, source-bound asset extraction, editable OOXML building,
  previews, assembly, and diagnostics as independently testable helpers.
- Do not introduce run-state, coverage, containment, Hybrid fallback, Office
  acceptance, or fixed repair-loop contracts into the public Skill surface.

## Contribution and release

- Use a feature branch and a pull request for non-trivial changes.
- Keep commit and PR titles in English and use Conventional Commit style.
- Add user-visible changes under `CHANGELOG.md` `## Unreleased`.
- Do not commit generated PPTX, images, caches, credentials, or local `.env`
  files.

## Verification

Run the Skill quick validator, CLI tests, `editppt doctor --json`, and at least
one real single-page Codex smoke before release. Confirm multi-page assembly
preserves page order.
