# Notesheet PGE — build plan (cycle 1: harness smoke)

> **This file is normally written by the planner agent.** I (Claude,
> bootstrapping the harness) wrote this initial plan by hand so we
> have a working example to point the planner at on the next cycle.
> The planner agent will overwrite this file when the next operator
> ask comes in.

## Operator ask

(See `OPERATOR_ASK.md`.) Smoke-test the PGE harness end-to-end on a
knowingly-broken trivial feature: a new Notesheet note should open
with cell A1 displaying `harness-smoke-OK` in red. The generator
must implement the seed; the evaluator must catch any visual
regression by capturing its OWN screenshot via Playwright MCP
attached to the running Joplin app.

## Features

### feature-1-smoke-red-cell

**Spec**

When the user creates a new Notesheet note via the existing **New
Spreadsheet** command (toolbar button, Tools menu, or
`Cmd/Ctrl+Shift+S`), the resulting note opens with cell A1
pre-populated with the literal text `harness-smoke-OK` rendered in
red (#FF0000). All other cells remain empty. Editing the note works
normally afterward; this is a one-time seed in the initial snapshot.

**Acceptance criteria**

- Screenshot of a freshly-created Notesheet note shows cell A1
  containing the text `harness-smoke-OK` (read from the screenshot
  with the human eye — must be the literal string, not a paraphrase).
- The text in cell A1 visibly renders in **red** (a strong, fully
  saturated red recognizable as #FF0000 — not pink, not maroon, not
  default black). The evaluator must specifically state the text
  color it observed.
- The screenshot is captured by the **evaluator** via Playwright MCP
  attached to the running Joplin desktop app, NOT a screenshot the
  generator dropped into `screenshots/`. (Generator screenshots are
  acceptable additional evidence but not sufficient on their own.)
- All cells other than A1 are empty in the freshly-created note.

**Out of scope**

- The seed only applies to newly-created notes from this session
  forward. Existing Notesheet notes do not migrate.
- Persistence after Joplin restart is fine but not specifically
  required for this smoke; the visual confirmation in the just-created
  note is enough.
- No tests need to be added to the Jest suite for this feature
  specifically. The point of the smoke is to validate the runtime
  harness, not to add Jest coverage. (Future features WILL add Jest
  tests — but the gate for "passing" is the evaluator screenshot, not
  Jest output.)
- Localization: the text is the literal ASCII string
  `harness-smoke-OK`; no translation involved.

**Suggested fixtures**

None — this feature creates a note from scratch via the New
Spreadsheet command, not from an .xlsx import. The generator does
not need to import any fixture for this feature.

**Related risks**

- The New Spreadsheet command lives in `src/index.ts` near the
  command-registration block (`createNewSpreadsheet` or similar) and
  the snapshot scaffold helper in `src/snapshot.ts`. Touching these
  could regress unrelated paths — keep the change minimal and don't
  refactor the surrounding code.
- The initial-snapshot shape needs to carry both the cell value AND
  the cell-level style (font color). Per M13 lessons, **putting the
  color on the snapshot data alone is not enough — the evaluator
  must see Univer actually render it red.** If Univer needs the
  color in a specific shape (e.g. `cl: { rgb: '#FF0000' }` on a
  `style-N` entry referenced by `cellData[0][0].s`), that's part of
  the implementation.

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file as the template. Each feature gets:
- `### feature-N-<kebab-id>` heading
- **Spec**: one paragraph naming the user-observable change
- **Acceptance criteria**: bulleted list of observable evidence (no
  data-shape assertions; no "code does X" — only outcomes)
- **Out of scope**: explicit non-goals
- **Suggested fixtures**: which files under
  `tests/ExcelBaseTestData/formatting-testdata/` exercise this
- **Related risks**: regression hot-spots

The lowest-numbered feature with `passes: false` is what the
generator works on next.
