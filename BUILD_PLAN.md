# Notesheet PGE — build plan (M19: Static-render & import gaps)

> **Prior plan.** The completed M17 plan (chart import + chart SVG export)
> that this file replaces lives in git history:
> `git log --follow -- BUILD_PLAN.md` (last M17 revision: `git show 045d3d2:BUILD_PLAN.md`).

> **Milestone assignment.** The operator has already assigned all of this
> work to **M19** (README "Milestones" row, `BACKLOG.md` groups D–E). Adding
> anything to or removing anything from that row, or moving an item to a
> different milestone, is the operator's call (AGENTS.md "Milestone
> assignment is the maintainer's call — always ask"). Generators do not
> self-assign.

> **Build order (explicit; it overrides "lowest-numbered first").**
> feature-1 → feature-2 → … → feature-8 → **feature-10** → feature-9.
> feature-10 was added on 2026-10-15, after the plan was written. It has to
> ship **before** the close-out feature-9, because feature-9 checks the docs
> for everything else. feature-9 keeps its id. In `test-results.json` the
> feature-10 row sits before the feature-9 row, so `scripts/pge/run-cycle.sh`
> (which takes the first `passes: false` key in file order) follows the same
> order. Generators: work the first `passes: false` feature **in this order**.

> **Evaluation is independent and black-box (since 2026-10-15).** The
> evaluator (`.claude/agents/evaluator.md`) reads only this spec. It deploys
> the commit under test itself with `scripts/pge/deploy-build.sh`, writes and
> runs its **own** probes, and drives the product with user-visible inputs
> only. It doesn't reuse the generator's tests, probes or screenshots, so
> commit before handing off. Each feature's **Acceptance checks (black-box)**
> section is the operator↔evaluator contract. **The generator must not edit
> it.** If a check is wrong, impossible, or contradicted by the platform, the
> generator stops, reports in `PROGRESS.md` `## In progress`, and the operator
> decides. The generator's own live proof is a self-check, and it should cover
> the same checks, boundaries included.
>
> Conventions in the Acceptance checks: **Mod** = Cmd on macOS, Ctrl on
> Windows/Linux. "As on `main`" means the evaluator records a `main` baseline
> for that exact action first and compares against it. A Windows/Linux variant
> that can't be run live because no such machine is available is recorded as
> `NOT RUN (platform unavailable)` in the evaluator's matrix. On its own it
> doesn't block PASS, provided the feature's Jest criterion for the
> Windows/Linux key mapping passes. The same applies to checks against an
> operator-supplied Excel reference PNG that hasn't been committed yet:
> `NOT RUN (reference missing)`.

## Cycle context

M19 closes the documented gaps in `BACKLOG.md` groups D (PDF / HTML / preview
renderer) and E (editor + `.xlsx` import). D1 (dataBar in HTML) already
shipped in M18. The operator approved all seven remaining items: D2, D3,
D4, D5, E1, E2, E3, E4. Research findings that the operator supplied are
folded into each feature's spec below.

Added 2026-10-15 (operator): feature-2's scope is extended so that
Cmd/Ctrl+K works everywhere inside the Notesheet editor, and feature-10 fixes the
New Spreadsheet shortcut's clash with Joplin's Focus-sidebar default.

Two different proof regimes apply:

- **Renderer features (D2–D5).** `src/contentScripts/notesheetRenderer.ts`
  reads the snapshot without booting Univer. The proof is (a) a Jest
  fidelity test anchored to an **external** ground truth (the fixture's own
  XML, a cited spec, or a real-Excel PNG under `screenshots/excel-reference/`),
  written **failing first** and committed before the fix (CONTRIBUTING
  "Testing"), and (b) a PGE **preview-pane** screenshot
  (`eval-screenshot.js` `previewPane` region, M16 pattern).
- **Editor / import features (E1–E4).** The proof is Jest plus a **live**
  probe in the running Joplin: a CDP probe (pattern `scripts/pge/move-selection.js`),
  a webview import (`scripts/pge/eval-import.js` — `import-fixture.sh` runs
  under Node and can't catch a webview-bundle break, per the M20 notes), and an
  editor-canvas screenshot. Any feature that touches `.xlsx` export also gets a
  file-content check on the **exported** artifact via
  `scripts/pge/eval-export.js` (planner HARD RULE).

### Proof before passing (applies to every feature)

Per `.claude/CLAUDE.md`, a row only flips to `passes: true` after this
sequence: `npm run typecheck && npm run lint && npm test` green → `npm run dist`
→ **quit Joplin → `install-plugin.sh` → `launch-joplin.sh`** (in that order,
because install against a live Joplin removes the plugin) → open/import the
fixture → capture the screenshot under `screenshots/<feature-id>/` → Read
the screenshot → the evaluator grades PASS from its **own** capture. Each new
feature ID has to be registered in `scripts/pge/eval-screenshot.js`'s
feature → region map (`previewPane` for D-features, editor canvas for
E-features). The harness work belongs to the feature.

**Stale build is the first suspect** for any render that looks unfixed
(AGENTS.md). Before debugging, check the `.jpl` mtime against the fix commit
and grep the unpacked bundle for a non-minified string token.

### Operator inputs this plan depends on (gaps the planner had to assume)

1. **Excel reference PNGs for D4/D5 and the non-3Arrows icon sets in D2.**
   `screenshots/excel-reference/ConditionalFormatting-Variants.png` already
   exists and covers 3Arrows (column I). No Excel render exists yet for the
   accounting patterns or the other icon families. Two PNGs need to be
   captured by the operator in real Excel and committed:
   `screenshots/excel-reference/NumberFormats-Accounting.png` (fixture authored
   in feature-4) and `screenshots/excel-reference/IconSets-Variants.png`
   (fixture authored in feature-6). The plan assumes the operator can do this
   while the cycle runs. The PNG-sampling tests use the M15 **skip-when-missing**
   pattern (`excelCanvasFidelity.test.ts` ~L686–735), so the cycle isn't blocked
   by a missing PNG. Each of those features also carries a mandatory
   spec-anchored Jest test (ECMA-376 / Microsoft number-format documentation,
   cited inline), so it still has an external anchor before the PNG lands. When
   feature-4 starts, the generator writes a `STEER`-style request to the operator
   in `PROGRESS.md` `## In progress` asking for both captures, so they can be
   done in one batch.
2. **E4 accelerator choice — resolved.** `CmdOrCtrl+Alt+K` was verified free
   in both default keymap tables at Joplin v3.7.21 and in every installed
   Univer binding (feature-2 Part A, `PROGRESS.md` Notes). It is recorded in
   README and CHANGELOG. On 2026-10-15 the operator extended E4: plain
   Cmd/Ctrl+K must also work everywhere inside the Notesheet editor
   (feature-2 Part B).
3. **New Spreadsheet replacement accelerator (feature-10).** The planner
   proposes candidates. The generator verifies them and picks the first one
   that passes every check in feature-10. If none passes, it stops and the
   operator picks.

### Test-suite no-regression discipline (all features)

The evaluator runs `git diff main -- tests/` and confirms the changes are
**only** new files (`tests/m19*.test.ts`, new fixtures, new
`tests/util/*` helpers) plus the specific existing-test edits listed below.
Any other content edit to an existing test file is a NEEDS_WORK.

| Existing test                                                                                                          | Allowed edit                                                                                                                                                             | Feature   |
| ---------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------- |
| `tests/m16NotesheetMarkdownRender.test.ts` — the krona pin (`'Krona accounting pattern → " 1,234.56 kr "'`, ~L546–560) | Rewrite it to assert the new Excel-anchored layout, **or** delete it and point at the canonical m19 fidelity test (CONTRIBUTING: don't keep two copies of one assertion) | feature-4 |
| `tests/m16NotesheetMarkdownRender.test.ts` — the three US-Accounting pins (~L523–544: positive / negative / zero)      | Same rule as the row above. These pin the "single space" approximation that D5 removes                                                                                   | feature-5 |
| `tests/m12FixtureRoundTrip.test.ts:194` `KNOWN SHORTCOMING — theme-tinted borders…`                                    | Flip it into a passing assertion that cites the m19 E2 fidelity test                                                                                                     | feature-8 |
| `tests/m12ImportRecovery.test.ts` header comments (~L23, ~L79)                                                         | Comment-only updates if the `xlsx-multi-table-unsupported` code is renamed. No assertion changes                                                                         | feature-7 |

`tests/m18AbsoluteRelTargets.test.ts` has to keep passing **unchanged**,
because feature-7 replaces the normalizer that test was written for.

`tests/m19*.test.ts` files created earlier on this branch (e.g.
`tests/m19InsertLinkShortcut.test.ts` from feature-2 Part A) are new files
relative to `main`. Later features (feature-2 Part B, feature-10) may extend
them. Extending them must not weaken an assertion that an earlier, passing
feature relied on.

---

## Features

Order: isolated, low-risk editor items first (E3, E4). Then the renderer items,
smallest first (D3 → D4 → D5 → D2). Then the two import/export items with the
widest blast radius (E1, E2). Then feature-10 (New Spreadsheet shortcut
conflict, added 2026-10-15). Last is the closeout, feature-9 (see "Build order"
at the top). Each feature is one PR-sized unit. feature-4 and feature-5 share a
formatter and **may** ship in one PR, but they are graded separately.

---

### feature-1-m19-arrow-keys-clamp-at-sheet-edge

_Backlog E3. Ours now: upstream Univer #6988 was closed as by-design, and the maintainer advises overriding._

**Spec**

In the Notesheet editor, pressing an arrow key at the sheet's edge leaves the
active cell where it is, as Excel does. Today, Left in column A wraps to the
last column of the previous row, and Left or Up at A1 jumps to the bottom-right
cell. Right in the last column and Down in the last row also wrap. Root cause:
`findNextRange` in `@univerjs/sheets-ui` `lib/es/index.js` (~L5060) defaults
`isGoBack = true`. Shift+arrow (`findNextRangeExpand`) and Ctrl/Cmd+arrow
(`findNextGapRange`) already clamp and must keep working. Committing a cell
edit with an arrow key calls `MoveSelectionCommand` directly (~L10744) and
bypasses `ShortcutService`, so overriding only the shortcut would miss that
path. The fix is a command-level veto. In `src/editorView.tsx`, next to the
existing `univerAPI.addEvent(univerAPI.Event.CommandExecuted, …)` (~L915),
register `univerAPI.addEvent(univerAPI.Event.BeforeCommandExecute, e => …)`.
When `e.id === 'sheet.command.move-selection'`, and `!e.params?.jumpOver`, and
`!e.params?.extra`, and the active cell has **no visible cell beyond it** in
`e.params.direction`, set `e.cancel = true`. Core `executeCommand` catches the
resulting `CanceledError` and returns `false`. The "beyond" check has to be
merge-aware (use the merge's outer edge on the movement side) and hidden-aware
(hidden rows/columns don't count as visible cells). Put the edge decision in a
pure helper (suggested: `src/editor/arrowEdgeClamp.ts`) so Jest can test it
without Univer.

**Acceptance criteria**

1. **Jest (pure helper), new file `tests/m19ArrowEdgeClamp.test.ts`.** It runs
   a fake worksheet model (row/col count, hidden rows/cols, merges) through the
   helper and returns cancel/allow for these cases. A5 Left → cancel. C1 Up →
   cancel. A1 Left → cancel. A1 Up → cancel. Last-col Right → cancel. Last-row
   Down → cancel. B5 Left → allow. Column A hidden with B5 active, Left →
   cancel. Active cell inside merge `A5:B6`, Left → cancel. Active cell inside
   merge `B5:C6`, Left → allow. Any `jumpOver` param → allow (Ctrl/Cmd+arrow
   is never vetoed). Any `extra` param → allow. The Direction enum values come
   from `@univerjs/core`'s exported `Direction` and are not hard-coded.
2. **Live CDP probe** (new `scripts/pge/arrow-edge-probe.js`, modelled on
   `move-selection.js`). It sends **real keystrokes** with Playwright
   `keyboard.press` into the webview frame, not `executeCommand`, because the
   bug lives in the shortcut path. For each case it records the active cell
   before and after (`univerAPI.getActiveWorkbook().getActiveSheet().getActiveRange()`).
   The JSON output in `screenshots/feature-1-m19-arrow-keys-clamp-at-sheet-edge/probe-*.json` shows:
    - A5 Left → A5. C1 Up → C1. A1 Left → A1. A1 Up → A1. Last-col Right → unchanged. Last-row Down → unchanged.
    - B5 Left → A5 (normal movement is not affected).
    - Shift+Left at A5 → selection unchanged, no wrap (already clamped, regression check).
    - Ctrl/Cmd+Left from D5 → jumps to the data edge as it does on `main` today (regression check; record the `main` result first).
    - Type `x` at A5, then Left → the edit commits (A5 value is `x`) and the active cell stays A5.
    - Column A hidden, B5 active, Left → stays B5.
    - A drawing (use fixture `tests/fixtures/images/TwoImages-MultiSheet.xlsx`) selected → arrow keys still nudge the drawing (its transform `left` changes). The veto must not swallow drawing moves.
3. **Screenshot** `screenshots/feature-1-m19-arrow-keys-clamp-at-sheet-edge/eval-*.png`
   of the editor canvas after Left from A1. The active-cell highlight is
   visibly on A1 (top-left), not at the bottom-right of the grid.
4. **Console signal.** The Joplin DevTools console shows no uncaught
   `CanceledError` and no `[Notesheet]` errors during the probe run.
5. **Docs.** The README "Known gaps → Editor" bullet for Left in column A / Up
   in row 1 is removed. A CHANGELOG entry and the `BACKLOG.md` E3 line are
   updated to say it's fixed in M19, with the citation to upstream #6988's
   by-design closure.

**Acceptance checks (black-box) — added 2026-10-15 for the record, after the
evaluator PASS.** This list records what was verified, including the corner
cases the operator asked for after the first PASS. It doesn't re-open the
feature. Default sheet = 20 columns (A–T) × 1000 rows.

- **AC1** (crit 2, 3): given A5 is selected / when the user presses Left / then A5 stays selected.
- **AC2** (crit 2): given C1 / when Up / then C1 stays selected.
- **AC3** (crit 2, 3): given A1 / when Left, and separately Up / then A1 stays selected, and the view doesn't scroll to the bottom-right.
- **AC4** (crit 2): given T5 (last column) / when Right / then T5 stays selected. Given C1000 (last row) / when Down / then C1000 stays selected.
- **AC5** (crit 2): given T1000 (bottom-right corner) / when Right, and separately Down / then T1000 stays selected.
- **AC6** (crit 2): given a sheet resized to 50 columns / when AX5 Right, and AX1000 Down / then the selection doesn't move. When T5 Right / then U5 is selected (the clamp follows the real width).
- **AC7** (crit 2): given B5 / when Left / then A5 is selected (normal movement unchanged).
- **AC8** (crit 2): given A5 / when Shift+Left / then the selection is unchanged and doesn't wrap.
- **AC9** (crit 2): given D5 with data to its left / when Mod+Left / then the selection jumps to the data edge, as on `main`.
- **AC10** (crit 2): given A5 / when the user types `x` then presses Left / then A5 shows `x` and A5 stays selected.
- **AC11** (crit 2): given column A hidden and B5 selected / when Left / then B5 stays selected.
- **AC12** (crit 2): given an image selected in `TwoImages-MultiSheet.xlsx` / when an arrow key is pressed / then the image moves.
- **AC13** (crit 4): during all of the above, the DevTools console has no uncaught `CanceledError` and no `[Notesheet]` error.

**Out of scope**

- Enter / Tab / Shift+Enter / Shift+Tab wrap behaviour (these use different commands).
- Ctrl/Cmd+arrow semantics (left as Univer implements them).
- Page Up/Down and Home/End.

**Suggested fixture(s)**

- A fresh empty notesheet (`scripts/pge/create-empty-notesheet.sh`), seeded via CDP with a value at D5 and a merge `A5:B6` for the merge cases.
- `tests/fixtures/images/TwoImages-MultiSheet.xlsx` for the drawing-nudge case.

**Related risks**

- Univer also fires `move-selection` from the drawing / float-DOM code. The `extra` guard and criterion 2's drawing case are there to cover it.
- Formula-bar / cell-editor arrow handling: inside an active edit, Left moves the caret, not the selection. Confirm the veto only fires when the command is actually dispatched.
- Research says this survives Univer 1.x (`BACKLOG.md` F5). Note the dependency on `BeforeCommandExecute` in the F5 upgrade checklist comment.

**Files likely touched**

`src/editorView.tsx`, new `src/editor/arrowEdgeClamp.ts`, new
`tests/m19ArrowEdgeClamp.test.ts`, new `scripts/pge/arrow-edge-probe.js`,
`scripts/pge/eval-screenshot.js` (feature id), `README.md`, `BACKLOG.md`,
`CHANGELOG.md`.

---

### feature-2-m19-insert-link-alt-shortcut

_Backlog E4. Operator decisions: "Alt shortcut + document" (2026-10-14).
**Scope extended 2026-10-15: "Cmd/Ctrl+K must work everywhere" inside the
Notesheet spreadsheet editor.** The id is unchanged; the feature now covers
Part A (built) plus Part B (new)._

**Status when the scope was extended (2026-10-15).** Part A (the alt shortcut,
the Tools menu item, and the item being disabled for markdown notes) is built
and the generator has self-checked it live, but the evaluator hasn't graded it.
Commits: `d2bafb1` (failing-first tests), `1725529` (implementation), `fa8e033`
(live proof + docs), `c18f673` (re-verify on build #42). Part B is not started.
The `test-results.json` row stays `passes: false`, with its existing evidence
list kept, until the evaluator grades **both** parts.

**Superseded research premise (kept for the record).** The original spec said:
_"At runtime Joplin finds no built-in editor and logs `Received command, but
editor is gone textLink`, so the keystroke never reaches the plugin iframe. A
plugin can't fix this."_ The generator's live probes contradict the first half
of that (macOS, Joplin 3.7.21; `main` at `2edd4b6` behaves the same):

- **Grid focused:** Univer's own `InsertLinkShortcut` (`@univerjs/sheets-hyper-link-ui`
  0.23 `lib/es/index.js` ~L1627; binding `KeyCode.K | MetaKeys.CTRL_COMMAND`,
  precondition `whenSheetEditorFocused`) handles the keydown first and opens its
  link panel. Joplin's `textLink` does not run.
- **In-cell editor focused (typing in a cell), or focus outside the editor:**
  Univer doesn't handle the key. Joplin's `textLink` app-menu accelerator runs,
  logs `Received command, but editor is gone textLink`, and nothing opens.
- **Formula bar focused:** not yet probed.

So the keydown **does** reach the editor webview whenever Univer is the focused
target, and a keydown that Univer handles stops Joplin's accelerator, at least
on macOS. Evidence: `screenshots/feature-2-m19-insert-link-alt-shortcut/probe-{main,branch}-*.json`
and `PROGRESS.md` Notes, "M19 feature-2 (E4) live-harness notes". The rest of the
original reasoning still holds: re-registering or overriding `textLink` would
break the markdown editors, and the plugin must never do it.

**Spec**

_Part A — built (keep as is)._ The command `notesheetInsertLink` ("Insert link
in spreadsheet", `enabledCondition: 'hasActivePluginEditor'`) is exposed as a
Tools menu item with the accelerator `CmdOrCtrl+Alt+K`. The accelerator was
verified free in both default keymap tables at Joplin v3.7.21 and in every
installed Univer binding. Running the command posts `{ type: 'insertLink' }` to
the active, visible editor view, and the webview runs Univer's insert-link
operation (`InsertLinkShortcut.id` = `sheet.operation.insert-hyper-link-toolbar`).
In markdown notes the menu item is greyed out.

_Part B — new._ Inside the Notesheet spreadsheet editor, Cmd+K (macOS) or
Ctrl+K (Windows/Linux) opens Univer's insert-link panel wherever keyboard focus
is in the editor:

- the **grid** (already works through Univer; must keep working),
- the **in-cell editor** while the user is typing in a cell,
- the **formula bar**.

Text typed into the cell or formula bar before the shortcut is pressed is never
lost. Outside the editor (note list, sidebar, search box, note title) and in
markdown notes (Markdown editor and Rich Text editor), Cmd/Ctrl+K behaves
exactly as on `main`. The plugin never registers, overrides or rebinds Joplin's
`textLink` command or its accelerator. The fix lives inside the editor webview.
Cmd/Ctrl+Alt+K and the Tools menu item stay as the alternative entry points, and
all three entry points (Mod+K, Mod+Alt+K, Tools menu) do the same thing in the
same focus state.

**What "insert link while editing a cell" does is decided by research, not
assumed.** The candidates are:

- **(i) Univer's in-editor link:** link the selected text, or insert a link at
  the caret, inside the cell being edited. Univer 0.23's hyperlink UI has
  editing-mode handling (`getEditingCustomRangePosition`, the `currentEditing`
  popup paths in `sheets-hyper-link-ui` ~L1110–1200), which suggests the toolbar
  operation already supports this.
- **(ii) Commit then open:** commit the edit, then open the panel for the cell.

The default is whatever Univer's own toolbar link button and Part A's Mod+Alt+K
already do in that state, so the entry points agree. The generator records the
choice, with evidence, in `PROGRESS.md` Notes and the PR. If research shows (i)
and (ii) are both impossible or both lose text, the generator stops and reports
(see below).

**Required research before coding Part B** (AGENTS.md "Verify external-platform
behavior with research"; cite file paths and doc URLs):

1. How Univer 0.23 routes Mod+K when the **cell editor** (a docs unit) or the
   **formula bar** has focus. Which shortcut preconditions apply there, whether
   `docs-hyper-link-ui` is loaded by our presets and binds Mod+K, and what
   `insert-hyper-link-toolbar` does when the cell editor is active.
2. Whether a **capture-phase** `keydown` handler in the editor webview can
   recognise Mod+K in those focus states, run Univer's insert-link for the
   current cell (or text), and `preventDefault` so Joplin's `textLink`
   accelerator doesn't also fire. This is observed for Univer's own handler on
   macOS. For Windows/Linux, confirm the order in which Electron handles
   app-menu accelerators relative to renderer/iframe keydown from Electron docs
   or source, because it may differ from macOS.
3. That the grid path, which already works through Univer, doesn't fire twice.
   The operation toggles: a second dispatch closes the panel it just opened.
4. That IME composition (`isComposing`) and auto-repeat don't trigger
   unintended opens or toggles.

**STOP rule.** If research or live probing shows Part B can't be done in some
in-editor focus state without touching `textLink` (e.g. on Windows/Linux the
accelerator fires before the webview sees the key), or only by losing typed
text, the generator stops. It writes the evidence to `PROGRESS.md`
`## In progress` and leaves the row `passes: false`. It doesn't ship a partial
behaviour silently and doesn't rebind `textLink`. The operator decides.

**Windows/Linux.** Part B covers Ctrl+K, and Part A covers Ctrl+Alt+K. Live
verification on Windows/Linux may not be possible (no machine). If so, the
generator says so in PROGRESS and the PR. It relies on the research from item 2
plus the Jest key-matching criterion, and the evaluator records those checks as
`NOT RUN (platform unavailable)` (see the header conventions).

**Acceptance criteria**

_Part A (built; still graded):_

1. **Jest** (`tests/m19InsertLinkShortcut.test.ts`, existing `api` mocks):
    - `onStart` registers `notesheetInsertLink` with `enabledCondition: 'hasActivePluginEditor'`.
    - A menu item is created with the accelerator `CmdOrCtrl+Alt+K`.
    - Running the command posts `{ type: 'insertLink' }` to the active editor handle.
    - It posts nothing when `isActive` returns false.
    - It never registers or overrides `textLink`: no `commands.register` call has the name `textLink`.
2. **Live, alt keystroke path.** With a Notesheet note open and B2 selected, a
   real OS keystroke Mod+Alt+K (System Events) opens Univer's hyperlink edit
   panel (link text / URL inputs visible), and Joplin's markdown "Insert
   Hyperlink" dialog is **not** showing. This holds with the grid focused and
   while typing in the cell. Screenshot
   `screenshots/feature-2-m19-insert-link-alt-shortcut/eval-*.png`.
3. **Live, menu path.** Tools → "Insert link in spreadsheet" opens the same panel.
4. **Live, no regression on markdown notes.** In a markdown note, Mod+K still
   opens Joplin's own link dialog, and the new menu item is greyed out.

_Part B (new):_

5. **Research recorded.** `PROGRESS.md` Notes and the PR body record the
   findings for research items 1–4 with citations, and the chosen in-cell
   semantics ((i) or (ii)) with the evidence for it.
6. **Jest, written failing first** (extend `tests/m19InsertLinkShortcut.test.ts`
   or add `tests/m19CmdKInEditor.test.ts`). It pins the key-matching rule for
   both platforms:
    - It matches macOS Meta+K and Windows/Linux Ctrl+K.
    - It does **not** match macOS Ctrl+K, Mod+Shift+K, or Mod+Alt+K (Mod+Alt+K stays Part A's accelerator path).
    - It does not match a keydown with `isComposing` set or an auto-repeat.
    - It does not dispatch when the grid itself is focused, where Univer's own shortcut already handles the key. This guards against a double toggle.
    - It never touches `textLink`.
7. **Live, macOS, real Cmd+K keystrokes** (System Events, as in Part A) open
   Univer's link panel with the **grid**, the **in-cell editor** and the
   **formula bar** focused. Joplin's "Insert Hyperlink" dialog does not open.
   Joplin's log and the DevTools console show no `Received command, but editor
is gone textLink` line for presses made inside the editor. A screenshot per
   focus state goes under `screenshots/feature-2-m19-insert-link-alt-shortcut/`.
8. **Live, no text loss and link applied.** Typed text survives the shortcut.
   Entering a URL and confirming produces a cell that displays a link to that
   URL. Cancelling leaves the typed text and adds no link.
9. **Live, unchanged outside the editor.** Markdown notes (both editors) and
   focus outside the Notesheet editor behave as on `main` for Mod+K.
10. **Docs (file-content).**
    - The README "Known gaps → Editor" Cmd/Ctrl+K bullet is rewritten. Mod+K now opens Univer's link panel in the grid, while typing in a cell, and in the formula bar. Mod+Alt+K and Tools → "Insert link in spreadsheet" remain as alternatives.
    - The bullet states the remaining residual: with a Notesheet open but focus outside the editor, Mod+K still goes to Joplin and does nothing.
    - The bullet states whether Windows/Linux was live-verified.
    - The `textLink` rebind workaround is removed, unless something still needs it.
    - `BACKLOG.md` E4 and the CHANGELOG Unreleased entry are updated to match. Released CHANGELOG entries are not edited.

**Acceptance checks (black-box)**

Setup: a Notesheet note (an empty one, plus `Hyperlinks-Variants.xlsx` imported)
and a plain markdown note. "Panel" = Univer's link edit panel with its URL input
visible. "Joplin's dialog" = Joplin's markdown "Insert Hyperlink" prompt.

- **AC1** (crit 2): given a Notesheet with B2 selected and the grid focused / when the user presses Mod+Alt+K / then the panel opens at B2, and Joplin's dialog doesn't appear.
- **AC2** (crit 2, 8): given the user has typed `abc` into B2 and is still editing / when Mod+Alt+K / then the panel opens and `abc` isn't lost.
- **AC3** (crit 3): given B2 selected / when the user clicks Tools → "Insert link in spreadsheet" / then the panel opens for B2.
- **AC4** (crit 4): given a markdown note is open / when the user opens the Tools menu / then "Insert link in spreadsheet" is greyed out. When Mod+Alt+K is pressed / then no panel and no new dialog appears.
- **AC5** (crit 7): given B2 selected and the grid focused / when Mod+K / then the panel opens at B2, and Joplin's dialog doesn't appear (regression of what works on `main`).
- **AC6** (crit 7, 8): given the user has typed `abc` into B2, caret at the end, still editing / when Mod+K / then the panel opens, `abc` is still visible (in the editor or committed to B2), and the Joplin log has no `editor is gone textLink` line for this press.
- **AC7** (crit 8): continuing AC6 / when the user enters `https://example.org` in the panel and confirms / then, once editing ends, B2 displays `abc` as a link, hovering it shows `https://example.org`, and no other cell changed.
- **AC8** (crit 8): given the user has typed `abc` into B2 and pressed Mod+K / when the user cancels the panel (Escape or clicking away) / then B2 still holds `abc` (still editing or committed, per the researched semantics) and has no link.
- **AC9** (crit 7, 8): given B2 selected and the user has clicked into the formula bar and typed `xyz` / when Mod+K / then the panel opens and `xyz` isn't lost. After confirming `https://example.org`, B2 shows a link.
- **AC10** (crit 7): given the user starts editing an empty B3 (editor open, nothing typed) / when Mod+K / then the panel opens with no error. Cancelling leaves B3 empty.
- **AC11** (crit 7): given a non-default position — the bottom-right cell T1000 (scrolled into view), and separately AX5 on a sheet widened to 50 columns, and separately B2 on the **second** sheet tab — / when Mod+K with the grid focused, and again while editing that cell / then the panel opens anchored at that cell, inside the visible editor area.
- **AC12** (crit 7): given the panel is open after a Mod+K / when Mod+K is pressed again / then it behaves as Mod+K does in the grid on `main` (the panel closes). At no point are two panels visible. A third press opens it again. The same holds when the presses start from in-cell editing.
- **AC13** (crit 7): given B2:C3 selected / when Mod+K, then (after closing) Mod+Alt+K, then (after closing) Univer's toolbar link button / then all three produce the same result.
- **AC14** (crit 7): given B2 selected in `Hyperlinks-Variants.xlsx` after import / when Mod+K is pressed and the panel is cancelled / then the fixture's existing links still render as links and still open their targets when clicked.
- **AC15** (crit 9): given a markdown note in the Markdown editor / when Mod+K / then Joplin's dialog opens, as on `main`. Given the same note in the Rich Text editor / when Mod+K / then the same link dialog appears as on `main`.
- **AC16** (crit 9): given a Notesheet note is open but focus is in the note list (and separately in the search box) / when Mod+K / then the result is the same as on `main` (no panel; nothing new appears).
- **AC17** (crit 6): on macOS / given the grid focused, and separately while editing a cell / when **Control**+K (not Cmd) / then the result is the same as on `main` (the panel does not open).
- **AC18** (crit 2, 6, 7): Windows/Linux variants of AC1, AC5, AC6, AC9 and AC15, using Ctrl+Alt+K and Ctrl+K. `NOT RUN (platform unavailable)` is allowed per the header conventions.
- **AC19** (crit 7): throughout AC1–AC17 the DevTools console has no uncaught error and no `[Notesheet]` error.
- **AC20** (crit 10): README "Known gaps → Editor" no longer says Mod+K fails while editing a cell. It names Mod+Alt+K and the Tools item as alternatives and states the focus-outside-the-editor residual. CHANGELOG Unreleased and `BACKLOG.md` E4 say the same.

**Out of scope**

- Making Mod+K open the panel when focus is outside the Notesheet editor.
- Changing, re-registering or rebinding Joplin's `textLink`.
- Other inputs inside the editor UI (find/replace box, sheet-tab rename, conditional-formatting side panel, comment editor). Mod+K there is not required to open the panel. It must not throw or lose typed text.
- Editing or removing an existing link via the shortcut. That is whatever Univer's panel does.
- Joplin Mobile.

**Suggested fixture(s)**

- An empty notesheet (`scripts/pge/create-empty-notesheet.sh`) for the insert cases. Widen it to 50 columns for AC11 (`scripts/pge/widen-columns.js` / the feature-1 probe pattern).
- `tests/fixtures/formatting-testdata/Hyperlinks-Variants.xlsx` (AC14), and `MultiSheet.xlsx` for the second-tab case.

**Related risks**

- **Double toggle.** Univer's grid shortcut and a new webview handler could both fire on one press, so the panel opens and immediately closes. Criterion 6 and AC5/AC12 guard this.
- **Cell-edit commit semantics.** Committing or cancelling the in-cell editor around the panel can drop typed text, or write it to the wrong cell after selection moves.
- **Electron accelerator order on Windows/Linux** may differ from macOS (research item 2).
- **Multiple open editor views** (split view or a second window): Part A targets the **active, visible** handle only. Part B handles the key in the focused webview only.
- The `onMessage` dispatcher in `editorView.tsx` also handles `snapshot` messages. The M18 RTE fence work is adjacent.
- **Univer 1.x upgrade (BACKLOG F5):** re-verify `InsertLinkShortcut.id`, the preconditions, and the cell-editor/formula-bar focus detection.

**Files likely touched**

`src/editorView.tsx`, `src/editor/insertLink.ts` (or a new pure key-matching
helper under `src/editor/`), `src/index.ts` (Part A, already done),
`tests/m19InsertLinkShortcut.test.ts` and/or new `tests/m19CmdKInEditor.test.ts`,
`scripts/pge/insert-link-probe.js` / `insert-link-lib.js` (generator self-check),
`README.md`, `BACKLOG.md`, `CHANGELOG.md`.

---

### feature-3-m19-rich-text-runs-in-html

_Backlog D3._

**Spec**

In Joplin's preview pane and in PDF / HTML export, a cell whose snapshot
carries rich-text runs (`cell.p.body.textRuns`) renders each run with its own
style. Today `renderCellValue` (`notesheetRenderer.ts` ~L713) flattens
`dataStream` to plain text. Each run becomes a `<span>` covering
`dataStream.slice(st, ed)` (escaped). The span's inline style maps Univer
`ITextStyle` as follows:

- `bl:1` → `font-weight:bold`
- `it:1` → `font-style:italic`
- `ul.s:1` → `text-decoration:underline`
- `st.s:1` → `text-decoration:line-through` (combined with underline when both are set)
- `cl.rgb` → `color`
- `fs` → `font-size:<n>pt`
- `ff` → `font-family` (quoted and escaped)
- `va` (sub/superscript) → `vertical-align:sub|super` only if the importer emits it

Text not covered by any run inherits the cell style. Paragraph breaks
(`\r\n` / `\r` / `\n` in the dataStream) keep rendering as `<br/>`, including
inside a run. All text and attribute values stay escaped. MaliciousValues-style
`<script>` in a run must render as literal text.

**Acceptance criteria**

1. **Jest fidelity test written failing first**, new
   `tests/m19RichTextHtml.test.ts`. The expected run boundaries and styles are
   parsed **independently from the fixture's own sheet XML** with JSZip + regex
   on `xl/worksheets/sheet1.xml` `<is><r><rPr>…</rPr><t>…</t></r>`. They are
   not taken from `xlsxBufferToSnapshot` output. For
   `RichTextInOneCell.xlsx`:
    - A1: "Hello" is inside an element whose computed style includes bold, and " world" is not bold.
    - A2: "Red" carries the color from the XML (`00FF0000` → `#FF0000`), " and " has no run color, "Blue" carries `#0000FF`, and " text" has no run color.

    Parse the rendered HTML with a DOM parser (jsdom is already available to
    Jest, or use an equivalent). Don't match byte strings, so the emit format
    can change without breaking the test. The first commit on the branch for
    this feature contains this test **failing** against the current renderer.

2. **Jest escape / safety.** A synthetic snapshot with a run whose text is
   `<img src=x onerror=alert(1)>`, and a run whose `ff` contains `";background:url(x)`,
   renders with no live element and no injected style. The text appears
   literally.
3. **Jest regression.** Cells with `cell.p` but **no** `textRuns`, and hyperlink
   cells (`buildHyperlinkCellP` shape: one empty-`ts` run plus `customRanges`),
   render exactly the same text as on `main`. The existing
   `m16NotesheetMarkdownRender` and `rteFenceRoundTrip` suites pass unchanged.
4. **PGE preview-pane screenshot** of `RichTextInOneCell.xlsx`, imported via
   the webview Import button (`eval-import.js`) and viewed in the preview pane,
   at `screenshots/feature-3-m19-rich-text-runs-in-html/eval-*.png`. In cell A1,
   "Hello" is visibly heavier than " world". In A2, "Red" is red and "Blue" is
   blue, and " and " / " text" are the default ink colour.
5. **Docs.** The README "Known gaps → PDF / HTML export" per-run rich-text
   bullet is removed, and `BACKLOG.md` D3 is marked done. The stale "Per-run
   formatting is out of scope per the M16 plan" comment in
   `notesheetRenderer.ts` (~L703–706) is updated.

**Acceptance checks (black-box)**

Setup: import fixtures through the Notesheet import UI (Tools → "Import .xlsx as
Notesheet", or the webview Import button) and view the note in Joplin's
**preview pane** unless a check says otherwise. Source runs, from the fixture
XML: A1 = **"Hello"** bold + " world" plain. A2 = "Red" `#FF0000` + " and " +
"Blue" `#0000FF` + " text". A3 = plain text "Visit example.com for more info"
with a hyperlink.

- **AC1** (crit 1, 4): given `RichTextInOneCell.xlsx` / when the user views the preview / then in A1 "Hello" is visibly heavier than " world", and " world" is regular weight.
- **AC2** (crit 1, 4): same note / then in A2 "Red" is red, "Blue" is blue, and " and " and " text" are the default ink colour. The four pieces read as one continuous line, "Red and Blue text", with no extra spaces or line breaks.
- **AC3** (crit 3): same note / then A3 reads "Visit example.com for more info" with the same text as on `main`. It isn't bold or coloured by a neighbouring run.
- **AC4** (crit 1): given the same note / when the user exports it to HTML (File → Export → HTML) / then in the exported file "Hello" in A1 sits inside an element styled bold and " world" doesn't, and "Red" / "Blue" carry `#FF0000` / `#0000FF` (file-content check).
- **AC5** (crit 4): given the same note / when the user exports it to PDF / then the PDF page shows the same bold and colour distinctions as the preview.
- **AC6** (crit 1): given an empty notesheet / when the user types `line1`, inserts an in-cell line break (Alt+Enter; Option+Enter on macOS), types `line2`, makes only `line2` bold in the cell editor, commits, and switches to the preview / then the cell shows two lines, and only `line2` is bold.
- **AC7** (crit 2): given `MaliciousValues.xlsx` / when previewed / then every script- or HTML-looking value appears as literal text. No alert or dialog appears, no image or element is injected, and the DevTools console shows no script error.
- **AC8** (crit 3): given `NumberFormats.xlsx` and `FormattingSmorgasboard.xlsx` / when previewed / then every cell's text and number formatting matches `main`, including cells with no rich text.
- **AC9** (crit 3): given `RichTextInOneCell.xlsx` open / when the user switches the note to the Rich Text editor and back (Joplin's toggle-editors button) / then the note still opens as a spreadsheet with A1–A3 unchanged (the fence survives).
- **AC10** (crit 5): README "Known gaps → PDF / HTML export" no longer lists per-run rich text, and `BACKLOG.md` D3 is marked done (file-content check).

**Out of scope**

- Rendering hyperlinks as `<a href>` in the preview (they render as text today and continue to).
- Paragraph-level styles (alignment, indent) within `cell.p`.
- Rich text in chart titles (charts have their own SVG path).
- Theme-font fallback inside a run beyond what `ff` names.

**Suggested fixture(s)**

`tests/fixtures/formatting-testdata/RichTextInOneCell.xlsx` (A1, A2).
`MaliciousValues.xlsx` for a spot check.

**Related risks**

- `renderCellValue` is shared by every cell. The numFmt path (`formatNumberWithPattern`) must still run for non-rich cells.
- RTE fence integrity (CONTRIBUTING "Rendered notesheet HTML must survive the Rich Text editor"). The new spans live inside the `joplin-editable` wrapper and must not change `wrapEditable`.

**Files likely touched**

`src/contentScripts/notesheetRenderer.ts`, new `tests/m19RichTextHtml.test.ts`,
`scripts/pge/eval-screenshot.js`, `README.md`, `BACKLOG.md`, `CHANGELOG.md`.

---

### feature-4-m19-accounting-symbol-position

_Backlog D4._

**Spec**

In the preview pane and in PDF / HTML export, an accounting-style number format
puts the currency symbol where the format code puts it. For the krona pattern
`_-* #,##0.00 "kr"_-`, the symbol goes **after** the number (`1,234.56 kr`), as
Excel does. The current `formatAccounting` (`notesheetRenderer.ts` ~L661)
pulls the first quoted literal and always emits `${symbol}${number}`. It also
emits `kr-` for zero under a single-section format, where Excel shows
`0.00 kr`. Replace the regex-and-symbol extraction with a **section tokenizer**
that keeps format-code order. Token kinds: quoted literal, `\X` escaped
literal, bare literal chars (`$ - + ( ) : space`), `_X` pad, `*X` fill, `??`
digit placeholders, and the numeric core (`#,##0.00`). The rendered output
follows token order, with Excel's section-selection rules:

- 1 section: applies to all values, and negatives get a leading `-`.
- 2 sections: positive/zero ; negative.
- 3 sections: positive ; negative ; zero.
- 4 sections: the fourth is text.

For this feature `_X` and `*X` keep their **current** single-space rendering.
feature-5 replaces that. This feature only changes ordering and section
selection.

**Acceptance criteria**

1. **New fixture** `tests/fixtures/formatting-testdata/NumberFormats-Accounting.xlsx`,
   generated by a new section in `generate_excel_samples.py`. Don't edit
   `NumberFormats.xlsx` (fixtures README rule). Rows:
    - krona single-section `_-* #,##0.00 "kr"_-` × {1234.56, −1234.56, 0}.
    - krona 4-section (Excel's built-in Swedish accounting) `_-* #,##0.00 "kr"_-;-* #,##0.00 "kr"_-;_-* "-"?? "kr"_-;_-@_-` × {1234.56, −1234.56, 0, "text"}.
    - US Accounting `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)` × {1234.56, −1234.56, 0, "text"}.
    - Euro suffix `#,##0.00 "€"` × {1234.56}, as a regression check.

    Add the fixture to the fixtures README table.

2. **Jest fidelity test written failing first**, new
   `tests/m19AccountingFormats.test.ts`. Expected **visible-text token order**
   (ignoring pad/fill whitespace) is written in the test with a citation to the
   Microsoft number-format-code documentation and ECMA-376 Part 1 §18.8.31.
   Examples: krona 1234.56 → `1,234.56 kr`. Krona −1234.56 (single section) →
   `-1,234.56 kr`. Krona 0 (single section) → `0.00 kr`. Krona 0 (4-section) →
   `- kr`. US 1234.56 → `$` … `1,234.56`. US −1234.56 → `$` … `(1,234.56)`.
   US 0 → `$` … `-`. Text in the 4th section → the text. The cell's visible text
   is read from the rendered DOM.
3. **Excel-reference geometric test (skip-when-missing)** against
   `screenshots/excel-reference/NumberFormats-Accounting.png` (operator input
   #1). For the krona rows, the "kr" ink's left edge is to the **right** of the
   digits' ink. For the US rows, the "$" ink sits in the left part of the cell
   and the digits sit at the right. Sample with `tests/util/pngSampler.ts`
   (pure stdlib), following the M15 skip-when-missing pattern.
4. **Pin migration.** The krona pin in `m16NotesheetMarkdownRender.test.ts`
   (~L546) is rewritten or removed per the allowed-edit table, and cites the m19
   test. No other m16 assertion changes.
5. **PGE preview-pane screenshot** of the imported `NumberFormats-Accounting.xlsx`
   at `screenshots/feature-4-m19-accounting-symbol-position/eval-*.png`. Every
   krona row reads number-then-`kr`. The single-section zero row reads `0.00 kr`.
   The US rows still read `$`-then-number.
6. **Docs.** The README "Known gaps → PDF / HTML export" krona bullet is removed,
   `BACKLOG.md` D4 is marked done, and the formatter header comment
   (~L345–358) is corrected.

**Acceptance checks (black-box)**

Setup: import `NumberFormats-Accounting.xlsx` (authored in this feature) and view
the preview pane. "Reads" means the visible characters in order, with
inter-token whitespace ignored. **The Excel reference screenshot
`screenshots/excel-reference/NumberFormats-Accounting.png` comes from the
operator** (captured in real Excel). Checks against it are
`NOT RUN (reference missing)` until it's committed.

- **AC1** (crit 2, 5): krona single-section, 1234.56 / then it reads `1,234.56 kr`, with the number first and `kr` last.
- **AC2** (crit 2, 5): krona single-section, −1234.56 / then it reads `-1,234.56 kr`.
- **AC3** (crit 2, 5): krona single-section, 0 / then it reads `0.00 kr`, not `kr-` or `kr0.00`.
- **AC4** (crit 2, 5): krona 4-section, 1234.56 / −1234.56 / 0 / text `text` / then they read `1,234.56 kr`, `-1,234.56 kr`, `- kr`, and `text`.
- **AC5** (crit 2, 5): US Accounting, 1234.56 / −1234.56 / 0 / `text` / then they read `$` … `1,234.56`, `$` … `(1,234.56)`, `$` … `-`, and `text`. `$` is always first.
- **AC6** (crit 2): Euro suffix 1234.56 / then it reads `1,234.56 €` (unchanged from `main`).
- **AC7** (crit 3): with the operator's Excel reference PNG committed / when the evaluator compares its preview capture to it / then for every krona row `kr` is right of the digits in both, and for every US row `$` is at the left of the cell and the digits are at the right in both.
- **AC8** (crit 5): given `NumberFormats.xlsx` / when previewed / then B9 (krona) reads number-then-`kr`, B16 (US accounting) reads `$`-then-number, and every other cell in the sheet, including any `[Red]`/`[Blue]` section-colour cells and the unknown-pattern cell, shows the same text and colour as on `main`.
- **AC9** (crit 5): given `NumberFormats-Accounting.xlsx` / when the note is exported to PDF / then the PDF shows the same token order as AC1–AC6.
- **AC10** (spec, unchanged neighbour): given the same note / when opened in the **spreadsheet editor** (not the preview) / then the cells display exactly as on `main`. This feature changes the static renderer only.
- **AC11** (crit 6): README "Known gaps → PDF / HTML export" no longer lists the krona-symbol position bullet. It lists `[$…]` locale-currency codes as a remaining gap. `BACKLOG.md` D4 is marked done (file-content check).

**Out of scope**

- `[$kr-41D]` locale-currency bracket codes and other `[$…]` currency tags beyond the existing `[$-409]` strip. List these as a remaining gap in README.
- Locale-specific digit grouping and decimal separators. Output stays `en-US` grouping, as today.
- Width semantics of `_X`, `*X`, and `??` (feature-5).

**Suggested fixture(s)**

New `NumberFormats-Accounting.xlsx`. Existing `NumberFormats.xlsx` B9 (krona)
and B16 (US accounting) as smoke checks.

**Related risks**

- `formatNumberWithPattern`'s branch order (~L380–480). The `[Red]…;[Blue]…` two-section colour branch runs before accounting. Make sure a 2-section accounting code without colour markers reaches the tokenizer, and that colour-marker codes still take the colour branch.
- The text-suffix branch (`@ "suffix"`) has to keep winning for its pattern.
- The unknown-pattern fall-through pin (`'???-NO-SUCH-PATTERN-???'` → raw) has to stay green.

**Files likely touched**

`src/contentScripts/notesheetRenderer.ts`, `tests/fixtures/formatting-testdata/generate_excel_samples.py`,
new fixture plus fixtures README row, new `tests/m19AccountingFormats.test.ts`,
`tests/m16NotesheetMarkdownRender.test.ts` (allowed pin only), `scripts/pge/eval-screenshot.js`,
`README.md`, `BACKLOG.md`, `CHANGELOG.md`.

---

### feature-5-m19-accounting-fill-and-padding-width

_Backlog D5. Builds on feature-4's tokenizer._

**Spec**

In the preview pane and in PDF / HTML export, Excel's column-alignment
constructs render with Excel's geometry instead of collapsing to one space:

- `_X` (pad) becomes blank space exactly the width of character X. Use an inline `<span style="visibility:hidden">X</span>` or an equivalent technique that keeps exact glyph width and survives PDF print.
- `??` digit placeholders become blank space the width of a digit (hidden `0`).
- `*<space>` (repeat-fill) makes the cell content span the full cell width, with the fill absorbing the slack. That puts the text before the fill flush left and the text after it flush right. Use an inline flex container with a `flex:1` spacer, so it works in Joplin's preview and in Chromium PDF export.

Together, these make a column of US-Accounting values line up as it does in
Excel: `$` at the left edge, the digits right-aligned, and positive values
inset by the width of `)` so they line up with the parenthesised negatives.

**Acceptance criteria**

1. **Jest fidelity test written failing first**, in
   `tests/m19AccountingFormats.test.ts` (new `describe`). The structure is
   anchored to the format code (ECMA-376 / Microsoft docs, cited):
    - For US positive 1234.56, the rendered cell contains a hidden pad whose text is exactly `)` after the digits, and a fill spacer between `$` and the digits.
    - For krona, there are hidden `-` pads at both ends.
    - For the 4-section zero `"-"??`, there are exactly two hidden digit-width pads after `-`.
    - The visible text (pads and fills stripped) equals feature-4's expectations, so the test confirms the two features compose.
    - No pad or fill markup leaks visible characters. `visibility:hidden` or equivalent is asserted.
2. **Alignment check (Jest + jsdom-free).** For a 3-row US-Accounting column
   (1234.56 / −1234.56 / 0), every rendered cell has the same structure:
   `[symbol][fill][number-or-dash][pad or ")"]`. The positive value's trailing
   pad character equals the negative value's closing `)`. That is the property
   that makes Excel's columns align.
3. **Excel-reference geometric test (skip-when-missing)**, same PNG as
   feature-4. For the US rows, the right edge of the last digit in the positive
   row and the right edge of the last digit inside the parentheses in the
   negative row are within Δ ≤ 3 px of each other (Excel's alignment property).
   The same property is checked on the evaluator's preview-pane capture.
4. **Pin migration.** The three US-Accounting pins in
   `m16NotesheetMarkdownRender.test.ts` (~L523–544) are rewritten or removed per
   the allowed-edit table.
5. **PGE preview-pane screenshot** of `NumberFormats-Accounting.xlsx` at
   `screenshots/feature-5-m19-accounting-fill-and-padding-width/eval-*.png`.
   In the US column, `$` hugs the cell's left edge and the digits of the
   positive and negative rows line up on the right. The krona rows are
   right-aligned with `kr` last. The same layout holds in a PDF export of that
   note (File → Export → PDF; screenshot of the PDF page, or `pdftotext -layout`
   output showing the alignment, captured as evidence).
6. **Docs.** The README "Known gaps → PDF / HTML export" `_X` underscore-fill
   bullet is removed, `BACKLOG.md` D5 is marked done, and the
   `formatAccounting` comment (~L650–660) about the "single space"
   approximation is removed.

**Acceptance checks (black-box)**

Setup: same note as feature-4 (`NumberFormats-Accounting.xlsx`, preview pane).
Pixel positions come from the evaluator's own preview capture. **The Excel
reference screenshot `screenshots/excel-reference/NumberFormats-Accounting.png`
comes from the operator.** Checks against it are `NOT RUN (reference missing)`
until it's committed.

- **AC1** (crit 1, 5): US Accounting column, rows 1234.56 / −1234.56 / 0 / then in all three rows `$` hugs the cell's left edge (left padding only).
- **AC2** (crit 2, 3, 5): same column / then the right edge of the last digit of `1,234.56` (positive row) and the right edge of the last digit inside `(1,234.56)` (negative row) are within 3 px of each other horizontally. The positive row has visible blank space where the negative row has `)`.
- **AC3** (crit 1, 5): US zero row / then the `-` sits left of where the positive row's last digit ends, by roughly two digit widths. It isn't flush against the right border.
- **AC4** (crit 1, 5): krona rows / then the content is right-aligned with `kr` last, and there's a small blank gap (about one `-` wide) between `kr` and the cell's right border.
- **AC5** (crit 1): every accounting cell / then no pad or fill characters are visible: no stray `)` on positive rows, and no `_`, `*` or `?` anywhere.
- **AC6** (crit 1, 4): visible text with blanks ignored is the same as feature-4's AC1–AC6 (the two features compose).
- **AC7** (spec, non-default dimension): given the US Accounting column widened in the editor to about twice its width / when the user switches back to the preview / then `$` still hugs the left edge, the digits still hug the right edge, and the gap between them grew.
- **AC8** (crit 5): given the same note / when exported to PDF / then the PDF shows the AC1/AC2 alignment (`pdftotext -layout` output or a PDF page capture).
- **AC9** (risk, unchanged neighbour): given `MergedCellsAndAlignment.xlsx`, and the non-accounting rows of `NumberFormats.xlsx` / when previewed / then wrapped cells, merged cells and row heights look the same as on `main`, and accounting rows are single-line height.
- **AC10** (crit 3): with the operator's Excel reference PNG committed / then the AC2 alignment property also holds in the Excel reference (it confirms that is Excel's behaviour).
- **AC11** (crit 6): README "Known gaps → PDF / HTML export" no longer lists the `_X` underscore-fill bullet. It lists non-space `*X` fill as a remaining gap. `BACKLOG.md` D5 is marked done (file-content check).

**Out of scope**

- `*X` with a non-space X (e.g. `*-` dash leaders). It renders without the repeated character and is documented as a remaining gap in README.
- `?` placeholders in fraction formats (`# ?/?`).
- Exact pixel parity of glyph widths with Excel's fonts (we match the construct, not Excel's font metrics).

**Suggested fixture(s)**

`NumberFormats-Accounting.xlsx` (from feature-4).

**Related risks**

- A `display:flex` container inside a `<td>` can change row height and wrapping. Check that cells with `wrapText` and merged cells around the accounting column look unchanged in the preview screenshot.
- PDF export path (Chromium print). Confirm the flex spacer still works and `print-color-adjust` isn't affected.

**Files likely touched**

`src/contentScripts/notesheetRenderer.ts`, `tests/m19AccountingFormats.test.ts`,
`tests/m16NotesheetMarkdownRender.test.ts` (allowed pins only), `scripts/pge/eval-screenshot.js`,
`README.md`, `BACKLOG.md`, `CHANGELOG.md`.

---

### feature-6-m19-iconset-glyphs-in-html

_Backlog D2._

**Spec**

In the preview pane and in PDF / HTML export, cells covered by an iconSet
conditional-formatting rule show the band's icon glyph to the left of the
value, as Excel does. Today the `iconSet` case falls through to `default` in
`evaluateCfFor` (`notesheetRenderer.ts` ~L1026). The importer
(`src/xlsx.ts` ~L1358–1427) emits Univer-shaped rules
`{ type:'iconSet', isShowValue, config:[{operator, value:cfvo, iconType, iconId}] }`.
Config is in **descending** threshold order: index 0 is the highest band, and
the last entry is a catch-all. `iconType` is the source `iconSet` name, passed
through verbatim, and each entry can name its own `iconType` (Univer-authored
rules may mix sets or use `EMPTY_ICON_TYPE`). The renderer:

1. Resolves each entry's cfvo against the range values. Reuse `resolveCfvo` for num / percent / percentile / min / max.
2. Picks the first entry whose `operator` / `value` the cell satisfies, which mirrors Univer's `IconSetCalculateUnit`.
3. Maps `(iconType, iconId)` to an inline-SVG glyph, following the ordering in Univer's `iconGroup` (`@univerjs/sheets-conditional-formatting` `lib/es/index.js` ~L2255; e.g. 3Arrows = [up-green, right-gold, down-red]).
4. Emits the glyph inline in the cell, before the value, sized to the cell font.
5. Honours `isShowValue === false` by omitting the value.

Inline SVG, not Unicode, so colours survive PDF export and don't depend on
emoji fonts.

**Supported first (in scope).** All 3-icon sets Univer and Excel share:
`3Arrows`, `3ArrowsGray`, `3TrafficLights1`, `3TrafficLights2`, `3Signs`,
`3Symbols`, `3Symbols2`, `3Flags`. `EMPTY_ICON_TYPE` renders no glyph.

**Acceptance criteria**

1. **Jest band-selection fidelity, written failing first**, new
   `tests/m19IconSetHtml.test.ts`. For `ConditionalFormatting-Variants.xlsx`
   column I (`I2:I11`, 3Arrows, cfvo percent 0/33/67, from the fixture's own
   `<cfRule type="iconSet">` XML, parsed independently), the expected band for
   each cell is computed from the XML cfvos with Excel's `>=` semantics
   (ECMA-376 §18.3.1.49 `gte` default true). The rendered HTML's glyph
   identity (a `data-icon="3Arrows:0"`-style attribute or an equivalent stable
   hook) matches: values in the top band → up-green, the middle band →
   right-gold, the bottom band → down-red.
2. **Jest colour anchoring.** The glyph fill colours for 3Arrows match the
   Excel reference `screenshots/excel-reference/ConditionalFormatting-Variants.png`
   within RGB Δ ≤ 8. Sample the I11 (green up), a mid-band cell, and I2 (red
   down) glyph regions with `tests/util/pngSampler.ts`, reusing the sentinel
   approach in `excelCanvasFidelity.test.ts` ~L888. Colours sampled from the
   PNG are the expected values, not constants copied from our SVG.
3. **New fixture** `tests/fixtures/formatting-testdata/IconSets-Variants.xlsx`
   (openpyxl `IconSetRule`, via `generate_excel_samples.py`). One column per
   in-scope set, values 0..90, plus one column with `showValue="0"`. Add it to
   the fixtures README. Jest asserts the band per cell for every in-scope set,
   from the fixture XML. A skip-when-missing PNG colour test runs against
   `screenshots/excel-reference/IconSets-Variants.png` (operator input #1).
4. **Jest out-of-scope behaviour.** A 4- or 5-icon set (e.g. `5Rating`) and an
   unknown name render the value with **no** glyph and no exception. This is
   pinned as a `KNOWN SHORTCOMING` test naming the follow-up.
5. **PGE preview-pane screenshot** of `ConditionalFormatting-Variants.xlsx` at
   `screenshots/feature-6-m19-iconset-glyphs-in-html/eval-*.png`. In column I,
   high-value rows show a green up-arrow, mid rows a gold right-arrow, low rows
   a red down-arrow, each to the left of the number. A second capture of
   `IconSets-Variants.xlsx` shows traffic-light circles, symbols, and flags in
   their columns. The dataBar, colorScale, cellIs, and top-N columns of
   ConditionalFormatting-Variants look unchanged from `main`.
6. **Docs.** README "Known gaps → PDF / HTML export" is updated so the iconSet
   bullet lists **only** the unsupported sets. The README "Compatibility" /
   renderer paragraph (~L48: "only iconSet is not yet rendered") is corrected,
   and `BACKLOG.md` D2 is updated.

**Acceptance checks (black-box)**

Setup: import fixtures and view the preview pane. In
`ConditionalFormatting-Variants.xlsx`, column I is I2:I11 = 0, 10, …, 90 with
3Arrows at percent thresholds 0/33/67 (`gte`). That puts I2–I4 (0–20) in the
bottom band, I5–I8 (30–60) in the middle band, and I9–I11 (70–90) in the top
band. `screenshots/excel-reference/ConditionalFormatting-Variants.png` is already
committed. **`screenshots/excel-reference/IconSets-Variants.png` comes from the
operator.** Checks against it are `NOT RUN (reference missing)` until it's
committed.

- **AC1** (crit 1, 5): `ConditionalFormatting-Variants.xlsx` / then I9, I10 and I11 each show a green up-arrow immediately left of the number.
- **AC2** (crit 1, 5): then I5–I8 show a gold/yellow right-arrow. Boundary cells: I5 (30, the lowest middle value) and I8 (60, the highest middle value).
- **AC3** (crit 1, 5): then I2–I4 show a red down-arrow. Boundary cells: I2 (0, the minimum) and I4 (20, the highest bottom value).
- **AC4** (crit 2): the arrow colours in the evaluator's capture match the same cells in `ConditionalFormatting-Variants.png` (green, gold and red families; canvas anti-aliasing tolerance applies).
- **AC5** (crit 3, 5): `IconSets-Variants.xlsx` / then each in-scope set's column (3Arrows, 3ArrowsGray, 3TrafficLights1, 3TrafficLights2, 3Signs, 3Symbols, 3Symbols2, 3Flags) shows that set's three glyphs. The lowest values show the "low" icon and the highest values show the "high" icon, matching the fixture's thresholds.
- **AC6** (crit 3, out of scope): in the `showValue="0"` column / then values are **still shown** next to the glyphs (expected until the importer follow-up lands).
- **AC7** (crit 3): with the operator's `IconSets-Variants.png` committed / then glyph shapes and colours per column match the Excel reference by colour family.
- **AC8** (crit 4): given a workbook with a `5Rating` icon set and one with an unknown icon-set name (the evaluator may author these with openpyxl under its own folder) / when imported and previewed / then the values show with **no** glyph, the import succeeds, and the console has no error.
- **AC9** (crit 5): `ConditionalFormatting-Variants.xlsx` / then the colour-scale (A), data-bar (C) and cell-is (E) columns, and any other conditional-formatting ranges on the sheet, look the same as on `main`.
- **AC10** (spec): given the same note / when exported to PDF / then the arrows appear in colour in the PDF, not as blank boxes or emoji.
- **AC11** (unchanged neighbour): row heights in column I are unchanged from `main`, and no glyph is taller than its row. The spreadsheet **editor's** own canvas icons look as on `main`.
- **AC12** (crit 6): README no longer says "only iconSet is not yet rendered". The iconSet Known-gaps bullet lists only the unsupported sets. `BACKLOG.md` D2 is updated (file-content check).

**Out of scope (explicit list)**

- 4- and 5-icon sets: `4Arrows`, `4ArrowsGray`, `5Arrows`, `5ArrowsGray`, `4RedToBlack`, `4TrafficLights`, `4Rating`, `5Rating`, `5Quarters`, `5Boxes`, `_5Felling`.
- x14-only 3-icon sets: `3Triangles`, `3Stars`.
- Custom per-threshold icon overrides from `<x14:cfIcon>` beyond what the importer already maps.
- Importer changes: `reverse="1"`, and mapping source `showValue="0"` into `isShowValue`. The importer hard-codes `isShowValue: true`. The renderer honours the snapshot flag, but fixing the importer is a follow-up. Note it in README, and note in criterion 3 that the `showValue="0"` column is expected to show values until that follow-up lands.
- Changing Univer's own canvas icon artwork (README already documents that gap).

**Suggested fixture(s)**

`ConditionalFormatting-Variants.xlsx` (column I) plus new `IconSets-Variants.xlsx`.

**Related risks**

- `evaluateCfFor` currently returns only per-cell **fills**. Glyphs need a second output channel (e.g. `Map<key, {fill?, iconHtml?, hideValue?}>`). Change the return type without changing dataBar / colorScale / cellIs / rank behaviour. Their existing m16 tests must stay green unchanged.
- Rule priority / `stopIfTrue` interplay with other rules on the same range.
- Bundle size of the content script. Keep the SVGs as small path strings, not base64 images (Univer's own icons are base64).

**Files likely touched**

`src/contentScripts/notesheetRenderer.ts` (plus optionally a new
`src/contentScripts/iconSetGlyphs.ts`), `generate_excel_samples.py`, new fixture
plus README row, new `tests/m19IconSetHtml.test.ts`, `scripts/pge/eval-screenshot.js`,
`README.md`, `BACKLOG.md`, `CHANGELOG.md`.

---

### feature-7-m19-table-parts-canonicalized-on-import

_Backlog E1. Reframed by research: multi-table workbooks already import. The failure depends on the form of the table relationship Target, not on how many tables there are._

**Spec**

A workbook whose table parts are wired up in valid but unusual OOXML forms
imports cleanly, and each table lands on the right sheet with its style. M18's
`normalizeAbsoluteRelTargets` (`src/xlsx.ts:785`, called ~L2205) fixes only
leading-`/` targets. exceljs 4.4.0 keys loaded tables by the literal
`../tables/${name}.xml` (`lib/xlsx/xlsx.js:166`) and only loads parts named
`xl/tables/table\d+.xml` (`:391`). It then looks up
`options.tables[rel.Target]` (`lib/xlsx/xform/sheet/worksheet-xform.js:520-522`),
gets `undefined`, and crashes in `lib/doc/worksheet.js:917-920`. A dangling
`rId` also crashes at `:521`. Upstream exceljs #1468 has been open since 2020
with no fix.

**Fix (Option A, operator-recommended).** Replace `normalizeAbsoluteRelTargets`
with `canonicalizeTableParts(zip)`. For each `xl/worksheets/_rels/*.rels` table
relationship:

1. Resolve its Target to a package path. Handle `/abs`, `../`, `./`, and bare relative forms.
2. If the part is missing, drop the `<tablePart r:id>` from the sheet XML and drop the relationship.
3. If the part isn't named `xl/tables/table\d+.xml`, rename it to an unused `tableK.xml` and fix its `[Content_Types].xml` `<Override PartName>`. Also move its own `_rels` part if present.
4. Rewrite the Target to `../tables/tableK.xml`.

Non-table absolute targets that the old normalizer fixed (anything
`m18AbsoluteRelTargets.test.ts` covers) must still be normalized. Either keep
that generic pass inside the new function or run it first.

Feed the **canonical** buffer to both exceljs and our own
`readTablesFromXlsxZip` (`src/xlsx.ts:1656`, called ~L2251; today it reads the
original `buffer`).

Also fix our reader's sheet-keying bug. It keys by the worksheet **filename**
number, but `src/xlsx.ts` ~L2407 looks up by `ws.id` (the `workbook.xml`
`sheetId`), and the two diverge on edited workbooks (M18 A1 note). Resolve the
filename to the sheetId with the existing `buildFilenameToSheetId`
(`src/drawings/sheetIdResolver.ts`). Replace the `xlsx-multi-table-unsupported`
classification (~L2237–2243) with a residual code that names the real cause
(suggested `xlsx-table-part-unresolved`, message "A table in this .xlsx
references a part Notesheet couldn't resolve…"), and fix the stale comments
(~L2214–2218, ~L2183).

**Acceptance criteria**

1. **New fixtures**, built by zip surgery from an existing two-table workbook
   (`FormulasAndStructuredRefs.xlsx` or `LargeWorkbook.xlsx`) with a committed
   generator script (suggested `tests/util/m19BuildTablePartVariants.ts`, or
   checked-in `.xlsx` files under `tests/fixtures/formatting-testdata/` with a
   README row each):
    - (a) **renamed part**: `xl/tables/Sales.xml` with matching Target and Content_Types Override.
    - (b) **`./` target**: `Target="./../tables/table2.xml"` or another `./`-prefixed form valid relative to the owner dir.
    - (c) **dangling rId**: a `<tablePart r:id="rId9">` with no matching relationship, and separately a relationship whose Target part is missing.
    - (d) **sheetId ≠ filename**: `sheet1.xml` ↔ `sheetId="3"`, `sheet2.xml` ↔ `sheetId="2"`, each with its own table.
2. **Jest written failing first**, new `tests/m19TablePartCanonicalize.test.ts`.
   On `main`, each of (a)–(d) either throws from `xlsxBufferToSnapshot` or puts
   a table on the wrong sheet. After the fix, every variant imports without
   throwing. The snapshot table resource lists each table under the subUnit
   whose sheet name matches the **source** `workbook.xml` sheet that owns it.
   Each table's `ref` / name equals the source table XML's `ref` / `name`,
   parsed independently from the fixture zip. For (c) the dangling table is
   absent and the other tables are present.
3. **Jest regression.** `tests/m18AbsoluteRelTargets.test.ts` and
   `tests/m12ImportRecovery.test.ts` pass with **no assertion edits**. The
   zip-transform "no nodebuffer" browser-safety guard in m18AbsoluteRelTargets
   covers the new function too, and it must use only browser-safe JSZip output
   types.
4. **Live webview import** (`scripts/pge/eval-import.js`) of fixtures (a) and
   (d) reaches "Imported." Screenshot
   `screenshots/feature-7-m19-table-parts-canonicalized-on-import/eval-*.png`
   shows the second sheet's table with its header band styled, on the sheet
   whose tab name matches the source. The DevTools console has no
   `Cannot read properties of undefined (reading 'name')`.
5. **Export file-content check** (`scripts/pge/eval-export.js`) on the imported
   fixture (d). The exported `.xlsx` contains one `xl/tables/table*.xml` per
   source table. Each one's `<table ref>` equals the source table's `ref`, and
   each is referenced from the worksheet whose `workbook.xml` sheet name matches
   the source owner.
6. **Docs.** `BACKLOG.md:52` and the README milestone row text (~L143,
   "multi-table workbooks") are reworded to the real diagnosis. Wording changes
   to the milestone row are the operator's call, so propose the text in the PR.
   The README "Known gaps" import section is updated. The remaining residual is
   unresolvable/corrupt table parts, which now skip the table instead of
   failing the import.

**Acceptance checks (black-box)**

Setup: fixtures (a)–(d) from criterion 1. Import each through the Notesheet
import UI in the running Joplin. "Styled header band" = the table's header row
shows its table-style fill.

- **AC1** (crit 2, 4): (a) renamed table part / when imported / then the import finishes with "Imported.", and both tables appear, each with a styled header band, on the sheet tab that owns it in the source.
- **AC2** (crit 2): (b) `./`-prefixed target / when imported / then the import succeeds and the table is present on its sheet.
- **AC3** (crit 2): (c) dangling `rId`, and separately a relationship whose part is missing / when imported / then the import succeeds and no error dialog appears. The broken table is absent, and the remaining tables are present.
- **AC4** (crit 2, 4): (d) sheetId ≠ filename / when imported / then each table appears on the tab whose **name** matches its source sheet, not on the tab with the same position or number.
- **AC5** (crit 5, HARD RULE): given (d) imported / when the user exports it to `.xlsx` (`scripts/pge/eval-export.js`) / then the exported file contains one `xl/tables/table*.xml` per source table, and each `<table ref>` equals the source `ref`. Each is referenced from the worksheet whose `workbook.xml` sheet name matches the source owner, and `[Content_Types].xml` has an Override for every table part.
- **AC6** (crit 5): given the exported file from AC5 / when the user imports it again / then the same tables appear on the same named tabs.
- **AC7** (crit 4): during AC1–AC6 the DevTools console has no `Cannot read properties of undefined (reading 'name')` and no `[Notesheet]` import error.
- **AC8** (crit 3, risk): `FormulasAndStructuredRefs.xlsx`, `LargeWorkbook.xlsx`, `MultiSheet.xlsx` and a workbook with no tables (`NumberFormats.xlsx`) / when imported / then the result looks as on `main`: same tabs, same tables, same header bands.
- **AC9** (risk): the full fixture sweep (`scripts/pge/bulk-import-check.cjs` over `tests/fixtures/`) reports zero import failures.
- **AC10** (crit 6): `BACKLOG.md` and README Known gaps describe the real cause (table-part relationship form), not "multi-table workbooks". The proposed milestone-row wording is in the PR, not committed (file-content check).

**Out of scope**

- Worksheets themselves not named `sheetN.xml`. The reader regex still assumes that. If it's found in the wild, note it as a follow-up.
- Pivot tables, query tables, and table `extLst` slicers.
- Fixing exceljs upstream (#1468). We work around it.

**Suggested fixture(s)**

The new variants (a)–(d). Existing `FormulasAndStructuredRefs.xlsx` and
`LargeWorkbook.xlsx` as regressions.

**Related risks**

- Every import passes through this function. A bug here breaks **all** imports. The full fixture sweep (`scripts/pge/bulk-import-check.cjs`) must stay green, and its output is attached as evidence.
- Chart / image / shape importers already re-read the original buffer and resolve sheetIds themselves. Don't change their inputs in this feature. Only exceljs and `readTablesFromXlsxZip` switch to the canonical buffer.
- The `[Content_Types].xml` Override has to stay consistent, or Excel will refuse the **exported** file. Criterion 5 guards this.

**Files likely touched**

`src/xlsx.ts`, possibly `src/drawings/sheetIdResolver.ts` (re-use only), new
fixtures plus README rows, new `tests/m19TablePartCanonicalize.test.ts`,
`tests/m12ImportRecovery.test.ts` (comments only, if the code is renamed),
`scripts/pge/eval-screenshot.js`, `README.md`, `BACKLOG.md`, `CHANGELOG.md`.

---

### feature-8-m19-theme-tinted-borders-roundtrip

_Backlog E2._

**Spec (planner survey of the current behaviour)**

- **Import** resolves `{theme, tint}` border colours against the **source** workbook's palette. `readThemeClrScheme` reads `xl/theme/theme*.xml`, and `resolveExceljsColor` + `applyOoxmlTint` turn the reference into RGB (`src/xlsx.ts` ~L350–375, ~L903–960, ~L1584–1655).
- The snapshot stores **only the resolved RGB** (`bd.<side>.cl.rgb`).
- **Export** (`applyStyleToCell`, ~L3126–3146) writes `{ argb }`. The theme reference is lost: the exported `styles.xml` carries `<color rgb="FF95B3D7"/>` where the source had `<color theme="4" tint="0.3999…"/>`. Changing the theme in Excel afterwards no longer recolours the border, and the exported file doesn't match the source XML.
- The only theme-tinted-border fixture (`BordersAndCellColors.xlsx` row 15) uses the Office-2007 palette, which is identical to exceljs's built-in default. So a palette mix-up can't be seen today. The `KNOWN SHORTCOMING` test at `tests/m12FixtureRoundTrip.test.ts:194` only asserts `/^#[0-9A-F]{6}$/`.

**Change.** Round-trip the theme reference. On import, record each themed
border side's `{theme, tint}` together with the RGB it resolved to. Suggested
mechanism: a side resource, e.g. `SHEET_NOTESHEET_BORDER_THEME_PLUGIN`
`{ [subUnitId]: { "r:c:side": {theme, tint, rgb} } }`, the same "keep original,
re-emit if unchanged" pattern as M18 A3's EMU anchors. Don't use an extra field
inside Univer's style object, because Univer may normalise it away; the
generator verifies whichever choice is made. On export, if the cell's current
side RGB still equals the recorded resolved RGB, emit
`color: { theme, tint }`. If the user changed the colour in the editor, emit
`argb` as today. The in-editor rendering keeps using the source-palette RGB,
which is already correct.

**Acceptance criteria**

1. **New fixture** `tests/fixtures/formatting-testdata/ThemeTintedBorders-Aptos.xlsx`.
   Its theme differs from exceljs's default: copy the Aptos `theme1.xml` from
   `FormattingSmorgasboard.xlsx` (accent1 `#156082`). It has borders using
   `<color theme="4" tint="0.39997558519241921"/>`, `theme="5" tint="-0.249977111117893"`,
   and `theme="1"` with no tint. Build it by zip surgery or openpyxl, with a
   committed script, and add a README row.
2. **Jest import fidelity (anchor = fixture XML + its own theme1.xml).** New
   `tests/m19ThemeTintedBorders.test.ts`. Expected RGB is computed **in the
   test** from the fixture's `theme1.xml` `<a:clrScheme>` and the ECMA-376
   Part 1 §18.8.19 / §20.1.2.3.33 HSL-luminance tint formula, written
   independently in the test. It does **not** import `applyOoxmlTint`. The
   snapshot border RGB matches within Δ ≤ 1 per channel. If this already passes
   on `main`, record it as a regression pin. The failing-first test for this
   feature is criterion 3.
3. **Jest export fidelity, written failing first.** `snapshotToXlsxBuffer` of
   the imported fixture produces `xl/styles.xml` whose border entry for each
   themed cell carries `theme="4"` and a tint equal to the source's to 1e-6,
   `theme="5"` with the negative tint, and `theme="1"`, with no `rgb`
   attribute. Map cell → `s` → `borderId` in both source and export XML. The
   exported `xl/theme/theme1.xml` `<a:clrScheme>` equals the source (already
   handled by `patchThemeClrScheme`; asserted here as a guard).
4. **Jest edit invalidation.** After setting one themed side's RGB to `#FF0000`
   in the snapshot (simulating an editor change), the export writes `rgb="FFFF0000"`
   for that side and keeps theme refs on the untouched sides.
5. **Jest re-import.** Re-importing the exported file yields the same border
   RGBs as the first import (bidirectional, M17 pattern).
6. **Live export file-content check (HARD RULE)** via `scripts/pge/eval-export.js`
   on the fixture imported into the running Joplin. The **exported artifact's**
   `xl/styles.xml` contains the `theme="4" tint="0.39997…"` border colour for
   the themed cell. Attach the eval-export JSON.
7. **Live screenshot** `screenshots/feature-8-m19-theme-tinted-borders-roundtrip/eval-*.png`
   of the editor canvas for the fixture. The themed-border cells show the Aptos
   accent-derived hue (a teal-blue family from `#156082`), not the Office-2007
   `#4F81BD` family. Pixel-sample the border strip against the test-computed RGB
   within Δ ≤ 23, since Univer's canvas anti-aliases colours by Δ14–23 (M16
   note).
8. **Pin flip and docs.** `m12FixtureRoundTrip.test.ts:194` `KNOWN SHORTCOMING`
   is flipped per the allowed-edit table. The README "Known gaps → .xlsx import
   / export" theme-tinted-borders bullet is removed or reduced to what's left
   (edited colours export as RGB). `BACKLOG.md` E2 is updated.

**Scope extension — font and fill colours (operator, 2026-10-14: current
milestone).** Font colours (`ff.cl`) and cell fills (`bg`) lose their
`{theme, tint}` reference on export in exactly the same way. It's an existing
bug with the same fix, so it ships in this feature. Use the same mechanism: the
side resource records `{theme, tint, rgb}` per `r:c:<kind>`, where `<kind>` is
a border side, `font` or `fill`. Each one re-emits the theme reference only
while its RGB is unchanged. Fixing export also fixes re-import: an exported
file then resolves its colours against its own theme again.

9. **Jest export fidelity for font + fill, written failing first.** Extend
   `ThemeTintedBorders-Aptos.xlsx` (or add a sibling fixture, and say which)
   with a theme-coloured font (e.g. `theme="1" tint="0.249977111117893"`) and a
   theme-coloured solid fill (e.g. `theme="4" tint="0.7999816888943144"`). The
   exported `xl/styles.xml` `<font><color theme tint/>` and
   `<fill><patternFill><fgColor theme tint/>` match the source to 1e-6, with
   no `rgb`. Expected values come from the source XML, the same anchor rule as
   criterion 3.
10. **Edit invalidation + re-import for font + fill.** These are criteria 4 and
    5 applied to font and fill: an edited colour exports as `rgb`, untouched
    ones keep their theme reference, and re-importing gives the same RGBs.
11. **Live check.** The criterion-6 export file-content check also asserts the
    font and fill theme references. The criterion-7 screenshot shows the fill
    in the Aptos accent family.

**Acceptance checks (black-box)**

Setup: import `ThemeTintedBorders-Aptos.xlsx` (authored in this feature; the
Aptos theme has accent1 `#156082`) in the running Joplin. Export through
`scripts/pge/eval-export.js` and inspect the exported `.xlsx`.

- **AC1** (crit 7): given the fixture open in the editor / then the theme-4-tinted border cells show a teal-blue colour derived from `#156082`, not the Office-2007 `#4F81BD` blue.
- **AC2** (crit 11): then the theme-coloured fill shows the Aptos accent family, and the theme-coloured font shows the colour its theme reference resolves to under the fixture's own theme.
- **AC3** (crit 3, 6, 9, 11, HARD RULE): given the fixture imported and **not edited** / when exported / then the exported `xl/styles.xml` has, for those cells: `theme="4"` with tint `0.39997…`; `theme="5"` with the negative tint; `theme="1"` with no tint; the theme font colour and the theme fill `fgColor` matching the source to 1e-6. None of these colours has an `rgb` attribute.
- **AC4** (crit 3): same export / then the exported `xl/theme/theme1.xml` colour scheme equals the source's.
- **AC5** (crit 4): given one side of a themed border recoloured to red in the editor's border tool / when exported / then that side is `rgb="FFFF0000"`, and the other untouched sides keep their theme references.
- **AC6** (crit 10): given the themed font recoloured and the themed fill recoloured in the editor / when exported / then those two colours are `rgb`, and untouched themed colours keep theme references.
- **AC7** (risk): given a row inserted above the themed cells in the editor / when exported / then every exported theme reference sits on a cell whose on-screen colour it reproduces. A moved cell either keeps the correct reference or falls back to `rgb`. No cell ever gets a different colour.
- **AC8** (crit 5, 10): given the AC3 export / when it is imported again / then the borders, font and fill look the same as after the first import.
- **AC9** (repeated action): given the fixture imported / when the note is closed, reopened (and Joplin restarted once), then exported / then AC3 still holds (the theme references survive saving the note).
- **AC10** (risk): given `FormulasAndStructuredRefs.xlsx` (table-style borders) / when exported / then its table-cell borders are written as on `main`, with no new theme references.
- **AC11** (crit 8, regression): `BordersAndCellColors.xlsx` row 15 looks the same as on `main` in the editor, and its export carries the source theme reference.
- **AC12** (crit 8): README "Known gaps → .xlsx import / export" no longer lists theme-tinted borders as lost. Only the residual remains (edited colours export as RGB; workbooks without a theme part). `BACKLOG.md` E2 is updated (file-content check).

**Out of scope**

- Workbooks with no theme part (the palette is null). Themed borders stay dropped as today. Document it.
- Live re-theming inside Joplin (there is no UI for changing the theme).
- Table-style synthesized borders (`synthesizeTableStyleAssignments`). They're synthesized, not source cell borders, and must not get theme refs.

**Suggested fixture(s)**

New `ThemeTintedBorders-Aptos.xlsx`. `BordersAndCellColors.xlsx` row 15 as a
regression check.

**Related risks**

- The synth-styles sidecar (`readSynthStylesSidecar`, ~L3193) and the `skip` set in `applyStyleToCell` (~L3136). Theme-ref emit must not apply to synthesized table-style borders.
- Univer re-keys style IDs on save. Don't key the side resource on style ID. Key it on cell position plus side, and guard with the RGB-equality check.
- Row/column insert or delete in the editor moves cells. A stale position key then hits a different cell. The RGB-equality guard makes the worst case "emit argb", never a wrong colour. State this in a code comment and cover it with a Jest case (insert a row above, then export: no wrong theme refs).

**Files likely touched**

`src/xlsx.ts`, new fixture plus builder script plus README row, new
`tests/m19ThemeTintedBorders.test.ts`, `tests/m12FixtureRoundTrip.test.ts`
(allowed flip only), `scripts/pge/eval-screenshot.js`, `README.md`,
`BACKLOG.md`, `CHANGELOG.md`.

---

### feature-10-m19-new-spreadsheet-shortcut-conflict

_Operator decision 2026-10-15. Spotted during feature-2 (`PROGRESS.md` Notes,
"Out-of-scope bug spotted"). Builds **before** feature-9 (see "Build order")._

**Spec**

The New Spreadsheet command (`newSpreadsheet`, Tools menu item
`notesheetNewMenu` in `src/index.ts`) currently ships the default accelerator
`CmdOrCtrl+Shift+S`. That is also Joplin's default for `focusElementSideBar`
("Focus sidebar"): `laurent22/joplin` `packages/lib/services/KeymapService.ts`
L45 `Shift+Cmd+S` (macOS table) and L100 `Ctrl+Shift+S` (Windows/Linux table)
at tag `v3.7.21` (`e41516e6`). Two app-menu items on one accelerator have
undefined behaviour, so one of the two commands silently loses its shortcut.
This feature gives New Spreadsheet a new default shortcut that is verified free,
so pressing it reliably creates a spreadsheet note and Joplin's own Focus-sidebar
shortcut works again. Tools → New Spreadsheet and both toolbar buttons stay as
they are. This is a user-visible shortcut change, so it is documented in README
and CHANGELOG.

**Choosing the accelerator (required before coding).** A candidate passes only
if **all** of these hold, each cited with a file path and commit SHA or doc URL:

1. Unbound in **both** of Joplin's default keymap tables (`KeymapService.ts`, darwin and default).
2. Not hard-coded in `packages/app-desktop/gui/MenuBar.tsx` (e.g. the `CommandOrControl+0/Plus/=/-` zoom items) or in an Electron role menu Joplin installs.
3. Not bound by any Univer shortcut in the installed `@univerjs/*` packages, so the key still reaches Joplin's menu when the grid or cell editor has focus. Pin this in Jest, the same way feature-2 pins `CmdOrCtrl+Alt+K`.
4. Not a macOS system-reserved shortcut and not a common Windows/Linux desktop-environment shortcut.
5. Not `CmdOrCtrl+Alt+K` (feature-2) or any other Notesheet accelerator.

Candidates, in order. The planner checked only rule 1 at `v3.7.21`, and both
are free in both tables:

1. `CmdOrCtrl+Shift+Alt+S`
2. `CmdOrCtrl+Shift+Alt+N`

The generator picks the first candidate that passes rules 1–5 and records why.
If neither passes, it **stops** and reports the alternatives it checked, and the
operator picks.

**Acceptance criteria**

1. **Verification recorded.** The PR body and `PROGRESS.md` Notes cite rules 1–5 for the chosen accelerator, and for any rejected candidate say which rule it failed.
2. **Jest, written failing first** (new `tests/m19NewSpreadsheetShortcut.test.ts`, or extend `tests/m19InsertLinkShortcut.test.ts`):
    - the `newSpreadsheet` menu item is created with the chosen accelerator;
    - no Notesheet menu item uses `CmdOrCtrl+Shift+S`;
    - the chosen accelerator is absent from a pinned copy of Joplin's two default keymap tables (citation inline) and from every installed Univer binding;
    - it differs from `CmdOrCtrl+Alt+K`.
3. **Live, new shortcut works.** A real OS keystroke of the new shortcut creates and opens a new spreadsheet note. It works with focus in a markdown note, in the note list, and in a Notesheet grid.
4. **Live, conflict gone.** Shift+Cmd+S (macOS) moves focus to the sidebar, which is Joplin's default, and creates **no** note.
5. **Live, entry points intact.** Tools → New Spreadsheet shows the new shortcut and creates a note. The editor-toolbar and note-toolbar New Spreadsheet buttons still create one.
6. **Docs (file-content).**
    - README (the install paragraph ~L9 and the feature paragraph ~L15) names the new shortcut and no longer says `Cmd/Ctrl+Shift+S` for New Spreadsheet.
    - The CHANGELOG **Unreleased** section has a "Changed" entry. It names the old and new shortcut, gives the reason (clash with Joplin's Focus sidebar), and explains how to rebind in Keyboard Shortcuts.
    - Released CHANGELOG entries (e.g. the `Cmd/Ctrl+Shift+S` mention ~L25) are history and are **not** edited.

**Acceptance checks (black-box)**

"New shortcut" = the accelerator chosen above (Mod variant per platform).
"Creates a spreadsheet note" = exactly one new note appears in the note list and
opens in the spreadsheet editor.

- **AC1** (crit 3): given a markdown note open with the cursor in the Markdown editor / when the user presses the new shortcut / then a spreadsheet note is created and opened.
- **AC2** (crit 3): given focus in the note list / when the new shortcut is pressed / then a spreadsheet note is created.
- **AC3** (crit 3): given a Notesheet note open with the grid focused / when the new shortcut is pressed / then a spreadsheet note is created, and the original sheet's contents are unchanged when it's reopened.
- **AC4** (crit 3): given the user has typed `abc` into B2 of a Notesheet and is still editing / when the new shortcut is pressed / then `abc` isn't lost: reopening the original note shows `abc` in B2. It's acceptable for the shortcut to do nothing while a cell is being edited. Whichever happens, the generator records which.
- **AC5** (crit 4): given a markdown note open / when the user presses Shift+Cmd+S (macOS) / then keyboard focus moves to the sidebar, which is Joplin's Focus-sidebar behaviour, and no new note is created.
- **AC6** (crit 3, repeated action): given focus in the note list / when the new shortcut is pressed twice, a second or so apart / then exactly two new spreadsheet notes exist, one per press, with no duplicates.
- **AC7** (crit 5): when the user opens the Tools menu / then "New Spreadsheet" shows the new shortcut. Clicking it creates a spreadsheet note. Clicking the New Spreadsheet button on the editor toolbar and on the note toolbar each creates one too.
- **AC8** (crit 5): when the user opens the Keyboard Shortcuts settings / then "New Spreadsheet" is listed with the new shortcut and can be rebound there.
- **AC9** (crit 3, unchanged neighbours): Joplin's nearby defaults still do their own thing and create no spreadsheet note: Cmd+S (sync), Option+Cmd+S (toggle sidebar) on macOS, Shift+Cmd+L (focus note list). In a Notesheet, Mod+Alt+K still opens the link panel (feature-2).
- **AC10** (crit 3, 4): Windows/Linux variants of AC1, AC2 and AC5 (Ctrl / Ctrl+Shift+S). `NOT RUN (platform unavailable)` is allowed per the header conventions.
- **AC11** (crit 6): README no longer says `Cmd/Ctrl+Shift+S` for New Spreadsheet and names the new shortcut. CHANGELOG Unreleased has the "Changed" entry naming old → new. The released entry ~L25 is byte-identical to `main`.

**Out of scope**

- Migrating users who already rebound New Spreadsheet in Keyboard Shortcuts. Joplin keeps user overrides; the generator confirms this in research and notes it in the CHANGELOG entry if relevant.
- Changing or removing Joplin's `focusElementSideBar` binding.
- Any other Notesheet command or shortcut.

**Suggested fixture(s)**

None needed. Use an existing markdown note, an empty notesheet
(`scripts/pge/create-empty-notesheet.sh`), and the note list.

**Related risks**

- A chord with Alt/Option can type a special character into a focused text input on macOS. Confirm the menu accelerator wins and nothing is typed (AC1, AC4).
- Univer could add a binding for the chosen chord in 1.x (BACKLOG F5). The Jest pin from criterion 2 catches that on upgrade.
- Leftover evaluator or generator notes created by AC1–AC7 clutter the dev profile. Delete them after the run (dev profile only).

**Files likely touched**

`src/index.ts`, new or extended `tests/m19*.test.ts`, `README.md`, `CHANGELOG.md`
(Unreleased only), optionally `scripts/pge/` for a keystroke probe.

---

### feature-9-m19-closeout-docs-and-regression

**Spec**

_Runs last, after feature-10 (see "Build order" at the top)._

With all M19 features merged on the branch, the docs say what the plugin now
does, and the full suite and the live smoke sweep are green. This feature adds
no new behaviour. It is the gate that the README, BACKLOG, CHANGELOG, and test
inventory agree with the shipped code.

**Acceptance criteria**

1. **File-content checks on docs.**
    - README "Known gaps" no longer lists anything fixed in features 1–8 or 10. The remaining residuals are listed exactly once:
        - Mod+K with focus outside the Notesheet editor still goes to Joplin, plus Windows/Linux live-verification status if it wasn't done (feature-2).
        - Unsupported icon sets.
        - `*X` non-space fill.
        - `[$…]` currency tags.
        - Edited colours export as RGB, and workbooks without a theme part drop themed colours (feature-8, which now covers borders, fonts and fills).
        - Unresolvable table parts.
        - Any residual feature-2 or feature-10 reported and the operator accepted.

        _(Updated 2026-10-15: the earlier list said "Cmd/Ctrl+K owned by Joplin" and "theme refs on fonts/fills". Both are now covered by feature-2 Part B and feature-8's scope extension.)_

    - `BACKLOG.md` groups D–E reflect done or residual status.
    - CHANGELOG has an M19 entry.
    - The README milestone row is **not** flipped to ✅ by the generator. Proposing that is fine; the operator decides.

2. **Every pin flip accounted for.** `git diff main -- tests/` shows only the
   allowed edits in the table at the top of this plan, plus new files.
3. **Gate.** `npm run typecheck && npm run lint && npm test` all pass. The test
   count is ≥ the `main` count plus the new m19 tests (nothing deleted except the
   pins the allowed-edit table permits).
4. **Live sweep.** `scripts/pge/bulk-import-check.cjs` over every fixture under
   `tests/fixtures/` reports zero import failures. Attach the output. A
   preview-pane screenshot of `FormattingSmorgasboard.xlsx` and of
   `ConditionalFormatting-Variants.xlsx` shows no regression against their M18
   evaluator captures (same table bands, CF fills, and chart SVGs where present).
5. **No employer / company-specific references** anywhere in the diff (public
   repo): no corporate hostnames, internal tool names, work aliases, or work
   e-mail addresses. Commits use the repo-local noreply identity.

**Acceptance checks (black-box)**

- **AC1** (crit 1): README "Known gaps" lists nothing fixed by features 1–8 or 10. Each residual from criterion 1 appears exactly once (file-content check).
- **AC2** (crit 1): `BACKLOG.md` groups D–E show done or residual status for D2–D5 and E1–E4. CHANGELOG has an M19 entry that mentions Mod+K in the editor and the New Spreadsheet shortcut change.
- **AC3** (crit 1): the README milestone row for M19 is **not** marked ✅ by the generator.
- **AC4** (crit 2): `git diff main -- tests/` shows only new files plus the edits in the allowed-edit table.
- **AC5** (crit 3): `npm run typecheck`, `npm run lint` and `npm test` pass on the commit under test. The test count is ≥ `main`'s count plus the new m19 tests.
- **AC6** (crit 4): `scripts/pge/bulk-import-check.cjs` over every fixture under `tests/fixtures/` reports zero failures.
- **AC7** (crit 4): preview-pane captures of `FormattingSmorgasboard.xlsx` and `ConditionalFormatting-Variants.xlsx` show the same table bands, CF fills and charts as their M18 evaluator captures, plus the M19 additions (icon glyphs, rich runs) where expected.
- **AC8** (crit 1): with a fresh deploy, the Tools menu shows "New Spreadsheet" (with feature-10's shortcut), "Import .xlsx as Notesheet", and "Insert link in spreadsheet" (Mod+Alt+K), and each one works once.
- **AC9** (crit 5): `git log main..HEAD` and `git diff main` contain no employer or company-specific hostnames, tool names, aliases or work e-mail addresses, and every commit author e-mail is the repo-local noreply address.

**Out of scope**

- Release tagging and npm publishing (operator's call, separate PR).

**Related risks**

- Doc drift between README "Known gaps" and the `KNOWN SHORTCOMING` Jest pins.

**Files likely touched**

`README.md`, `BACKLOG.md`, `CHANGELOG.md`, `PROGRESS.md`.

---

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file and its predecessors in git history as templates. Each feature
gets a stable `### feature-N-<kebab-id>` heading, a one-paragraph **Spec**, and
**Acceptance criteria** phrased as observable evidence. The criteria are
anchored to external ground truth: source XML, a cited spec, or a real-Excel
reference PNG. They are never anchored to our own emit, and any export outcome
gets a file-content check on the exported artifact. Each feature also gets
**Out of scope**, **Suggested fixture(s)**, and **Related risks**, plus the
allowed-edit table for existing tests. Since 2026-10-15 every feature also gets
**Acceptance checks (black-box)**: given/when/then in user-visible terms, with
stable `ACn` ids mapped to criteria and the boundaries listed explicitly. The
generator never edits them. The generator works on the first feature with
`passes: false` in the plan's stated build order. Without a stated order, that
is the lowest-numbered one.
