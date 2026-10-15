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

## Cycle context

M19 closes the documented gaps in `BACKLOG.md` groups D (PDF / HTML / preview
renderer) and E (editor + `.xlsx` import). D1 (dataBar in HTML) already
shipped in M18. The operator approved all seven remaining items: D2, D3,
D4, D5, E1, E2, E3, E4. Research findings that the operator supplied are
folded into each feature's spec below.

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
2. **E4 accelerator choice.** The operator chose "alt shortcut + document".
   The plan proposes `CmdOrCtrl+Alt+K` but requires the generator to verify,
   against the upstream Joplin default keymap, that it is unbound. If it is
   bound, the generator picks the first free candidate from the list in
   feature-2 and records the citation. The choice is not open for debate after
   that. It is recorded in README and CHANGELOG.

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

---

## Features

Order: isolated, low-risk editor items first (E3, E4). Then the renderer items,
smallest first (D3 → D4 → D5 → D2). Then the two import/export items with the
widest blast radius (E1, E2). Last is a closeout. Each feature is one PR-sized
unit. feature-4 and feature-5 share a formatter and **may** ship in one PR, but
they are graded separately.

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

_Backlog E4. Operator decision: "Alt shortcut + document"._

**Spec**

A dedicated keyboard shortcut opens Univer's own insert-link panel for the
selected cell while a Notesheet note is active. Cmd/Ctrl+K stays a documented
known gap. Background from research: Joplin binds Cmd/Ctrl+K as an Electron
app-menu accelerator to `textLink` (`laurent22/joplin`
`packages/lib/services/KeymapService.ts` ~L39/L94;
`packages/app-desktop/gui/MenuBar.tsx` ~L732).
`editorCommandDeclarations.ts` enables that accelerator with
`|| hasActivePluginEditor`. At runtime Joplin finds no built-in editor and logs
`Received command, but editor is gone textLink`, so the keystroke never reaches
the plugin iframe. A plugin can't fix this. Re-registering `textLink` would
break the markdown editor, and duplicate accelerators have undefined behaviour.
The fix:

- Register a plugin command (suggested name `notesheetInsertLink`, label "Insert link in spreadsheet") with `enabledCondition: 'hasActivePluginEditor'`.
- Expose it as a `joplin.views.menuItems` entry with its own free accelerator. This follows the existing `CmdOrCtrl+Shift+S` pattern in `src/index.ts` ~L225.
- When the command runs, check `joplin.views.editors.isActive(handle)` and send `joplin.views.editors.postMessage(handle, { type: 'insertLink' })`.
- In the webview, the `onMessage` handler in `src/editorView.tsx` (~L1029) calls `univerAPI.executeCommand('sheet.operation.insert-hyper-link-toolbar')`. That command ID is verified present in `@univerjs/sheets-hyper-link-ui` 0.23 and is the one `InsertLinkShortcut` uses (~L1627).

**Accelerator verification (required before coding).** Confirm that
`CmdOrCtrl+Alt+K` is not bound in Joplin's default keymap (`KeymapService.ts`,
macOS and Windows/Linux tables) or in `MenuBar.tsx`, and cite the file and
commit SHA in the PR. If it is bound, use the first free candidate from
`CmdOrCtrl+Alt+L` → `CmdOrCtrl+Shift+Alt+K` and record why. Plugin menu items
appear in Tools → Options → Keyboard Shortcuts, so users can rebind it.

**Acceptance criteria**

1. **Jest** (new `tests/m19InsertLinkShortcut.test.ts`, using the existing `api` mocks):
    - `onStart` registers a command named `notesheetInsertLink` with `enabledCondition: 'hasActivePluginEditor'`.
    - A menu item is created with the verified accelerator string.
    - Running the command posts `{ type: 'insertLink' }` to the active editor handle.
    - It posts nothing when `isActive` returns false.
    - It never re-registers or overrides `textLink`: assert no `commands.register` call has the name `textLink`.
2. **Live, keystroke path.** With a Notesheet note open and cell B2 selected,
   the harness sends the verified accelerator as a real OS keystroke via
   `osascript` System Events. The harness already holds Accessibility permission
   for `prep-joplin-window.sh`. Screenshot
   `screenshots/feature-2-m19-insert-link-alt-shortcut/eval-*.png` shows Univer's
   hyperlink edit panel open over the grid (link text / URL inputs visible), and
   Joplin's markdown "Insert link" dialog is **not** showing. A CDP DOM probe
   records the panel's root element as present. The generator identifies a
   stable selector and documents it in `PROGRESS.md` Notes.
3. **Live, menu path.** Clicking Tools → "Insert link in spreadsheet" opens the
   same panel (probe JSON in the same folder).
4. **Live, no regression on markdown notes.** In a plain markdown note, Cmd/Ctrl+K
   still opens Joplin's own link dialog, and the new menu item is
   disabled/greyed (its `enabledCondition` is false). The screenshot or probe
   records both.
5. **Docs.**
    - README "Known gaps → Editor" Cmd/Ctrl+K bullet is rewritten. It says Cmd/Ctrl+K is owned by Joplin's `textLink` app-menu accelerator and can't be intercepted by a plugin (cite the KeymapService.ts / MenuBar.tsx / editorCommandDeclarations.ts paths).
    - It names the new shortcut.
    - It gives the workaround for users who want Cmd/Ctrl+K itself: Tools → Options → Keyboard Shortcuts, clear or rebind `textLink`, then bind "Insert link in spreadsheet" to Cmd/Ctrl+K.
    - `BACKLOG.md` E4 and CHANGELOG are updated to match.

**Out of scope**

- Making Cmd/Ctrl+K itself open Univer's panel (not possible from a plugin, per the research above).
- Editing or removing an existing link via the shortcut (Univer's panel behaviour is whatever it does).
- Joplin Mobile.

**Suggested fixture(s)**

- `tests/fixtures/formatting-testdata/Hyperlinks-Variants.xlsx`, imported, to confirm existing links still render and click through after the change.
- An empty notesheet for the insert case.

**Related risks**

- Multiple open editor views (split / second window): the command must target the **active** handle only. Track handles in a set in `src/index.ts` and pick via `isActive`.
- The `onMessage` dispatcher in `editorView.tsx` already handles `snapshot`-type messages. Add the new `type` branch without disturbing the snapshot-load path (M18 RTE fence work is adjacent).

**Files likely touched**

`src/index.ts`, `src/editorView.tsx`, new `tests/m19InsertLinkShortcut.test.ts`,
possibly `tests/__mocks__/*`, `scripts/pge/eval-screenshot.js`, `README.md`,
`BACKLOG.md`, `CHANGELOG.md`.

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

**Out of scope**

- Theme references on **font** and **fill** colours. These have the same loss on export and are a follow-up candidate; the operator decides the milestone.
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

### feature-9-m19-closeout-docs-and-regression

**Spec**

With all M19 features merged on the branch, the docs say what the plugin now
does, and the full suite and the live smoke sweep are green. This feature adds
no new behaviour. It is the gate that the README, BACKLOG, CHANGELOG, and test
inventory agree with the shipped code.

**Acceptance criteria**

1. **File-content checks on docs.**
    - README "Known gaps" no longer lists anything fixed in features 1–8. The remaining residuals are listed exactly once: Cmd/Ctrl+K owned by Joplin, unsupported icon sets, `*X` non-space fill, `[$…]` currency tags, theme refs on fonts/fills, unresolvable table parts.
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
allowed-edit table for existing tests. The generator works on the
lowest-numbered feature with `passes: false` in `test-results.json`.
