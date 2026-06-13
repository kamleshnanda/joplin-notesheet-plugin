# Contributing to Notesheet

This is the **single source of truth** for how we build, test, and maintain
Notesheet. It applies to **everyone** — human contributors and AI coding
agents alike. If you are an AI agent, also read [`AGENTS.md`](./AGENTS.md)
for the agent-specific operational layer; it does **not** repeat anything
here, so this file still applies to you in full.

> **No-duplication rule.** Every guideline lives in exactly one place.
> Shared rules (this file) are never restated in `AGENTS.md` or
> `.claude/CLAUDE.md` — those reference this file. If you find a rule
> duplicated, collapse it to one home and link the other.

## Getting started

```bash
npm install        # install deps (also runs the build via `prepare`)
npm run dist       # build the .jpl into publish/
npm test           # Jest unit tests
npm run typecheck  # tsc --noEmit (no emit, type errors only)
npm run lint       # ESLint
npm run format     # Prettier write
```

A passing local gate before you commit is: `npm run typecheck && npm run lint && npm test`.
The pre-commit hook runs lint + typecheck on staged files automatically;
CI runs the full set on every PR (see [Enforcement](#enforcement)).

- **Node 20.9+** (CI matrices 20.x and 22.x). Node 18 is EOL.
- **Desktop Joplin 3.5+** is the only supported host.

## Code quality

- **TypeScript must typecheck clean** (`npm run typecheck`, zero errors).
  This is enforced in CI and pre-commit. A build that webpack-bundles but
  has type errors is not acceptable.
- **ESLint + Prettier** are authoritative for style; don't hand-fight the
  formatter. Run `npm run format` before committing.
- **Match the surrounding code.** Comment density, naming, and idiom
  should look like the file you're editing, not like a different project.
- **Prefer flat / immutable / sparse-dict data shapes over mutable
  object-model mutation** where there's a choice. Flat shapes make the
  diff between two states trivially inspectable (a JSON diff); live-mutating
  object graphs hide updates in getter/setter chains. This bit us once with
  exceljs's `BorderStyle` type drift — minimal-to-no side effects is the goal.
- **No dead code or orphaned files.** If you add a source file, it must be
  imported and committed (we once shipped an uncommitted `trackedCharts.ts`
  that production imported). If you remove a feature, remove its dead tests.

## Testing

- **Every behavioral change ships with a test.** Jest is the baseline gate.
- **Fidelity tests anchor to the EXTERNAL ground truth, never to our own
  output.** For "match Excel" features, assert against an artifact captured
  _from Excel_ (a reference `.xlsx`'s own XML, or a real-Excel screenshot
  under `screenshots/excel-reference/` sampled by a pure-stdlib PNG reader
  within RGB Δ ≤ 8) — **not** against what our code happens to emit.
  Anchoring a test to our emit means the test passes even when the output
  is wrong; this shipped wrong colours through a green suite twice.
- **Write the fidelity test failing first, then make it pass.** Commit
  ordering proves the gap was real, not retro-fitted to existing output.
- **Don't keep two copies of the same assertion alive.** A pin-down derived
  from a catalog should cite the canonical fidelity test as its source;
  duplicated assertions drift apart silently.
- **`// approximation` / `// matches Excel within a few units` comments are
  claims, not facts** — treat them as TODOs until a test verifies them.
- **Test fixtures** live under `tests/fixtures/` — `tests/fixtures/charts/`
  for chart workbooks, `tests/fixtures/formatting-testdata/` for formatting
  workbooks. Keep new reference `.xlsx` files there.

## Rendered notesheet HTML must survive the Rich Text editor

The Markdown-It content script (`src/contentScripts/notesheetRenderer.ts`)
renders the `notesheet v=1` fence into an HTML `<table>`. That output is
**load-bearing for data integrity** and must stay wrapped in Joplin's
`joplin-editable` / `joplin-source` convention:

- The outer element carries class `joplin-editable` (TinyMCE's
  `noneditable_class`), so the block is atomic in the Rich Text editor.
- A hidden `<pre class="joplin-source" hidden data-joplin-language="notesheet"
data-joplin-source-open="…" data-joplin-source-close="…">` carries the
  **verbatim original fence body** and precedes the visible render.

**Why this is non-negotiable:** without the wrapper, opening a Notesheet
note in Joplin's **Rich Text (TinyMCE) editor** and saving makes TinyMCE
serialize the rendered `<table>` back to a plain Markdown table —
**destroying the fence and the entire Univer snapshot** (styles, charts,
formulas). That is silent, total data loss. It was a real regression
(introduced by the M16 HTML render, fixed in M18); `tests/rteFenceRoundTrip.test.ts`
guards it.

- The source-open/close delimiters **must mirror `wrapSnapshot()` in
  `src/snapshot.ts` exactly** (` ```notesheet v=1\n<json>\n``` `), with
  newlines encoded as `&#10;` in the attributes — any drift corrupts the
  round-trip.
- There is **no Joplin API signal** to tell the Rich Text editor apart from
  the preview/PDF/HTML render at render time, so you cannot "only wrap in the
  editor" — the wrapper is the canonical approach and is harmless on the
  read-only surfaces. (Every built-in Joplin renderer — mermaid, katex,
  fountain — does the same.)
- If you change the rendered HTML structure, **keep the `joplin-editable`
  wrapper and the verbatim `joplin-source` body.** Never strip them.

## Dependency hygiene

Dependency drift has bitten this project (exceljs silently moved 4.4.0 →
3.4.0 during an unrelated `npm install`, and the resulting type error got
patched in a test instead of investigated). Hold the line:

1. **Audit every `package.json` / `package-lock.json` diff** before
   committing — `git diff package.json package-lock.json`. Surface _every_
   version change, not just the one you intended. npm's transitive
   resolution decisions are not implicitly approved.
2. **Downgrades are blocked by default.** Any `from > to` (semver) requires
   (a) an explicit documented blocking reason and (b) maintainer approval
   before it lands. "npm decided" / "to make CI pass" are never sufficient.
   The dependency-downgrade guard enforces this in CI and pre-commit.
3. **Major-version upgrades** must call out the breaking-change surface
   (Node engines, peer deps, public API) in the PR/commit message.
4. **Debug from the version, not the symptom.** When a typecheck error or
   test failure appears right after a dependency change, the first
   hypothesis is "the dependency changed in a way I didn't expect" — check
   the resolved version and its types _before_ editing app code or tests to
   make the symptom disappear.
5. **When CI fails but it builds locally,** suspect environment drift (Node
   version, OS, lockfile) before blaming the dependency. Fix the environment
   to match the dependency's real requirements; only swap the dependency if
   its requirement is itself broken.

## High standards over shortcuts

- **Fix the root cause, not the symptom.** When the choice is between fixing
  it properly and suppressing/ignoring/working-around, default to the proper
  fix. Suppressing a warning, downgrading to dodge a type error, or patching
  a test to match wrong output are all shortcuts we reject.
- **Prefer a documented known-shortcoming over an unexpected bug.** Shipping
  a series of point-fixes where each one introduces the next regression is
  worse than scoping the gap honestly, documenting it in the README's
  "Known gaps", and pinning it with a test. _"I'll happily take known
  shortcomings over disappointing users with unexpected bugs."_
- **Non-functional health is first-class delivery.** Stale dependencies,
  CVEs, memory leaks, and performance regressions are features customers
  care about over long use. They earn their own milestones (see below), not
  a footnote in a feature PR.

## Milestones, branches, and PRs

- **`main` is protected.** Never commit or push directly to `main`. Branch,
  open a PR, let CI pass, merge.
- **Every change maps to exactly one milestone** in the README "Milestones"
  table — feature, fix, refactor, or health work. The milestone is decided
  by the maintainer (current / new / future), not assumed. Health and
  large-refactor work that doesn't deliver an existing milestone's feature
  gets its **own** milestone row (precedents: M11 Dependency hygiene, M14
  SheetJS spike, M20 Codebase health).
- **List every contributing PR in its milestone row** (e.g. `#32 #33`),
  space-separated, ascending PR-number order. Nothing is left unlisted.
- **Pending scope** is tracked in [`BACKLOG.md`](./BACKLOG.md) with stable
  item IDs (A1, C3, …). Add newly-deferred work there.
- **Commit messages** describe the _why_, not just the what. Reference the
  milestone / issue where relevant.

## Enforcement

These rules are not honor-system where we can help it — the earliest
practical checkpoint enforces them:

| Rule                  | Pre-commit hook          | CI                                   |
| --------------------- | ------------------------ | ------------------------------------ |
| Type errors           | ✅ `typecheck` on staged | ✅ `npm run typecheck`               |
| Lint / format         | ✅ `lint-staged`         | ✅ `npm run lint`                    |
| Tests pass            | — (too slow for commit)  | ✅ `npm test`                        |
| Build succeeds        | —                        | ✅ `npm run dist` (Node 20.x + 22.x) |
| Dependency downgrades | ✅ guard script          | ✅ guard script                      |

If a hook or CI check blocks you, **fix the cause** — do not bypass with
`--no-verify` or by weakening the check. If a check is itself wrong, fix the
check in its own PR and say why.
