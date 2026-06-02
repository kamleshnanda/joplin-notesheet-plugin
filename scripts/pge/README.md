# PGE harness — operator quickstart

The Notesheet planner-generator-evaluator (PGE) harness drives
long-running sessions where:

- a **planner** turns a one-line ask into `BUILD_PLAN.md` +
  `test-results.json`,
- a **generator** builds one feature per session,
- an **evaluator** runs in a fresh process, captures its OWN
  screenshot of the running Joplin app, and grades PASS or NEEDS_WORK.

The evaluator's screenshot is the gate. It catches the failure mode
where Jest tests pass on snapshot data but the Univer renderer is
broken — the M13 lesson that prompted this harness.

## One-time setup (do these once per machine)

1. **Install Playwright.**
   ```sh
   npm install --save-dev playwright
   npx playwright install chromium
   ```
   The Electron driver Playwright bundles by default; no extra
   package needed.

2. **Confirm Joplin desktop installed at the expected path.**
   ```sh
   ls /Applications/Joplin.app
   ```
   If yours is elsewhere, set `JOPLIN_BIN` in your shell profile.

3. **Set up the dev profile.** First time only:
   ```sh
   /Applications/Joplin.app/Contents/MacOS/Joplin --env dev
   ```
   This creates `~/.config/joplindev-desktop/`. Quit when it opens.

4. **Enable Web Clipper service in the dev profile.** Launch Joplin
   with `--env dev` again, go to **Tools → Options → Web Clipper**,
   click **Enable Web Clipper service**. The setting persists.
   Verify with:
   ```sh
   curl http://localhost:41184/ping
   # → "JoplinClipperServer"
   ```

5. **Quit Joplin.** The harness manages launching it from now on.

## Each cycle

```sh
# 1. Author/update the operator ask:
$EDITOR OPERATOR_ASK.md

# 2. Run the planner once to produce BUILD_PLAN.md +
#    test-results.json + initial PROGRESS.md:
claude --agent planner -p "Read OPERATOR_ASK.md and produce the build plan."

# 3. For each feature, run a cycle:
./scripts/pge/run-cycle.sh
```

Each `run-cycle.sh` invocation:
- picks the lowest-numbered `passes:false` feature,
- runs ONE generator session (it builds, installs, screenshots),
- runs ONE evaluator session in a separate process (it captures its
  OWN screenshot via Playwright),
- writes `NEXT_FINDINGS.md` if NEEDS_WORK,
- exits 0 on PASS, 1 on NEEDS_WORK.

Re-run as needed.

## Operator controls during a cycle

- **Steer**: write a one-line redirect to `STEER.md`. The next agent
  tool call surfaces it as `OPERATOR STEERING:` and clears the file.
  Use for "switch to feature-2 instead", "stop after this commit",
  or "the spec is wrong, here's what I actually meant."

- **Hard stop**: `touch AGENT_STOP`. Blocks every tool call until you
  `rm AGENT_STOP`. Use when something is going wrong and you need
  the agent to stop NOW rather than at a clean checkpoint.

- **Inspect screenshots**: every PNG under `screenshots/<feature-id>/`
  is committed to git. Generator screenshots are prefixed `gen-`,
  evaluator screenshots `eval-`. The evaluator-captured ones are
  authoritative.

## File map

| File | Owner | Purpose |
|---|---|---|
| `.claude/CLAUDE.md` | repo | Generator's runtime contract |
| `.claude/agents/planner.md` | repo | Planner brief |
| `.claude/agents/generator.md` | repo | Generator brief |
| `.claude/agents/evaluator.md` | repo | Evaluator brief |
| `.claude/hooks/*.sh` | repo | Default-FAIL evidence gate, kill switch, etc. |
| `.claude/settings.json` | repo | Hook-event wiring |
| `OPERATOR_ASK.md` | operator | One-line(ish) task for this cycle |
| `BUILD_PLAN.md` | planner | Per-feature spec + acceptance criteria |
| `PROGRESS.md` | generator | Session-to-session handoff |
| `test-results.json` | planner seeds, generator flips bits | The default-FAIL contract |
| `NEXT_FINDINGS.md` | run-cycle.sh writes after evaluator NEEDS_WORK | Generator reads on next session |
| `STEER.md` | operator (rare) | Mid-session redirect |
| `AGENT_STOP` | operator (emergency) | Hard kill switch |
| `screenshots/<feature-id>/` | gen + eval | Visual evidence (committed) |

## Why the gate works

Five mechanisms compound to catch the M13-style failure:

1. **Default-FAIL contract.** Every row starts `passes: false`. The
   generator can never claim done by inaction.
2. **Evidence-Read gate.** `verify-gate.sh` denies Write/Edit to
   `test-results.json` until evidence files have been opened. After
   the gate fires, the evidence log is wiped, so each pass needs
   fresh proof.
3. **Fresh-context evaluator.** `claude --agent evaluator -p` spawns
   a new process. None of the generator's reasoning is in its
   context.
4. **No write tools on evaluator.** It cannot launder a fix into a
   pass.
5. **Open the bytes, judge the bytes.** The evaluator captures its
   OWN screenshot via Playwright + Joplin Electron, opens the PNG
   via Read, and grades from what it actually sees.

If any of these mechanisms is absent, the gate is weaker than it
looks.
