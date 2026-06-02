#!/usr/bin/env bash
# One PGE cycle, on demand. Operator runs this:
#
#     ./scripts/pge/run-cycle.sh
#
# The cycle:
#   1. Pick the lowest-numbered passes:false feature from
#      test-results.json.
#   2. Invoke `claude` as the GENERATOR agent — it builds the
#      feature, captures evidence, flips the row to passes:true.
#   3. Invoke `claude --agent evaluator` in a SEPARATE process — it
#      captures its own screenshot via eval-screenshot.sh, opens it,
#      grades PASS or NEEDS_WORK.
#   4. If PASS: update evaluator_verdict, exit 0.
#      If NEEDS_WORK: write findings to NEXT_FINDINGS.md, revert the
#      row to passes:false, exit 1.
#
# This script does NOT auto-loop. Each invocation runs at most one
# generator session and one evaluator pass. Operator decides whether
# to re-run.
#
# IMPORTANT: this script does NOT install the .jpl or relaunch Joplin
# on its own — the generator agent does that mid-session per
# CLAUDE.md (which calls install-plugin.sh + launch-joplin.sh after
# its build). The wrapper just orchestrates the agent processes.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$REPO_ROOT"

# 1. Pick a feature.
FEATURE_ID=$(node -e '
const r = JSON.parse(require("fs").readFileSync("test-results.json","utf8"));
const id = Object.keys(r).find(k => r[k] && r[k].passes === false);
if (!id) { console.error("All features pass. Nothing to do."); process.exit(2); }
process.stdout.write(id);
')
echo "run-cycle: feature is $FEATURE_ID"

# 2. Generator session.
echo "run-cycle: invoking generator..."
claude -p \
    "You are the generator (.claude/agents/generator.md). Build feature $FEATURE_ID per BUILD_PLAN.md. Follow the proof-before-passing sequence. Do NOT skip the screenshot step." \
    || { echo "run-cycle: generator failed" >&2; exit 1; }

# 3. Evaluator session (FRESH PROCESS).
echo "run-cycle: invoking evaluator..."
VERDICT_FILE=$(mktemp -t pge-verdict.XXXX)
trap 'rm -f "$VERDICT_FILE"' EXIT
claude --agent evaluator -p \
    "Grade the most recent generator session's claim that feature $FEATURE_ID passes. Read BUILD_PLAN.md. Capture your own screenshot via 'bash scripts/pge/eval-screenshot.sh $FEATURE_ID'. Then PASS or NEEDS_WORK." \
    > "$VERDICT_FILE"

VERDICT=$(head -1 "$VERDICT_FILE" | tr -d '[:space:]')
echo "run-cycle: verdict is $VERDICT"

# 4. Process verdict.
if [ "$VERDICT" = "PASS" ]; then
    node -e '
        const fs = require("fs");
        const r = JSON.parse(fs.readFileSync("test-results.json","utf8"));
        r["'"$FEATURE_ID"'"].evaluator_verdict = "PASS";
        fs.writeFileSync("test-results.json", JSON.stringify(r, null, 2) + "\n");
    '
    rm -f NEXT_FINDINGS.md
    echo "run-cycle: PASS — feature $FEATURE_ID locked in."
    exit 0
fi

# NEEDS_WORK or unknown verdict.
echo "run-cycle: NEEDS_WORK — reverting row and writing NEXT_FINDINGS.md"
{
    echo "# Evaluator findings for $FEATURE_ID"
    echo
    echo "Generated $(date -u '+%Y-%m-%dT%H:%M:%SZ') by run-cycle.sh."
    echo
    cat "$VERDICT_FILE"
} > NEXT_FINDINGS.md

node -e '
    const fs = require("fs");
    const r = JSON.parse(fs.readFileSync("test-results.json","utf8"));
    r["'"$FEATURE_ID"'"].passes = false;
    r["'"$FEATURE_ID"'"].evaluator_verdict = "NEEDS_WORK";
    fs.writeFileSync("test-results.json", JSON.stringify(r, null, 2) + "\n");
'

exit 1
