#!/usr/bin/env bash
# Bash entrypoint to eval-screenshot.js. Evaluator agents run THIS
# (Bash is in their tool list; running node directly is also fine
# but this gives a stable interface).
set -euo pipefail
REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
exec node "$REPO_ROOT/scripts/pge/eval-screenshot.js" "$@"
