#!/usr/bin/env bash
# Headless `Import .xlsx as Notesheet`. Compiles import-fixture.ts via
# tsc to a temp dir and runs it under Node. Prints the new note ID.
#
# Why compile-on-demand? `xlsxBufferToSnapshot` lives in TS and we want
# the harness to call the SAME conversion the plugin calls at runtime,
# not a re-implementation. tsc is already a project dependency.
#
# Usage:
#   ./scripts/pge/import-fixture.sh <fixture-name> [--title <title>]
#
# Examples:
#   ./scripts/pge/import-fixture.sh MergedCellsAndAlignment.xlsx
#   ./scripts/pge/import-fixture.sh MergedCellsAndAlignment.xlsx --title "M13/C eval"
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
SCRIPT_DIR="$REPO_ROOT/scripts/pge"
TSC="$REPO_ROOT/node_modules/.bin/tsc"

if [ "${1:-}" = "" ]; then
    echo "usage: import-fixture.sh <fixture-name> [--title <title>]" >&2
    exit 2
fi

if [ ! -x "$TSC" ]; then
    echo "import-fixture: $TSC not found. Run 'npm ci' first." >&2
    exit 1
fi

# Compile to a stable temp dir (idempotent). We use --outDir + --rootDir
# so the relative `../../src/xlsx` import in the .ts file resolves to a
# JS sibling layout the runtime can load.
OUT_DIR="$(mktemp -d -t pge-fixture-XXXX)"
trap 'rm -rf "$OUT_DIR"' EXIT

# Compile the helper plus the project files it depends on.
# --skipLibCheck because we don't need to typecheck @types/* under
# node_modules just to run a CLI.
"$TSC" \
    --target ES2020 \
    --module commonjs \
    --moduleResolution node \
    --esModuleInterop \
    --resolveJsonModule \
    --skipLibCheck \
    --outDir "$OUT_DIR" \
    --rootDir "$REPO_ROOT" \
    "$SCRIPT_DIR/import-fixture.ts" \
    >&2

# Run the compiled output. The relative `require('../../src/xlsx')`
# in the .js becomes `require('../../src/xlsx')` in the OUT_DIR
# layout, which resolves to OUT_DIR/src/xlsx.js (tsc emits the
# whole dependency graph alongside). Node's resolver walks up from
# OUT_DIR for non-relative deps (exceljs, jszip, etc.); we point it
# at the repo's node_modules via NODE_PATH so they're found.
PGE_REPO_ROOT="$REPO_ROOT" NODE_PATH="$REPO_ROOT/node_modules" \
    node "$OUT_DIR/scripts/pge/import-fixture.js" "$@"
