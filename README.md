# Joplin Univer Spreadsheet Plugin

Embed interactive spreadsheets directly in your Joplin notes. Create, edit, and persist tabular data with a dialog-based editor that works across Markdown and Rich Text editors.

## Version: 2.3.0

## Features

- Insert spreadsheets via toolbar button, menu item, or keyboard shortcut (Ctrl/Cmd+Shift+U)
- Click-to-open dialog editor with editable cell grid
- Formula support: SUM, AVERAGE, COUNT, MIN, MAX, IF, CONCAT, ABS, ROUND, UPPER, LOWER, LEN, TRIM
- Cell references (A1, B2) and ranges (A1:C3) in formulas
- Arithmetic expressions: =A1+B1*2, =A1/B1-C1
- Comparison operators in IF: =IF(A1>10,"high","low")
- Formula bar showing active cell reference and formula
- Circular reference detection
- Data stored in Joplin's userData API — syncs across devices automatically
- Works in both Markdown Editor and Rich Text Editor (joplin-editable pattern)
- Spreadsheet links render as styled clickable spans in the note viewer
- Keyboard accessible (Enter/Space to open)
- Input validation: cell length limits, JSON size limits, grid dimension caps

## How It Works

1. **Insert**: Use the toolbar button or Ctrl/Cmd+Shift+U to insert a spreadsheet link
2. **Click**: Click the rendered spreadsheet link in the note viewer to open the editor
3. **Edit**: Modify cells directly in the table grid — type = to start a formula
4. **Save**: Click Save and Close to persist changes back to the note's userData

## Formula Examples

| Formula | Description |
|---------|-------------|
| =A1+B1 | Add two cells |
| =SUM(A1:A10) | Sum a range |
| =AVERAGE(B1:B5) | Average of a range |
| =COUNT(A1:C3) | Count numeric cells |
| =MIN(A1:A10) | Minimum value |
| =MAX(A1:A10) | Maximum value |
| =IF(A1>100,"over","under") | Conditional |
| =ROUND(A1,2) | Round to 2 decimals |
| =CONCAT(A1," ",B1) | Join text |
| =UPPER(A1) | Uppercase |

## Installation

### From .jpl File

1. Download the .jpl from the publish/ directory
2. In Joplin: Tools, Options, Plugins, Install plugin, select the .jpl file
3. Restart Joplin

### From Source

```bash
npm install
npm run dist
```

## Architecture

```
src/
  index.ts                 - Plugin entry: commands, dialog, message handlers
  contentScript.ts         - Markdown-it plugin: renders links as spans
  contentScriptTinyMCE.ts  - CodeMirror stub (no-op)
  dialogScript.js          - Dialog webview: spreadsheet table + formula engine
  univerWebviewHandler.js  - Note viewer webview: click handler
  manifest.json            - Plugin metadata
```

## Development

```bash
npm run dist
```

Requires Node.js v22.x with --openssl-legacy-provider.

## Compatibility

- Joplin 3.5+
- Markdown Editor and Rich Text Editor
- Windows, macOS, Linux

## License

MIT
