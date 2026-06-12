# Implementation Tasks: Joplin Univer Plugin v2.3.0

> Dialog-based architecture with userData storage, span-based clicks,
> joplin-editable RTE support, and client-side formula engine.

## 1. Project Setup and Configuration

- [x] 1.1 Project structure (Yeoman generator)
- [x] 1.2 Dependencies installed
- [x] 1.3 Webpack configured for content scripts
- [x] 1.4 TypeScript configuration verified
- [x] 1.5 plugin.config.json configured

## 2. Core Plugin Infrastructure

- [x] 2.1 Plugin registration in src/index.ts
    - [x] 2.1.1 Register MarkdownItPlugin content script
    - [x] 2.1.2 Register CodeMirror content script (no-op stub)
    - [x] 2.1.3 Register Insert Spreadsheet command
    - [x] 2.1.4 Register Open Spreadsheet command with dialog editor
    - [x] 2.1.5 Create toolbar button (EditorToolbar)
    - [x] 2.1.6 Create menu item (Tools, Ctrl/Cmd+Shift+U)
    - [x] 2.1.7 Message handlers for content script communication
    - [x] 2.1.8 Error handling for plugin initialization

- [x] 2.2 manifest.json metadata

## 3. Content Script (contentScript.ts)

- [x] 3.1 Intercept univer-sheet:// links
- [x] 3.2 Render as span (not a tag)
- [x] 3.3 joplin-editable + joplin-source wrapper for RTE
- [x] 3.4 CSS assets
- [x] 3.5 JS asset: univerWebviewHandler.js

## 4. Click Handling (univerWebviewHandler.js)

- [x] 4.1 Event delegation (capture phase)
- [x] 4.2 DOM walk to find .univer-link
- [x] 4.3 Extract sheet ID from data attributes
- [x] 4.4 webviewApi.postMessage()
- [x] 4.5 Keyboard accessibility
- [x] 4.6 Duplicate handler guard

## 5. Dialog Editor (dialogScript.js)

- [x] 5.1 Parse JSON from hidden div
- [x] 5.2 Render editable HTML table with form inputs
- [x] 5.3 Column headers and row numbers
- [x] 5.4 Minimum grid: 10 rows x 5 columns
- [x] 5.5 In-memory data sync on cell events
- [x] 5.6 Handle double-stringified JSON

## 6. Data Persistence (userData API)

- [x] 6.1 Store via userDataSet() on insert
- [x] 6.2 Load via userDataGet() on open
- [x] 6.3 Robust parsing (string vs object)
- [x] 6.4 Validate data structure
- [x] 6.5 Reconstruct cellData from form inputs on save
- [x] 6.6 Persist via userDataSet() on Save and Close
- [x] 6.7 Preserve formula cells {f, v} during save

## 7. Rich Text Editor Compatibility

- [x] 7.1 joplin-editable div wrapper
- [x] 7.2 Original markdown in joplin-source
- [x] 7.3 CodeMirror no-op stub
- [x] 7.4 Verified round-trip across editors

## 8. Error Handling and Validation

- [x] 8.1 Try-catch around plugin initialization
- [x] 8.2 Try-catch around insert and open commands
- [x] 8.3 User-friendly error dialogs
- [x] 8.4 Console logging
- [x] 8.5 Validate userDataGet() return value
- [x] 8.6 Input validation for cell content
    - [x] 8.6.1 Cell content length (10K char max)
    - [x] 8.6.2 JSON size before save (5 MB limit)
    - [x] 8.6.3 Spreadsheet dimensions (500 rows, 52 cols)
    - [x] 8.6.4 HTML escaping for cell display values

## 9. Formula Engine (dialogScript.js)

- [x] 9.1 Formula detection (= prefix)
- [x] 9.2 Cell reference resolution (A1, B2)
- [x] 9.3 Range expansion (A1:C3)
- [x] 9.4 Arithmetic: +, -, \*, / with operator precedence
- [x] 9.5 Comparison operators: >, <, >=, <=, =, <>
- [x] 9.6 SUM, AVERAGE, COUNT, MIN, MAX
- [x] 9.7 IF(condition, true_val, false_val)
- [x] 9.8 CONCAT, UPPER, LOWER, LEN, TRIM
- [x] 9.9 ABS, ROUND
- [x] 9.10 Circular reference detection
- [x] 9.11 Error values: #DIV/0!, #VALUE!, #NAME?, #CIRC!, #ERROR!
- [x] 9.12 Formula bar (cell ref + formula display)
- [x] 9.13 Focus: show formula; blur: show computed value
- [x] 9.14 Recalc all cells on any cell change
- [x] 9.15 Formula cells stored as {f: "=...", v: ""}

## 10. Documentation

- [x] 10.1 README.md with formula docs
- [x] 10.2 Code comments in all source files
- [ ]\* 10.3 User documentation with screenshots
- [ ]\* 10.4 Developer/contribution guide

## 11. Build and Deployment

- [x] 11.1 Build with npm run dist
- [x] 11.2 .jpl file generation
- [x] 11.3 package.json metadata
- [x] 11.4 Version 2.3.0 in manifest.json and package.json
- [ ]\* 11.5 Publish to npm
- [ ]\* 11.6 GitHub release

## 12. Optional Advanced Features

- [ ]\* 12.1 Excel import/export
- [ ]\* 12.2 Spreadsheet templates
- [ ]\* 12.3 Cell formatting (colors, bold, borders)
- [ ]\* 12.4 Chart support
- [ ]\* 12.5 Auto-save mechanism
- [x] 12.6 Add/delete row/column buttons

## Notes

- Tasks marked with \* are optional enhancements
- Formula engine is client-side JS, no external dependencies
- Formulas stored in cellData as {f: "=SUM(A1:A3)", v: ""}
