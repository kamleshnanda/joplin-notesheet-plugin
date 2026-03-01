# Testing Guide for Joplin Univer Plugin

## Quick Start Testing

### 1. Install the Plugin

```bash
# Build the plugin
npm run dist

# The plugin file will be at:
# publish/com.kamleshnanda.joplin-univer-plugin.jpl
```

### 2. Load in Joplin

1. Open Joplin
2. Go to **Tools > Options > Plugins**
3. Click **"Install plugin"**
4. Navigate to `publish/com.kamleshnanda.joplin-univer-plugin.jpl`
5. Click **"Open"**
6. **Restart Joplin**

### 3. Verify Plugin is Loaded

After restart:
1. Go to **Tools > Options > Plugins**
2. You should see "Joplin Univer Spreadsheet Plugin" in the list
3. Verify it's enabled (checkbox should be checked)

### 4. Test Basic Functionality

#### Test 1: Toolbar Button
1. Create a new note
2. Look for the spreadsheet icon (📊) in the editor toolbar
3. Click it
4. Verify a code block is inserted with `univer-sheet` syntax

#### Test 2: Menu Item
1. Create a new note
2. Go to **Tools > Insert Univer Spreadsheet**
3. Verify a code block is inserted

#### Test 3: Keyboard Shortcut
1. Create a new note
2. Press **Ctrl+Shift+U** (Windows/Linux) or **Cmd+Shift+U** (Mac)
3. Verify a code block is inserted

#### Test 4: Content Rendering
1. Create a note with a spreadsheet block (use any method above)
2. Switch to **view mode** (toggle editor/viewer)
3. You should see a placeholder box with:
   - "Univer Spreadsheet Placeholder"
   - "Plugin is installed and working!"
   - The spreadsheet ID

#### Test 5: Multiple Spreadsheets
1. Create a note
2. Insert 2-3 spreadsheet blocks
3. Switch to view mode
4. Verify all spreadsheets render correctly

## Expected Output

### In Edit Mode

You should see a markdown code block:

```markdown
```univer-sheet
{
  "id": "sheet-1709078400000",
  "name": "Spreadsheet",
  "sheetOrder": ["sheet-01"],
  "appVersion": "0.1.0",
  "locale": "en-US",
  "styles": {},
  "sheets": {
    "sheet-01": {
      "id": "sheet-01",
      "name": "Sheet1",
      "cellData": {
        "0": {
          "0": { "v": "Column A" },
          "1": { "v": "Column B" },
          "2": { "v": "Column C" }
        }
      },
      "rowCount": 100,
      "columnCount": 26,
      "defaultColumnWidth": 93,
      "defaultRowHeight": 27
    }
  }
}
```
```

### In View Mode

You should see a styled box with:
- Border (dashed, gray)
- Centered text
- Heading: "Univer Spreadsheet Placeholder"
- Message: "Plugin is installed and working!"
- Spreadsheet ID displayed

## Troubleshooting

### Plugin doesn't appear after installation

**Solution:**
1. Verify the .jpl file was created: `ls -lh publish/`
2. Check Joplin's plugin directory:
   - Windows: `%USERPROFILE%\.config\joplin-desktop\plugins`
   - Mac: `~/.config/joplin-desktop/plugins`
   - Linux: `~/.config/joplin-desktop/plugins`
3. Restart Joplin completely (quit and reopen)

### Toolbar button doesn't appear

**Solution:**
1. Check if plugin is enabled in Tools > Options > Plugins
2. Try toggling the plugin off and on
3. Restart Joplin

### Spreadsheet doesn't render in view mode

**Solution:**
1. Open Developer Tools: Help > Toggle Development Tools
2. Check the Console tab for errors
3. Verify the fence block syntax is exactly `univer-sheet` (no spaces)
4. Verify the JSON is valid (use a JSON validator)

### Build fails

**Solution:**
```bash
# Clean and rebuild
rm -rf node_modules dist publish
npm install
npm run dist
```

### "No note selected" error

**Solution:**
- Make sure you have a note open and selected before clicking the toolbar button

## Developer Testing

### Check Console Output

1. Open Joplin
2. Go to Help > Toggle Development Tools
3. Click Console tab
4. Look for messages:
   - "Joplin Univer Sheets plugin starting..."
   - "Univer Sheets plugin registration complete"
   - "UniverSheetManager initialized (placeholder mode)"

### Verify Files

Check that these files exist in `dist/`:
```bash
ls -lh dist/
# Should show:
# - index.js (main plugin)
# - contentScript.js (content script)
# - manifest.json (plugin metadata)
# - univer-bundle.css (Univer styles)
```

### Test Error Handling

1. Create a note with invalid JSON:
```markdown
```univer-sheet
{ invalid json }
```
```

2. Switch to view mode
3. Should see error message in red

## Next Steps

Once basic functionality is verified:

1. **Report Issues**: Document any bugs or unexpected behavior
2. **Test Edge Cases**: Try unusual inputs, large datasets, etc.
3. **Performance**: Test with multiple spreadsheets in one note
4. **Compatibility**: Test on different platforms (Windows, Mac, Linux)

## Success Criteria

✅ Plugin installs without errors
✅ Toolbar button appears and works
✅ Menu item appears and works
✅ Keyboard shortcut works
✅ Spreadsheet blocks render in view mode
✅ Multiple spreadsheets work in one note
✅ No console errors during normal operation

## Known Limitations (Current Version)

- Spreadsheets are not interactive (placeholder only)
- No editing capabilities yet
- No formula support yet
- No data persistence yet
- Full Univer integration pending build system resolution

These will be addressed in future updates.
