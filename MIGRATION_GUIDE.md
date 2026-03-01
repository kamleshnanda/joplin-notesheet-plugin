# Migration Guide: v1.0.x to v1.1.0

## Overview

Version 1.1.0 introduces a new storage mechanism that solves the data loss issue when switching between Markdown and Rich Text editors.

## What's Different?

### Old Approach (v1.0.x)
Spreadsheets were stored as JSON inside markdown fence blocks:

```markdown
```univer-sheet
{
  "id": "sheet-123",
  "name": "Budget",
  "sheets": { ... }
}
```
```

**Problem**: Switching to Rich Text Editor converted this to a markdown table, destroying the JSON.

### New Approach (v1.1.0)
Spreadsheets are stored as markdown links with embedded data:

```markdown
[📊 Spreadsheet: Budget](univer://sheet-123 "eyJpZCI6InNoZWV0LTEyMyIsLi4ufQ==")
```

**Solution**: Links survive editor mode switches. Data is embedded (base64) and backed up in userData.

## Migration Steps

### For Existing Spreadsheets

If you have spreadsheets created with v1.0.x:

1. **Backup your notes** (always a good idea!)

2. **Identify affected notes**
   - Look for notes with ` ```univer-sheet` fence blocks
   - These are vulnerable to data loss in Rich Text Editor

3. **Recreate important spreadsheets**
   - Open the note in Markdown editor
   - Copy the cell data you want to preserve
   - Delete the old fence block
   - Insert a new spreadsheet (Ctrl+Shift+U / Cmd+Shift+U)
   - Manually re-enter the data (or wait for future import feature)

4. **Leave less important spreadsheets**
   - Old fence blocks still work in Markdown editor
   - Just avoid switching to Rich Text Editor for those notes

### For New Spreadsheets

All new spreadsheets automatically use the safe link approach. No action needed!

## Testing the New Version

1. **Install v1.1.0**
   - Tools > Options > Plugins
   - Install the new `.jpl` file
   - Restart Joplin

2. **Create a test spreadsheet**
   - Create a new note
   - Click the spreadsheet toolbar button
   - You should see a markdown link inserted

3. **Test editor switching**
   - Switch to viewer mode - see the preview
   - Switch to Rich Text Editor - link remains intact
   - Switch back to Markdown editor - link still there
   - Switch to viewer mode again - preview still works!

4. **Verify data persistence**
   - Close and reopen Joplin
   - Spreadsheet should still render correctly
   - If you sync, it should appear on other devices

## Troubleshooting

### "Failed to load spreadsheet data" message

This means the base64 data in the link title couldn't be decoded. Possible causes:
- Link was manually edited
- Data corruption during sync
- Plugin version mismatch

**Solution**: Recreate the spreadsheet

### Old fence blocks not rendering

Old fence blocks (v1.0.x) should still work in Markdown editor. If they don't:
- Check console for errors (Help > Toggle Development Tools)
- Verify plugin is v1.1.0 or later
- Try disabling and re-enabling the plugin

### Link shows as plain text

If the link appears as plain text instead of rendering a preview:
- Make sure you're in viewer mode (not editor mode)
- Check that the link format is correct: `[text](univer://id "data")`
- Verify the plugin is loaded (check console)

## Benefits of Upgrading

✅ **No more data loss** - Switch editors freely
✅ **Better sync** - Data stored in userData + embedded in link
✅ **Cleaner UX** - Links look intentional
✅ **Same preview** - Beautiful table rendering preserved

## Questions?

If you encounter issues:
1. Check the console for errors
2. Verify the link format is correct
3. Try recreating the spreadsheet
4. Report issues on GitHub

## Future Improvements

Coming in future versions:
- Interactive editing (panel/dialog)
- Excel import/export
- Formula support
- Cell formatting
- Automatic migration tool for old fence blocks
