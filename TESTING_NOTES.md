# Testing Notes for Version 2.0.0

## Quick Test Checklist

### Basic Functionality
- [ ] Plugin loads without errors
- [ ] Toolbar button appears
- [ ] Menu item appears (Tools > Insert Univer Spreadsheet)
- [ ] Keyboard shortcut works (Ctrl+Shift+U / Cmd+Shift+U)

### Insert Spreadsheet
- [ ] Click toolbar button
- [ ] Link is inserted: `[📊 Spreadsheet: Spreadsheet](:/univer/sheet-123)`
- [ ] Success message appears
- [ ] Link is styled (looks like a button)

### Open Dialog
- [ ] Click the spreadsheet link
- [ ] Dialog opens
- [ ] Dialog shows spreadsheet data
- [ ] Table has editable cells
- [ ] Can type in cells
- [ ] "Save & Close" and "Cancel" buttons visible

### Editor Mode Switching (CRITICAL TEST)
- [ ] Create spreadsheet in Markdown editor
- [ ] Switch to Rich Text Editor
- [ ] **Link should remain as link** (not converted to table!)
- [ ] Switch back to Markdown editor
- [ ] Link still intact
- [ ] Click link - dialog opens with data

### Data Persistence
- [ ] Create spreadsheet
- [ ] Edit some cells in dialog
- [ ] Click "Save & Close"
- [ ] Close and reopen Joplin
- [ ] Click link - data should be preserved

### Sync (if applicable)
- [ ] Create spreadsheet on device 1
- [ ] Sync
- [ ] Open note on device 2
- [ ] Click link - data should appear

## Known Issues to Watch For

### Issue 1: Link Not Clickable
**Symptom**: Link appears but doesn't open dialog when clicked

**Possible Causes**:
- Content script not loaded
- Link format incorrect
- Command not registered

**Check**:
- Open console (Help > Toggle Development Tools)
- Look for "UNIVER CONTENT SCRIPT v2.0.0 LOADED"
- Look for "Univer Sheets plugin registration complete"

### Issue 2: Dialog Shows "Data Not Found"
**Symptom**: Dialog opens but shows error message

**Possible Causes**:
- Data not stored in userData
- Sheet ID mismatch
- Note ID changed

**Check**:
- Console should show "Storing spreadsheet data in userData with ID: sheet-xxx"
- Verify sheet ID in link matches stored data key

### Issue 3: Link Converted to Table
**Symptom**: After switching to Rich Text Editor, link becomes a table

**This should NOT happen in v2.0.0!** If it does:
- Verify you're using v2.0.0 (check manifest.json)
- Check that content script is NOT rendering HTML preview
- Content script should only add styling, not render content

### Issue 4: Data Not Syncing
**Symptom**: Spreadsheet appears on one device but not another

**Possible Causes**:
- Sync not configured
- userData not syncing
- Different note versions

**Check**:
- Verify Joplin sync is working for regular notes
- userData should sync automatically with note content

## Console Messages to Look For

### Successful Startup
```
=== UNIVER PLUGIN v2.0.0 STARTING ===
Joplin Univer Sheets plugin starting...
Registering content script...
=== CONTENT SCRIPT REGISTERED SUCCESSFULLY ===
Univer Sheets plugin registration complete
```

### Successful Insert
```
Storing spreadsheet data in userData with ID: sheet-1234567890
Inserting markdown link: [📊 Spreadsheet: Spreadsheet](:/univer/sheet-1234567890)
Univer spreadsheet inserted successfully
```

### Successful Open
```
Opening spreadsheet editor for ID: sheet-1234567890
Dialog closed with result: { id: 'save' }
```

### Content Script Load
```
=== UNIVER CONTENT SCRIPT v2.0.0 LOADED ===
Univer content script function called with context: [object]
Univer markdown-it plugin initializing...
```

## Debugging Tips

### Enable Verbose Logging
Open console and run:
```javascript
localStorage.setItem('debug', 'joplin:*');
```

### Check UserData
To verify data is stored, you can inspect it via Joplin's Data API (requires developer tools).

### Inspect Link Format
The link should look exactly like this:
```markdown
[📊 Spreadsheet: Name](:/univer/sheet-1234567890)
```

Key points:
- Protocol: `:/univer/` (not `univer://`)
- No spaces in URL
- No embedded data in title attribute

### Check Dialog HTML
If dialog opens but looks broken:
- Check console for JavaScript errors
- Verify HTML is rendering correctly
- Check if styles are applied

## Performance Notes

### Expected Behavior
- Insert: Instant
- Open dialog: < 1 second
- Edit cells: Instant
- Save: < 1 second

### If Slow
- Check spreadsheet size (large cellData)
- Check console for errors
- Verify no infinite loops in dialog script

## Comparison with Previous Versions

### v1.0.x (Fence Block)
- ❌ Data loss when switching editors
- ✅ Nice preview in viewer
- ❌ No interactive editing

### v1.1.0 (Markdown Link with Preview)
- ❌ Data loss when switching editors
- ✅ Nice preview in viewer
- ❌ No interactive editing

### v2.0.0 (Dialog-Based)
- ✅ No data loss (immune to editor switches)
- ⚠️ No inline preview (by design)
- ✅ Interactive editing in dialog
- ✅ Clean, intentional UX

## Success Criteria

Version 2.0.0 is successful if:

1. ✅ Spreadsheet links survive Rich Text Editor switches
2. ✅ Data persists across Joplin restarts
3. ✅ Dialog opens and shows data correctly
4. ✅ Basic cell editing works
5. ✅ Data syncs across devices (if sync configured)

## Reporting Issues

If you find bugs, please report:
1. Joplin version
2. Plugin version (should be 2.0.0)
3. Operating system
4. Steps to reproduce
5. Console error messages
6. Expected vs actual behavior

## Next Testing Phase

Once basic functionality is confirmed, test:
- [ ] Large spreadsheets (100+ cells)
- [ ] Special characters in cells
- [ ] Empty spreadsheets
- [ ] Multiple spreadsheets in one note
- [ ] Spreadsheets in different notes
- [ ] Copy/paste notes with spreadsheets
- [ ] Export notes with spreadsheets
