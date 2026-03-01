# Changelog

## Version 1.1.0 - Markdown Link Approach (2024-02-28)

### Major Changes

**Switched from fence blocks to markdown links for spreadsheet storage**

This version fundamentally changes how spreadsheets are stored and rendered to solve the data loss issue when switching between Markdown and Rich Text editors.

### What Changed

#### Before (v1.0.x - Fence Block Approach)
```markdown
```univer-sheet
{
  "id": "sheet-123",
  "sheets": { ... }
}
```
```

**Problem**: When switching to Rich Text Editor, Joplin converts the fence block to a markdown table, permanently destroying the JSON data.

#### After (v1.1.0 - Markdown Link Approach)
```markdown
[📊 Spreadsheet: Budget 2024](univer://sheet-123 "base64encodedJSON")
```

**Solution**: 
- Markdown links survive editor mode switches
- Data is embedded in the link's title attribute (base64 encoded)
- Data is also stored in note's userData for backup/sync
- Users can freely switch between editors without data loss

### Technical Implementation

1. **Data Storage (Dual approach)**
   - Primary: Embedded in markdown link title attribute (base64 encoded JSON)
   - Backup: Stored in note's userData via `joplin.data.userDataSet()`
   - Both sync across devices via Joplin's sync mechanism

2. **Link Format**
   ```
   [📊 Spreadsheet: Name](univer://sheet-id "base64data")
   ```
   - Protocol: `univer://` for identification
   - Sheet ID: Unique identifier
   - Title attribute: Base64 encoded spreadsheet JSON

3. **Content Script Changes**
   - Intercepts markdown links with `univer://` protocol
   - Decodes base64 data from title attribute
   - Renders spreadsheet preview inline
   - No more fence block interception needed

### Benefits

✅ **Editor Mode Safe**: Switch freely between Markdown and Rich Text editors
✅ **Data Persistence**: Data stored in userData + embedded in link
✅ **Sync Compatible**: Works with all Joplin sync methods (Dropbox, etc.)
✅ **Clean UX**: Link looks intentional, not like a bug
✅ **Preview Rendering**: Same beautiful table preview as before

### User Experience

**Inserting a Spreadsheet**
1. Click toolbar button or press `Ctrl+Shift+U` / `Cmd+Shift+U`
2. A markdown link is inserted: `[📊 Spreadsheet: Spreadsheet](univer://sheet-123 "...")`
3. Success message confirms data is safely stored

**Viewing a Spreadsheet**
- In viewer mode, the link renders as a full spreadsheet preview
- Shows table with cell data, headers, and metadata
- Same visual appearance as v1.0.x

**Editing Modes**
- ✅ Markdown Editor: Link visible, preview in viewer
- ✅ Rich Text Editor: Link preserved as clickable link
- ✅ Viewer Mode: Full spreadsheet preview rendered

### Migration Notes

**Existing Spreadsheets (v1.0.x fence blocks)**
- Old fence blocks will continue to work in Markdown editor
- ⚠️ Still vulnerable to Rich Text Editor conversion
- Recommendation: Manually recreate important spreadsheets with v1.1.0

**New Spreadsheets (v1.1.0 links)**
- All new spreadsheets use the link approach
- Fully protected from editor mode switches

### Files Changed

- `src/index.ts`: Updated to create markdown links and store in userData
- `src/contentScript.ts`: Rewritten to intercept links instead of fence blocks
- `src/manifest.json`: Version bumped to 1.1.0
- `package.json`: Version bumped to 1.1.0

### Known Limitations

- Interactive editing not yet implemented (coming in future version)
- Preview is read-only
- Maximum 10x10 cells shown in preview (performance optimization)

### Next Steps (Future Versions)

- [ ] Add panel/dialog for interactive editing
- [ ] Implement full Univer SDK integration
- [ ] Add Excel import/export
- [ ] Support formulas and calculations
- [ ] Add cell formatting options
