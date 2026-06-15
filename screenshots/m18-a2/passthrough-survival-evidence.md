# M18 A2 — shape passthrough-survival live evidence (2026-06-15)

Verified the preserve-only shape resource survives Univer's editor save path,
the load-bearing claim of A2.

Fixture: tests/fixtures/shapes/ShapeTextBox-SingleSheet.xlsx (a roundRect text
box "Preserved Shape", anchored D2:H8, FFE699 fill / BF8F00 outline).

Imported as note 9a90950d... in the running dev Joplin, opened in the Custom
Editor (Univer mounted), then called `workbook.save()` in the editor frame —
this invokes getResources(), the exact mechanism that silently DROPS any
resource without a registered passthrough hook.

Result (saved snapshot resource names):
  ... SHEET_NOTESHEET_SHAPES_PLUGIN ...   <-- present
  shapeSurvivesSave: true
  shapeHasContent:   true   (the "Preserved Shape" anchor XML is intact)

Conclusion: the editorView.tsx passthrough hook for NOTESHEET_SHAPES_RESOURCE
works — shapes survive editor edit/save/reload. Stored note body (Data API)
also confirmed to retain the SHEET_NOTESHEET_SHAPES_PLUGIN resource + the
shape text at baseline and after editor mount.
