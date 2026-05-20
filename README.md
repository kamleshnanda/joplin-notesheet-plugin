# Notesheet — Spreadsheets for Joplin

Notesheet turns a Joplin note into a real spreadsheet. Powered by the [Univer SDK](https://github.com/dream-num/univer), it gives Joplin first-class support for formulas, formatting, sorting, filtering, charts, and `.xlsx` import/export — all inside the note editor pane you already use.

## Status

**v1 is under active development.** The plugin currently scaffolds the v1 architecture (clean break from the v0 popup model). Spreadsheet functionality lands in the M1 milestone; see the milestone plan in `/Users/kamleshn/.claude/plans/abstract-frolicking-sloth.md` (project-internal) for the roadmap.

## Architecture (v1)

- A **Spreadsheet note** is a regular Joplin note whose body is a Univer snapshot wrapped in a fenced markdown code block tagged ```` ```notesheet ````. When the active note matches that shape, Joplin's editor pane shows the Univer editor instead of the markdown editor (via Joplin's Custom Editor API).
- A **regular markdown note** can link to a Spreadsheet note with the standard `[label](:/<noteId>)` syntax. The Notesheet content script renders an inline read-only preview of the linked spreadsheet in the markdown preview pane. Click the preview to open the spreadsheet for editing.
- A **fullscreen toggle** inside the Univer editor collapses Joplin's sidebar and notes list when you need maximum room.

## Development

```bash
npm install
npm run dist     # builds the .jpl into publish/
npm test         # runs Jest unit tests
```

Requires Node.js 18+. The legacy OpenSSL flag is set in the build scripts.

## Compatibility

- Joplin 3.5+
- Desktop only (Joplin Mobile's plugin model isn't yet ready for the Custom Editor API)

## License

MIT. The Univer SDK is Apache-2.0.
