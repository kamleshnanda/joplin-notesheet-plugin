# Changelog

All notable changes will be documented here. Versioning follows [SemVer](https://semver.org/).

## Unreleased

### M0 — Notesheet rebrand and architectural reset

- Rebranded the plugin from "Univer Worksheet Plugin" to **Notesheet**.
- Plugin id changed to `com.kamleshnanda.joplin-notesheet`; npm package name to `joplin-plugin-notesheet`.
- Removed the v0 popup-spreadsheet-embedded-in-markdown model (custom URL scheme, fenced-markdown mirror, dialog UI, custom HTML/JS spreadsheet implementation).
- Added Jest + GitHub Actions CI scaffolding for v1.

This release contains no end-user spreadsheet functionality on its own. v1 is built incrementally over subsequent milestones (M1 reintroduces the Univer SDK as the engine and adds the Custom Editor view; M2+ adds inline preview, formatting, sort/filter, .xlsx import/export, charts, and tables).
