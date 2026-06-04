// Mock @univerjs/sheets-table so importing src/univerTableTheme doesn't
// drag in the whole Univer ESM graph (lodash-es etc.) through jest's
// CJS transform. The helper is identity-tested; the sentinel works
// just as well as the real plugin reference.
jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

// Pin-down regression tests for M12 (formatting fidelity polish).
//
// Each pin-down captures a structural invariant that, if violated, would
// reproduce a bug we already shipped or fixed during M12 development.
// The named "WOULD HAVE CAUGHT" comments reference the actual regressions
// users reported during 2026-05-30/2026-05-31 manual testing.
//
// Anchored on project-owned fixtures in tests/ExcelBaseTestData/formatting-testdata/.
// **No personal data, no /Users/ paths, no skipped-on-missing tests.** A
// future contributor must be able to clone the repo, run `npm test`, and
// see these pass without fetching anything external.
//
// When adding a new pin-down: state the bug it captures in the body of
// the test as well, not just in a commit message.

import { readFileSync } from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata');
const APTOS = path.join(FIXTURES_DIR, 'FormattingSmorgasboard.xlsx');
const CLASSIC = path.join(FIXTURES_DIR, 'FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx');

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, {
            v?: unknown; t?: number; s?: string;
            p?: {
                body?: {
                    dataStream?: string;
                    paragraphs?: Array<{ startIndex: number }>;
                    sectionBreaks?: Array<{ startIndex: number }>;
                    customRanges?: Array<{ rangeType?: number; properties?: { url?: string } }>;
                };
                documentStyle?: { pageSize?: { width?: number; height?: number } };
            };
        }>>;
    }>;
    styles: Record<string, Record<string, unknown>>;
    defaultStyle?: { ff?: string };
    resources?: Array<{ name: string; data: string }>;
}

async function importFixture(file: string): Promise<SnapshotShape> {
    const buf = readFileSync(file);
    return await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;
}

async function roundTrip(snap: SnapshotShape): Promise<{ zip: JSZip; wb: ExcelJS.Workbook }> {
    const out = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
    const zip = await JSZip.loadAsync(out as ArrayBuffer);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out as unknown as Parameters<typeof wb.xlsx.load>[0]);
    return { zip, wb };
}

async function readThemeXml(zip: JSZip): Promise<string> {
    const themePath = Object.keys(zip.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p))!;
    return await zip.files[themePath].async('string');
}

async function readTableXml(zip: JSZip): Promise<string | null> {
    const tablePath = Object.keys(zip.files).find((p) => /^xl\/tables\/table\d+\.xml$/.test(p));
    return tablePath ? await zip.files[tablePath].async('string') : null;
}

// ─── Smoke: both fixtures must import without throwing ─────────────────

describe('M12 fixtures — basic smoke', () => {
    test('Aptos fixture imports + round-trips without throwing', async () => {
        const snap = await importFixture(APTOS);
        expect(snap.sheetOrder.length).toBe(1);
        // The fixture has 9 data rows (rows 2-9 from Excel) + 1 header + 1 totals → 10 rows in cellData.
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const rows = Object.keys(sheet.cellData).length;
        expect(rows).toBe(10);
        // Round-trip must succeed.
        await roundTrip(snap);
    });

    test('Classic fixture imports + round-trips without throwing', async () => {
        const snap = await importFixture(CLASSIC);
        expect(snap.sheetOrder.length).toBe(1);
        await roundTrip(snap);
    });
});

// ─── Pin-down: hyperlink cell.p shape (Bug 1 from 2026-05-30 evening) ──
//
// REGRESSION HISTORY: Yesterday's first attempt at fixing the
// "Cannot read properties of undefined (reading 'height')" crash removed
// `documentStyle.pageSize` from the hyperlink cell.p entirely. That
// produced a documentSkeleton with zero pages, and Univer's
// `_calcActiveCell` → `calcPadding` still crashed reading `pages[0].height`.
//
// The correct fix mirrors `createDocumentModelWithStyle` in
// @univerjs/engine-render: dataStream ends with `\r\n`, paragraphs +
// sectionBreaks present, finite pageSize. This pin-down asserts every
// piece. If anyone ever simplifies `buildHyperlinkCellP` and breaks one
// of these, the test must fail.

describe('M12 pin-down — hyperlink documentSkeleton', () => {
    // We build the workbook in-process here (Pattern A: cell.value =
    // {text, hyperlink}) rather than rely on a fixture, because the test
    // is about cell.p shape — independent of which import path produced
    // the hyperlink. Pattern B (named-Hyperlink cellStyle) is exercised
    // separately by the FormattingSmorgasboard pin-down in this file.
    test('hyperlink cell.p has finite pageSize + paragraphs + sectionBreaks + dataStream "\\r\\n" terminator', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = { text: 'click me', hyperlink: 'https://example.com/' };
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const linkCell = sheet.cellData[0]?.[0];
        expect(linkCell?.p).toBeDefined();
        const body = linkCell!.p!.body!;
        const ds = linkCell!.p!.documentStyle!;

        // dataStream must end with \r\n — the empty-document terminator
        // Univer's layout pipeline expects.
        expect(body.dataStream!.endsWith('\r\n')).toBe(true);

        // Paragraphs + sectionBreaks must be present at the right offsets.
        const textLen = body.dataStream!.length - 2;
        expect(body.paragraphs).toEqual([
            expect.objectContaining({ startIndex: textLen }),
        ]);
        expect(body.sectionBreaks).toEqual([{ startIndex: textLen + 1 }]);

        // pageSize must be finite. JSON.stringify of Infinity → null, so
        // the value MUST be a real number that survives round-trip.
        const w = ds.pageSize?.width;
        const h = ds.pageSize?.height;
        expect(typeof w).toBe('number');
        expect(typeof h).toBe('number');
        expect(Number.isFinite(w)).toBe(true);
        expect(Number.isFinite(h)).toBe(true);

        // JSON round-trip preserves these numbers (this is the actual
        // invariant — the symptom of breaking it is the runtime crash).
        const cloned = JSON.parse(JSON.stringify(linkCell));
        expect(cloned.p.documentStyle.pageSize.width).toBe(w);
        expect(cloned.p.documentStyle.pageSize.height).toBe(h);

        // Hyperlink customRange present.
        const ranges = body.customRanges ?? [];
        const hyperlink = ranges.find((r) => r.rangeType === 0);
        expect(hyperlink).toBeDefined();
        expect(hyperlink!.properties?.url).toMatch(/^https?:\/\//);
    });
});

// ─── Pin-down: theme palette round-trip ───────────────────────────────
//
// REGRESSION HISTORY: Before NOTESHEET_THEME_CLR_SCHEME_PLUGIN existed,
// exceljs emitted its own Office-2007 default <a:clrScheme> on export.
// The Aptos fixture's TableStyleMedium4 resolves accent3 against the
// active theme. With the wrong palette in the exported file, the
// round-tripped sheet rendered visibly different colors than the
// original even though `tableStyleInfo` carried the same name.
//
// Both fixtures are checked: Aptos accents (#156082 / #196B24) and the
// Classic fixture's accents (#4472C4 / #A5A5A5).

describe('M12 pin-down — theme palette round-trips', () => {
    test('Aptos fixture: exported theme1.xml carries source accents (#156082, #196B24, hlink #467886)', async () => {
        const snap = await importFixture(APTOS);
        const { zip } = await roundTrip(snap);
        const xml = await readThemeXml(zip);

        expect(xml).toContain('val="156082"'); // accent1
        expect(xml).toContain('val="196B24"'); // accent3
        expect(xml).toContain('val="467886"'); // hlink

        // Negative: exceljs's own Office 2007 default accent3 (#9BBB59
        // olive) must NOT be in the exported theme.
        expect(xml).not.toContain('val="9BBB59"');
    });

    test('Classic fixture: exported theme1.xml carries source accents (#4472C4, not #4F81BD)', async () => {
        const snap = await importFixture(CLASSIC);
        const { zip } = await roundTrip(snap);
        const xml = await readThemeXml(zip);

        // The Classic fixture uses Office 2013-style accent palette.
        expect(xml).toContain('val="4472C4"'); // accent1
        expect(xml).toContain('val="A5A5A5"'); // accent3
        // Negative: exceljs's own accent1 (#4F81BD Office-2007).
        expect(xml).not.toContain('val="4F81BD"');
    });

    test('Aptos fixture: exported theme1.xml carries Aptos Narrow as minor font', async () => {
        const snap = await importFixture(APTOS);
        const { zip } = await roundTrip(snap);
        const xml = await readThemeXml(zip);
        expect(xml).toContain('typeface="Aptos Narrow"');
    });
});

// ─── Pin-down: synthesized fields are tagged + stripped on export ─────
//
// REGRESSION HISTORY: Without the SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN
// sidecar, our import-time table-style synthesis baked `bg`, `cl`, `bd`
// into per-cell records → on export those wrote as explicit cell-level
// fills/borders → Excel painted them ON TOP of TableStyleMedium4's own
// rendering. The bands looked twice as saturated/dark vs the source.
//
// The fix tags each synthesized field path (`bg`, `cl`, `bl`, `bd.t`,
// `bd.b`, `bd.l`, `bd.r`) into a sidecar resource. On export,
// applyStyleToCell skips those fields when writing to exceljs.

describe('M12 pin-down — synth-styles sidecar', () => {
    test('import emits SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN with header tags for table cells', async () => {
        const snap = await importFixture(APTOS);
        const synth = snap.resources?.find((r) => r.name === 'SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN');
        expect(synth).toBeDefined();
        const sidecar = JSON.parse(synth!.data) as Record<string, Record<string, string[]>>;
        const sheetSidecar = sidecar[snap.sheetOrder[0]];
        expect(sheetSidecar).toBeDefined();
        // Header cell (0:0) was synthesized: bg + bl + 3 borders (t, b, l)
        // because it's the top-left corner of the table.
        const headerTags = sheetSidecar['0:0'];
        expect(headerTags).toEqual(expect.arrayContaining(['bg', 'bl', 'bd.t', 'bd.b', 'bd.l']));
    });

    test('import emits SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN with the source clrScheme', async () => {
        const snap = await importFixture(APTOS);
        const themeRes = snap.resources?.find((r) => r.name === 'SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN');
        expect(themeRes).toBeDefined();
        // Should contain accent1 #156082.
        expect(themeRes!.data).toContain('156082');
        expect(themeRes!.data).toMatch(/<a:clrScheme\b/);
    });
});

// ─── Pin-down: header export does NOT flip applyFont=1 for synth-only cells ─
//
// REGRESSION HISTORY (2026-05-31): I shipped a fix that wrote
// `cell.font = { name: 'Aptos Narrow' }` on every cell that didn't
// already have an `ff` in its style — as a "default font" fallback. But
// for synthesized header cells (bg/cl/bl synthesized, no source font),
// this set `applyFont="1"` on the cell's xf, which disables Excel's
// TableStyle dxf overrides. Result: white-on-color header text rendered
// black on the exported xlsx.
//
// The correct behavior: cells with no explicit source font get NO
// per-cell font on export. Theme1.xml's minorFont rewrite covers the
// "what font does this cell render in" case via Excel's normal theme
// inheritance.
//
// This pin-down builds a synthetic table programmatically (because both
// of our hand-saved fixtures already have explicit cell colors on
// every cell), exports it, and asserts the header cell's xf does NOT
// carry applyFont="1" or any font color.

describe('M12 pin-down — header cells DO NOT flip applyFont on export', () => {
    test('synthesized header (no source cell font) exports without per-cell font', async () => {
        // Build a minimal table with NO per-cell formatting on the header
        // — Excel's TableStyleMedium4 should paint header bg + white
        // text via dxf at render time. If our exporter writes
        // applyFont="1" on the header cell, Excel SKIPS the dxf and the
        // header renders black-on-color (the regression).
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.addTable({
            name: 'TM4',
            ref: 'A1',
            headerRow: true,
            style: { theme: 'TableStyleMedium4', showRowStripes: true } as ExcelJS.TableProperties['style'],
            columns: [{ name: 'Col1' }, { name: 'Col2' }],
            rows: [['x', 1], ['y', 2], ['z', 3]],
        });
        const buf0 = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf0 as unknown as Buffer) as unknown as SnapshotShape;
        const out = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);

        // Inspect the raw OOXML of the exported workbook. exceljs's
        // re-parse interprets <font/> entries lazily, so we go to the
        // XML directly to verify the header cell's xf has no applyFont
        // and the header's referenced font (if any) has no color.
        const zip = await JSZip.loadAsync(out as ArrayBuffer);
        const stylesXml = await zip.files['xl/styles.xml'].async('string');
        const sheetXml = await zip.files['xl/worksheets/sheet1.xml'].async('string');

        // Find the header cell A1's `s` attribute (xf index).
        const headerMatch = /<c\s+r="A1"[^>]*\bs="(\d+)"/.exec(sheetXml);
        // It's also valid for A1 to have NO `s` attribute (= xf 0,
        // default style, no font).
        const xfIndex = headerMatch ? parseInt(headerMatch[1], 10) : 0;

        // Walk cellXfs and pull the xfIndex'th entry.
        const cellXfsMatch = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
        expect(cellXfsMatch).not.toBeNull();
        const xfList = cellXfsMatch![1].match(/<xf\b[^/]*\/>/g) ?? [];
        const headerXf = xfList[xfIndex];
        expect(headerXf).toBeDefined();

        // applyFont="1" combined with a font-color-less <font/> is the
        // exact shape that triggers the regression. We assert NEITHER:
        //
        //   - applyFont="1" is forbidden when the cell had no source
        //     font (the synthesizer should leave the cell with no font
        //     reference at all, letting Excel inherit theme.minorFont)
        //
        //   - if applyFont IS "1" (which only happens when the source
        //     cell had a real font), the referenced font MUST carry an
        //     explicit color so Excel's TableStyle dxf isn't expected
        //     to provide one.
        //
        // Implementation: just check applyFont. The current exporter
        // path writes nothing on the header xf for synth-only cells
        // because we removed the defaultFontName fallback.
        expect(headerXf).not.toMatch(/\bapplyFont="1"/);
    });
});

// ─── Pin-down: passthrough table theme is wired ────────────────────────
//
// REGRESSION HISTORY: Univer's table plugin auto-applies
// `table-default-0` (lavender) over any cell that's part of an
// ITableJson, even when our snapshot already has correct synthesized
// colors. The fix registers a no-op theme as `userThemes[0]` and
// `defaultThemeIndex: 0`, so Univer's resolver picks our empty theme.
//
// This is a unit-level pin-down for the helper function — full
// integration is exercised by editorView at runtime.

describe('M12 pin-down — passthrough table theme', () => {
    test('FLAT_TABLE_THEME_CONFIG keeps userThemes[0] empty + defaultThemeIndex 0', async () => {
        // eslint-disable-next-line @typescript-eslint/no-require-imports
        const mod = await import('../src/univerTableTheme');
        expect(mod.FLAT_TABLE_THEME_CONFIG.defaultThemeIndex).toBe(0);
        expect(mod.FLAT_TABLE_THEME_CONFIG.userThemes).toHaveLength(1);
        expect(mod.FLAT_TABLE_THEME_CONFIG.userThemes[0].style).toEqual({});
        expect(mod.FLAT_TABLE_THEME_CONFIG.userThemes[0].name).toBe(mod.FLAT_TABLE_THEME_NAME);
    });
});

// ─── Round-trip golden: exported table + sheet structure matches source ─
//
// Loosely checked because exceljs adds harmless metadata (xr:uid, etc.).
// We only assert what the user sees in Excel: table style name, stripes
// flag, header row name, totals row presence, and number of data rows.

describe('M12 round-trip golden — table structure', () => {
    test('Aptos fixture: round-trip preserves table name, style, stripes, totals row, columns', async () => {
        const snap = await importFixture(APTOS);
        const { zip, wb } = await roundTrip(snap);
        const tableXml = await readTableXml(zip);

        expect(tableXml).toContain('name="ProjectTracker"');
        expect(tableXml).toContain('TableStyleMedium4');
        expect(tableXml).toContain('showRowStripes="1"');

        // exceljs's parsed view.
        const ws = wb.getWorksheet('Sheet1')!;
        const tables = (ws as unknown as { getTables: () => Array<{ table: { name: string; columns: Array<{ name: string }> } }> }).getTables();
        expect(tables).toHaveLength(1);
        expect(tables[0].table.name).toBe('ProjectTracker');
        expect(tables[0].table.columns.map((c) => c.name)).toEqual([
            'Project', 'Website', 'Budget', 'Spent', '% Complete', 'Start Date', 'Status',
        ]);
    });

    test('Classic fixture: round-trip preserves table name + ProductCatalog columns', async () => {
        const snap = await importFixture(CLASSIC);
        const { zip, wb } = await roundTrip(snap);
        const tableXml = await readTableXml(zip);
        expect(tableXml).toContain('name="ProductCatalog"');

        const ws = wb.getWorksheet('Sheet1')!;
        const tables = (ws as unknown as { getTables: () => Array<{ table: { name: string; columns: Array<{ name: string }> } }> }).getTables();
        expect(tables).toHaveLength(1);
        expect(tables[0].table.columns.map((c) => c.name)).toEqual([
            'Product', 'Website', 'Price', 'Discount', 'Launch Date', 'Revenue',
        ]);
    });

    // Pattern B (named-Hyperlink cellStyle) round-trip. The Aptos
    // fixture's column B cells use Excel's built-in "Hyperlink" cell
    // style (Format → Cell Styles → Hyperlink) rather than the
    // {text, hyperlink} value shape. Our import recognizes the cellStyle
    // chain and synthesizes cell.p; on export we re-emit Pattern A
    // (`<hyperlinks>` block + cell value `{text, hyperlink}`), which
    // Excel renders identically. So the round-tripped cells now read as
    // `cell.isHyperlink === true` even though the source used Pattern B.
    test('Aptos fixture: round-trip preserves all 8 named-Hyperlink cells in column B', async () => {
        const snap = await importFixture(APTOS);
        const { wb } = await roundTrip(snap);
        const ws = wb.getWorksheet('Sheet1')!;
        let linkCount = 0;
        const seenUrls = new Set<string>();
        ws.eachRow({ includeEmpty: false }, (row, rowIdx) => {
            if (rowIdx === 1 || rowIdx === 10) return; // skip header + totals
            const c = row.getCell(2); // B = column 2
            if (c.isHyperlink) {
                linkCount++;
                const url = (c.value as { hyperlink?: string }).hyperlink;
                if (url) seenUrls.add(url);
            }
        });
        expect(linkCount).toBe(8);
        expect(seenUrls.size).toBe(8);
    });

    test('Aptos fixture: theme=10 hyperlink color resolves to the source workbook hlink (Aptos #467886)', async () => {
        // REGRESSION HISTORY: Before the theme+tint resolver landed
        // (2026-05-31 evening), `font.color = {theme: 10}` resolved to
        // undefined → fell back to no color → cell rendered black-on-bg.
        // Symptom in the user's InvestmentSummary.xlsx: Vendor/Start Date
        // hyperlinks looked like plain black underlined text instead of
        // the dark teal-blue (Aptos hlink). The resolver maps theme=10
        // to the workbook clrScheme's <a:hlink> element.
        const snap = await importFixture(APTOS);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const cell = sheet.cellData[1]?.[1];
        const style = cell?.s ? snap.styles[cell.s] : null;
        // ul (underline) AND cl (color) both come from the named-Hyperlink
        // font (font 1 in styles.xml: <font><u/><sz/><color theme="10"/>
        // <name val="Aptos Narrow"/>...</font>).
        expect((style?.ul as { s: number } | undefined)?.s).toBe(1);
        const cl = style?.cl as { rgb: string } | undefined;
        // Aptos workbook theme: <a:hlink><a:srgbClr val="467886"/>.
        // Anything other than #467886 means the resolver fell back to
        // black or skipped the theme reference — both are P1 regressions.
        expect(cl?.rgb).toBe('#467886');
    });

    test('Aptos fixture: theme=10 hyperlink color survives round-trip', async () => {
        const snap = await importFixture(APTOS);
        const { wb } = await roundTrip(snap);
        const ws = wb.getWorksheet('Sheet1')!;
        // Snapshot row 1 col 1 = Excel B2.
        const c = ws.getCell('B2');
        const argb = c.font?.color?.argb;
        // exceljs serializes the resolved RGB as 'FFRRGGBB' (alpha=FF).
        expect(argb?.toUpperCase()).toBe('FF467886');
    });

    test('Classic fixture: theme=10 hyperlink color resolves against Classic clrScheme (#0563C1)', async () => {
        // Classic theme's <a:hlink> is #0563C1. Same code path as Aptos
        // but a different RGB result; this pin-down ensures the
        // resolver isn't accidentally hardcoded to one workbook's
        // palette.
        const snap = await importFixture(CLASSIC);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const cell = sheet.cellData[1]?.[1];
        const style = cell?.s ? snap.styles[cell.s] : null;
        const cl = style?.cl as { rgb: string } | undefined;
        expect(cl?.rgb).toBe('#0563C1');
    });

    test('Aptos fixture: import emits cell.p for Pattern B (named-Hyperlink) cells in B2', async () => {
        // The B2 cell (snapshot row 1, col 1) is plain string-valued in
        // exceljs (`isHyperlink === false`) but uses the named-Hyperlink
        // cellStyle. Our import detects this via xl/styles.xml's
        // cellStyles + cellXfs chain and synthesizes a hyperlink cell.p.
        // Without that detection (the bug we shipped before adding
        // readNamedHyperlinkCells), Vendor/Start Date columns rendered
        // as black-underlined plain text instead of clickable links.
        const snap = await importFixture(APTOS);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const cell = sheet.cellData[1]?.[1];
        expect(cell?.p).toBeDefined();
        const ranges = cell!.p!.body!.customRanges ?? [];
        const link = ranges.find((r) => r.rangeType === 0);
        expect(link?.properties?.url).toBe('https://example.com/alpha');
    });
});

// ─── Pin-down: M13/E theme-aware banding synthesis ────────────────────
//
// REGRESSION HISTORY (2026-06-03): Before M13/E, `synthesizeTableStyleAssignments`
// looked styles up in the static `EXCEL_TABLE_STYLE_BY_NAME` catalog,
// which is hardcoded against the Office 2016+ Aptos accent palette.
// When a workbook shipped its own non-Aptos `<a:clrScheme>` (e.g. the
// 2013-era "Classic" palette whose accent3 is `#A5A5A5` grey), the
// same `TableStyleMedium4` baked **green** Aptos accent3 colours into
// per-cell styles instead of the Classic grey.
//
// FIRST FIX (PR #22): Routed the catalog lookup through a parallel
// recipe table that names each slot's accent index + HSL-L tint, and
// resolved the recipe against the source workbook's `<a:clrScheme>`.
// That fixed the obvious "always Aptos green" regression but the recipe
// values themselves were eyeballed against the catalog, NOT against
// real Excel renders — so we shipped header `#196B24` (raw accent3)
// when Excel actually paints `#34692E`, and banded `#84E291` when
// Excel paints `#CAEFCB`.
//
// SECOND FIX (this PR): Added an empirical-override map keyed by
// `(styleName, accentHex)` whose RGBs were sampled from
// `screenshots/excel-reference/*.png`. The recipe formula remains the
// fallback path for unmeasured accents.
//
// CANONICAL TEST: `tests/excelReferenceFidelity.test.ts` is the source
// of truth for "does the synthesizer match Excel?" These pin-downs
// echo a subset of the same constraints with simpler shape assertions
// for fast failure when round-trip pipelines drift.

describe('M13/E pin-down — theme-aware banding synthesis', () => {
    test('Aptos fixture: ProjectTracker header bg matches Excel render at #34692E (regression sentinel)', async () => {
        const snap = await importFixture(APTOS);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        // ProjectTracker spans A1:G10 with TableStyleMedium4. Header is
        // row 0. Per the empirical override for (Medium4, #196B24
        // Aptos accent3), the rendered header bg is #34692E — measured
        // from the Excel reference PNG, NOT a hand-derivation. See
        // tests/excelReferenceFidelity.test.ts for the ground-truth gate.
        const headerCell = sheet.cellData[0]?.[0];
        expect(headerCell?.s).toBeDefined();
        const headerStyle = snap.styles[headerCell!.s!];
        const bg = (headerStyle.bg as { rgb: string }).rgb;
        expect(bg).toBe('#34692E');
        // Negative: the raw Aptos accent3 must NOT be what we paint
        // for the header — that was the PR #22 symptom.
        expect(bg).not.toBe('#196B24');
    });

    test('Classic fixture: ProductCatalog header bg matches Excel render at #A5A5A5 (failure-mode sentinel)', async () => {
        const snap = await importFixture(CLASSIC);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        // ProductCatalog spans A1:F10 with the SAME TableStyleMedium4
        // as the Aptos fixture. The Classic workbook's accent3 is
        // #A5A5A5 (grey). Theme-aware synthesis must paint grey, not
        // Aptos's #196B24 / #34692E green. The empirical override for
        // (Medium4, #A5A5A5) lands the header at raw #A5A5A5 — same
        // value Excel paints.
        const headerCell = sheet.cellData[0]?.[0];
        expect(headerCell?.s).toBeDefined();
        const headerStyle = snap.styles[headerCell!.s!];
        const bg = (headerStyle.bg as { rgb: string }).rgb;
        expect(bg).toBe('#A5A5A5');
        // Negative: the Aptos accent3 / its TableStyleMedium4 derivative
        // must NOT bleed in. This is the exact symptom of the original
        // M13/E bug.
        expect(bg).not.toBe('#196B24');
        expect(bg).not.toBe('#34692E');
    });

    test('Classic fixture: banded data row bg matches Excel render at #EDEDED (greyscale)', async () => {
        // Even data rows on the Classic ProductCatalog table render at
        // #EDEDED in Excel — a light grey. With the old Aptos-hardcoded
        // path, the bandedEvenBg was #84E291 (a green). With PR #22's
        // tint(+0.6) it became #DBDBDB. The empirical override now
        // lands the right value.
        const snap = await importFixture(CLASSIC);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const bandedCell = sheet.cellData[1]?.[0];
        expect(bandedCell?.s).toBeDefined();
        const style = snap.styles[bandedCell!.s!];
        const bg = (style.bg as { rgb: string } | undefined)?.rgb;
        expect(bg).toBe('#EDEDED');
        // Sanity: it's still a greyscale; this is the structural shape
        // the original spec called out as a generic invariant.
        expect(bg).toMatch(/^#([0-9A-F]{2})\1\1$/i);
    });

    test('Aptos fixture: totals-row carries top AND bottom borders in the Excel separator (#72D068 green, MEDIUM, both sides)', async () => {
        // PR #22's recipe shape did not model the totals-row separators
        // at all — the `borderColor` slot only covers the table outline.
        // M13/E adds `totalsTopBorder` and `totalsBottomBorder` slots.
        //
        // Style choice MEDIUM (s=8) NOT DOUBLE (s=7): Excel's render
        // shows a single 2px strip in the lighter accent on each side,
        // separated by the totals-row body height. Pixel sampling at
        // `screenshots/excel-reference/FormattingSmorgasboard-Aptos.png`
        // confirms strips at y=424-425 (top) and y=472-473 (bottom),
        // both `#72D068`, with white body in between.
        //
        // The totals BOTTOM strip replaces the table outline's thin
        // frame on the totals row's bottom edge — Excel paints the
        // accent-coloured strip across the full width, not the outline
        // colour.
        const snap = await importFixture(APTOS);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        // ProjectTracker has totalsRowCount=1; totals row is row 9.
        const totalsCell = sheet.cellData[9]?.[0];
        expect(totalsCell?.s).toBeDefined();
        const style = snap.styles[totalsCell!.s!];
        const bd = style.bd as
            | { t?: { s: number; cl: { rgb: string } }; b?: { s: number; cl: { rgb: string } } }
            | undefined;
        expect(bd?.t).toBeDefined();
        expect(bd!.t!.cl.rgb).toBe('#72D068');
        expect(bd!.t!.s).toBe(8); // BorderStyleTypes.MEDIUM
        expect(bd?.b).toBeDefined();
        expect(bd!.b!.cl.rgb).toBe('#72D068');
        expect(bd!.b!.s).toBe(8);
    });

    test('Classic fixture: totals-row carries top AND bottom borders in the Excel separator (#C9C9C9 grey, MEDIUM, both sides)', async () => {
        // Same shape as Aptos, different accent. Empirical override
        // for Classic accent3 #A5A5A5 sets both sides to #C9C9C9.
        const snap = await importFixture(CLASSIC);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const totalsCell = sheet.cellData[9]?.[0];
        expect(totalsCell?.s).toBeDefined();
        const style = snap.styles[totalsCell!.s!];
        const bd = style.bd as
            | { t?: { s: number; cl: { rgb: string } }; b?: { s: number; cl: { rgb: string } } }
            | undefined;
        expect(bd?.t).toBeDefined();
        expect(bd!.t!.cl.rgb).toBe('#C9C9C9');
        expect(bd!.t!.s).toBe(8);
        expect(bd?.b).toBeDefined();
        expect(bd!.b!.cl.rgb).toBe('#C9C9C9');
        expect(bd!.b!.s).toBe(8);
    });
});
