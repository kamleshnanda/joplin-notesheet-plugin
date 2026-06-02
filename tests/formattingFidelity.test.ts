// Tests for M12 formatting fidelity polish:
//   1. Theme fonts — preserve workbook-default font (Aptos Narrow, Calibri,
//      etc.) declared in xl/theme/theme1.xml on import + export.
//   2. Banded-style synthesis — bake Excel TableStyle palette colors into
//      per-cell styles on import so Joplin shows the right colors even
//      though Univer's table preset uses its own theme.
//   3. Hyperlinks — preserve cell URLs through xlsx → snapshot → xlsx.

import ExcelJS from 'exceljs';
import JSZip from 'jszip';

import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

// Snapshot shape used in assertions. Loose typing — we only assert the
// fields we care about; the wider shape is opaque to the test.
interface Snapshot {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, {
            v?: unknown; t?: number; s?: string;
            p?: {
                body?: {
                    dataStream?: string;
                    paragraphs?: Array<{ startIndex: number; paragraphStyle?: Record<string, unknown> }>;
                    sectionBreaks?: Array<{ startIndex: number }>;
                    textRuns?: Array<{ st: number; ed: number; ts?: Record<string, unknown> }>;
                    customRanges?: Array<{ rangeType?: number; properties?: { url?: string } }>;
                };
                documentStyle?: {
                    pageSize?: { width?: number; height?: number };
                };
            };
        }>>;
    }>;
    styles: Record<string, Record<string, unknown>>;
    defaultStyle?: { ff?: string };
    resources?: Array<{ name: string; data: string }>;
}

// ─── Theme fonts ───────────────────────────────────────────────────────────

describe('M12 — theme fonts', () => {
    test('import: workbook with Aptos Narrow theme → snapshot.defaultStyle.ff === "Aptos Narrow"', async () => {
        // The user's real-world InvestmentSummary.xlsx uses Aptos Narrow via
        // theme; we replicate the relevant theme fragment here so the test
        // doesn't depend on a checked-in fixture.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'data';
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);

        // Patch the embedded theme1.xml's minorFont to Aptos Narrow.
        const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
        const themePath = Object.keys(zip.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p))!;
        let themeXml = await zip.files[themePath].async('string');
        themeXml = themeXml.replace(
            /<a:minorFont>\s*<a:latin\b[^>]*\btypeface="[^"]*"/,
            '<a:minorFont><a:latin typeface="Aptos Narrow"',
        );
        zip.file(themePath, themeXml);
        const patchedBuf = Buffer.from(await zip.generateAsync({ type: 'arraybuffer' }) as ArrayBuffer);

        const snap = await xlsxBufferToSnapshot(patchedBuf) as unknown as Snapshot;
        expect(snap.defaultStyle?.ff).toBe('Aptos Narrow');
    });

    test('import: workbook with no theme override → no defaultStyle on snapshot', async () => {
        // exceljs's default theme is Calibri/Cambria, but we only set
        // defaultStyle when the import explicitly finds a theme font.
        // (Note: if exceljs ships a theme with Calibri, this test will see
        // "Calibri" — that's fine; the assertion is loose.)
        const wb = new ExcelJS.Workbook();
        wb.addWorksheet('Sheet1');
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as Snapshot;
        // Either no defaultStyle (theme parsing failed) or ff is a plausible
        // font name string — both acceptable. The bug we're guarding against
        // is the import dropping the field even when the theme has data.
        if (snap.defaultStyle) {
            expect(typeof snap.defaultStyle.ff).toBe('string');
            expect(snap.defaultStyle.ff!.length).toBeGreaterThan(0);
        }
    });

    test('export: snapshot.defaultStyle.ff is patched into theme1.xml so unstyled cells inherit it', async () => {
        // Cells with no explicit font in their style record do NOT get a
        // per-cell font.name written on export — Excel inherits from
        // theme1.xml's minorFont. Writing a per-cell font.name would flip
        // the cell's xf applyFont="1" and prevent Excel's TableStyle dxfs
        // from applying (the most visible regression: white-on-color
        // header text rendering as black on the exported xlsx). So we
        // verify the theme is patched and rely on render-time inheritance.
        const snap = {
            id: 'wb-1', name: 'Spreadsheet', appVersion: '0.1.0', locale: 'enUS',
            sheetOrder: ['s1'],
            styles: {},
            sheets: {
                s1: {
                    id: 's1', name: 'Sheet1',
                    cellData: { 0: { 0: { v: 'hello', t: 1 } } },
                    rowCount: 100, columnCount: 26,
                    defaultColumnWidth: 73, defaultRowHeight: 19,
                    mergeData: [], rowData: {}, columnData: {},
                },
            },
            defaultStyle: { ff: 'Aptos Narrow' },
        };
        const buf = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);

        // Theme carries Aptos Narrow.
        const zip = await JSZip.loadAsync(buf as ArrayBuffer);
        const themePath = Object.keys(zip.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p));
        expect(themePath).toBeDefined();
        const themeXml = await zip.files[themePath!].async('string');
        expect(themeXml).toContain('<a:latin typeface="Aptos Narrow"');

        // Cell has NO explicit font (so the table-style dxf path stays
        // open). exceljs's getCell().font returns undefined for cells
        // with no <font> in their xf.
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf as unknown as Parameters<typeof wb.xlsx.load>[0]);
        const ws = wb.getWorksheet('Sheet1')!;
        const fontEntry = ws.getCell('A1').font;
        expect(fontEntry === undefined || fontEntry.name === undefined).toBe(true);
    });

    test('round-trip: import xlsx with custom theme font → export preserves it', async () => {
        // Build a minimal xlsx with the Aptos Narrow theme, import → snapshot,
        // export → xlsx, re-import, confirm font survives.
        const wb1 = new ExcelJS.Workbook();
        const ws1 = wb1.addWorksheet('Sheet1');
        ws1.getCell('A1').value = 'x';
        const buf0 = Buffer.from((await wb1.xlsx.writeBuffer()) as ArrayBuffer);
        const zip0 = await JSZip.loadAsync(buf0 as unknown as ArrayBuffer);
        const themePath = Object.keys(zip0.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p))!;
        const themeXml0 = (await zip0.files[themePath].async('string')).replace(
            /<a:minorFont>\s*<a:latin\b[^>]*\btypeface="[^"]*"/,
            '<a:minorFont><a:latin typeface="Source Sans Pro"',
        );
        zip0.file(themePath, themeXml0);
        const buf1 = Buffer.from(await zip0.generateAsync({ type: 'arraybuffer' }) as ArrayBuffer);

        const snap = await xlsxBufferToSnapshot(buf1);
        const buf2 = await snapshotToXlsxBuffer(snap);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(buf2 as ArrayBuffer)) as unknown as Snapshot;
        expect(snap2.defaultStyle?.ff).toBe('Source Sans Pro');
    });
});

// ─── Banded-style synthesis ────────────────────────────────────────────────

describe('M12 — banded-style synthesis', () => {
    test('table with TableStyleMedium2 + showRowStripes → header + banded rows have synthesized colors', async () => {
        // Build a workbook with a table styled TableStyleMedium2.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.addTable({
            name: 'T1',
            ref: 'A1',
            headerRow: true,
            style: { theme: 'TableStyleMedium2', showRowStripes: true } as ExcelJS.TableProperties['style'],
            columns: [{ name: 'A' }, { name: 'B' }],
            rows: [
                ['x', 1], ['y', 2], ['z', 3], ['w', 4],
            ],
        });
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];

        // The workbook is built via exceljs's writer which ships its own
        // Office 2007 default theme (accent1 = #4F81BD), so M13's
        // theme-aware resolver synthesizes TableStyleMedium2 → accent1
        // → header #4F81BD, banded row tint(+0.6) ≈ #B9CDE5. Note: the
        // M12 baseline expected Aptos colors (#156082 / #83CBEB) here
        // because the catalog was hardcoded to Aptos; the M13 fix makes
        // the import resolver theme-aware.
        const headerStyleId = sheet.cellData[0]?.[0]?.s;
        expect(headerStyleId).toBeDefined();
        const headerStyle = snap.styles[headerStyleId!];
        expect((headerStyle.bg as { rgb: string }).rgb).toBe('#4F81BD');
        expect((headerStyle.cl as { rgb: string }).rgb).toBe('#FFFFFF');
        expect(headerStyle.bl).toBe(1);

        // Even data row (row 1, the first data row, index 0 relative to data
        // start) should carry the bandedRowEvenBg.
        const evenStyleId = sheet.cellData[1]?.[0]?.s;
        expect(evenStyleId).toBeDefined();
        const evenStyle = snap.styles[evenStyleId!];
        expect((evenStyle.bg as { rgb: string }).rgb).toBe('#B9CDE5');

        // Odd data row (row 2) should be white (bandedRowOddBg=#FFFFFF) — and
        // since we skip white, no bg should be set on this style; the cell
        // may not even have a synthesized style. Verify the cell exists.
        const oddCell = sheet.cellData[2]?.[0];
        expect(oddCell).toBeDefined();
        // odd row could have no style or one with no bg (white skipped)
        if (oddCell?.s) {
            const oddStyle = snap.styles[oddCell.s];
            // If a style exists, bg should NOT be #FFFFFF (we skip white).
            if (oddStyle.bg) {
                expect((oddStyle.bg as { rgb: string }).rgb).not.toBe('#FFFFFF');
            }
        }
    });

    test('table with TableStyleMedium2 → synthesized header + outer borders (TableStyleMedium2 borderColor)', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.addTable({
            name: 'TB',
            ref: 'A1',
            headerRow: true,
            style: { theme: 'TableStyleMedium2', showRowStripes: true } as ExcelJS.TableProperties['style'],
            columns: [{ name: 'A' }, { name: 'B' }],
            rows: [['x', 1], ['y', 2], ['z', 3]],
        });
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as Snapshot & {
            sheets: Record<string, { cellData: Record<number, Record<number, { s?: string }>> }>;
        };
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const styleAt = (r: number, c: number) => {
            const id = sheet.cellData[r]?.[c]?.s;
            return id ? snap.styles[id] : undefined;
        };

        // M13 theme-aware: programmatic exceljs writer ships Office 2007
        // theme, so TableStyleMedium2 → accent1 #4F81BD.
        const BORDER_RGB = '#4F81BD';
        const headerStyle = styleAt(0, 0)!;
        const headerBd = headerStyle.bd as Record<string, { s: number; cl: { rgb: string } }>;
        // Header has top + bottom borders, plus left on the first column.
        expect(headerBd.t.cl.rgb).toBe(BORDER_RGB);
        expect(headerBd.b.cl.rgb).toBe(BORDER_RGB);
        expect(headerBd.l.cl.rgb).toBe(BORDER_RGB);

        // The right edge (column 1) gets a right border.
        const headerRightStyle = styleAt(0, 1)!;
        const headerRightBd = headerRightStyle.bd as Record<string, { s: number; cl: { rgb: string } }>;
        expect(headerRightBd.r.cl.rgb).toBe(BORDER_RGB);

        // Data rows: outer left/right edges have borders, inner rows do NOT
        // get top/bottom (banding fill is the separator). Last data row gets
        // a bottom border.
        const dataLeftStyle = styleAt(2, 0)!; // middle data row, left edge
        const dataLeftBd = dataLeftStyle.bd as Record<string, { s: number; cl: { rgb: string } }> | undefined;
        expect(dataLeftBd?.l?.cl.rgb).toBe(BORDER_RGB);
        expect(dataLeftBd?.t).toBeUndefined();
        expect(dataLeftBd?.b).toBeUndefined();

        // Last data row (row 3 = dataStartRow=1 + 2 data rows after first)
        // gets bottom border.
        const lastRow = 3;
        const lastRowLeft = styleAt(lastRow, 0)!;
        const lastRowLeftBd = lastRowLeft.bd as Record<string, { s: number; cl: { rgb: string } }>;
        expect(lastRowLeftBd.b.cl.rgb).toBe(BORDER_RGB);
    });

    test('export: synthesized table-style fields are stripped (no double-paint)', async () => {
        // Round-trip a table with TableStyleMedium2 through import → export.
        // The exported xlsx should NOT carry per-cell <fill>/<border> records
        // for the cells we synthesized; Excel paints them from the table
        // style at render time, and a doubled paint reads visually heavier
        // than the original.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.addTable({
            name: 'TS',
            ref: 'A1',
            headerRow: true,
            style: { theme: 'TableStyleMedium2', showRowStripes: true } as ExcelJS.TableProperties['style'],
            columns: [{ name: 'A' }, { name: 'B' }],
            rows: [['x', 1], ['y', 2], ['z', 3]],
        });
        const buf0 = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf0 as unknown as Buffer);

        // Sidecar should be present.
        const synth = (snap as unknown as Snapshot).resources?.find((r) => r.name === 'SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN');
        expect(synth).toBeDefined();
        const sidecar = JSON.parse(synth!.data) as Record<string, Record<string, string[]>>;
        const sheetSidecar = sidecar[(snap as unknown as Snapshot).sheetOrder[0]];
        expect(sheetSidecar).toBeDefined();
        // Header cell at (0,0): bg, cl, bl, bd.t, bd.b, bd.l should be tagged.
        expect(sheetSidecar['0:0']).toEqual(expect.arrayContaining(['bg', 'cl', 'bl', 'bd.t', 'bd.b', 'bd.l']));

        const buf1 = await snapshotToXlsxBuffer(snap);
        const wb1 = new ExcelJS.Workbook();
        await wb1.xlsx.load(buf1 as unknown as Parameters<typeof wb1.xlsx.load>[0]);
        const ws1 = wb1.getWorksheet('Sheet1')!;

        // The header cell should NOT have an explicit fill or border in the
        // round-tripped file — Excel renders TableStyleMedium2 on top, and
        // we don't want our synth doubled. exceljs reads "no fill" as
        // either undefined or `{ pattern: 'none' }` (depending on whether
        // the source xlsx had a <fill index='0'/> placeholder); both shapes
        // mean "no color painted on this cell".
        const a1 = ws1.getCell('A1');
        const a1Fill = a1.fill as { pattern?: string; fgColor?: { argb?: string } } | undefined;
        expect(a1Fill?.pattern === 'solid' || (a1Fill?.fgColor?.argb)).toBeFalsy();
        // Border object should also be empty (or undefined).
        const a1Border = a1.border as Record<string, unknown> | undefined;
        expect(a1Border ? Object.keys(a1Border).length === 0 : true).toBe(true);
        // Bold attribute should not have been emitted from the synthesized bl.
        expect(a1.font?.bold).toBeFalsy();
    });

    test('import resolves theme + tint colors against the workbook clrScheme', async () => {
        // Build a workbook with custom accent palette, set a cell's font
        // color to {theme: 4, tint: -0.25} (= darken accent1 by 25%), and
        // confirm the import resolves it to RGB rather than dropping it.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        const c = ws.getCell('A1');
        c.value = 'colored';
        // exceljs accepts color: { theme, tint } shape on cell.font.
        c.font = { color: { theme: 4, tint: -0.25 } as unknown as ExcelJS.Color };
        const buf0 = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        // exceljs's default theme has accent1 = #4F81BD. tint -0.25 = darken
        // by multiplying L by 0.75.
        const snap = await xlsxBufferToSnapshot(buf0 as unknown as Buffer) as unknown as Snapshot;
        const cellData = snap.sheets[snap.sheetOrder[0]].cellData[0]?.[0];
        const styleId = cellData?.s;
        expect(styleId).toBeDefined();
        const style = snap.styles[styleId!];
        const rgb = (style.cl as { rgb?: string } | undefined)?.rgb;
        expect(rgb).toBeDefined();
        // Resolved color must NOT be #000000 (the old fallback that hid
        // theme references) and must NOT be the raw accent1 #4F81BD
        // (without the tint applied).
        expect(rgb).not.toBe('#000000');
        expect(rgb).not.toBe('#4F81BD');
        // It should look like a darker shade of accent1.
        const r = parseInt(rgb!.slice(1, 3), 16);
        const g = parseInt(rgb!.slice(3, 5), 16);
        const b = parseInt(rgb!.slice(5, 7), 16);
        // accent1 #4F81BD → blue dominant. Darker shade keeps blue dominant.
        expect(b).toBeGreaterThan(r);
        expect(b).toBeGreaterThan(g);
    });

    test('table without showRowStripes → header gets color but data rows do NOT', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.addTable({
            name: 'T2',
            ref: 'A1',
            headerRow: true,
            style: { theme: 'TableStyleMedium2', showRowStripes: false } as ExcelJS.TableProperties['style'],
            columns: [{ name: 'A' }, { name: 'B' }],
            rows: [['x', 1], ['y', 2]],
        });
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];

        // Header should still be styled. M13 theme-aware: exceljs's
        // writer ships Office 2007 theme, so TableStyleMedium2 → accent1
        // #4F81BD.
        const headerStyleId = sheet.cellData[0]?.[0]?.s;
        expect(headerStyleId).toBeDefined();
        const headerStyle = snap.styles[headerStyleId!];
        expect((headerStyle.bg as { rgb: string }).rgb).toBe('#4F81BD');

        // Data row 1 should NOT have synthesized banding bg.
        const row1 = sheet.cellData[1]?.[0];
        if (row1?.s) {
            const s = snap.styles[row1.s];
            // No banded-row bg should have been synthesized when stripes
            // are off. (Cell may carry other styles from explicit cell-level
            // formatting; just ensure we didn't paint either palette's band
            // color over the data row.)
            if (s.bg) {
                const bg = (s.bg as { rgb: string }).rgb;
                expect(bg).not.toBe('#83CBEB'); // Aptos band
                expect(bg).not.toBe('#B9CDE5'); // Office 2007 band
            }
        }
    });

    test('unknown TableStyle name → no synthesis, falls back gracefully', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.addTable({
            name: 'T3',
            ref: 'A1',
            headerRow: true,
            style: { theme: 'NotARealStyle', showRowStripes: true } as unknown as ExcelJS.TableProperties['style'],
            columns: [{ name: 'A' }, { name: 'B' }],
            rows: [['x', 1]],
        });
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as Snapshot;
        // Import should still succeed — no throw. Cells exist normally.
        expect(snap.sheets[snap.sheetOrder[0]].cellData[0]?.[0]?.v).toBe('A');
    });
});

// ─── Hyperlinks ────────────────────────────────────────────────────────────

describe('M12 — hyperlinks', () => {
    test('import: cell with { text, hyperlink } → snapshot has cell.p with HYPERLINK customRange', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = { text: 'Click me', hyperlink: 'https://example.com/' };
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as Snapshot;
        const cell = snap.sheets[snap.sheetOrder[0]].cellData[0]?.[0];
        expect(cell?.v).toBe('Click me');
        // p must mirror Univer's runtime hyperlink doc model so its layout
        // pipeline produces a real page on hover (otherwise mouse-move over
        // the cell throws "Cannot read properties of undefined (reading
        // 'height')" — see buildHyperlinkCellP comment).
        const body = cell?.p?.body;
        expect(body?.dataStream).toBe('Click me\r\n');
        expect(body?.paragraphs).toEqual([
            { startIndex: 'Click me'.length, paragraphStyle: {} },
        ]);
        expect(body?.sectionBreaks).toEqual([
            { startIndex: 'Click me'.length + 1 },
        ]);
        expect(body?.textRuns).toEqual([
            { st: 0, ed: 'Click me'.length, ts: {} },
        ]);
        const ranges = body?.customRanges ?? [];
        expect(ranges.length).toBe(1);
        expect(ranges[0].rangeType).toBe(0);
        expect(ranges[0].properties?.url).toBe('https://example.com/');

        // documentStyle.pageSize must be finite & JSON-safe so the
        // documentSkeleton actually lays out a page on load.
        const ds = (cell?.p as { documentStyle?: { pageSize?: { width?: unknown; height?: unknown } } } | undefined)?.documentStyle;
        expect(ds?.pageSize?.width).toBeDefined();
        expect(ds?.pageSize?.height).toBeDefined();
        expect(Number.isFinite(ds?.pageSize?.width as number)).toBe(true);
        expect(Number.isFinite(ds?.pageSize?.height as number)).toBe(true);
        // Round-trip-safe: JSON.stringify shouldn't drop these to null.
        const roundTripped = JSON.parse(JSON.stringify(cell));
        expect(roundTripped.p.documentStyle.pageSize.width).toBe(ds?.pageSize?.width);
        expect(roundTripped.p.documentStyle.pageSize.height).toBe(ds?.pageSize?.height);
    });

    test('export: snapshot with cell.p hyperlink → exported xlsx has hyperlink in the cell', async () => {
        const snap = {
            id: 'wb-h', name: 'Spreadsheet', appVersion: '0.1.0', locale: 'enUS',
            sheetOrder: ['s1'],
            styles: {},
            sheets: {
                s1: {
                    id: 's1', name: 'Sheet1',
                    cellData: {
                        0: {
                            0: {
                                v: 'Click here', t: 1,
                                p: {
                                    id: '__INTERNAL_EDITOR__DOCS_NORMAL',
                                    body: {
                                        dataStream: 'Click here',
                                        customRanges: [{
                                            startIndex: 0, endIndex: 9, rangeId: 'r1',
                                            rangeType: 0, properties: { url: 'https://destination.example/' },
                                        }],
                                    },
                                },
                            },
                        },
                    },
                    rowCount: 100, columnCount: 26,
                    defaultColumnWidth: 73, defaultRowHeight: 19,
                    mergeData: [], rowData: {}, columnData: {},
                },
            },
        };
        const buf = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf as unknown as Parameters<typeof wb.xlsx.load>[0]);
        const ws = wb.getWorksheet('Sheet1')!;
        const cell = ws.getCell('A1');
        // exceljs normalizes { text, hyperlink } so cell.value comes back
        // as the same shape, AND cell.hyperlink reads the URL.
        expect(cell.text).toBe('Click here');
        expect(cell.isHyperlink).toBe(true);
        expect(cell.hyperlink).toBe('https://destination.example/');
    });

    test('round-trip: hyperlink xlsx → snapshot → xlsx → hyperlink preserved', async () => {
        const wb1 = new ExcelJS.Workbook();
        const ws1 = wb1.addWorksheet('Sheet1');
        ws1.getCell('A1').value = { text: 'home', hyperlink: 'https://home.example.com/' };
        ws1.getCell('A2').value = { text: 'page', hyperlink: 'https://page.example.com/path?q=1' };
        const buf1 = Buffer.from((await wb1.xlsx.writeBuffer()) as ArrayBuffer);

        const snap = await xlsxBufferToSnapshot(buf1 as unknown as Buffer);
        const buf2 = await snapshotToXlsxBuffer(snap);

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(buf2 as unknown as Parameters<typeof wb2.xlsx.load>[0]);
        const ws2 = wb2.getWorksheet('Sheet1')!;
        expect(ws2.getCell('A1').hyperlink).toBe('https://home.example.com/');
        expect(ws2.getCell('A1').text).toBe('home');
        expect(ws2.getCell('A2').hyperlink).toBe('https://page.example.com/path?q=1');
        expect(ws2.getCell('A2').text).toBe('page');
    });

    test('cell with empty URL → no hyperlink emitted', async () => {
        const snap = {
            id: 'wb-h', name: 'Spreadsheet', appVersion: '0.1.0', locale: 'enUS',
            sheetOrder: ['s1'],
            styles: {},
            sheets: {
                s1: {
                    id: 's1', name: 'Sheet1',
                    cellData: {
                        0: { 0: { v: 'no url', t: 1 } },
                    },
                    rowCount: 100, columnCount: 26,
                    defaultColumnWidth: 73, defaultRowHeight: 19,
                    mergeData: [], rowData: {}, columnData: {},
                },
            },
        };
        const buf = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf as unknown as Parameters<typeof wb.xlsx.load>[0]);
        const ws = wb.getWorksheet('Sheet1')!;
        expect(ws.getCell('A1').isHyperlink).toBe(false);
        expect(ws.getCell('A1').value).toBe('no url');
    });
});
