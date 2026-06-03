// M13 workstream D: rich-text within a single cell.
//
// Excel can store multi-run formatting INSIDE a single cell — e.g. the
// word "Hello" bold + " world" plain, all in one cell. exceljs surfaces
// this as `cell.value = { richText: [{font, text}, ...] }`. Univer's
// model carries the same shape via `cell.p.body.textRuns`, an array of
// { st, ed, ts } entries where `ts` is an ITextStyle (extends IStyleBase
// with the same ff/fs/bl/it/ul/cl/... fields we already use for
// cell-level styles).
//
// Before M13: our import flattened richText to plain text, losing all
// per-run formatting. KNOWN SHORTCOMING tests in m12FixtureRoundTrip
// pinned that down.
//
// After M13: import emits cell.p.body.textRuns + paragraphs +
// sectionBreaks + finite pageSize (same shape as buildHyperlinkCellP).
// Export reverses the flow when cell.p has multiple textRuns AND no
// hyperlink customRange.
//
// Edge cases this file pins down:
//   1. Bold-only run + plain run (the canonical "Hello bold" case)
//   2. Multi-color runs (red + plain + blue + plain)
//   3. Mixed format: bold + italic + colored on the same run
//   4. Underline run (must NOT collide with hyperlink path)
//   5. cell.v still carries the plain-text concatenation so existing
//      string-only readers keep working
//   6. Round-trip: rich-text → snapshot → export → re-import preserves
//      the run boundaries and per-run formatting
//   7. A single-run "rich text" (richText with one element) collapses
//      back to plain text on import — no need to emit cell.p

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';

import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURE = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata', 'RichTextInOneCell.xlsx');

interface Body {
    dataStream?: string;
    paragraphs?: Array<{ startIndex: number }>;
    sectionBreaks?: Array<{ startIndex: number }>;
    textRuns?: Array<{ st: number; ed: number; ts?: Record<string, unknown> }>;
    customRanges?: Array<{ rangeType?: number; properties?: { url?: string } }>;
}

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, {
            v?: unknown;
            t?: number;
            s?: string;
            p?: { body?: Body; documentStyle?: { pageSize?: { width?: number; height?: number } } };
        }>>;
    }>;
    styles: Record<string, Record<string, unknown>>;
}

async function importFixture(file: string): Promise<SnapshotShape> {
    const buf = readFileSync(file);
    return await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;
}

async function importBuf(buf: Buffer): Promise<SnapshotShape> {
    return await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;
}

describe('M13 — rich-text within a single cell', () => {
    test('fixture A1: "Hello"(bold) + " world"(plain) → 2 textRuns with bl=1 on first', async () => {
        const snap = await importFixture(FIXTURE);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const cell = sheet.cellData[0]?.[0];
        // Plain-text concatenation still on cell.v so existing readers
        // (snapshot dumpers, html exporters, etc.) keep working.
        expect(cell?.v).toBe('Hello world');
        // Multi-run rich text MUST emit cell.p with proper documentSkeleton
        // shape — same invariants as buildHyperlinkCellP so Univer's
        // layout pipeline doesn't crash on hover.
        expect(cell?.p).toBeDefined();
        const body = cell!.p!.body!;
        expect(body.dataStream).toBe('Hello world\r\n');
        expect(body.paragraphs).toEqual([{ startIndex: 'Hello world'.length, paragraphStyle: {} }]);
        expect(body.sectionBreaks).toEqual([{ startIndex: 'Hello world'.length + 1 }]);
        expect(Number.isFinite(cell!.p!.documentStyle!.pageSize!.width!)).toBe(true);
        // Two textRuns covering "Hello" (bold) and " world" (plain).
        const runs = body.textRuns!;
        expect(runs).toHaveLength(2);
        expect(runs[0]).toMatchObject({ st: 0, ed: 5, ts: { bl: 1 } });
        expect(runs[1]).toMatchObject({ st: 5, ed: 11 });
        // Plain run carries no formatting (or an empty ts).
        const plainTs = runs[1].ts ?? {};
        expect(plainTs.bl).toBeUndefined();
    });

    test('fixture A2: "Red"(red) + " and " + "Blue"(blue) + " text" → 4 textRuns with cl on color runs', async () => {
        const snap = await importFixture(FIXTURE);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const cell = sheet.cellData[1]?.[0];
        expect(cell?.v).toBe('Red and Blue text');
        const runs = cell?.p?.body?.textRuns ?? [];
        expect(runs).toHaveLength(4);
        // Run 0: "Red" with cl #FF0000.
        expect(runs[0]).toMatchObject({ st: 0, ed: 3 });
        expect((runs[0].ts?.cl as { rgb?: string } | undefined)?.rgb).toBe('#FF0000');
        // Run 1: " and " with no color.
        expect(runs[1]).toMatchObject({ st: 3, ed: 8 });
        expect(runs[1].ts?.cl).toBeUndefined();
        // Run 2: "Blue" with cl #0000FF.
        expect(runs[2]).toMatchObject({ st: 8, ed: 12 });
        expect((runs[2].ts?.cl as { rgb?: string } | undefined)?.rgb).toBe('#0000FF');
        // Run 3: " text" with no color.
        expect(runs[3]).toMatchObject({ st: 12, ed: 17 });
    });

    test('fixture A3: hyperlink+plain (Pattern A) keeps cell.p with hyperlink customRange — NOT richText', async () => {
        // A3 is a single-format string with cell.hyperlink set. exceljs
        // reports cell.value = "Visit example.com for more info" plus
        // cell.hyperlink = "https://example.com" — Pattern A. The
        // import path produces the hyperlink cell.p shape (NOT a
        // multi-run rich-text shape). This test guards against the
        // rich-text code accidentally consuming Pattern A cells.
        const snap = await importFixture(FIXTURE);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const cell = sheet.cellData[2]?.[0];
        expect(cell?.p).toBeDefined();
        // Has a HYPERLINK customRange (rangeType=0).
        const ranges = cell?.p?.body?.customRanges ?? [];
        const link = ranges.find((r) => r.rangeType === 0);
        expect(link?.properties?.url).toBe('https://example.com');
        // textRuns is either absent or carries the same single run
        // covering the whole text — both shapes mean "single format,
        // not multi-run rich text".
        const runs = cell?.p?.body?.textRuns ?? [];
        if (runs.length > 0) {
            expect(runs).toHaveLength(1);
        }
    });

    test('round-trip: rich-text exports back to richText, re-imports identical', async () => {
        const snap = await importFixture(FIXTURE);
        const buf2 = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf2 as unknown as Parameters<typeof wb.xlsx.load>[0]);
        const ws = wb.getWorksheet(1)!;
        // A1: bold + plain.
        const a1 = ws.getCell('A1').value as { richText?: Array<{ font?: { bold?: boolean }; text: string }> };
        expect(a1?.richText).toBeDefined();
        expect(a1.richText).toHaveLength(2);
        expect(a1.richText![0]).toMatchObject({ text: 'Hello', font: expect.objectContaining({ bold: true }) });
        expect(a1.richText![1]).toMatchObject({ text: ' world' });
        // A2: red + plain + blue + plain.
        const a2 = ws.getCell('A2').value as { richText?: Array<{ font?: { color?: { argb?: string } }; text: string }> };
        expect(a2?.richText).toHaveLength(4);
        expect(a2.richText![0].font?.color?.argb?.toUpperCase().slice(2)).toBe('FF0000');
        expect(a2.richText![2].font?.color?.argb?.toUpperCase().slice(2)).toBe('0000FF');
    });

    test('mixed format on one run: bold + italic + color all preserve', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = {
            richText: [
                {
                    font: { bold: true, italic: true, color: { argb: 'FF00AA00' } } as ExcelJS.Font,
                    text: 'fancy',
                },
                { text: ' plain' },
            ],
        };
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf);
        const cell = snap.sheets[snap.sheetOrder[0]].cellData[0]?.[0];
        const runs = cell?.p?.body?.textRuns ?? [];
        expect(runs).toHaveLength(2);
        const ts0 = runs[0].ts ?? {};
        expect(ts0.bl).toBe(1);
        expect(ts0.it).toBe(1);
        expect((ts0.cl as { rgb?: string } | undefined)?.rgb).toBe('#00AA00');
    });

    test('underline run inside rich text: ul=1 in textRun.ts, NOT mistaken for a hyperlink', async () => {
        // Underline-only is the canonical "looks like a link but isn't"
        // case. Rich-text import must emit ul on the run's ts; the
        // hyperlink customRange must NOT be added.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = {
            richText: [
                { font: { underline: true } as ExcelJS.Font, text: 'underlined' },
                { text: ' rest' },
            ],
        };
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf);
        const cell = snap.sheets[snap.sheetOrder[0]].cellData[0]?.[0];
        const runs = cell?.p?.body?.textRuns ?? [];
        expect(runs).toHaveLength(2);
        const ul = runs[0].ts?.ul as { s?: number } | undefined;
        expect(ul?.s).toBe(1);
        // No hyperlink customRange.
        const ranges = cell?.p?.body?.customRanges ?? [];
        const link = ranges.find((r) => r.rangeType === 0);
        expect(link).toBeUndefined();
    });

    test('single-run richText collapses to plain text (no cell.p, no textRuns)', async () => {
        // exceljs sometimes emits richText with one entry even when the
        // cell has no multi-run formatting. We treat these as plain
        // strings — emitting cell.p just for one uniformly-styled run
        // would bloat the snapshot for no benefit.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = {
            richText: [{ text: 'just plain' }],
        };
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf);
        const cell = snap.sheets[snap.sheetOrder[0]].cellData[0]?.[0];
        expect(cell?.v).toBe('just plain');
        expect(cell?.p).toBeUndefined();
    });

    test('round-trip: pure plain-text cell still exports as plain string (no false richText)', async () => {
        // Sanity: the export-side rich-text emission must NOT trigger
        // for ordinary cells. Otherwise every plain-text cell after a
        // round-trip would be wrapped in {richText: [{text}]} which
        // exceljs's reader handles but is a snapshot bloat we don't
        // want.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'plain';
        const buf0 = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf0);
        const buf1 = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const wb1 = new ExcelJS.Workbook();
        await wb1.xlsx.load(buf1 as unknown as Parameters<typeof wb1.xlsx.load>[0]);
        const ws1 = wb1.getWorksheet('Sheet1')!;
        const v = ws1.getCell('A1').value;
        expect(typeof v).toBe('string');
        expect(v).toBe('plain');
    });
});
