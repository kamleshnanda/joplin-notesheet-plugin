// Regression tests for export-correctness bugs surfaced during M18 manual
// testing (operator round-tripped real fixtures through import → export →
// reopen-in-Excel and reported broken output). Each test anchors to the
// SOURCE fixture XML, not to our own emit, and drives the real
// xlsxBufferToSnapshot → snapshotToXlsxBuffer pipeline.
//
//   #4 — array (CSE) formula `=A2:A6*2` became `=@A2:A6*2` in Excel.
//        Source F7 is <f t="array" ref="F7:F11">A2:A6*2</f>; the array
//        marker + spill ref must survive export or Excel reinterprets it as
//        an implicit-intersection formula.
//   #5 — table totals row (SUBTOTAL) produced a circular-reference error.
//        Source columns carry totalsRowFunction="sum"/"average"/"count";
//        dropping them (writing totalsRowFunction="none" + no totals
//        formula) makes the structured ref include the totals cell itself.

import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

const FX = path.resolve(__dirname, 'fixtures/formatting-testdata');
const IMG_FX = path.resolve(__dirname, 'fixtures/images');

async function roundTripDir(dir: string, name: string): Promise<JSZip> {
    const buf = fs.readFileSync(path.join(dir, name));
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
    const snap = await xlsxBufferToSnapshot(ab);
    const out = await snapshotToXlsxBuffer(snap);
    return JSZip.loadAsync(out);
}

async function roundTripToZip(name: string): Promise<JSZip> {
    return roundTripDir(FX, name);
}

describe('#4 — array (CSE) formula survives export', () => {
    test('F7 exports as <f t="array" ref="F7:F11"> not a plain formula', async () => {
        const zip = await roundTripToZip('FormulasAndStructuredRefs.xlsx');
        const sheet1 = (await zip.file('xl/worksheets/sheet1.xml')?.async('string')) ?? '';
        // The F7 cell must carry the array-formula marker + spill ref.
        const f7 = sheet1.match(/<c r="F7"[^>]*>[\s\S]*?<\/c>/)?.[0] ?? '';
        expect(f7).toMatch(/<f\b[^>]*\bt="array"/);
        expect(f7).toMatch(/\bref="F7:F11"/);
        // And the formula body must NOT carry an implicit-intersection @.
        expect(f7).toContain('A2:A6*2');
        expect(f7).not.toContain('@');
    });
});

describe('#5 — table totals-row functions survive export (no circular ref)', () => {
    test('Smorgasboard ProjectTracker columns keep their totalsRowFunction', async () => {
        const zip = await roundTripToZip('FormattingSmorgasboard.xlsx');
        const tableXml = (await zip.file('xl/tables/table1.xml')?.async('string')) ?? '';
        // Source declares: Budget=sum, Spent=sum, % Complete=average,
        // Status=count, and a totalsRowLabel on the first column.
        expect(tableXml).toMatch(/name="Budget"[^>]*totalsRowFunction="sum"/);
        expect(tableXml).toMatch(/name="Spent"[^>]*totalsRowFunction="sum"/);
        expect(tableXml).toMatch(/name="% Complete"[^>]*totalsRowFunction="average"/);
        expect(tableXml).toMatch(/name="Status"[^>]*totalsRowFunction="count"/);
        expect(tableXml).toMatch(/totalsRowLabel=/);
    });

    test('Status totals cell exports a SUBTOTAL scoped to the table column', async () => {
        const zip = await roundTripToZip('FormattingSmorgasboard.xlsx');
        const sheet1 = (await zip.file('xl/worksheets/sheet1.xml')?.async('string')) ?? '';
        // G10 (Status totals) must be a SUBTOTAL over the structured column,
        // which Excel scopes to the data body (excludes the totals cell) —
        // this is what prevents the circular-reference error.
        expect(sheet1).toMatch(/SUBTOTAL\(103,ProjectTracker\[Status\]\)/);
    });

    test('totals-row formula cells keep their CACHED VALUE (no recalc-on-open)', async () => {
        // The earlier version of this fix asserted only the formula TEXT and
        // still shipped the circular-ref bug: exceljs's Table.store() rewrote
        // the totals cells as <f>…</f> with NO <v>, and a value-less totals
        // formula forces Excel to recalc on open, which mis-resolves the
        // structured ref into a self-reference. The cached <v> is the actual
        // fix — assert it survives for every totals-row formula cell.
        const zip = await roundTripToZip('FormattingSmorgasboard.xlsx');
        const sheet1 = (await zip.file('xl/worksheets/sheet1.xml')?.async('string')) ?? '';
        const row10 = sheet1.match(/<row r="10"[\s\S]*?<\/row>/)?.[0] ?? '';
        // Every totals cell that carries an <f> must also carry a <v>.
        const cellsWithFormula =
            row10.match(/<c\b[^>]*>(?:(?!<\/c>)[\s\S])*?<f>[\s\S]*?<\/c>/g) ?? [];
        expect(cellsWithFormula.length).toBeGreaterThanOrEqual(4); // C/D/E/G totals
        for (const cell of cellsWithFormula) {
            expect(cell).toMatch(/<f>[\s\S]*?<\/f>\s*<v>[\s\S]*?<\/v>/);
        }
        // Status totals specifically: SUBTOTAL(103,…) with cached count 8.
        expect(row10).toMatch(/SUBTOTAL\(103,ProjectTracker\[Status\]\)<\/f><v>8<\/v>/);
    });

    test('workbook calcId is not the stale exceljs default (no forced recalc)', async () => {
        const zip = await roundTripToZip('FormattingSmorgasboard.xlsx');
        const wb = (await zip.file('xl/workbook.xml')?.async('string')) ?? '';
        // exceljs hard-codes calcId="171027" (older than modern Excel); that
        // stale stamp forces a recalc-on-open. We bump it so Excel trusts the
        // cached results.
        expect(wb).not.toMatch(/calcId="171027"/);
        expect(wb).toMatch(/calcId="191029"/);
    });
});

describe('#1 — styled image glow/shadow frame survives export', () => {
    // Multi-sheet-StlyedImages.xlsx drawing1 (Styled Image sheet) carries an
    // <a:effectLst> with a glow (accent2) + outer shadow. Univer 0.23 can't
    // render image effects live, but the round-tripped .xlsx must still show
    // the frame in Excel (preserve-only).
    async function findDrawingWithEffects(zip: JSZip): Promise<string> {
        const drawingPaths = Object.keys(zip.files).filter((p) =>
            /^xl\/drawings\/drawing\d+\.xml$/i.test(p),
        );
        const bodies: string[] = [];
        for (const p of drawingPaths) {
            bodies.push((await zip.file(p)?.async('string')) ?? '');
        }
        return bodies.join('\n');
    }

    test('exported drawing keeps the <a:effectLst> glow + shadow', async () => {
        const zip = await roundTripDir(IMG_FX, 'Multi-sheet-StlyedImages.xlsx');
        const drawings = await findDrawingWithEffects(zip);
        expect(drawings).toContain('<a:effectLst>');
        // The glow radius + outer shadow from the source must survive.
        expect(drawings).toMatch(/<a:glow\b/);
        expect(drawings).toMatch(/<a:outerShdw\b/);
    });
});
