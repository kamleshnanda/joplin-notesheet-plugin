// Bidirectional round-trip integrity test.
//
// PURPOSE
//   The existing fidelity tests verify that a fixture imports the same
//   way every time, AND that exporting+re-importing returns to the
//   same snapshot. Both gates are static — the workbook never changes.
//
//   Real users do change things. They open a Joplin notesheet, edit
//   cell values, change colours, then export to .xlsx. They open
//   that .xlsx in Excel, edit more cells, save, then drag the file
//   back into Joplin. The workbook flows through both editors at
//   least once each.
//
//   This test exercises that bidirectional flow:
//
//     Aptos.xlsx
//       → snap1 (initial import)
//       → mutate snap1: simulate Joplin user edits (content + style)
//       → snap1Edited
//       → export to xlsx buffer (buf1)
//       → mutate buf1 via exceljs: simulate Excel user edits
//       → buf2
//       → re-import → snap2
//       → assert: snap2 reflects BOTH the Joplin edits AND the Excel
//         edits, in the right cells, with no other cells drifting.
//
//   Then for safety: do the cycle TWICE and confirm a tracked cell's
//   value doesn't drift across cycles (synth-stripping is idempotent).
//
// LIMITATIONS
//   - "Excel edits" are simulated via exceljs. exceljs does not have
//     Excel.app's full rendering / formula recalc, but it deterministi-
//     cally models the file-level mutations a user would produce.
//   - This test does not open the file in Excel.app. The visual final
//     check requires a manual eyeball pass; the assertion here is
//     strictly about file-level data fidelity.

import path from 'path';
import { readFileSync } from 'fs';
import ExcelJS from 'exceljs';
import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const APTOS = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata', 'FormattingSmorgasboard.xlsx');

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, { v?: unknown; s?: string; p?: unknown }>>;
    }>;
    styles: Record<string, Record<string, unknown>>;
}

function getCell(snap: SnapshotShape, row: number, col: number): { v?: unknown; s?: string; p?: unknown } | undefined {
    return snap.sheets[snap.sheetOrder[0]].cellData[row]?.[col];
}

function getStyle(snap: SnapshotShape, row: number, col: number): Record<string, unknown> | undefined {
    const cell = getCell(snap, row, col);
    return cell?.s ? snap.styles[cell.s] : undefined;
}

function getCellValueText(cell: { v?: unknown; p?: unknown } | undefined): string | undefined {
    if (!cell) return undefined;
    if (cell.v !== undefined && cell.v !== null) return String(cell.v);
    // Rich text: pull text out of p.body.dataStream (Univer's IDocumentData shape).
    const p = cell.p as { body?: { dataStream?: string } } | undefined;
    if (p?.body?.dataStream) {
        // Strip the trailing \r\n marker Univer appends.
        return p.body.dataStream.replace(/\r\n?$/, '');
    }
    return undefined;
}

describe('Bidirectional round-trip integrity (Aptos TableStyleMedium4)', () => {
    test('Joplin edit + Excel edit flow through both directions correctly', async () => {
        // ---- Stage 1: initial import ----
        const buf0 = readFileSync(APTOS);
        const snap1 = await xlsxBufferToSnapshot(buf0 as unknown as Buffer) as unknown as SnapshotShape;

        // Capture original state for cells we'll edit so we can detect drift.
        const origB2Value = getCellValueText(getCell(snap1, 1, 1));   // "https://example.com/alpha"
        const origC2Value = getCellValueText(getCell(snap1, 1, 2));   // 50000 (Budget)

        expect(origB2Value).toContain('alpha');  // sanity
        expect(origC2Value).toBe('50000');

        // ---- Stage 2: simulate Joplin user edits (content + style) ----
        const sheet = snap1.sheets[snap1.sheetOrder[0]];

        // Joplin edit 1: change B2 value to a new URL.
        // Univer cell-data shape: keep s, replace v.
        const b2 = sheet.cellData[1][1] as { v?: unknown; s?: string };
        b2.v = 'https://newurl.com/joplin-edited';

        // Joplin edit 2: add a custom font colour to C2 (red text).
        // We have to materialise this as a NEW style id since the existing
        // style is shared. The simplest path: create a new style record.
        const newRedStyleId = 'joplin-red-' + Math.floor(Date.now() % 1e9);
        snap1.styles[newRedStyleId] = {
            ...(getStyle(snap1, 1, 2) ?? {}),
            cl: { rgb: '#FF0000' },
        };
        sheet.cellData[1][2] = { ...(sheet.cellData[1][2] ?? {}), s: newRedStyleId };

        // ---- Stage 3: export ----
        const buf1 = await snapshotToXlsxBuffer(snap1 as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);

        // ---- Stage 4: simulate Excel user edits (content + style) ----
        const wb1 = new ExcelJS.Workbook();
        await wb1.xlsx.load(buf1 as unknown as Parameters<typeof wb1.xlsx.load>[0]);
        const ws1 = wb1.getWorksheet('Sheet1') ?? wb1.worksheets[0];

        // Excel edit 1: change D2 value (Spent) from 32000 → 99999.
        // Excel rows are 1-indexed; D = col 4. Snapshot row 1 = Excel row 2.
        ws1.getCell('D2').value = 99999;

        // Excel edit 2: add a yellow fill to E2 (% Complete).
        ws1.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' },
        };

        const buf2 = Buffer.from((await wb1.xlsx.writeBuffer()) as ArrayBuffer);

        // ---- Stage 5: re-import ----
        const snap2 = await xlsxBufferToSnapshot(buf2 as unknown as Buffer) as unknown as SnapshotShape;

        // ---- Stage 6: verify both sets of edits flowed through ----
        // Joplin's B2 edit must be present.
        const b2Round = getCellValueText(getCell(snap2, 1, 1));
        expect(b2Round).toBe('https://newurl.com/joplin-edited');

        // Joplin's C2 red-font edit must be present.
        const c2Style = getStyle(snap2, 1, 2);
        const c2cl = c2Style?.cl as { rgb?: string } | undefined;
        expect(c2cl?.rgb).toBe('#FF0000');

        // Excel's D2 value edit must be present.
        const d2Round = getCellValueText(getCell(snap2, 1, 3));
        expect(d2Round).toBe('99999');

        // Excel's E2 yellow fill must be present.
        const e2Style = getStyle(snap2, 1, 4);
        const e2bg = e2Style?.bg as { rgb?: string } | undefined;
        expect(e2bg?.rgb).toBe('#FFFF00');

        // Negative: cells we did NOT touch should not drift.
        // F2 (Start Date) stays at its original 46037 (epoch days).
        const f2Round = getCellValueText(getCell(snap2, 1, 5));
        expect(f2Round).toBe('46037');

        // The header row (row 0) should still be styled by TableStyleMedium4
        // (re-synthesized from the table-style declaration); header bg is
        // still #34692E.
        const headerStyle2 = getStyle(snap2, 0, 0);
        expect((headerStyle2?.bg as { rgb?: string } | undefined)?.rgb).toBe('#34692E');

        // The totals row's bd shape is re-synthesized: bd.t #34692E DOUBLE,
        // bd.b #72D068 MEDIUM. Synth fields from the first cycle were
        // stripped on export; the table-style declaration drove re-synthesis
        // on re-import.
        const totalsStyle2 = getStyle(snap2, 9, 0);
        const totalsBd = totalsStyle2?.bd as { t?: { s: number; cl: { rgb: string } }; b?: { s: number; cl: { rgb: string } } } | undefined;
        expect(totalsBd?.t?.cl.rgb).toBe('#34692E');
        expect(totalsBd?.t?.s).toBe(7);
        expect(totalsBd?.b?.cl.rgb).toBe('#72D068');
        expect(totalsBd?.b?.s).toBe(8);

        // Row 1 col 0 still has its user-set #F4B183 border (predates this
        // test's edits; was already in the source workbook).
        const userBorderStyle2 = getStyle(snap2, 1, 0);
        const ub = userBorderStyle2?.bd as { t?: { cl: { rgb: string } } } | undefined;
        expect(ub?.t?.cl.rgb).toBe('#F4B183');

        // Inter-row strip on row 2 col 1 is still re-synthesized.
        const interRowStyle2 = getStyle(snap2, 2, 1);
        const irBd = interRowStyle2?.bd as { t?: { s: number; cl: { rgb: string } } } | undefined;
        expect(irBd?.t?.cl.rgb).toBe('#72D068');
        expect(irBd?.t?.s).toBe(8);
    });

    test('synth fields do not accumulate or drift across two round-trips', async () => {
        // Two cycles import → export → re-import → export → re-import
        // without any user edits. The final snapshot should equal the
        // first import for the cells we check.
        const buf0 = readFileSync(APTOS);
        const snap1 = await xlsxBufferToSnapshot(buf0 as unknown as Buffer) as unknown as SnapshotShape;
        const buf1 = await snapshotToXlsxBuffer(snap1 as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const snap2 = await xlsxBufferToSnapshot(buf1 as unknown as Buffer) as unknown as SnapshotShape;
        const buf2 = await snapshotToXlsxBuffer(snap2 as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const snap3 = await xlsxBufferToSnapshot(buf2 as unknown as Buffer) as unknown as SnapshotShape;

        // Check: header bg is still the synth value, not "synth on top of
        // synth on top of synth".
        const h1 = getStyle(snap1, 0, 0);
        const h3 = getStyle(snap3, 0, 0);
        expect((h1?.bg as { rgb: string }).rgb).toBe('#34692E');
        expect((h3?.bg as { rgb: string }).rgb).toBe('#34692E');

        // Inter-row strip stays identical across cycles.
        const ir1 = getStyle(snap1, 2, 1);
        const ir3 = getStyle(snap3, 2, 1);
        const ir1Bd = ir1?.bd as { t?: { s: number; cl: { rgb: string } } } | undefined;
        const ir3Bd = ir3?.bd as { t?: { s: number; cl: { rgb: string } } } | undefined;
        expect(ir1Bd?.t?.cl.rgb).toBe('#72D068');
        expect(ir3Bd?.t?.cl.rgb).toBe('#72D068');
        expect(ir1Bd?.t?.s).toBe(8);
        expect(ir3Bd?.t?.s).toBe(8);

        // Totals row bd stays identical.
        const t1 = getStyle(snap1, 9, 0);
        const t3 = getStyle(snap3, 9, 0);
        const t1Bd = t1?.bd as { t?: { s: number; cl: { rgb: string } }; b?: { s: number; cl: { rgb: string } } } | undefined;
        const t3Bd = t3?.bd as { t?: { s: number; cl: { rgb: string } }; b?: { s: number; cl: { rgb: string } } } | undefined;
        expect(t1Bd?.t?.cl.rgb).toBe('#34692E');
        expect(t3Bd?.t?.cl.rgb).toBe('#34692E');
        expect(t1Bd?.t?.s).toBe(7);
        expect(t3Bd?.t?.s).toBe(7);
        expect(t1Bd?.b?.cl.rgb).toBe('#72D068');
        expect(t3Bd?.b?.cl.rgb).toBe('#72D068');

        // User-set border on row 1 col 0 still preserved across two cycles.
        const ub1 = getStyle(snap1, 1, 0);
        const ub3 = getStyle(snap3, 1, 0);
        const ub1Bd = ub1?.bd as { t?: { cl: { rgb: string } } } | undefined;
        const ub3Bd = ub3?.bd as { t?: { cl: { rgb: string } } } | undefined;
        expect(ub1Bd?.t?.cl.rgb).toBe('#F4B183');
        expect(ub3Bd?.t?.cl.rgb).toBe('#F4B183');
    });

    test('Excel-edit-only flow: synth fields do not block external edits to data cells', async () => {
        // No Joplin edits this time — just round-trip with an Excel edit
        // to a cell that DOES have a synth-applied bd.t (inter-row strip).
        // Confirms that adding a user fill to a cell with a synth border
        // doesn't get the synth border re-emitted as user formatting on
        // the next export, and the user fill does survive.
        const buf0 = readFileSync(APTOS);
        const snap1 = await xlsxBufferToSnapshot(buf0 as unknown as Buffer) as unknown as SnapshotShape;
        const buf1 = await snapshotToXlsxBuffer(snap1 as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);

        // Open in exceljs, add a fill to D5 (a banded data cell with a
        // synth bd.t).
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf1 as unknown as Parameters<typeof wb.xlsx.load>[0]);
        const ws = wb.getWorksheet('Sheet1') ?? wb.worksheets[0];
        ws.getCell('D5').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF00FFFF' },  // cyan
        };
        const buf2 = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);

        const snap2 = await xlsxBufferToSnapshot(buf2 as unknown as Buffer) as unknown as SnapshotShape;
        const d5Style = getStyle(snap2, 4, 3);

        // The user's cyan fill survives.
        expect((d5Style?.bg as { rgb?: string } | undefined)?.rgb).toBe('#00FFFF');

        // The inter-row strip (bd.t #72D068) is STILL re-synthesized on
        // re-import — synth runs every import, so D5's bd.t is back even
        // though we exported with it stripped.
        const d5Bd = d5Style?.bd as { t?: { s: number; cl: { rgb: string } } } | undefined;
        expect(d5Bd?.t?.cl.rgb).toBe('#72D068');
        expect(d5Bd?.t?.s).toBe(8);
    });
});
