// Pin-down: Univer is the source of truth for formula evaluation.
//
// The HTML / PDF / preview-pane renderer
// (`src/contentScripts/notesheetRenderer.ts`) reads `cell.v` (the
// cached evaluated value) and does NOT parse or evaluate `cell.f`
// (the formula text). This is a deliberate decision documented in
// the README "Known shortcomings" block: re-implementing Excel's
// 470+ functions in a renderer-side evaluator would mean two
// engines drifting apart and a multi-MB bundle on every Joplin
// note open.
//
// What this test pins:
//
//   1. Import preserves `f` AND `v` together. `xlsxBufferToSnapshot`
//      copies the source workbook's cached `result` value into
//      `cell.v` and the formula text into `cell.f`. Verified against
//      the project-owned fixtures.
//
//   2. Stale cached values flow through unchanged — exceljs does NOT
//      recalculate at parse time. If Excel saved with a wrong
//      cached `result`, `cell.v` holds that wrong value.
//
//   3. The export-side does NOT touch formula cells' values when
//      synthesizing table styles. (The synthesizer applies bg/cl/bd
//      overlays without modifying `f` or `v`, so a totals cell
//      keeps its formula even after we paint a totals-top border on
//      it.)
//
//   4. Manually-edited markdown is the only narrow case where the
//      renderer can show a stale value — and the fix is "open the
//      note in the Notesheet editor once," because Univer recalcs
//      on snapshot load.

import path from 'path';
import { readFileSync } from 'fs';
import ExcelJS from 'exceljs';
import { xlsxBufferToSnapshot } from '../src/xlsx';

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, { f?: string; v?: unknown; t?: number; s?: string }>>;
    }>;
    styles: Record<string, Record<string, unknown>>;
}

const APTOS = path.join(__dirname, 'fixtures', 'formatting-testdata', 'FormattingSmorgasboard.xlsx');

describe('M16 formula source-of-truth contract', () => {
    test('Aptos fixture: every formula cell has BOTH f and v populated at import', async () => {
        const buf = readFileSync(APTOS);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;

        let formulaCount = 0;
        let withV = 0;
        for (const sid of snap.sheetOrder) {
            const sheet = snap.sheets[sid];
            for (const rowKey of Object.keys(sheet.cellData)) {
                const row = sheet.cellData[Number(rowKey)];
                for (const colKey of Object.keys(row)) {
                    const cell = row[Number(colKey)];
                    if (typeof cell.f === 'string' && cell.f) {
                        formulaCount++;
                        if (cell.v !== undefined && cell.v !== null) withV++;
                    }
                }
            }
        }
        expect(formulaCount).toBeGreaterThan(0);
        // Every formula cell carries a cached v for the renderer.
        expect(withV).toBe(formulaCount);
    });

    test('xlsxBufferToSnapshot preserves stale cached results — does NOT recalculate at import', async () => {
        // Build a workbook with a deliberately STALE cached result for
        // the formula. If exceljs were recalculating at parse time,
        // we'd get v=60. Since it preserves whatever Excel saved, we
        // get v=999.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('S');
        ws.getCell('A1').value = 10;
        ws.getCell('A2').value = 20;
        ws.getCell('A3').value = 30;
        ws.getCell('B1').value = { formula: 'SUM(A1:A3)', result: 999 };

        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;

        const b1 = snap.sheets[snap.sheetOrder[0]].cellData[0]?.[1];
        expect(b1?.f).toBe('=SUM(A1:A3)');
        // The stale cached result flows through unchanged. The
        // renderer will display 999 until Univer is opened on the
        // note (Univer recalcs on snapshot load and writes the
        // correct 60 back into cell.v on save).
        expect(b1?.v).toBe(999);
    });

    test('Aptos fixture: synthesizeTableStyleAssignments does not touch formula cells\' v or f', async () => {
        // The totals row has formulas (=SUBTOTAL(...)). The synth
        // applies bd.t (DOUBLE, header colour) + bd.b (MEDIUM, lighter
        // accent) overlays. It must NOT modify v or f on those cells.
        const buf = readFileSync(APTOS);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;

        // Totals row is row 9 (0-indexed) in ProjectTracker A1:G10.
        // C9 (col 2, 0-indexed) has =SUBTOTAL(109, ProjectTracker[Budget]).
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const totalsBudget = sheet.cellData[9]?.[2];
        expect(totalsBudget?.f).toBe('=SUBTOTAL(109,ProjectTracker[Budget])');
        // Cached value from the source workbook (Excel evaluated at
        // its last save). 50000+75000+120000+30000+95000+60000+45000+85000 = 560000.
        expect(totalsBudget?.v).toBe(560000);
        // The synth overlays applied a style with bd.t / bd.b on this
        // cell — verify the style id is set AND the bd structure is
        // present, but v / f stay intact.
        expect(totalsBudget?.s).toBeDefined();
        const totalsStyle = snap.styles[totalsBudget!.s!];
        expect(totalsStyle.bd).toBeDefined();
        const bd = totalsStyle.bd as { t?: { s: number; cl: { rgb: string } }; b?: { s: number; cl: { rgb: string } } };
        expect(bd.t?.s).toBe(7);  // DOUBLE
        expect(bd.b?.s).toBe(8);  // MEDIUM
    });

    test('renderer contract: HTML output of a formula cell shows cell.v, not cell.f', () => {
        // The renderer's `renderCellValue` (in
        // src/contentScripts/notesheetRenderer.ts) returns escapeHtml
        // of String(cell.v) for non-rich-text cells. It never reads
        // cell.f. This test is a code-shape sentinel: if a future
        // change to renderCellValue starts evaluating cell.f, the
        // README "Known shortcomings" entry needs updating AND the
        // bundle-cost / two-engine-drift discussion must reopen.
        //
        // Read the renderer source and assert the contract.
        const rendererSrc = readFileSync(
            path.join(__dirname, '..', 'src', 'contentScripts', 'notesheetRenderer.ts'),
            'utf8',
        );
        // The function exists.
        expect(rendererSrc).toMatch(/function renderCellValue\b/);
        // Within renderCellValue's body, it must reference cell.v but
        // NOT cell.f. (We assert by carving out the function body and
        // checking. A simple grep — the function body up to the next
        // function declaration — is sufficient because the file is
        // structured.)
        const fnStart = rendererSrc.indexOf('function renderCellValue');
        expect(fnStart).toBeGreaterThan(-1);
        // Find the closing `}` of the function — track brace depth
        // starting from the first `{` after the signature.
        let i = rendererSrc.indexOf('{', fnStart);
        let depth = 1;
        i++;
        while (depth > 0 && i < rendererSrc.length) {
            const ch = rendererSrc[i];
            if (ch === '{') depth++;
            else if (ch === '}') depth--;
            i++;
        }
        const body = rendererSrc.slice(fnStart, i);
        // Renderer reads cell.v / cell.p but must NOT read cell.f.
        expect(body).toMatch(/cell\.v\b/);
        expect(body).not.toMatch(/cell\.f\b/);
    });
});
