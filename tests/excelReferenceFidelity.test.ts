// Reference-anchored fidelity gate for `synthesizeTableStyleAssignments`.
//
// PURPOSE
//
// Pin Notesheet's table-style synthesis against ground-truth Excel
// renders, NOT against what our own code emits. The pin-downs in
// `m12FixturePinDowns.test.ts` previously asserted things like
// `bg.rgb === '#196B24'` — but #196B24 is the raw accent3, NOT what
// Excel actually paints for `TableStyleMedium4`'s header. Excel
// renders the header at #34692E (a darker, less-saturated derivative
// of accent3) and the banded row at #CAEFCB. Asserting our own emit
// against a hand-derived recipe that nobody compared to a real Excel
// render is what allowed PR #22 to ship visibly wrong colours.
//
// This test instead reads the operator-captured Excel screenshots
// under `screenshots/excel-reference/` and samples the dominant fill
// of three regions per fixture: header band, banded data row, and
// (for Aptos) totals-row top double-line border. It then asserts
// our `xlsxBufferToSnapshot` output for the same fixture matches
// each of those reference RGBs within Δ ≤ 8 per channel.
//
// The fixtures are the two project-owned smorgasboard files. Both use
// `TableStyleMedium4`; they differ only in their `<a:clrScheme>`
// (Aptos accent3 #196B24 vs Classic accent3 #A5A5A5). The same
// catalog entry must produce two different rendered outputs.
//
// REFERENCE SCREENSHOTS
//
// `screenshots/excel-reference/FormattingSmorgasboard-Aptos.png` and
// `…-Classic.png` were captured by the operator from real Excel
// renders. They are the source of truth for "what Excel paints" in
// the regression-vs-fix conversation. Re-capturing them requires
// opening the corresponding `.xlsx` in Excel on a system with the
// matching theme installed and screenshotting the visible table.
//
// SAMPLING REGIONS
//
// Sampled empirically with `tools/util/pngSampler.ts:dominantColor`.
// The y-bands were eyeballed from a column probe printed during
// Phase 2 of M13/E. The reference PNGs are stable; the y-bands won't
// drift unless someone re-captures the screenshots, in which case
// they should re-derive these constants and update the test.
//
//   Aptos (1780 × 658):
//     header band  : x=200..1500, y=132..148  → dominant #34692E
//     banded data 1: x=200..1500, y=188..228  → dominant #CAEFCB
//     totals top border (double line): two strips at y=616 and y=520
//        actually the double-line border appears between data row
//        and totals row at y=472 and y=520 (separating banded → white).
//        Easier sentinel: y=472..472 (single horizontal pixel row of
//        the border) → dominant #72D068.
//
//   Classic (1378 × 618):
//     header band  : x=200..1200, y=130..162  → dominant #A5A5A5
//     banded data 1: x=200..1200, y=172..208  → dominant #EDEDED
//     totals top border: y=168 → dominant #C9C9C9
//
// TOLERANCE
//
// Δ ≤ 8 per channel. Excel's render anti-aliases at cell edges, so a
// few pixels along glyph rasters drift; the dominant fill is rock-
// stable. With `dominantColor` returning the highest-frequency RGB,
// the tolerance is mostly slack — the typical match is Δ = 0.

import path from 'path';
import { readFileSync } from 'fs';

import { xlsxBufferToSnapshot } from '../src/xlsx';
import { decodePng, dominantColor, hexToRgb, rgbDelta } from './util/pngSampler';

const REFERENCES_DIR = path.join(__dirname, '..', 'screenshots', 'excel-reference');
const FIXTURES_DIR = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata');

const APTOS_PNG = path.join(REFERENCES_DIR, 'FormattingSmorgasboard-Aptos.png');
const CLASSIC_PNG = path.join(REFERENCES_DIR, 'FormattingSmorgasboard-Classic.png');
const APTOS_XLSX = path.join(FIXTURES_DIR, 'FormattingSmorgasboard.xlsx');
const CLASSIC_XLSX = path.join(FIXTURES_DIR, 'FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx');

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, { v?: unknown; s?: string }>>;
    }>;
    styles: Record<string, {
        bg?: { rgb: string };
        bd?: Partial<Record<'t' | 'r' | 'b' | 'l', { s: number; cl: { rgb: string } }>>;
    }>;
}

const TOLERANCE = 8;

async function loadSnapshot(file: string): Promise<SnapshotShape> {
    const buf = readFileSync(file);
    return await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;
}

function snapBg(snap: SnapshotShape, row: number, col: number): string | undefined {
    const sheet = snap.sheets[snap.sheetOrder[0]];
    const cell = sheet.cellData[row]?.[col];
    if (!cell?.s) return undefined;
    const style = snap.styles[cell.s];
    return style?.bg?.rgb;
}

function snapBorderTop(snap: SnapshotShape, row: number, col: number): { rgb: string; style: number } | undefined {
    const sheet = snap.sheets[snap.sheetOrder[0]];
    const cell = sheet.cellData[row]?.[col];
    if (!cell?.s) return undefined;
    const style = snap.styles[cell.s];
    const t = style?.bd?.t;
    if (!t) return undefined;
    return { rgb: t.cl.rgb, style: t.s };
}

function expectRgbWithin(actualHex: string | undefined, expectedHex: string, tol: number, label: string): void {
    expect(actualHex).toBeDefined();
    const a = hexToRgb(actualHex!);
    const e = hexToRgb(expectedHex);
    const d = rgbDelta(a, e);
    if (d.max > tol) {
        throw new Error(
            `${label}: expected ${expectedHex} ±${tol}, got ${actualHex.toUpperCase()} ` +
            `(ΔR=${d.dR}, ΔG=${d.dG}, ΔB=${d.dB})`,
        );
    }
}

describe('Excel reference fidelity — TableStyleMedium4 (M13/E)', () => {
    describe('Aptos fixture (accent3 #196B24)', () => {
        let snap: SnapshotShape;
        beforeAll(async () => { snap = await loadSnapshot(APTOS_XLSX); });

        test('header bg matches Excel render at #34692E (Δ ≤ 8)', () => {
            const ref = decodePng(APTOS_PNG);
            const headerSample = dominantColor(ref, 200, 1500, 132, 148);
            // Sanity: anchor the reference to the operator-known #34692E
            // so an updated reference PNG fails loudly.
            expect(headerSample.hex).toBe('#34692E');

            const bg = snapBg(snap, 0, 0);
            expectRgbWithin(bg, headerSample.hex, TOLERANCE, 'Aptos header bg');
        });

        test('banded row even bg matches Excel render at #CAEFCB (Δ ≤ 8)', () => {
            const ref = decodePng(APTOS_PNG);
            const bandSample = dominantColor(ref, 200, 1500, 188, 228);
            expect(bandSample.hex).toBe('#CAEFCB');

            // First banded data row is row 1 (header at row 0). Even-row
            // band fills the full row; sample mid-table col 1.
            const bg = snapBg(snap, 1, 1);
            expectRgbWithin(bg, bandSample.hex, TOLERANCE, 'Aptos banded row even bg');
        });

        test('totals row top border matches Excel double-line at #72D068 (Δ ≤ 8)', () => {
            const ref = decodePng(APTOS_PNG);
            // Pick the y of the data-to-totals separator. The border lies
            // at y=472 (above the "Totals" row when shown as double line);
            // dominant is #72D068.
            const borderSample = dominantColor(ref, 200, 1500, 472, 472);
            expect(borderSample.hex).toBe('#72D068');

            // ProjectTracker spans A1:G10 with totalsRowCount=1; the
            // totals row is row 9 in the snapshot (0-based). The top
            // border slot must be present and within tolerance.
            const top = snapBorderTop(snap, 9, 0);
            expect(top).toBeDefined();
            expectRgbWithin(top!.rgb, borderSample.hex, TOLERANCE, 'Aptos totals top border colour');
        });
    });

    describe('Classic fixture (accent3 #A5A5A5)', () => {
        let snap: SnapshotShape;
        beforeAll(async () => { snap = await loadSnapshot(CLASSIC_XLSX); });

        test('header bg matches Excel render at #A5A5A5 (Δ ≤ 8)', () => {
            const ref = decodePng(CLASSIC_PNG);
            const headerSample = dominantColor(ref, 200, 1200, 130, 162);
            expect(headerSample.hex).toBe('#A5A5A5');

            const bg = snapBg(snap, 0, 0);
            expectRgbWithin(bg, headerSample.hex, TOLERANCE, 'Classic header bg');
        });

        test('banded row even bg matches Excel render at #EDEDED (Δ ≤ 8)', () => {
            const ref = decodePng(CLASSIC_PNG);
            const bandSample = dominantColor(ref, 200, 1200, 172, 208);
            expect(bandSample.hex).toBe('#EDEDED');

            const bg = snapBg(snap, 1, 1);
            expectRgbWithin(bg, bandSample.hex, TOLERANCE, 'Classic banded row even bg');
        });

        test('totals row top border matches Excel double-line at #C9C9C9 (Δ ≤ 8)', () => {
            const ref = decodePng(CLASSIC_PNG);
            const borderSample = dominantColor(ref, 200, 1200, 168, 168);
            expect(borderSample.hex).toBe('#C9C9C9');

            // ProductCatalog spans A1:F10; totals at row 9.
            const top = snapBorderTop(snap, 9, 0);
            expect(top).toBeDefined();
            expectRgbWithin(top!.rgb, borderSample.hex, TOLERANCE, 'Classic totals top border colour');
        });
    });
});
