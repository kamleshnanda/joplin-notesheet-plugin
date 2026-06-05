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

import JSZip from 'jszip';

// Univer's CF plugin identifies its snapshot resource via this constant.
// The const is exported as `SHEET_CONDITIONAL_FORMATTING_PLUGIN` from
// `@univerjs/sheets-conditional-formatting` (verified in
// `lib/types/base/const.d.ts`), but importing that package from Jest
// pulls in `lodash-es` (ESM-only) which jest-runtime cannot parse
// without a transformer override. Hard-coding the string here is a
// deliberate trade-off: the test must fail loud if Univer renames the
// constant. We re-export the same string from `src/xlsx.ts` (also as a
// hard-coded const) — the Jest snapshot fidelity tests below assert the
// snapshot's resource entry matches THIS literal, the runtime preset
// reads from the const exported from `@univerjs/sheets-conditional-formatting`,
// and a single-test sanity gate compares them. If Univer ever renames
// the constant, the runtime side gets the new value but our string
// stays the old; the sanity gate's diff fires loudly in CI.
const SHEET_CONDITIONAL_FORMATTING_PLUGIN = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN';

import { xlsxBufferToSnapshot } from '../src/xlsx';
import { decodePng, dominantColor, hexToRgb, rgbDelta } from './util/pngSampler';

const REFERENCES_DIR = path.join(__dirname, '..', 'screenshots', 'excel-reference');
const FIXTURES_DIR = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata');

const APTOS_PNG = path.join(REFERENCES_DIR, 'FormattingSmorgasboard-Aptos.png');
const CLASSIC_PNG = path.join(REFERENCES_DIR, 'FormattingSmorgasboard-Classic.png');
const APTOS_XLSX = path.join(FIXTURES_DIR, 'FormattingSmorgasboard.xlsx');
const CLASSIC_XLSX = path.join(FIXTURES_DIR, 'FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx');
const CF_VARIANTS_XLSX = path.join(FIXTURES_DIR, 'ConditionalFormatting-Variants.xlsx');

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

function snapBorderBottom(snap: SnapshotShape, row: number, col: number): { rgb: string; style: number } | undefined {
    const sheet = snap.sheets[snap.sheetOrder[0]];
    const cell = sheet.cellData[row]?.[col];
    if (!cell?.s) return undefined;
    const style = snap.styles[cell.s];
    const b = style?.bd?.b;
    if (!b) return undefined;
    return { rgb: b.cl.rgb, style: b.s };
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

        test('totals row TOP border matches Excel render at #72D068 (Δ ≤ 8)', () => {
            const ref = decodePng(APTOS_PNG);
            // Excel paints two distinct strips on the totals row in
            // FormattingSmorgasboard-Aptos.png:
            //   - top strip at y=424-425 (just below row 9's banded fill)
            //   - bottom strip at y=472-473 (just above the white area
            //     below the table)
            // Both at #72D068. Earlier this test sampled y=472 and
            // labelled it "top" — it's the BOTTOM strip; the assertion
            // happened to pass because both sides share a colour.
            const borderSample = dominantColor(ref, 200, 1500, 424, 425);
            expect(borderSample.hex).toBe('#72D068');

            // ProjectTracker spans A1:G10 with totalsRowCount=1; the
            // totals row is row 9 in the snapshot (0-based).
            const top = snapBorderTop(snap, 9, 0);
            expect(top).toBeDefined();
            expectRgbWithin(top!.rgb, borderSample.hex, TOLERANCE, 'Aptos totals top border colour');
        });

        test('totals row BOTTOM border matches Excel render at #72D068 (Δ ≤ 8)', () => {
            const ref = decodePng(APTOS_PNG);
            // Bottom strip at y=472-473 — the second of the two #72D068
            // strips that frame the totals-row body (~46 px apart).
            const borderSample = dominantColor(ref, 200, 1500, 472, 473);
            expect(borderSample.hex).toBe('#72D068');

            const bottom = snapBorderBottom(snap, 9, 0);
            expect(bottom).toBeDefined();
            expectRgbWithin(bottom!.rgb, borderSample.hex, TOLERANCE, 'Aptos totals bottom border colour');
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

        test('totals row TOP border matches Excel render at #C9C9C9 (Δ ≤ 8)', () => {
            const ref = decodePng(CLASSIC_PNG);
            // Top strip at y=422-423 (just below row 9's banded fill).
            const borderSample = dominantColor(ref, 200, 1200, 422, 423);
            expect(borderSample.hex).toBe('#C9C9C9');

            // ProductCatalog spans A1:F10; totals at row 9.
            const top = snapBorderTop(snap, 9, 0);
            expect(top).toBeDefined();
            expectRgbWithin(top!.rgb, borderSample.hex, TOLERANCE, 'Classic totals top border colour');
        });

        test('totals row BOTTOM border matches Excel render at #C9C9C9 (Δ ≤ 8)', () => {
            const ref = decodePng(CLASSIC_PNG);
            // Bottom strip at y=464-465 — the second of the two #C9C9C9
            // strips that frame the totals-row body.
            const borderSample = dominantColor(ref, 200, 1200, 464, 465);
            expect(borderSample.hex).toBe('#C9C9C9');

            const bottom = snapBorderBottom(snap, 9, 0);
            expect(bottom).toBeDefined();
            expectRgbWithin(bottom!.rgb, borderSample.hex, TOLERANCE, 'Classic totals bottom border colour');
        });
    });
});

// ─── M15: Conditional Formatting fidelity ────────────────────────────────
//
// Pin the imported snapshot's CF resource (keyed by Univer's
// `SHEET_CONDITIONAL_FORMATTING_PLUGIN`) against the source workbook's
// `xl/worksheets/sheet1.xml`. Anchored UPSTREAM of our parser per
// `feedback_pge_fidelity_test_gap.md` — we read the source XML's cfRule
// elements directly (via JSZip + a minimal regex-based parser; no
// reliance on exceljs's CF surface). Each test asserts the rule's
// translated Univer shape carries the same ref / cfvo / colour /
// operator / rank as the source.

interface ParsedCfRule {
    type: string;
    sqref: string;
    priority: number;
    operator?: string;
    rank?: number;
    bottom?: boolean;
    formulae: string[];
    cfvo: Array<{ type: string; val?: string }>;
    colors: string[];           // each entry like "#RRGGBB"
    iconSet?: string;
    dxfBgArgb?: string;         // raw "FFRRGGBB" if dxfId resolves to a dxf with fill bgColor
    dxfId?: number;
}

function parseSourceCfRules(zip: JSZip, sheetXmlPath: string, stylesXmlPath: string): ParsedCfRule[] {
    const sheetEntry = zip.file(sheetXmlPath);
    const stylesEntry = zip.file(stylesXmlPath);
    if (!sheetEntry || !stylesEntry) throw new Error(`missing ${sheetXmlPath} or ${stylesXmlPath} in fixture`);
    // Synchronous read via Node's zip helper — JSZip is async, but
    // for tests we wrap the async call in beforeAll. We surface a
    // sync-shape parsing function and the caller awaits separately.
    throw new Error('parseSourceCfRules is the sync core; use parseSourceCfRulesAsync');
}

async function loadCfFixtureXml(filePath: string): Promise<{
    sheetXml: string;
    stylesXml: string;
}> {
    const buf = readFileSync(filePath);
    const zip = await JSZip.loadAsync(buf);
    const sheetXmlPath = Object.keys(zip.files).find((p) => /^xl\/worksheets\/sheet1\.xml$/.test(p));
    const stylesXmlPath = Object.keys(zip.files).find((p) => /^xl\/styles\.xml$/.test(p));
    if (!sheetXmlPath) throw new Error('missing xl/worksheets/sheet1.xml in fixture');
    if (!stylesXmlPath) throw new Error('missing xl/styles.xml in fixture');
    const sheetXml = await zip.files[sheetXmlPath].async('string');
    const stylesXml = await zip.files[stylesXmlPath].async('string');
    return { sheetXml, stylesXml };
}

// Pull <dxfs> entries' fill bgColor argb in dxfId order. For
// ConditionalFormatting-Variants.xlsx: dxf 0 has bgColor #C6EFCE
// (top-3 highlight green), dxf 1 has bgColor #FFC7CE (cellIs > 50
// pink). Order is asserted in the test, not assumed.
function parseDxfFillBgArgbs(stylesXml: string): string[] {
    // Match the <dxfs> block.
    const dxfsMatch = stylesXml.match(/<dxfs[^>]*>([\s\S]*?)<\/dxfs>/);
    if (!dxfsMatch) return [];
    const inner = dxfsMatch[1];
    // Each <dxf>...</dxf> block. Inside, find bgColor rgb (or
    // theme-indexed; we only handle rgb for the M15 fixture).
    const out: string[] = [];
    const dxfRe = /<dxf>([\s\S]*?)<\/dxf>/g;
    let m: RegExpExecArray | null;
    while ((m = dxfRe.exec(inner))) {
        const body = m[1];
        const bgMatch = body.match(/<bgColor\s+rgb="([0-9A-Fa-f]{8})"/);
        out.push(bgMatch ? bgMatch[1].toUpperCase() : '');
    }
    return out;
}

function parseCfRulesFromSheet(sheetXml: string, dxfBgArgbs: string[]): ParsedCfRule[] {
    const out: ParsedCfRule[] = [];
    // Each <conditionalFormatting sqref="...">...</conditionalFormatting> block.
    const blockRe = /<conditionalFormatting\s+sqref="([^"]+)">([\s\S]*?)<\/conditionalFormatting>/g;
    let bm: RegExpExecArray | null;
    while ((bm = blockRe.exec(sheetXml))) {
        const sqref = bm[1];
        const inner = bm[2];
        // Each <cfRule .../> or <cfRule>...</cfRule>.
        const ruleRe = /<cfRule\b([^>]*)(?:\/>|>([\s\S]*?)<\/cfRule>)/g;
        let rm: RegExpExecArray | null;
        while ((rm = ruleRe.exec(inner))) {
            const attrs = rm[1];
            const body = rm[2] ?? '';
            const typeAttr = /type="([^"]+)"/.exec(attrs);
            const priorityAttr = /priority="(\d+)"/.exec(attrs);
            const opAttr = /operator="([^"]+)"/.exec(attrs);
            const rankAttr = /rank="(\d+)"/.exec(attrs);
            const bottomAttr = /bottom="(1|true)"/.exec(attrs);
            const dxfIdAttr = /dxfId="(\d+)"/.exec(attrs);
            const formulae: string[] = [];
            const formulaeRe = /<formula>([\s\S]*?)<\/formula>/g;
            let fm: RegExpExecArray | null;
            while ((fm = formulaeRe.exec(body))) formulae.push(fm[1].trim());
            const cfvo: Array<{ type: string; val?: string }> = [];
            const cfvoRe = /<cfvo\s+([^/]*?)\/>/g;
            let cm: RegExpExecArray | null;
            while ((cm = cfvoRe.exec(body))) {
                const ca = cm[1];
                const t = /type="([^"]+)"/.exec(ca);
                const v = /val="([^"]+)"/.exec(ca);
                cfvo.push({ type: t ? t[1] : '', val: v ? v[1] : undefined });
            }
            const colors: string[] = [];
            const colorRe = /<color\s+rgb="([0-9A-Fa-f]{8})"/g;
            let xm: RegExpExecArray | null;
            while ((xm = colorRe.exec(body))) {
                colors.push('#' + xm[1].slice(2).toUpperCase());
            }
            const iconSetAttrMatch = /<iconSet\s+iconSet="([^"]+)"/.exec(body);
            const dxfId = dxfIdAttr ? Number(dxfIdAttr[1]) : undefined;
            const dxfBgArgb = dxfId !== undefined ? dxfBgArgbs[dxfId] : undefined;
            out.push({
                type: typeAttr ? typeAttr[1] : '',
                sqref,
                priority: priorityAttr ? Number(priorityAttr[1]) : 0,
                operator: opAttr ? opAttr[1] : undefined,
                rank: rankAttr ? Number(rankAttr[1]) : undefined,
                bottom: !!bottomAttr,
                formulae,
                cfvo,
                colors,
                iconSet: iconSetAttrMatch ? iconSetAttrMatch[1] : undefined,
                dxfBgArgb: dxfBgArgb || undefined,
                dxfId,
            });
        }
    }
    // Sort by priority ascending — Excel applies higher-priority rules
    // first, but our snapshot maps them to an array in priority order.
    out.sort((a, b) => a.priority - b.priority);
    return out;
}

interface CfSnapshotShape {
    resources?: Array<{ name?: string; data?: string }>;
    sheetOrder: string[];
}

async function loadCfSnapshot(file: string): Promise<{
    snapshot: CfSnapshotShape;
    rulesBySubUnit: Record<string, Array<{
        cfId?: string;
        ranges: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
        rule: Record<string, unknown>;
    }>>;
}> {
    const buf = readFileSync(file);
    const snapshot = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as CfSnapshotShape;
    const entry = (snapshot.resources ?? []).find((r) => r?.name === SHEET_CONDITIONAL_FORMATTING_PLUGIN);
    if (!entry || typeof entry.data !== 'string') {
        return { snapshot, rulesBySubUnit: {} };
    }
    const parsed = JSON.parse(entry.data) as Record<string, Array<{
        cfId?: string;
        ranges: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
        rule: Record<string, unknown>;
    }>>;
    return { snapshot, rulesBySubUnit: parsed };
}

function rangesToSqref(ranges: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>): string {
    return ranges.map((r) => {
        const cl = (idx: number): string => {
            let n = idx;
            let s = '';
            while (n >= 0) {
                s = String.fromCharCode(65 + (n % 26)) + s;
                n = Math.floor(n / 26) - 1;
            }
            return s;
        };
        const tl = cl(r.startColumn) + (r.startRow + 1);
        const br = cl(r.endColumn) + (r.endRow + 1);
        return tl === br ? tl : `${tl}:${br}`;
    }).join(' ');
}

describe('M15 CF reference fidelity — ConditionalFormatting-Variants.xlsx', () => {
    let sourceRules: ParsedCfRule[];
    let snapRulesByPriority: Array<{
        ranges: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
        rule: Record<string, unknown>;
    }>;

    beforeAll(async () => {
        const { sheetXml, stylesXml } = await loadCfFixtureXml(CF_VARIANTS_XLSX);
        const dxfBgArgbs = parseDxfFillBgArgbs(stylesXml);
        sourceRules = parseCfRulesFromSheet(sheetXml, dxfBgArgbs);
        const { rulesBySubUnit } = await loadCfSnapshot(CF_VARIANTS_XLSX);
        // The snapshot keys by subUnitId (sheet id). The fixture has a
        // single sheet; pick the only entry. Snapshot rules come back
        // in the on-disk order — Univer's `addRule` reverses on load,
        // so we DON'T re-sort here; we just match by sqref.
        const subUnitIds = Object.keys(rulesBySubUnit);
        snapRulesByPriority = subUnitIds.length > 0 ? rulesBySubUnit[subUnitIds[0]] : [];
    });

    test('emits a SHEET_CONDITIONAL_FORMATTING_PLUGIN resource with 5 rules', () => {
        expect(sourceRules).toHaveLength(5);
        expect(snapRulesByPriority).toHaveLength(5);
    });

    test('Rule 0 (priority 1, sqref A2:A11): colorScale with 3 cfvo + 3 colours', () => {
        const src = sourceRules.find((r) => r.priority === 1);
        expect(src).toBeDefined();
        expect(src!.type).toBe('colorScale');
        expect(src!.sqref).toBe('A2:A11');
        expect(src!.cfvo).toHaveLength(3);
        expect(src!.cfvo[0].type).toBe('min');
        expect(src!.cfvo[1].type).toBe('percentile');
        expect(src!.cfvo[1].val).toBe('50');
        expect(src!.cfvo[2].type).toBe('max');
        expect(src!.colors).toEqual(['#F8696B', '#FFEB84', '#63BE7B']);

        // Find the matching snapshot rule by ranges → sqref.
        const snap = snapRulesByPriority.find((r) => rangesToSqref(r.ranges) === 'A2:A11');
        expect(snap).toBeDefined();
        const sr = snap!.rule;
        expect(sr.type).toBe('colorScale');
        const config = sr.config as Array<{ index?: number; color: string; value: { type: string; value?: number } }>;
        expect(config).toHaveLength(3);
        expect(config[0].value.type).toBe('min');
        expect(config[0].color.toUpperCase()).toBe('#F8696B');
        expect(config[1].value.type).toBe('percentile');
        expect(config[1].value.value).toBe(50);
        expect(config[1].color.toUpperCase()).toBe('#FFEB84');
        expect(config[2].value.type).toBe('max');
        expect(config[2].color.toUpperCase()).toBe('#63BE7B');
    });

    test('Rule 1 (priority 2, sqref C2:C11): dataBar with 2 cfvo + 1 colour', () => {
        const src = sourceRules.find((r) => r.priority === 2);
        expect(src).toBeDefined();
        expect(src!.type).toBe('dataBar');
        expect(src!.sqref).toBe('C2:C11');
        expect(src!.cfvo[0].type).toBe('min');
        expect(src!.cfvo[1].type).toBe('max');
        expect(src!.colors).toEqual(['#638EC6']);

        const snap = snapRulesByPriority.find((r) => rangesToSqref(r.ranges) === 'C2:C11');
        expect(snap).toBeDefined();
        const sr = snap!.rule as { type: string; config: { min: { type: string }; max: { type: string }; positiveColor: string; isGradient?: boolean } };
        expect(sr.type).toBe('dataBar');
        expect(sr.config.min.type).toBe('min');
        expect(sr.config.max.type).toBe('max');
        expect(sr.config.positiveColor.toUpperCase()).toBe('#638EC6');
    });

    test('Rule 2 (priority 3, sqref E2:E11): cellIs > 50, fill #FFC7CE → highlightCell number', () => {
        const src = sourceRules.find((r) => r.priority === 3);
        expect(src).toBeDefined();
        expect(src!.type).toBe('cellIs');
        expect(src!.sqref).toBe('E2:E11');
        expect(src!.operator).toBe('greaterThan');
        expect(src!.formulae).toEqual(['50']);
        expect(src!.dxfBgArgb).toBe('FFFFC7CE');

        const snap = snapRulesByPriority.find((r) => rangesToSqref(r.ranges) === 'E2:E11');
        expect(snap).toBeDefined();
        const sr = snap!.rule as { type: string; subType: string; operator: string; value?: number; style?: { bg?: { rgb: string } } };
        expect(sr.type).toBe('highlightCell');
        expect(sr.subType).toBe('number');
        expect(sr.operator).toBe('greaterThan');
        expect(sr.value).toBe(50);
        expect((sr.style?.bg?.rgb ?? '').toUpperCase()).toBe('#FFC7CE');
    });

    test('Rule 3 (priority 4, sqref G2:G11): top10 rank=3, fill #C6EFCE → highlightCell rank', () => {
        const src = sourceRules.find((r) => r.priority === 4);
        expect(src).toBeDefined();
        expect(src!.type).toBe('top10');
        expect(src!.sqref).toBe('G2:G11');
        expect(src!.rank).toBe(3);
        expect(src!.bottom).toBe(false);
        expect(src!.dxfBgArgb).toBe('FFC6EFCE');

        const snap = snapRulesByPriority.find((r) => rangesToSqref(r.ranges) === 'G2:G11');
        expect(snap).toBeDefined();
        const sr = snap!.rule as { type: string; subType: string; isBottom: boolean; isPercent?: boolean; value: number; style?: { bg?: { rgb: string } } };
        expect(sr.type).toBe('highlightCell');
        expect(sr.subType).toBe('rank');
        expect(sr.isBottom).toBe(false);
        expect(sr.value).toBe(3);
        expect((sr.style?.bg?.rgb ?? '').toUpperCase()).toBe('#C6EFCE');
    });

    test('Rule 4 (priority 5, sqref I2:I11): iconSet 3Arrows with 3 cfvo (percent 0/33/67)', () => {
        const src = sourceRules.find((r) => r.priority === 5);
        expect(src).toBeDefined();
        expect(src!.type).toBe('iconSet');
        expect(src!.sqref).toBe('I2:I11');
        expect(src!.iconSet).toBe('3Arrows');
        expect(src!.cfvo).toEqual([
            { type: 'percent', val: '0' },
            { type: 'percent', val: '33' },
            { type: 'percent', val: '67' },
        ]);

        const snap = snapRulesByPriority.find((r) => rangesToSqref(r.ranges) === 'I2:I11');
        expect(snap).toBeDefined();
        const sr = snap!.rule as { type: string; config: Array<{ iconType: string; iconId: string; value: { type: string; value: number } }> };
        expect(sr.type).toBe('iconSet');
        expect(sr.config).toHaveLength(3);
        // Every entry should reference the 3Arrows iconType.
        for (const item of sr.config) expect(item.iconType).toBe('3Arrows');
        // iconId values cover the 3-arrow set indices "0", "1", "2"
        // (one of each); we don't pin order here because the
        // translator may emit them as ascending-by-threshold or
        // descending depending on Univer's expected matching order.
        const ids = sr.config.map((c) => c.iconId).sort();
        expect(ids).toEqual(['0', '1', '2']);
        // Univer's iconSet uses N-1 thresholds for an N-icon set:
        // for 3Arrows the first two threshold cfvos (33, 67) carry
        // over from the source; the third entry is the catch-all
        // covering everything below the lowest threshold (Univer's
        // IconSetCalculateUnit returns the last entry unconditionally
        // when no higher-priority match fires). The catch-all uses
        // num/MAX_SAFE_INTEGER as a placeholder.
        const percentEntries = sr.config.filter((c) => c.value.type === 'percent');
        expect(percentEntries).toHaveLength(2);
        const percents = percentEntries.map((c) => c.value.value).sort((a, b) => a - b);
        expect(percents).toEqual([33, 67]);
        // Catch-all entry is the lowest-icon (red-down for 3Arrows,
        // iconId="2") with operator lessThanOrEqual.
        const catchAll = sr.config.find((c) => c.value.type === 'num');
        expect(catchAll).toBeDefined();
        expect(catchAll!.iconId).toBe('2');
    });
});
