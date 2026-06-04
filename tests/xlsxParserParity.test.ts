// M14 spike — parser parity matrix between exceljs (current) and
// xlsx-js-style (candidate). This test runs both
// `src/xlsx.ts:xlsxBufferToSnapshot` and `src/xlsxSheetJS.ts:xlsxBufferToSnapshot`
// against every fixture under `tests/ExcelBaseTestData/` and writes a
// per-fixture-per-dimension matrix to
// `tests/golden-snapshots/parity-matrix.json`.
//
// The test exits 0 even when divergences exist — divergences are the spike's
// signal. The matrix is consumed by `docs/m14-sheetjs-spike.md`'s capability
// matrix (rendered as a markdown table inline in the decision doc) and
// becomes input to the GO / NO-GO / CONDITIONAL recommendation.
//
// Status values per cell:
//   - 'match'           — both parsers produce structurally equivalent output
//   - 'divergence'      — values present on both sides, differ in shape/value
//   - 'sheetjs-blocked' — xlsx-js-style returns nothing where exceljs does
//   - 'exceljs-blocked' — exceljs throws / drops where xlsx-js-style does not
//   - 'not-applicable'  — the fixture doesn't exercise this dimension
//
// **No symptom-patching.** When the SheetJS side returns less than exceljs,
// the matrix records 'sheetjs-blocked' or 'divergence' — it does NOT alter
// the assertion to call SheetJS the new truth. The whole point of the spike
// is to surface these gaps for the operator to judge.

import * as fs from 'fs';
import * as path from 'path';

import * as exceljsParser from '../src/xlsx';
import * as sheetJSParser from '../src/xlsxSheetJS';

interface CellRecord {
    v?: string | number | boolean;
    f?: string;
    t?: number;
    s?: string;
    p?: Record<string, unknown>;
}

interface Sheet {
    id: string;
    name: string;
    cellData: Record<number, Record<number, CellRecord>>;
    rowCount: number;
    columnCount: number;
    mergeData?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
}

interface Snapshot {
    sheetOrder: string[];
    sheets: Record<string, Sheet>;
    styles: Record<string, Record<string, unknown>>;
    defaultStyle?: Record<string, unknown>;
    resources?: Array<{ name: string; data: string }>;
}

type Status = 'match' | 'divergence' | 'sheetjs-blocked' | 'exceljs-blocked' | 'not-applicable';

interface MatrixEntry {
    status: Status;
    note?: string;
}

const FORMATTING_DIR = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata');
const CHART_DIR = path.join(__dirname, 'ExcelBaseTestData', 'chart-testdata');
const MATRIX_OUT = path.join(__dirname, 'golden-snapshots', 'parity-matrix.json');

function listFixtures(dir: string): string[] {
    if (!fs.existsSync(dir)) return [];
    return fs
        .readdirSync(dir)
        .filter((f) => f.toLowerCase().endsWith('.xlsx'))
        .sort();
}

function countCells(sheet: Sheet | undefined): number {
    if (!sheet) return 0;
    let n = 0;
    for (const r of Object.keys(sheet.cellData ?? {})) {
        n += Object.keys(sheet.cellData[Number(r)] ?? {}).length;
    }
    return n;
}

function styleCount(snap: Snapshot | null): number {
    return snap ? Object.keys(snap.styles ?? {}).length : 0;
}

function hasField(snap: Snapshot | null, predicate: (style: Record<string, unknown>) => boolean): boolean {
    if (!snap) return false;
    for (const id of Object.keys(snap.styles ?? {})) {
        if (predicate(snap.styles[id])) return true;
    }
    return false;
}

function compareSheetSizes(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const excSheetCount = exc.sheetOrder?.length ?? 0;
    const sjsSheetCount = sjs.sheetOrder?.length ?? 0;
    if (excSheetCount !== sjsSheetCount) {
        return { status: 'divergence', note: `sheetCount exc=${excSheetCount} sjs=${sjsSheetCount}` };
    }
    const sheetIdsExc = exc.sheetOrder.map((id) => exc.sheets[id]?.name).filter(Boolean);
    const sheetIdsSjs = sjs.sheetOrder.map((id) => sjs.sheets[id]?.name).filter(Boolean);
    const namesMatch = sheetIdsExc.every((n, i) => n === sheetIdsSjs[i]);
    if (!namesMatch) {
        return { status: 'divergence', note: `sheetNames exc=[${sheetIdsExc.join(',')}] sjs=[${sheetIdsSjs.join(',')}]` };
    }
    // Compare per-sheet cell counts.
    let excTotal = 0, sjsTotal = 0;
    for (const id of exc.sheetOrder) {
        excTotal += countCells(exc.sheets[id]);
    }
    for (const id of sjs.sheetOrder) {
        sjsTotal += countCells(sjs.sheets[id]);
    }
    if (excTotal === sjsTotal) return { status: 'match', note: `${excTotal} cells across ${excSheetCount} sheet(s)` };
    return { status: 'divergence', note: `cellCount exc=${excTotal} sjs=${sjsTotal}` };
}

function compareCellValues(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    // Walk every cell present in exc; compare value to sjs's same address.
    let matches = 0, divergences = 0, missing = 0, total = 0;
    for (const sheetId of exc.sheetOrder ?? []) {
        const excSheet = exc.sheets[sheetId];
        const sjsSheet = sjs.sheets[sheetId];
        if (!excSheet) continue;
        for (const rk of Object.keys(excSheet.cellData ?? {})) {
            const r = Number(rk);
            const row = excSheet.cellData[r];
            for (const ck of Object.keys(row ?? {})) {
                const c = Number(ck);
                total++;
                const excCell = row[c];
                const sjsCell = sjsSheet?.cellData?.[r]?.[c];
                if (!sjsCell) {
                    if (excCell.v !== undefined) missing++;
                    continue;
                }
                if (sameValue(excCell.v, sjsCell.v)) matches++;
                else divergences++;
            }
        }
    }
    if (total === 0) return { status: 'not-applicable', note: 'no cells' };
    if (divergences === 0 && missing === 0) return { status: 'match', note: `${matches}/${total} values equal` };
    return {
        status: 'divergence',
        note: `${matches}/${total} match, ${divergences} differ, ${missing} missing on sjs`,
    };
}

function sameValue(a: unknown, b: unknown): boolean {
    if (a === b) return true;
    if (a === undefined || b === undefined) return false;
    if (typeof a === 'string' && typeof b === 'string') return a === b;
    // Number tolerance for date/serial divergence (exceljs sometimes
    // surfaces dates as Date objects, sjs as ISO strings).
    if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9;
    return false;
}

function compareFormulas(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    let excFormulas = 0, sjsFormulas = 0, matched = 0;
    for (const sheetId of exc.sheetOrder ?? []) {
        const excSheet = exc.sheets[sheetId];
        const sjsSheet = sjs.sheets[sheetId];
        if (!excSheet) continue;
        for (const rk of Object.keys(excSheet.cellData ?? {})) {
            const r = Number(rk);
            const row = excSheet.cellData[r];
            for (const ck of Object.keys(row ?? {})) {
                const c = Number(ck);
                const excCell = row[c];
                const sjsCell = sjsSheet?.cellData?.[r]?.[c];
                if (excCell.f) {
                    excFormulas++;
                    if (sjsCell?.f && normalizeFormula(excCell.f) === normalizeFormula(sjsCell.f)) matched++;
                }
                if (sjsCell?.f) sjsFormulas++;
            }
        }
    }
    if (excFormulas === 0 && sjsFormulas === 0) return { status: 'not-applicable', note: 'no formulas' };
    if (excFormulas === sjsFormulas && matched === excFormulas) {
        return { status: 'match', note: `${matched}/${excFormulas} formulas equal` };
    }
    return {
        status: 'divergence',
        note: `exc=${excFormulas} sjs=${sjsFormulas} matched=${matched}`,
    };
}

function normalizeFormula(f: string): string {
    // Strip leading '=' if present; both parsers emit it.
    return f.startsWith('=') ? f.slice(1).trim() : f.trim();
}

function compareMerges(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    let excMerges = 0, sjsMerges = 0;
    for (const id of exc.sheetOrder ?? []) {
        excMerges += exc.sheets[id]?.mergeData?.length ?? 0;
    }
    for (const id of sjs.sheetOrder ?? []) {
        sjsMerges += sjs.sheets[id]?.mergeData?.length ?? 0;
    }
    if (excMerges === 0 && sjsMerges === 0) return { status: 'not-applicable', note: 'no merges' };
    if (excMerges === sjsMerges) return { status: 'match', note: `${excMerges} merges` };
    return { status: 'divergence', note: `exc=${excMerges} sjs=${sjsMerges}` };
}

function compareStyleRecords(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const excCount = styleCount(exc);
    const sjsCount = styleCount(sjs);
    if (excCount === 0 && sjsCount === 0) return { status: 'not-applicable', note: 'no styles' };
    // sjs <= exc/2 indicates the major styling-loss issue surfaced by the
    // spike (xlsx-js-style returns empty {} for borders, drops alignment,
    // drops most font formatting from the styles.xml-driven path).
    if (sjsCount === 0 && excCount > 0) return { status: 'sheetjs-blocked', note: `exc=${excCount} sjs=0` };
    if (sjsCount < Math.ceil(excCount / 2)) {
        return { status: 'divergence', note: `style count exc=${excCount} sjs=${sjsCount} (sjs lost >50%)` };
    }
    if (sjsCount === excCount) return { status: 'match', note: `${excCount} styles` };
    return { status: 'divergence', note: `exc=${excCount} sjs=${sjsCount}` };
}

function compareThemeClrScheme(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const excHas = (exc.resources ?? []).some(
        (r) => r?.name === 'SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN',
    );
    const sjsHas = (sjs.resources ?? []).some(
        (r) => r?.name === 'SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN',
    );
    if (!excHas && !sjsHas) return { status: 'not-applicable', note: 'no theme clrScheme' };
    if (excHas && !sjsHas) return { status: 'sheetjs-blocked', note: 'exc captured, sjs missed' };
    if (!excHas && sjsHas) return { status: 'divergence', note: 'sjs captured, exc missed' };
    // Both present: compare the raw XML strings.
    const excXml = (exc.resources ?? []).find((r) => r?.name === 'SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN')?.data ?? '';
    const sjsXml = (sjs.resources ?? []).find((r) => r?.name === 'SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN')?.data ?? '';
    if (excXml === sjsXml) return { status: 'match', note: 'identical clrScheme XML' };
    return { status: 'divergence', note: 'clrScheme present on both, XML differs' };
}

function compareDefaultStyle(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const excDef = exc.defaultStyle;
    const sjsDef = sjs.defaultStyle;
    if (!excDef && !sjsDef) return { status: 'not-applicable', note: 'no defaultStyle on either' };
    if (excDef && !sjsDef) {
        return { status: 'sheetjs-blocked', note: `exc has ff="${excDef.ff}", sjs has none` };
    }
    if (!excDef && sjsDef) return { status: 'divergence', note: 'sjs added defaultStyle, exc did not' };
    if (JSON.stringify(excDef) === JSON.stringify(sjsDef)) return { status: 'match', note: JSON.stringify(excDef) };
    return { status: 'divergence', note: `exc=${JSON.stringify(excDef)} sjs=${JSON.stringify(sjsDef)}` };
}

function compareTables(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const excTables = (exc.resources ?? []).find((r) => r?.name === 'SHEET_TABLE_PLUGIN');
    const sjsTables = (sjs.resources ?? []).find((r) => r?.name === 'SHEET_TABLE_PLUGIN');
    if (!excTables && !sjsTables) return { status: 'not-applicable', note: 'no tables' };
    if (excTables && !sjsTables) return { status: 'sheetjs-blocked', note: 'exc captured tables, sjs missing' };
    if (!excTables && sjsTables) return { status: 'divergence', note: 'sjs added tables' };
    return { status: 'match', note: 'both captured' };
}

function compareRichText(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    // A rich-text cell has cell.p.body.textRuns with length >= 2.
    const hasRichText = (snap: Snapshot): number => {
        let n = 0;
        for (const sid of snap.sheetOrder ?? []) {
            const sheet = snap.sheets[sid];
            if (!sheet) continue;
            for (const rk of Object.keys(sheet.cellData ?? {})) {
                for (const ck of Object.keys(sheet.cellData[Number(rk)] ?? {})) {
                    const cell = sheet.cellData[Number(rk)][Number(ck)];
                    const body = (cell.p as { body?: { textRuns?: unknown[] } } | undefined)?.body;
                    if (body && Array.isArray(body.textRuns) && body.textRuns.length >= 2) n++;
                }
            }
        }
        return n;
    };
    const excN = hasRichText(exc);
    const sjsN = hasRichText(sjs);
    if (excN === 0 && sjsN === 0) return { status: 'not-applicable', note: 'no multi-run cells' };
    if (excN === sjsN) return { status: 'match', note: `${excN} multi-run cells on both` };
    if (sjsN < excN) return { status: 'divergence', note: `exc=${excN} sjs=${sjsN}` };
    return { status: 'divergence', note: `exc=${excN} sjs=${sjsN}` };
}

function compareHyperlinks(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const countHyper = (snap: Snapshot): number => {
        let n = 0;
        for (const sid of snap.sheetOrder ?? []) {
            const sheet = snap.sheets[sid];
            if (!sheet) continue;
            for (const rk of Object.keys(sheet.cellData ?? {})) {
                for (const ck of Object.keys(sheet.cellData[Number(rk)] ?? {})) {
                    const cell = sheet.cellData[Number(rk)][Number(ck)];
                    const ranges = (cell.p as { body?: { customRanges?: Array<{ rangeType?: number }> } } | undefined)
                        ?.body?.customRanges;
                    if (Array.isArray(ranges) && ranges.some((r) => r?.rangeType === 0)) n++;
                }
            }
        }
        return n;
    };
    const excN = countHyper(exc);
    const sjsN = countHyper(sjs);
    if (excN === 0 && sjsN === 0) return { status: 'not-applicable', note: 'no hyperlinks' };
    if (excN === sjsN) return { status: 'match', note: `${excN} hyperlinked cells` };
    return { status: 'divergence', note: `exc=${excN} sjs=${sjsN}` };
}

function compareRotatedText(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const excHas = hasField(exc, (s) => !!s.tr);
    const sjsHas = hasField(sjs, (s) => !!s.tr);
    if (!excHas && !sjsHas) return { status: 'not-applicable', note: 'no rotation' };
    if (excHas && !sjsHas) return { status: 'sheetjs-blocked', note: 'exc captured tr, sjs missed' };
    if (!excHas && sjsHas) return { status: 'divergence', note: 'sjs added rotation' };
    return { status: 'match', note: 'both have rotation' };
}

function compareBorders(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const excHas = hasField(exc, (s) => !!s.bd);
    const sjsHas = hasField(sjs, (s) => !!s.bd);
    if (!excHas && !sjsHas) return { status: 'not-applicable', note: 'no borders' };
    if (excHas && !sjsHas) return { status: 'sheetjs-blocked', note: 'exc captured borders, sjs missed' };
    if (!excHas && sjsHas) return { status: 'divergence', note: 'sjs has borders, exc missed' };
    return { status: 'match', note: 'both have borders' };
}

function compareNumberFormats(exc: Snapshot, sjs: Snapshot): MatrixEntry {
    const excHas = hasField(exc, (s) => !!(s.n as { pattern?: string } | undefined)?.pattern);
    const sjsHas = hasField(sjs, (s) => !!(s.n as { pattern?: string } | undefined)?.pattern);
    if (!excHas && !sjsHas) return { status: 'not-applicable', note: 'no numFmts' };
    if (excHas !== sjsHas) {
        return { status: 'divergence', note: `exc=${excHas} sjs=${sjsHas}` };
    }
    return { status: 'match', note: 'both have numFmts' };
}

function compareCharts(_exc: Snapshot, _sjs: Snapshot, fixturePath: string): MatrixEntry {
    // Charts ride along in the .xlsx zip as xl/charts/* and xl/drawings/*.
    // Neither parser surfaces them as snapshot resources; both pass them
    // through opaquely. The Phase-2 chart-export post-processor's
    // compatibility with sjs-emitted zips is a separate Phase-2 question
    // not measured here. We just check whether the fixture has charts
    // and report 'pass-through' in both cases.
    const isChartFixture = fixturePath.includes('chart-testdata');
    if (!isChartFixture) return { status: 'not-applicable', note: 'not a chart fixture' };
    return { status: 'match', note: 'opaque pass-through (zip-level; data-shape parity untested)' };
}

interface Run {
    fixture: string;
    excSnap: Snapshot | null;
    sjsSnap: Snapshot | null;
    excError?: string;
    sjsError?: string;
}

async function runBoth(fixturePath: string): Promise<Run> {
    const buf = fs.readFileSync(fixturePath);
    const out: Run = { fixture: path.basename(fixturePath), excSnap: null, sjsSnap: null };
    try {
        out.excSnap = (await exceljsParser.xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
    } catch (e) {
        out.excError = (e as Error).message;
    }
    try {
        out.sjsSnap = (await sheetJSParser.xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
    } catch (e) {
        out.sjsError = (e as Error).message;
    }
    return out;
}

const DIMENSIONS: Array<{
    key: string;
    fn: (exc: Snapshot, sjs: Snapshot, fixturePath: string) => MatrixEntry;
}> = [
    { key: 'sheetSizes', fn: compareSheetSizes },
    { key: 'cellValues', fn: compareCellValues },
    { key: 'formulas', fn: compareFormulas },
    { key: 'mergedCells', fn: compareMerges },
    { key: 'styleRecords', fn: compareStyleRecords },
    { key: 'borders', fn: compareBorders },
    { key: 'rotatedText', fn: compareRotatedText },
    { key: 'numberFormats', fn: compareNumberFormats },
    { key: 'hyperlinks', fn: compareHyperlinks },
    { key: 'richText', fn: compareRichText },
    { key: 'themeClrScheme', fn: compareThemeClrScheme },
    { key: 'defaultStyle', fn: compareDefaultStyle },
    { key: 'tables', fn: compareTables },
    { key: 'charts', fn: compareCharts as (a: Snapshot, b: Snapshot, p: string) => MatrixEntry },
];

interface MatrixRow {
    fixture: string;
    excelImported: boolean;
    sheetjsImported: boolean;
    excError?: string;
    sjsError?: string;
    dimensions: Record<string, MatrixEntry>;
}

const formattingFixtures = listFixtures(FORMATTING_DIR);
const chartFixtures = listFixtures(CHART_DIR);

describe('xlsx parser parity (M14 spike)', () => {
    const matrix: MatrixRow[] = [];

    test.each([
        ...formattingFixtures.map((f) => [path.join(FORMATTING_DIR, f), 'formatting'] as const),
        ...chartFixtures.map((f) => [path.join(CHART_DIR, f), 'chart'] as const),
    ])('matrix: %s', async (fixturePath, _category) => {
        const run = await runBoth(fixturePath);
        const row: MatrixRow = {
            fixture: run.fixture,
            excelImported: !!run.excSnap,
            sheetjsImported: !!run.sjsSnap,
            excError: run.excError,
            sjsError: run.sjsError,
            dimensions: {},
        };
        if (run.excSnap && run.sjsSnap) {
            for (const dim of DIMENSIONS) {
                try {
                    row.dimensions[dim.key] = dim.fn(run.excSnap, run.sjsSnap, fixturePath);
                } catch (e) {
                    row.dimensions[dim.key] = { status: 'divergence', note: 'comparator threw: ' + (e as Error).message };
                }
            }
        }
        matrix.push(row);
        // The test always passes — divergences are data, not failures. The
        // matrix is the artefact.
        expect(true).toBe(true);
    });

    afterAll(() => {
        // Sort rows by fixture name so the matrix is deterministic across CI runs.
        matrix.sort((a, b) => a.fixture.localeCompare(b.fixture));
        const dir = path.dirname(MATRIX_OUT);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        const summary = computeSummary(matrix);
        const out = { summary, rows: matrix };
        fs.writeFileSync(MATRIX_OUT, JSON.stringify(out, null, 2) + '\n');
    });
});

function computeSummary(rows: MatrixRow[]): Record<string, Record<string, number>> {
    const summary: Record<string, Record<string, number>> = {};
    for (const dim of DIMENSIONS) {
        summary[dim.key] = { match: 0, divergence: 0, 'sheetjs-blocked': 0, 'exceljs-blocked': 0, 'not-applicable': 0 };
    }
    summary.imports = { excSucceeded: 0, sjsSucceeded: 0, both: 0, excelOnly: 0, sheetjsOnly: 0, neither: 0 };
    for (const row of rows) {
        if (row.excelImported) summary.imports.excSucceeded++;
        if (row.sheetjsImported) summary.imports.sjsSucceeded++;
        if (row.excelImported && row.sheetjsImported) summary.imports.both++;
        else if (row.excelImported && !row.sheetjsImported) summary.imports.excelOnly++;
        else if (!row.excelImported && row.sheetjsImported) summary.imports.sheetjsOnly++;
        else summary.imports.neither++;
        for (const dim of DIMENSIONS) {
            const entry = row.dimensions[dim.key];
            if (!entry) continue;
            summary[dim.key][entry.status]++;
        }
    }
    return summary;
}
