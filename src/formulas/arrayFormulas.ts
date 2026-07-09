// M18 manual-test finding #4: array (CSE) formulas lose their array marker.
//
// A cell like  <c r="F7"><f t="array" ref="F7:F11">A2:A6*2</f></c>  is a
// legacy Ctrl-Shift-Enter array formula that spills over F7:F11. exceljs's
// cell model has no representation for `t="array"` + `ref` (its FormulaType is
// only None/Master/Shared), so on export the formula is written as a plain
// `<f>A2:A6*2</f>`. Excel then reinterprets a range-valued formula in a single
// cell as an implicit-intersection formula and rewrites it to `=@A2:A6*2`,
// which changes the result. The user has to delete the `@` by hand.
//
// We can't fix this through exceljs. Instead — same architecture as the chart
// / image / table readers — we read the array-formula cells zip-direct from
// the ORIGINAL workbook at import, stash them on a snapshot sidecar resource,
// and re-inject the `t="array" ref="..."` attributes into the exported
// worksheet XML with a post-pass after exceljs has written the buffer.
//
// Scope: we preserve the array marker for cells that the source authored as
// array formulas AND that we re-export with the same formula body. We do NOT
// synthesize new array formulas, and we intentionally handle only the master
// cell of the spill (the top-left cell that carries the `<f>` element); the
// spill followers carry no `<f>` in OOXML and Excel recomputes them.

import JSZip from 'jszip';
import { buildFilenameToSheetId } from '../drawings/sheetIdResolver';

// Resource name for the array-formula sidecar. Univer treats unknown
// resources as opaque and round-trips them through reload cycles.
export const NOTESHEET_ARRAY_FORMULAS_RESOURCE = 'SHEET_NOTESHEET_ARRAY_FORMULAS_PLUGIN';

// One captured array formula. `ref` is the spill range (OOXML `ref` attr, e.g.
// "F7:F11"); `formula` is the formula body WITHOUT a leading `=` (matches how
// the snapshot cell stores it minus the `=` we add elsewhere).
export interface ArrayFormulaEntry {
    row: number; // 0-based
    col: number; // 0-based
    ref: string;
    formula: string;
}

// sheetId (`sheet-<n>`) → list of array-formula cells on that sheet.
export type ArrayFormulaMap = Record<string, ArrayFormulaEntry[]>;

// A1 column letters → 0-based column index.
function colLettersToIndex(letters: string): number {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
}

// Parse an A1 cell ref like "F7" into { row, col } (0-based). Returns null on
// anything unexpected.
function parseA1Cell(ref: string): { row: number; col: number } | null {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    if (!m) return null;
    return { col: colLettersToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
}

function decodeXmlText(raw: string): string {
    return raw
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
        .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
        .replace(/&amp;/g, '&');
}

// A1 index → letters (0-based column).
function colIndexToLetters(index: number): string {
    let n = index + 1;
    let s = '';
    while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

// Read every array-formula cell in the workbook, grouped by our snapshot
// sheetId. Zip-direct + regex, tolerant of namespaced/plain markup. Fail-soft:
// any error yields an empty map (import proceeds without the sidecar).
export async function readArrayFormulasFromXlsxZip(
    buffer: ArrayBuffer | Uint8Array | Buffer,
): Promise<ArrayFormulaMap> {
    const out: ArrayFormulaMap = {};
    try {
        const zip = await JSZip.loadAsync(buffer as ArrayBuffer);
        const filenameToSheetId = await buildFilenameToSheetId(zip);

        const sheetPaths = Object.keys(zip.files).filter((p) =>
            /^xl\/worksheets\/sheet\d+\.xml$/i.test(p),
        );
        for (const sheetPath of sheetPaths) {
            const filenameNum = parseInt(
                sheetPath.replace(/^xl\/worksheets\/sheet(\d+)\.xml$/i, '$1'),
                10,
            );
            const resolvedId = filenameToSheetId.get(filenameNum) ?? filenameNum;
            const sheetId = `sheet-${resolvedId}`;
            const xml = await zip.files[sheetPath].async('string');

            // Match each <c r="F7" ...>…<f t="array" ref="F7:F11">body</f>…</c>.
            const cellMatches = xml.match(/<c\b[^>]*\br="[A-Z]+\d+"[^>]*>[\s\S]*?<\/c>/g) ?? [];
            for (const cellXml of cellMatches) {
                const refAttr = cellXml.match(/\br="([A-Z]+\d+)"/);
                if (!refAttr) continue;
                const fMatch = cellXml.match(/<f\b([^>]*)>([\s\S]*?)<\/f>/);
                if (!fMatch) continue;
                const fAttrs = fMatch[1];
                if (!/\bt="array"/.test(fAttrs)) continue;
                const refMatch = fAttrs.match(/\bref="([^"]+)"/);
                if (!refMatch) continue;
                const cell = parseA1Cell(refAttr[1]);
                if (!cell) continue;
                const entry: ArrayFormulaEntry = {
                    row: cell.row,
                    col: cell.col,
                    ref: refMatch[1],
                    formula: decodeXmlText(fMatch[2]),
                };
                if (!out[sheetId]) out[sheetId] = [];
                out[sheetId].push(entry);
            }
        }
    } catch (e) {
        console.warn(
            '[Notesheet] readArrayFormulasFromXlsxZip failed; array markers not preserved',
            e,
        );
        return {};
    }
    return out;
}

// Read the sidecar back off the snapshot at export time.
export function readArrayFormulaSidecar(snapshot: {
    resources?: Array<{ name?: string; data?: string }>;
}): ArrayFormulaMap {
    const resources = snapshot.resources;
    if (!Array.isArray(resources)) return {};
    const entry = resources.find((r) => r?.name === NOTESHEET_ARRAY_FORMULAS_RESOURCE);
    if (!entry || typeof entry.data !== 'string') return {};
    // Univer emits an empty-string `data` for a registered-but-empty resource
    // (no array formulas). That's "nothing to read", not corruption — skip.
    if (entry.data.trim() === '') return {};
    try {
        const parsed = JSON.parse(entry.data);
        if (!parsed || typeof parsed !== 'object') return {};
        return parsed as ArrayFormulaMap;
    } catch (e) {
        console.warn('[Notesheet] array-formula sidecar JSON parse failed; markers dropped', e);
        return {};
    }
}

// Post-pass: rewrite each stashed cell's exported `<f>...</f>` to carry
// `t="array" ref="...">`. Runs AFTER exceljs has written the buffer. Only
// touches cells whose exported formula body matches what we captured (so we
// never mark a formula the user changed). Fail-soft: returns input on error.
export async function injectArrayFormulasIntoZip(
    buffer: ArrayBuffer,
    snapshot: { resources?: Array<{ name?: string; data?: string }>; sheetOrder?: string[] },
): Promise<ArrayBuffer> {
    const sidecar = readArrayFormulaSidecar(snapshot);
    if (Object.keys(sidecar).length === 0) return buffer;
    const sheetOrder = snapshot.sheetOrder ?? [];
    try {
        const zip = await JSZip.loadAsync(buffer);
        let changed = false;

        for (const [sheetId, entries] of Object.entries(sidecar)) {
            if (!entries || entries.length === 0) continue;
            // Export writes worksheets in sheetOrder; file is sheet<idx+1>.xml.
            const idx0 = sheetOrder.indexOf(sheetId);
            if (idx0 < 0) continue;
            const sheetPath = `xl/worksheets/sheet${idx0 + 1}.xml`;
            const file = zip.files[sheetPath];
            if (!file) continue;
            let xml = await file.async('string');
            let fileChanged = false;

            for (const entry of entries) {
                const cellRef = colIndexToLetters(entry.col) + (entry.row + 1);
                // Find this specific cell and its <f>…</f>.
                const cellPattern = new RegExp(
                    `(<c\\b[^>]*\\br="${cellRef}"[^>]*>)([\\s\\S]*?)(</c>)`,
                );
                const cmatch = xml.match(cellPattern);
                if (!cmatch) continue;
                const cellBody = cmatch[2];
                // The <f> must be a plain formula (no t="array" already) whose
                // body equals what we captured — otherwise skip (user edited it).
                const fmatch = cellBody.match(/<f\b([^>]*)>([\s\S]*?)<\/f>/);
                if (!fmatch) continue;
                if (/\bt="array"/.test(fmatch[1])) continue; // already marked
                if (decodeXmlText(fmatch[2]) !== entry.formula) continue; // changed
                const newF = `<f t="array" ref="${entry.ref}"${fmatch[1]}>${fmatch[2]}</f>`;
                const newBody = cellBody.replace(fmatch[0], newF);
                const newCell = cmatch[1] + newBody + cmatch[3];
                xml = xml.replace(cmatch[0], newCell);
                fileChanged = true;
            }

            if (fileChanged) {
                zip.file(sheetPath, xml);
                changed = true;
            }
        }

        if (!changed) return buffer;
        return (await zip.generateAsync({ type: 'arraybuffer' })) as ArrayBuffer;
    } catch (e) {
        console.warn('[Notesheet] injectArrayFormulasIntoZip failed; array markers not written', e);
        return buffer;
    }
}
