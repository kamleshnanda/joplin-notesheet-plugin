// M14 spike — parallel xlsx parser built on xlsx-js-style@1.2.0 (SheetJS Community fork).
//
// **Dead code at runtime.** This module is NOT imported by `src/index.ts` or
// `src/editorView.tsx`; only `tests/xlsxParserParity.test.ts` and
// `tests/m14SheetJSCapability.test.ts` exercise it. The production import path
// stays on `src/xlsx.ts` (exceljs).
//
// Public surface mirrors `src/xlsx.ts`:
//   - xlsxBufferToSnapshot(buffer): Promise<UniverSnapshot>
//   - snapshotToXlsxBuffer(snapshot): Promise<ArrayBuffer>
//   - class NotesheetImportError extends Error
//   - NOTESHEET_SYNTH_STYLES_RESOURCE
//   - NOTESHEET_THEME_CLR_SCHEME_RESOURCE
//
// Spike scope: prove what xlsx-js-style can carry of M9/M12/M13's surface area,
// not a complete drop-in replacement. Wherever a code path in `src/xlsx.ts`
// depends on a SheetJS capability the spike could not crack within the time
// budget, this module emits a `// SPIKE-GAP:` marker and a degraded-but-valid
// snapshot field. Every gap is enumerated in `docs/m14-sheetjs-spike.md`'s
// capability matrix.
//
// Reused parser-agnostic helpers (verbatim, imported from src/xlsx.ts via a
// non-default import — they read raw OOXML XML through JSZip and never touch
// exceljs):
//   - readThemeClrScheme — captures <a:clrScheme> for theme-tinted resolution
//   - readNamedHyperlinkCells — Pattern B hyperlink detection
//
// Resource constants are duplicated literally so the m12FixturePinDowns sentinel
// at line ~226 / ~239 still pins exact `r.name === '...'` matches.

import JSZip from 'jszip';
import * as XLSX from 'xlsx-js-style';

import type { UniverSnapshot } from './snapshot';

// CRITICAL: these strings MUST match `src/xlsx.ts` byte-for-byte. The
// `tests/m12FixturePinDowns.test.ts` sentinels do an exact `r.name === '...'`
// lookup in the snapshot resources array, so a divergence here would pass
// xlsxSheetJS-side parity but fail snapshot consumption in the production
// editor on a future migration.
export const NOTESHEET_SYNTH_STYLES_RESOURCE = 'SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN';
export const NOTESHEET_THEME_CLR_SCHEME_RESOURCE = 'SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN';

// Same shape as `src/xlsx.ts:NotesheetImportError`. The `code` and `cause`
// fields are part of the public contract — UI code in `src/index.ts` and
// `src/editorView.tsx` keys off the `code` strings to render user-actionable
// dialog messages.
export class NotesheetImportError extends Error {
    public readonly code: string;
    public readonly cause?: unknown;
    constructor(code: string, message: string, cause?: unknown) {
        super(message);
        this.name = 'NotesheetImportError';
        this.code = code;
        this.cause = cause;
    }
}

// Univer numeric enums (same as src/xlsx.ts — kept local so the spike module
// doesn't depend on @univerjs/core in jest's node environment).
const HORIZONTAL: Record<string, number> = { left: 1, center: 2, right: 3 };
const VERTICAL: Record<string, number> = { top: 1, center: 2, bottom: 3 };
const WRAP_STRATEGY_WRAP = 3;
const VALUE_STRING = 1;
const VALUE_NUMBER = 2;
const VALUE_BOOLEAN = 3;

// xlsx-js-style border type → Univer numeric (mirror of BORDER_STYLE_TO_UNIVER).
const BORDER_STYLE_TO_UNIVER: Record<string, number> = {
    thin: 1,
    hair: 2,
    dotted: 3,
    dashed: 4,
    dashDot: 5,
    dashDotDot: 6,
    double: 7,
    medium: 8,
    mediumDashed: 9,
    mediumDashDot: 10,
    mediumDashDotDot: 11,
    slantDashDot: 12,
    thick: 13,
};

interface ThemePalette {
    raw: string;
    rgb: string[];
}

// Pure helper — could be imported from src/xlsx.ts but the spike duplicates it
// to keep the module self-contained for the parity test bed (avoids a circular
// dependency between the two parser modules during ts-jest type resolution).
async function readThemeClrScheme(buffer: ArrayBuffer | Uint8Array | Buffer): Promise<ThemePalette | null> {
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return null;
    }
    const themePath = Object.keys(zip.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p));
    if (!themePath) return null;
    const xml = await zip.files[themePath].async('string');
    const m = /<a:clrScheme\b[^>]*>[\s\S]*?<\/a:clrScheme>/.exec(xml);
    if (!m) return null;
    const raw = m[0];
    const ELEMENT_ORDER: Array<'lt1' | 'dk1' | 'lt2' | 'dk2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'hlink' | 'folHlink'> = [
        'lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink',
    ];
    const rgb: string[] = [];
    for (const elName of ELEMENT_ORDER) {
        const elRe = new RegExp(`<a:${elName}\\b[^>]*>([\\s\\S]*?)</a:${elName}>`);
        const elMatch = elRe.exec(raw);
        if (!elMatch) { rgb.push('#000000'); continue; }
        const inner = elMatch[1];
        const srgb = /<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(inner);
        const sys = /<a:sysClr\b[^>]*\blastClr="([0-9A-Fa-f]{6})"/.exec(inner);
        const hex = (srgb?.[1] ?? sys?.[1] ?? '000000').toUpperCase();
        rgb.push('#' + hex);
    }
    return { raw, rgb };
}

// Pure helper — Pattern B hyperlink detection via raw zip read. Same shape
// as src/xlsx.ts:readNamedHyperlinkCells.
async function readNamedHyperlinkCells(
    buffer: ArrayBuffer | Uint8Array | Buffer,
): Promise<Map<number, Set<string>>> {
    const result = new Map<number, Set<string>>();
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return result;
    }

    const stylesPath = 'xl/styles.xml';
    if (!zip.files[stylesPath]) return result;
    const stylesXml = await zip.files[stylesPath].async('string');

    const cellStylesMatch = /<cellStyles\b[^>]*>([\s\S]*?)<\/cellStyles>/.exec(stylesXml);
    if (!cellStylesMatch) return result;
    const namedHyperlinkXfIds = new Set<string>();
    const styleEntryRe = /<cellStyle\b[^>]*\/>/g;
    let m: RegExpExecArray | null;
    while ((m = styleEntryRe.exec(cellStylesMatch[1])) !== null) {
        const tag = m[0];
        const isHyperlink = /\bname="Hyperlink"/.test(tag) || /\bbuiltinId="8"/.test(tag);
        if (!isHyperlink) continue;
        const xfIdMatch = /\bxfId="(\d+)"/.exec(tag);
        if (xfIdMatch) namedHyperlinkXfIds.add(xfIdMatch[1]);
    }
    if (namedHyperlinkXfIds.size === 0) return result;

    const cellXfsMatch = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
    if (!cellXfsMatch) return result;
    const xfRe = /<xf\b[^>]*\/>/g;
    const linkCellXfIndices = new Set<number>();
    let xfIdx = 0;
    let xfMatch: RegExpExecArray | null;
    while ((xfMatch = xfRe.exec(cellXfsMatch[1])) !== null) {
        const tag = xfMatch[0];
        const xfIdAttr = /\bxfId="(\d+)"/.exec(tag);
        if (xfIdAttr && namedHyperlinkXfIds.has(xfIdAttr[1])) {
            linkCellXfIndices.add(xfIdx);
        }
        xfIdx++;
    }
    if (linkCellXfIndices.size === 0) return result;

    for (const fpath of Object.keys(zip.files)) {
        const sheetMatch = /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(fpath);
        if (!sheetMatch) continue;
        const sheetIdx = parseInt(sheetMatch[1], 10);
        const sheetXml = await zip.files[fpath].async('string');
        const cellRe = /<c\b[^>]*\/?>/g;
        const a1Set = new Set<string>();
        let cellMatch: RegExpExecArray | null;
        while ((cellMatch = cellRe.exec(sheetXml)) !== null) {
            const tag = cellMatch[0];
            const sAttr = /\bs="(\d+)"/.exec(tag);
            if (!sAttr) continue;
            if (!linkCellXfIndices.has(parseInt(sAttr[1], 10))) continue;
            const rAttr = /\br="([A-Z]+\d+)"/.exec(tag);
            if (rAttr) a1Set.add(rAttr[1]);
        }
        if (a1Set.size > 0) result.set(sheetIdx, a1Set);
    }

    return result;
}


function rgbHex(rgb: string | undefined): string | undefined {
    if (!rgb || typeof rgb !== 'string') return undefined;
    const trimmed = rgb.replace(/^#/, '');
    if (trimmed.length === 6) return '#' + trimmed.toUpperCase();
    if (trimmed.length === 8) return '#' + trimmed.slice(2).toUpperCase();
    return undefined;
}

function applyOoxmlTint(hex: string, tint: number): string {
    const r = parseInt(hex.slice(1, 3), 16) / 255;
    const g = parseInt(hex.slice(3, 5), 16) / 255;
    const b = parseInt(hex.slice(5, 7), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    let l = (max + min) / 2;
    if (max !== min) {
        const d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
        else if (max === g) h = (b - r) / d + 2;
        else h = (r - g) / d + 4;
        h /= 6;
    }
    if (tint < 0) l = l * (1 + tint);
    else l = l * (1 - tint) + tint;
    const hue2rgb = (p: number, q: number, t: number): number => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
    };
    let r2: number, g2: number, b2: number;
    if (s === 0) { r2 = g2 = b2 = l; }
    else {
        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        r2 = hue2rgb(p, q, h + 1 / 3);
        g2 = hue2rgb(p, q, h);
        b2 = hue2rgb(p, q, h - 1 / 3);
    }
    const toHex = (x: number) => Math.round(x * 255).toString(16).padStart(2, '0');
    return ('#' + toHex(r2) + toHex(g2) + toHex(b2)).toUpperCase();
}

// xlsx-js-style colour shape: { rgb: 'RRGGBB' } | { theme: N, tint: T, rgb?: 'pre-resolved' }.
// Note: with `cellStyles: true` the parser sometimes pre-resolves theme refs
// to an `rgb` field (see Fonts entries in BordersAndCellColors.xlsx — font[2]
// has both `theme: 7` AND `rgb: 'B3A2C7'`). When that's present, prefer the
// pre-resolved rgb. Otherwise fall back to our palette resolver.
type SheetJSColor = { rgb?: string; theme?: number; tint?: number } | null | undefined;
function resolveSheetJSColor(color: SheetJSColor, palette: ThemePalette | null): string | undefined {
    if (!color || typeof color !== 'object') return undefined;
    // Pre-resolved rgb — xlsx-js-style does this for theme refs in Fonts.
    if (color.rgb) {
        const h = rgbHex(color.rgb);
        if (h) return h;
    }
    if (typeof color.theme === 'number' && palette) {
        const base = palette.rgb[color.theme];
        if (!base) return undefined;
        return typeof color.tint === 'number' && color.tint !== 0 ? applyOoxmlTint(base, color.tint) : base;
    }
    return undefined;
}

// xlsx-js-style cell-style object → Univer per-cell style record.
//
// **SPIKE-GAP (critical, drives NO-GO recommendation):** xlsx-js-style@1.2.0
// does NOT propagate borders, alignment, or font formatting from the styles
// registry (`wb.Styles.CellXf`) into the per-cell `c.s` field for cells that
// reference a styles.xml-indexed cellXf via `s="N"` (the OOXML standard).
// Verified empirically against `BordersAndCellColors.xlsx` and
// `MergedCellsAndAlignment.xlsx`: cells styled in Excel come back with
// `c.s = { patternType: 'none' }` regardless of their actual border / font
// / alignment / rotation. Only fills (background colour) propagate, and only
// for cells whose fill was set inside Excel proper (not table-style synthesis).
//
// To work around this without abandoning xlsx-js-style entirely, this function
// can ALSO consume a `resolvedStyleByRef` map (built by walking the raw
// styles.xml + sheetN.xml ourselves — same shape as `readNamedHyperlinkCells`
// but generalised). That walker isn't implemented in this spike — the matrix
// flags it as a 1-2 day Phase-2 cost. Without it, the snapshot's `styles`
// registry is materially under-populated versus exceljs.
function buildStyleFromSheetJSCell(
    cell: XLSX.CellObject,
    palette: ThemePalette | null,
): Record<string, unknown> | null {
    const style: Record<string, unknown> = {};
    const s = cell.s as Record<string, unknown> | undefined;

    if (s && typeof s === 'object') {
        // Font.
        const font = s.font as Record<string, unknown> | undefined;
        if (font && typeof font === 'object') {
            if (typeof font.name === 'string') style.ff = font.name;
            if (typeof font.sz === 'number') style.fs = font.sz;
            if (font.bold) style.bl = 1;
            if (font.italic) style.it = 1;
            if (font.underline) style.ul = { s: 1 };
            if (font.strike) style.st = { s: 1 };
            const fontColor = resolveSheetJSColor(font.color as SheetJSColor, palette);
            if (fontColor) style.cl = { rgb: fontColor };
        }

        // Fill (background). xlsx-js-style mirrors fgColor in the solid-fill case.
        const fill = s.fill as Record<string, unknown> | undefined;
        const patternType = s.patternType ?? fill?.patternType;
        const fgColor = (s.fgColor as SheetJSColor) ?? (fill?.fgColor as SheetJSColor);
        if (patternType === 'solid' && fgColor) {
            const bg = resolveSheetJSColor(fgColor, palette);
            if (bg) style.bg = { rgb: bg };
        }

        // Alignment. SPIKE-GAP: not present on cells we tested (xlsx-js-style
        // drops alignment from the styles.xml-driven path); kept structural
        // for forward compatibility if a Phase-2 walker fills it in.
        const align = s.alignment as Record<string, unknown> | undefined;
        if (align && typeof align === 'object') {
            const horizontal = align.horizontal as string | undefined;
            const vertical = align.vertical as string | undefined;
            if (horizontal && HORIZONTAL[horizontal] !== undefined) style.ht = HORIZONTAL[horizontal];
            // xlsx-js-style uses 'center' for vertical, exceljs uses 'middle'.
            if (vertical) {
                const v = vertical === 'middle' ? 'center' : vertical;
                if (VERTICAL[v] !== undefined) style.vt = VERTICAL[v];
            }
            if (align.wrapText) style.tb = WRAP_STRATEGY_WRAP;
            const rot = align.textRotation;
            if (rot === 255) style.tr = { a: 0, v: 1 };
            else if (typeof rot === 'number' && rot !== 0) style.tr = { a: rot };
        }

        // Borders. SPIKE-GAP: xlsx-js-style ships an empty {} for every
        // border record in `wb.Styles.Borders` for every Microsoft-Excel-
        // generated fixture we tested. Self-roundtrip works (fork-written
        // borders survive a fork-read), but parser interop is broken for
        // imported borders. Code path retained for the rare case the source
        // file has borders xlsx-js-style happens to recognise.
        const border = s.border as Record<string, unknown> | undefined;
        if (border && typeof border === 'object') {
            const bd: Record<string, { s: number; cl: { rgb: string } }> = {};
            const SIDES: Array<['t' | 'r' | 'b' | 'l', string]> = [
                ['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left'],
            ];
            for (const [univerKey, sjsKey] of SIDES) {
                const side = border[sjsKey] as Record<string, unknown> | undefined;
                if (!side || typeof side !== 'object') continue;
                const styleName = side.style as string | undefined;
                if (!styleName) continue;
                const styleNum = BORDER_STYLE_TO_UNIVER[styleName];
                if (styleNum === undefined) continue;
                const colorHex = resolveSheetJSColor(side.color as SheetJSColor, palette) ?? '#000000';
                bd[univerKey] = { s: styleNum, cl: { rgb: colorHex } };
            }
            if (Object.keys(bd).length > 0) style.bd = bd;
        }

        // Number format.
        const numFmt = (s.numFmt as string | undefined) ?? (cell.z as string | undefined);
        if (numFmt && numFmt !== 'General' && numFmt !== '0') {
            style.n = { pattern: numFmt };
        }
    } else if (cell.z && cell.z !== 'General') {
        // Even when xlsx-js-style strips s, the per-cell z (number format)
        // is preserved on the cell object directly.
        style.n = { pattern: cell.z as string };
    }

    return Object.keys(style).length > 0 ? style : null;
}

// Build Univer cell.p from a hyperlink + display text. Same shape as
// src/xlsx.ts:buildHyperlinkCellP — Univer's hyperlink layer reads
// customRanges[].rangeType=0 (HYPERLINK) and the dataStream slice it
// addresses. Kept minimal here; richer cases (run formatting inside the
// link) are out of scope for the spike.
function buildHyperlinkCellP(text: string, url: string): Record<string, unknown> {
    const dataStream = text + '\r\n';
    const en = text.length;
    return {
        id: 'p',
        body: {
            dataStream,
            customRanges: [
                {
                    rangeId: 'r-1',
                    rangeType: 0,
                    startIndex: 0,
                    endIndex: en > 0 ? en - 1 : 0,
                    properties: { url },
                },
            ],
            paragraphs: [{ startIndex: en }],
        },
        documentStyle: {},
    };
}

// Read inline rich-text from the raw <c r="..." t="inlineStr"><is>...</is></c>.
// SheetJS collapses <r><rPr>...<t>...</t></r> runs into a single string in
// `c.v` and a stringified HTML in `c.h`, so per-run formatting is lost on the
// SheetJS side. This helper parses the raw XML to recover runs — used as a
// shimming layer when M13/D fidelity matters. Returns an array of runs or null.
//
// **SPIKE-GAP:** even with this shim, building a Univer rich-text cell.p from
// the runs requires the same logic as src/xlsx.ts:buildRichTextCellP (~70
// lines). The spike does NOT port that logic — instead it documents the
// presence of runs as a divergence in the parity matrix and falls back to a
// single-run plain-text cell. Phase 2 must port buildRichTextCellP verbatim.
interface RawRun {
    text: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    rgb?: string;
    fontName?: string;
    fontSize?: number;
}
function parseInlineRichRuns(xml: string): RawRun[] | null {
    // Match <c r="X" t="inlineStr"><is>...</is></c>. The caller is
    // expected to pass the inner <is>...</is> body. If we can't find <r>
    // tags, return null (fall through to plain text).
    const runRe = /<r\b[^>]*>([\s\S]*?)<\/r>/g;
    const runs: RawRun[] = [];
    let m: RegExpExecArray | null;
    while ((m = runRe.exec(xml)) !== null) {
        const inner = m[1];
        const textMatch = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(inner);
        if (!textMatch) continue;
        const run: RawRun = { text: decodeXmlEntities(textMatch[1]) };
        const rPrMatch = /<rPr\b[^>]*>([\s\S]*?)<\/rPr>/.exec(inner);
        if (rPrMatch) {
            const rPr = rPrMatch[1];
            if (/<b(?:\s|\/|>)/.test(rPr)) run.bold = true;
            if (/<i(?:\s|\/|>)/.test(rPr)) run.italic = true;
            if (/<u(?:\s|\/|>)/.test(rPr)) run.underline = true;
            const colorMatch = /<color\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/.exec(rPr);
            if (colorMatch) {
                const rgb = colorMatch[1].length === 8 ? colorMatch[1].slice(2) : colorMatch[1];
                run.rgb = '#' + rgb.toUpperCase();
            }
            const nameMatch = /<rFont\b[^>]*\bval="([^"]+)"/.exec(rPr);
            if (nameMatch) run.fontName = nameMatch[1];
            const szMatch = /<sz\b[^>]*\bval="([0-9.]+)"/.exec(rPr);
            if (szMatch) run.fontSize = parseFloat(szMatch[1]);
        }
        runs.push(run);
    }
    return runs.length > 0 ? runs : null;
}

function decodeXmlEntities(s: string): string {
    return s
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
}

// Read the raw sharedStrings.xml + worksheet XML to recover rich-text runs
// for cells xlsx-js-style flattened. Returns sheet1-based index → A1 → runs.
async function readInlineRichTextRuns(
    buffer: ArrayBuffer | Uint8Array | Buffer,
): Promise<Map<number, Map<string, RawRun[]>>> {
    const result = new Map<number, Map<string, RawRun[]>>();
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return result;
    }

    // Two cases: inline strings (`t="inlineStr"`) and shared strings (`t="s"`
    // pointing at sharedStrings.xml index). Both can carry <r><rPr> runs.
    let ssEntries: Array<RawRun[] | null> | null = null;
    if (zip.files['xl/sharedStrings.xml']) {
        const ssXml = await zip.files['xl/sharedStrings.xml'].async('string');
        ssEntries = [];
        const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
        let sm: RegExpExecArray | null;
        while ((sm = siRe.exec(ssXml)) !== null) {
            const inner = sm[1];
            const runs = parseInlineRichRuns(inner);
            ssEntries.push(runs && runs.length >= 2 ? runs : null);
        }
    }

    for (const fpath of Object.keys(zip.files)) {
        const sheetMatch = /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(fpath);
        if (!sheetMatch) continue;
        const sheetIdx = parseInt(sheetMatch[1], 10);
        const sheetXml = await zip.files[fpath].async('string');
        const cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
        const sheetMap = new Map<string, RawRun[]>();
        let cm: RegExpExecArray | null;
        while ((cm = cellRe.exec(sheetXml)) !== null) {
            const attrs = cm[1];
            const body = cm[2];
            const rAttr = /\br="([A-Z]+\d+)"/.exec(attrs);
            if (!rAttr) continue;
            const tAttr = /\bt="([^"]+)"/.exec(attrs);
            const cellType = tAttr?.[1];
            if (cellType === 'inlineStr') {
                const isMatch = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(body);
                if (!isMatch) continue;
                const runs = parseInlineRichRuns(isMatch[1]);
                if (runs && runs.length >= 2) sheetMap.set(rAttr[1], runs);
            } else if (cellType === 's' && ssEntries) {
                const vMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body);
                if (!vMatch) continue;
                const idx = parseInt(vMatch[1], 10);
                const runs = ssEntries[idx];
                if (runs) sheetMap.set(rAttr[1], runs);
            }
        }
        if (sheetMap.size > 0) result.set(sheetIdx, sheetMap);
    }
    return result;
}

// SPIKE-GAP: building a Univer rich-text cell.p from RawRun[] is the same
// logic as src/xlsx.ts:buildRichTextCellP. The spike provides a minimal
// version that emits the textRuns + dataStream so Univer's renderer can pick
// up colour/bold runs. Hyperlink-bearing runs are not handled (Pattern A
// hyperlink emission wins on export anyway).
function buildRichTextCellPFromRuns(runs: RawRun[]): Record<string, unknown> {
    let stream = '';
    const textRuns: Array<{ st: number; ed: number; ts: Record<string, unknown> }> = [];
    for (const r of runs) {
        const st = stream.length;
        stream += r.text;
        const ed = stream.length;
        const ts: Record<string, unknown> = {};
        if (r.fontName) ts.ff = r.fontName;
        if (typeof r.fontSize === 'number') ts.fs = r.fontSize;
        if (r.bold) ts.bl = 1;
        if (r.italic) ts.it = 1;
        if (r.underline) ts.ul = { s: 1 };
        if (r.rgb) ts.cl = { rgb: r.rgb };
        textRuns.push({ st, ed, ts });
    }
    return {
        id: 'p',
        body: {
            dataStream: stream + '\r\n',
            textRuns,
            paragraphs: [{ startIndex: stream.length }],
        },
        documentStyle: {},
    };
}

// Map xlsx-js-style cell types to Univer's CellValueType enum.
function mapCellValueType(t: string | undefined): number | undefined {
    switch (t) {
        case 's': return VALUE_STRING;
        case 'str': return VALUE_STRING;
        case 'n': return VALUE_NUMBER;
        case 'b': return VALUE_BOOLEAN;
        case 'd': return VALUE_NUMBER; // dates serialize as numbers in Univer's snapshot
        case 'inlineStr': return VALUE_STRING;
        default: return undefined;
    }
}

// Coerce a SheetJS cell value to the {string|number|boolean} shape Univer
// snapshots expect.
function coerceCellValue(cell: XLSX.CellObject): string | number | boolean | undefined {
    if (cell.v === null || cell.v === undefined) return undefined;
    if (cell.v instanceof Date) {
        // Match exceljs by emitting an ISO-ish string; future Phase-2 work
        // can switch to OOXML serial date if Univer prefers that. Documented
        // as a divergence in the matrix.
        return cell.v.toISOString();
    }
    return cell.v;
}

interface CellRecord {
    v?: string | number | boolean;
    f?: string;
    t?: number;
    s?: string;
    p?: Record<string, unknown>;
}

interface SheetRecord {
    id: string;
    name: string;
    cellData: Record<number, Record<number, CellRecord>>;
    rowCount: number;
    columnCount: number;
    defaultColumnWidth: number;
    defaultRowHeight: number;
    mergeData?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
    rowData?: Record<number, { h?: number }>;
    columnData?: Record<number, { w?: number }>;
}

function sortedJsonStringify(value: unknown): string {
    if (value === null || typeof value !== 'object') return JSON.stringify(value);
    if (Array.isArray(value)) return '[' + value.map(sortedJsonStringify).join(',') + ']';
    const obj = value as Record<string, unknown>;
    const keys = Object.keys(obj).sort();
    return '{' + keys.map((k) => JSON.stringify(k) + ':' + sortedJsonStringify(obj[k])).join(',') + '}';
}

export async function xlsxBufferToSnapshot(buffer: ArrayBuffer | Uint8Array | Buffer): Promise<UniverSnapshot> {
    let wb: XLSX.WorkBook;
    try {
        // Convert ArrayBuffer to Buffer in Node — xlsx-js-style accepts both
        // but type === 'buffer' wants Node Buffer, type === 'array' wants
        // Uint8Array. Normalise via Buffer.from when possible.
        const data = buffer instanceof ArrayBuffer
            ? Buffer.from(new Uint8Array(buffer))
            : Buffer.isBuffer(buffer)
                ? buffer
                : Buffer.from(buffer);
        wb = XLSX.read(data, {
            type: 'buffer',
            cellStyles: true,
            cellNF: true,
            cellDates: true,
            cellFormula: true,
            cellHTML: false,
            sheetStubs: false,
            // bookFiles=true is needed to access wb.Styles for borders in a
            // future Phase-2 walker. Not used in this spike's snapshot output
            // (borders come back empty regardless), but harmless.
            bookFiles: true,
        });
    } catch (err) {
        const e = err as Error;
        throw new NotesheetImportError('xlsx-import-failed', e?.message ?? String(err), err);
    }

    const themeClrScheme = await readThemeClrScheme(buffer);
    const namedHyperlinkCellsBySheet = await readNamedHyperlinkCells(buffer);
    const inlineRichTextBySheet = await readInlineRichTextRuns(buffer);

    const sheetOrder: string[] = [];
    const sheets: Record<string, SheetRecord> = {};
    const styles: Record<string, Record<string, unknown>> = {};
    const styleIdByKey = new Map<string, string>();
    let nextStyleId = 1;

    function internStyle(style: Record<string, unknown> | null): string | undefined {
        if (!style) return undefined;
        const key = sortedJsonStringify(style);
        const existing = styleIdByKey.get(key);
        if (existing) return existing;
        const id = 'style-' + nextStyleId++;
        styleIdByKey.set(key, id);
        styles[id] = style;
        return id;
    }

    const sheetNames = wb.SheetNames || [];
    for (let sIdx = 0; sIdx < sheetNames.length; sIdx++) {
        const sheetName = sheetNames[sIdx];
        const ws = wb.Sheets[sheetName];
        if (!ws) continue;
        // Sheet id format mirrors src/xlsx.ts: 'sheet-' + 1-based index.
        // exceljs's ws.id == array index + 1; xlsx-js-style doesn't expose
        // ws.id but the SheetNames array order is the canonical sheet order.
        const sheetId = 'sheet-' + (sIdx + 1);
        sheetOrder.push(sheetId);

        const cellData: Record<number, Record<number, CellRecord>> = {};
        let maxRow = 0;
        let maxCol = 0;
        const namedHyperlinkCells = namedHyperlinkCellsBySheet.get(sIdx + 1) ?? new Set<string>();
        const richTextCells = inlineRichTextBySheet.get(sIdx + 1) ?? new Map<string, RawRun[]>();

        for (const k of Object.keys(ws)) {
            if (k.startsWith('!')) continue;
            const cell = ws[k] as XLSX.CellObject;
            if (!cell) continue;
            const addr = XLSX.utils.decode_cell(k);
            const r = addr.r;
            const c = addr.c;
            const record: CellRecord = {};

            // Hyperlink Pattern A — xlsx-js-style sets c.l.Target.
            const hyperlinkUrl = (cell.l && typeof cell.l === 'object'
                && typeof (cell.l as { Target?: unknown }).Target === 'string')
                ? (cell.l as { Target: string }).Target
                : undefined;

            // Rich-text via raw-XML shim. Wins over the plain string from
            // xlsx-js-style for cells we recovered runs for.
            const richRuns = richTextCells.get(k);

            const value = coerceCellValue(cell);
            if (cell.f) {
                record.f = cell.f.startsWith('=') ? cell.f : '=' + cell.f;
                if (value !== undefined) record.v = value;
                const t = mapCellValueType(cell.t);
                if (t !== undefined) record.t = t;
            } else if (hyperlinkUrl && value !== undefined) {
                record.v = String(value);
                record.t = VALUE_STRING;
                record.p = buildHyperlinkCellP(String(value), hyperlinkUrl);
            } else if (richRuns) {
                record.v = String(value ?? '');
                record.t = VALUE_STRING;
                record.p = buildRichTextCellPFromRuns(richRuns);
            } else if (value !== undefined) {
                record.v = value;
                const t = mapCellValueType(cell.t);
                if (t !== undefined) record.t = t;
            }

            // Pattern B hyperlink synthesis (named-style xfId pointing at builtin Hyperlink).
            if (!record.p && typeof record.v === 'string' && namedHyperlinkCells.has(k)) {
                record.p = buildHyperlinkCellP(record.v, record.v);
            }

            const style = buildStyleFromSheetJSCell(cell, themeClrScheme);
            const styleId = internStyle(style);
            if (styleId) record.s = styleId;
            if (Object.keys(record).length === 0) continue;
            if (!cellData[r]) cellData[r] = {};
            cellData[r][c] = record;
            if (r > maxRow) maxRow = r;
            if (c > maxCol) maxCol = c;
        }

        const mergeData: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }> = [];
        const merges = ws['!merges'] as XLSX.Range[] | undefined;
        if (Array.isArray(merges)) {
            for (const m of merges) {
                if (m && m.s && m.e) {
                    mergeData.push({
                        startRow: m.s.r,
                        endRow: m.e.r,
                        startColumn: m.s.c,
                        endColumn: m.e.c,
                    });
                }
            }
        }

        // Row heights / column widths.
        const rowData: Record<number, { h?: number }> = {};
        const rowsInfo = ws['!rows'] as XLSX.RowInfo[] | undefined;
        if (Array.isArray(rowsInfo)) {
            rowsInfo.forEach((info, idx) => {
                if (!info) return;
                if (typeof info.hpt === 'number' && info.hpt > 0) {
                    rowData[idx] = { h: Math.round(info.hpt * (96 / 72)) };
                } else if (typeof info.hpx === 'number' && info.hpx > 0) {
                    rowData[idx] = { h: Math.round(info.hpx) };
                }
            });
        }

        const columnData: Record<number, { w?: number }> = {};
        const colsInfo = ws['!cols'] as XLSX.ColInfo[] | undefined;
        if (Array.isArray(colsInfo)) {
            colsInfo.forEach((info, idx) => {
                if (!info) return;
                if (typeof info.wpx === 'number' && info.wpx > 0) {
                    columnData[idx] = { w: Math.round(info.wpx) };
                } else if (typeof info.wch === 'number' && info.wch > 0) {
                    columnData[idx] = { w: Math.round(info.wch * 7 + 5) };
                } else if (typeof info.width === 'number' && info.width > 0) {
                    columnData[idx] = { w: Math.round(info.width * 7 + 5) };
                }
            });
        }

        // Use !ref if present to determine the sheet's active range — this
        // matches Excel's cell-region bound, which can be wider than the
        // populated cells (trailing empty columns inside a table, etc.).
        const refStr = ws['!ref'] as string | undefined;
        let refMaxRow = maxRow;
        let refMaxCol = maxCol;
        if (refStr) {
            try {
                const range = XLSX.utils.decode_range(refStr);
                refMaxRow = Math.max(refMaxRow, range.e.r);
                refMaxCol = Math.max(refMaxCol, range.e.c);
            } catch {
                // ignore — !ref malformed, fall back to populated bounds.
            }
        }

        sheets[sheetId] = {
            id: sheetId,
            name: sheetName,
            cellData,
            rowCount: Math.max(100, refMaxRow + 1),
            columnCount: Math.max(26, refMaxCol + 1),
            defaultColumnWidth: 73,
            defaultRowHeight: 19,
            mergeData,
            rowData,
            columnData,
        };
    }

    if (sheetOrder.length === 0) {
        const id = 'sheet-1';
        sheetOrder.push(id);
        sheets[id] = {
            id,
            name: 'Sheet1',
            cellData: {},
            rowCount: 100,
            columnCount: 26,
            defaultColumnWidth: 73,
            defaultRowHeight: 19,
            mergeData: [],
            rowData: {},
            columnData: {},
        };
    }

    const resources: Array<{ name: string; data: string }> = [];
    if (themeClrScheme) {
        resources.push({
            name: NOTESHEET_THEME_CLR_SCHEME_RESOURCE,
            data: themeClrScheme.raw,
        });
    }
    // SPIKE-GAP: SHEET_TABLE_PLUGIN and NOTESHEET_SYNTH_STYLES_RESOURCE
    // require porting the exceljs-driven table parser AND the
    // synthesizeTableStyleAssignments pipeline. The spike does not implement
    // either — Phase 2 must port both. Documented in the matrix.

    return {
        id: 'workbook-' + Date.now(),
        sheetOrder,
        name: 'Spreadsheet',
        appVersion: '0.1.0',
        locale: 'enUS',
        styles,
        sheets,
        ...(resources.length > 0 ? { resources } : {}),
    } as unknown as UniverSnapshot;
}

// SPIKE-GAP: snapshotToXlsxBuffer is a minimal port that handles cell values,
// formulas, merges, and basic styling (font, fill, alignment, border, numFmt).
// It does NOT handle:
//   - Tables (xlsx-js-style does not have an addTable equivalent — the
//     table.xml + relationships must be written directly into the zip via
//     JSZip).
//   - Theme palette splice (the parser-agnostic patchThemeClrScheme helper
//     could be reused, but the Phase-2 port hasn't been written).
//   - Chart injection (src/charts/xlsxChart.ts patches the zip exceljs
//     emits; the SheetJS-emitted zip has a different internal structure
//     that the patcher must be retargeted to).
//   - Rich-text per-run export.
//   - Pattern B hyperlinks (named-style cellStyle).
//
// This export is sufficient to round-trip simple value-only fixtures for the
// parity test bed. It is NOT sufficient to ship a Phase-2 production swap.
export async function snapshotToXlsxBuffer(snapshot: UniverSnapshot): Promise<ArrayBuffer> {
    const sheetOrder = (snapshot as { sheetOrder?: string[] }).sheetOrder ?? [];
    const sheets = (snapshot as { sheets?: Record<string, SheetRecord> }).sheets ?? {};
    const stylesMap = (snapshot as { styles?: Record<string, Record<string, unknown>> }).styles ?? {};

    const wb = XLSX.utils.book_new();
    for (const sheetId of sheetOrder) {
        const sheet = sheets[sheetId];
        if (!sheet) continue;
        const ws: XLSX.WorkSheet = {};
        const cellData = sheet.cellData ?? {};
        let maxR = 0, maxC = 0;
        for (const rowKey of Object.keys(cellData)) {
            const r = Number(rowKey);
            const row = cellData[r];
            if (!row) continue;
            for (const colKey of Object.keys(row)) {
                const c = Number(colKey);
                const data = row[c];
                if (!data) continue;
                const addr = XLSX.utils.encode_cell({ r, c });
                const cellObj: XLSX.CellObject = { t: 's', v: '' };
                if (data.f) {
                    const formula = data.f.startsWith('=') ? data.f.slice(1) : data.f;
                    cellObj.f = formula;
                    if (data.v !== undefined) cellObj.v = data.v;
                    cellObj.t = typeof data.v === 'number' ? 'n' : 'n';
                } else if (data.v !== undefined) {
                    cellObj.v = data.v;
                    if (typeof data.v === 'boolean') cellObj.t = 'b';
                    else if (typeof data.v === 'number') cellObj.t = 'n';
                    else cellObj.t = 's';
                }
                // Hyperlink (Pattern A).
                const hyperlinkUrl = extractHyperlinkFromCellP(data.p);
                if (hyperlinkUrl) {
                    cellObj.l = { Target: hyperlinkUrl } as XLSX.Hyperlink;
                }
                // Style.
                if (data.s) {
                    const style = stylesMap[data.s];
                    if (style) {
                        const sjsStyle = buildSheetJSStyleFromUniver(style);
                        if (sjsStyle && Object.keys(sjsStyle).length > 0) {
                            cellObj.s = sjsStyle;
                        }
                    }
                }
                ws[addr] = cellObj;
                if (r > maxR) maxR = r;
                if (c > maxC) maxC = c;
            }
        }
        const rangeEnd = XLSX.utils.encode_cell({ r: Math.max(maxR, 0), c: Math.max(maxC, 0) });
        ws['!ref'] = 'A1:' + rangeEnd;
        if (Array.isArray(sheet.mergeData) && sheet.mergeData.length > 0) {
            ws['!merges'] = sheet.mergeData.map((m) => ({
                s: { r: m.startRow, c: m.startColumn },
                e: { r: m.endRow, c: m.endColumn },
            }));
        }
        XLSX.utils.book_append_sheet(wb, ws, sheet.name || sheetId);
    }

    if (!wb.SheetNames || wb.SheetNames.length === 0) {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[]]), 'Sheet1');
    }

    const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
    // Convert Buffer/Uint8Array → ArrayBuffer to match src/xlsx.ts return shape.
    if (out instanceof ArrayBuffer) return out;
    if (Buffer.isBuffer(out)) {
        const ab = new ArrayBuffer(out.length);
        new Uint8Array(ab).set(out);
        return ab;
    }
    if (out instanceof Uint8Array) {
        const ab = new ArrayBuffer(out.byteLength);
        new Uint8Array(ab).set(out);
        return ab;
    }
    // String fallback.
    return new TextEncoder().encode(String(out)).buffer as ArrayBuffer;
}

function extractHyperlinkFromCellP(p: unknown): string | null {
    if (!p || typeof p !== 'object') return null;
    const body = (p as { body?: { customRanges?: Array<{ rangeType?: number; properties?: { url?: unknown } }> } }).body;
    const ranges = body?.customRanges;
    if (!Array.isArray(ranges)) return null;
    for (const r of ranges) {
        if (r?.rangeType !== 0) continue;
        const url = r?.properties?.url;
        if (typeof url === 'string' && url) return url;
    }
    return null;
}

function buildSheetJSStyleFromUniver(style: Record<string, unknown>): Record<string, unknown> | null {
    const out: Record<string, unknown> = {};
    const font: Record<string, unknown> = {};
    if (typeof style.ff === 'string') font.name = style.ff;
    if (typeof style.fs === 'number') font.sz = style.fs;
    if (style.bl === 1) font.bold = true;
    if (style.it === 1) font.italic = true;
    if ((style.ul as { s?: number } | undefined)?.s === 1) font.underline = true;
    if ((style.st as { s?: number } | undefined)?.s === 1) font.strike = true;
    const cl = style.cl as { rgb?: string } | undefined;
    if (cl?.rgb) font.color = { rgb: cl.rgb.replace(/^#/, '') };
    if (Object.keys(font).length > 0) out.font = font;

    const bg = style.bg as { rgb?: string } | undefined;
    if (bg?.rgb) {
        out.fill = { patternType: 'solid', fgColor: { rgb: bg.rgb.replace(/^#/, '') } };
    }

    const alignment: Record<string, unknown> = {};
    if (style.ht === 1) alignment.horizontal = 'left';
    else if (style.ht === 2) alignment.horizontal = 'center';
    else if (style.ht === 3) alignment.horizontal = 'right';
    if (style.vt === 1) alignment.vertical = 'top';
    else if (style.vt === 2) alignment.vertical = 'center';
    else if (style.vt === 3) alignment.vertical = 'bottom';
    if (style.tb === WRAP_STRATEGY_WRAP) alignment.wrapText = true;
    const tr = style.tr as { a?: number; v?: number } | undefined;
    if (tr) {
        if (tr.v === 1) alignment.textRotation = 255;
        else if (typeof tr.a === 'number') alignment.textRotation = tr.a;
    }
    if (Object.keys(alignment).length > 0) out.alignment = alignment;

    const numFmt = (style.n as { pattern?: string } | undefined)?.pattern;
    if (numFmt) out.numFmt = numFmt;

    const bd = style.bd as Record<string, { s: number; cl?: { rgb?: string } }> | undefined;
    if (bd) {
        const border: Record<string, unknown> = {};
        const REVERSE_BORDER = Object.fromEntries(
            Object.entries(BORDER_STYLE_TO_UNIVER).map(([k, v]) => [v, k]),
        );
        const SIDES: Array<['t' | 'r' | 'b' | 'l', string]> = [
            ['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left'],
        ];
        for (const [uk, sjs] of SIDES) {
            const side = bd[uk];
            if (!side || typeof side.s !== 'number') continue;
            const styleName = REVERSE_BORDER[side.s];
            if (!styleName) continue;
            const colorRgb = side.cl?.rgb?.replace(/^#/, '') ?? '000000';
            border[sjs] = { style: styleName, color: { rgb: colorRgb } };
        }
        if (Object.keys(border).length > 0) out.border = border;
    }

    return Object.keys(out).length > 0 ? out : null;
}
