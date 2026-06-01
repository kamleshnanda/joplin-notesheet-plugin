// xlsx ↔ Univer snapshot converters.
//
// exceljs is the underlying engine. Two pure functions are exposed:
//   xlsxBufferToSnapshot(buffer) → IWorkbookData-shaped snapshot
//   snapshotToXlsxBuffer(snapshot) → Promise<ArrayBuffer>
//
// Coverage: cell values (string/number/boolean), formulas, font (family/size/
// bold/italic/underline/color), fill (background), alignment (horizontal/
// vertical/wrap), number format, merged cells, borders (M9), and named tables
// (M9 — synthesized into the snapshot's resources field so Univer's formula
// engine resolves Table[[#This Row],[Col]] structured references natively).
// Charts, pivots, conditional formatting are out of scope.
//
// Universe of enums we care about (mirrored as numeric literals so we don't
// pull @univerjs/core into Jest's node-environment unit tests):
//   HorizontalAlign:  LEFT=1, CENTER=2, RIGHT=3
//   VerticalAlign:    TOP=1, MIDDLE=2, BOTTOM=3
//   WrapStrategy:     OVERFLOW=1, CLIP=2, WRAP=3
//   BooleanNumber:    FALSE=0, TRUE=1
//   CellValueType:    STRING=1, NUMBER=2, BOOLEAN=3

import ExcelJS from 'exceljs';
import JSZip from 'jszip';

import type { UniverSnapshot } from './snapshot';
import { injectChartsIntoZip } from './charts/xlsxChart';
import { EXCEL_TABLE_STYLE_BY_NAME, type ExcelTableStyle } from './charts/excelTableStyles';

const HORIZONTAL = { left: 1, center: 2, right: 3 } as const;
const VERTICAL = { top: 1, middle: 2, bottom: 3 } as const;
const WRAP_STRATEGY_WRAP = 3;
const WRAP_STRATEGY_CLIP = 2;
const VALUE_STRING = 1;
const VALUE_NUMBER = 2;
const VALUE_BOOLEAN = 3;

// Border style mapping. exceljs uses string names ('thin', 'medium', etc.);
// Univer's BorderStyleTypes is a numeric enum. We mirror the small subset
// of the enum we care about here so the converters don't need to import
// @univerjs/core just for the constants.
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
const BORDER_STYLE_TO_EXCELJS: Record<number, string> = Object.fromEntries(
    Object.entries(BORDER_STYLE_TO_UNIVER).map(([k, v]) => [v, k]),
);

// Univer's @univerjs/sheets-table plugin reads/writes its state through this
// resource entry name. Confirmed in @univerjs/sheets-table/lib/types/const.d.ts.
// IMPORTANT: the `data` field must be a JSON-stringified string, not an
// object — the plugin uses JSON.parse on load and silently drops malformed
// entries.
const TABLE_PLUGIN_NAME = 'SHEET_TABLE_PLUGIN';
// Schema-version marker stamped onto the table resource so a future Notesheet
// version can detect 0.23-shaped data and migrate. Lives inside the data
// blob as a sibling to the per-subUnit table maps. Sheet ids cannot start
// with `_`, so this key never collides.
const TABLE_RESOURCE_SCHEMA = '0.23';

// Notesheet-private resource that records which per-cell style fields the
// M12 table-style synthesizer added during import. The map shape is
//   { [sheetId]: { [`${row}:${col}`]: ['bg', 'cl', 'bd.t', ...] } }
// On export, applyStyleToCell consults this map and skips writing any
// listed fields, so Excel paints the table style cleanly without our
// synthesized decoration doubling on top.
// IResourceName must match `${'SHEET'|'DOC'}_${string}_PLUGIN` per Univer's
// type definition (services/resource-manager/type.d.ts). The resource
// manager's `getResources()` produces the snapshot's `resources` array by
// iterating registered hook objects; entries that don't come from a
// registered hook are stripped on save. So Notesheet registers hooks at
// editor boot for these names, with a per-unitId in-memory map serving as
// the storage. The names below are the same strings used as the hook's
// pluginName.
export const NOTESHEET_SYNTH_STYLES_RESOURCE = 'SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN';

// Captures the workbook's <a:clrScheme> on import (raw XML fragment) so
// the export can splice it back into theme1.xml. Without this, exceljs
// emits its own Office-2007 default palette, and the same TableStyle name
// (e.g. TableStyleMedium4 = accent3) renders against a different RGB —
// which is why a round-tripped file looks "darker green" than the source
// even though both declare the same style.
export const NOTESHEET_THEME_CLR_SCHEME_RESOURCE = 'SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN';

interface CellRecord {
    v?: string | number | boolean;
    f?: string;
    t?: number;
    s?: string;
    // Univer rich-text body. Carries hyperlinks via customRanges with
    // rangeType=CustomRangeType.HYPERLINK (=0). Constructed on import for
    // cells that had {text,hyperlink} in the source xlsx.
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

// Mirrors @univerjs/sheets-table's ITableJson. Kept loose (Record<string, unknown>
// for filters / meta / style) because exact shapes carry many optional fields
// we don't need; Univer's fromJSON copies these verbatim and tolerates empties.
interface TableColumnJson {
    id: string;
    displayName: string;
    dataType: string; // 'string' | 'number' | 'date' | 'bool' | 'checkbox' | 'list' | 'none'
    formula: string;
    meta: Record<string, unknown>;
    style: Record<string, unknown>;
}
// Notesheet-specific meta keys we stash on each table so the export side can
// reconstruct the original Excel table-style (name + stripe flag) byte-for-byte.
// Univer's sheets-table plugin treats `meta` as opaque and round-trips it
// through the snapshot, so this survives reload cycles.
interface NotesheetTableMeta {
    notesheetExcelStyleName?: string;
    notesheetShowRowStripes?: boolean;
}
interface TableJson {
    id: string;
    name: string;
    range: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    options: { showHeader?: boolean; showFooter?: boolean; tableStyleId?: string };
    filters: { tableColumnFilterList?: unknown[] };
    columns: TableColumnJson[];
    meta: NotesheetTableMeta;
}

// argb is "AARRGGBB" (ExcelJS); Univer wants "#RRGGBB". Drop alpha; if all zero,
// treat as no color so we don't collapse "no fill" into "#000000".
function argbToHex(argb: string | undefined): string | undefined {
    if (!argb || typeof argb !== 'string') return undefined;
    const trimmed = argb.replace(/^#/, '');
    if (trimmed.length === 8) {
        const rgb = trimmed.slice(2);
        if (/^0+$/.test(rgb)) return undefined;
        return '#' + rgb.toUpperCase();
    }
    if (trimmed.length === 6) return '#' + trimmed.toUpperCase();
    return undefined;
}

function hexToArgb(hex: string | undefined): string | undefined {
    if (!hex) return undefined;
    const trimmed = hex.replace(/^#/, '');
    if (trimmed.length === 6) return ('FF' + trimmed).toUpperCase();
    if (trimmed.length === 8) return trimmed.toUpperCase();
    return undefined;
}

// Style canonicalization. Two cells with the same visual style should share a
// style id so the snapshot is compact and round-trips deterministically.
//
// The second arg of JSON.stringify acts as a recursive *whitelist* of keys —
// any key not in the list is silently dropped at every level of nesting.
// Earlier this used `Object.keys(style).sort()` as that whitelist, which
// silently truncated nested objects (e.g. cells with different border sides
// `bd: { t, l }` vs `bd: { t, r }` both collapsed to `bd: {}`, so the
// interner gave them the same id).
//
// `sortedJsonStringify` walks objects recursively, sorting each level's
// keys, and produces a deterministic key without dropping any data.
function sortedJsonStringify(value: unknown): string {
    if (value === null || typeof value !== 'object') return JSON.stringify(value);
    if (Array.isArray(value)) {
        return '[' + value.map(sortedJsonStringify).join(',') + ']';
    }
    const obj = value as Record<string, unknown>;
    const keys = Object.keys(obj).sort();
    return '{' + keys.map((k) => JSON.stringify(k) + ':' + sortedJsonStringify(obj[k])).join(',') + '}';
}

function styleKey(style: Record<string, unknown>): string {
    return sortedJsonStringify(style);
}

function buildStyleFromExcelCell(cell: ExcelJS.Cell, themePalette: ThemePalette | null = null): Record<string, unknown> | null {
    const style: Record<string, unknown> = {};

    const font = cell.font;
    if (font) {
        if (font.name) style.ff = font.name;
        if (typeof font.size === 'number') style.fs = font.size;
        if (font.bold) style.bl = 1;
        if (font.italic) style.it = 1;
        if (font.underline) style.ul = { s: 1 };
        if (font.strike) style.st = { s: 1 };
        const fontColor = resolveExceljsColor(font.color as ExceljsColor, themePalette);
        if (fontColor) style.cl = { rgb: fontColor };
    }

    const fill = cell.fill;
    if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
        const bgColor = resolveExceljsColor(fill.fgColor as ExceljsColor, themePalette);
        if (bgColor) style.bg = { rgb: bgColor };
    }

    const align = cell.alignment;
    if (align) {
        if (align.horizontal && align.horizontal in HORIZONTAL) {
            style.ht = HORIZONTAL[align.horizontal as keyof typeof HORIZONTAL];
        }
        if (align.vertical && align.vertical in VERTICAL) {
            style.vt = VERTICAL[align.vertical as keyof typeof VERTICAL];
        }
        if (align.wrapText) style.tb = WRAP_STRATEGY_WRAP;
    }

    if (cell.numFmt && cell.numFmt !== 'General') {
        style.n = { pattern: cell.numFmt };
    }

    const border = cell.border;
    if (border) {
        const bd: Record<string, { s: number; cl: { rgb: string } }> = {};
        const SIDES: Array<['t' | 'r' | 'b' | 'l', keyof ExcelJS.Borders]> = [
            ['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left'],
        ];
        for (const [univerKey, exceljsKey] of SIDES) {
            const side = border[exceljsKey] as Partial<ExcelJS.Border> | undefined;
            if (!side?.style) continue;
            const styleNum = BORDER_STYLE_TO_UNIVER[side.style];
            if (styleNum === undefined) continue;
            const resolved = resolveExceljsColor(side.color as ExceljsColor, themePalette);
            const hasColorRef = side.color && typeof side.color === 'object' && Object.keys(side.color).length > 0;
            // If color is fully absent, default to Excel's "automatic" =
            // black border. If color IS specified but theme-based and we
            // can't resolve it (e.g. theme=4 tint=0.4 with no theme
            // palette), drop the border rather than render bold black —
            // that looks like a stray underline against a banded row,
            // while "no border" is closer to what Excel paints when the
            // tint resolves to a near-bg hue.
            const rgb = resolved ?? (hasColorRef ? null : '#000000');
            if (!rgb) continue;
            bd[univerKey] = { s: styleNum, cl: { rgb } };
        }
        if (Object.keys(bd).length > 0) style.bd = bd;
    }

    return Object.keys(style).length > 0 ? style : null;
}

// Excel's date epoch is 1899-12-30 UTC (with the 1900 leap-year bug
// preserved by skipping that day in the serial sequence). For dates after
// 1900-03-01 — which is essentially every practical date — adding the
// day count to that epoch and dividing by ms-per-day gives the serial.
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;
function dateToExcelSerial(d: Date): number {
    return (d.getTime() - EXCEL_EPOCH_MS) / MS_PER_DAY;
}

// Build a Univer cell.p (IDocumentData) carrying a hyperlink. The text
// becomes body.dataStream; the URL lives on a CustomRange of type
// HYPERLINK (=0). On snapshot load, sheets-hyper-link's controllers walk
// the cell matrix, find these ranges, and register them with
// RefRangeService — making the link clickable in Univer's editor without
// any further wiring on our side.
//
// Shape mirrors `createDocumentModelWithStyle` in @univerjs/engine-render
// (the helper Univer's runtime hyperlink controller calls) so the
// documentSkeleton lays out a real page when sheets-ui's
// `_calcActiveCell` → `calcPadding` reads `skeletonData.pages[0].height`
// during mouse hover. Required pieces:
//
//   - dataStream ends with "\r\n" (DEFAULT_EMPTY_DOCUMENT_VALUE in core)
//   - a paragraph mark + sectionBreak at the right offsets
//   - a finite documentStyle.pageSize that survives JSON serialization
//
// The runtime version uses `pageSize: { width: Infinity, height: Infinity }`
// but that path lives in memory only. JSON.stringify converts Infinity to
// `null`, after which the documentSkeleton lays out zero pages and any
// mouse-move over a cell with our `p` throws "Cannot read properties of
// undefined (reading 'height')". 1e9 is well below MAX_SAFE_INTEGER and
// far larger than any practical cell, so it acts as effectively "no wrap"
// while round-tripping cleanly.
const HYPERLINK_PAGE_SIZE = 1_000_000_000;
function buildHyperlinkCellP(text: string, url: string): Record<string, unknown> {
    const len = text.length;
    return {
        id: '__INTERNAL_EDITOR__DOCS_NORMAL',
        body: {
            dataStream: text + '\r\n',
            textRuns: [{ st: 0, ed: len, ts: {} }],
            paragraphs: [{ startIndex: len, paragraphStyle: {} }],
            sectionBreaks: [{ startIndex: len + 1 }],
            customRanges: [{
                startIndex: 0,
                endIndex: Math.max(0, len - 1),
                rangeId: 'lnk-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8),
                rangeType: 0, // CustomRangeType.HYPERLINK
                properties: { url },
            }],
        },
        documentStyle: {
            pageSize: {
                width: HYPERLINK_PAGE_SIZE,
                height: HYPERLINK_PAGE_SIZE,
            },
        },
    };
}

function extractCellValue(cell: ExcelJS.Cell): { v?: string | number | boolean; f?: string; t?: number; p?: Record<string, unknown> } {
    const raw = cell.value;
    if (raw === null || raw === undefined) return {};

    // Formula cell: { formula: '...', result: ... }
    if (typeof raw === 'object' && 'formula' in raw && raw.formula) {
        const result = (raw as { result?: unknown }).result;
        const out: { f: string; v?: string | number | boolean; t?: number } = { f: '=' + (raw.formula as string) };
        if (typeof result === 'number') {
            out.v = result;
            out.t = VALUE_NUMBER;
        } else if (typeof result === 'string') {
            out.v = result;
            out.t = VALUE_STRING;
        } else if (typeof result === 'boolean') {
            out.v = result;
            out.t = VALUE_BOOLEAN;
        }
        return out;
    }

    // Shared formula: only the master holds .formula; followers carry sharedFormula.
    if (typeof raw === 'object' && 'sharedFormula' in raw && raw.sharedFormula) {
        const shared = raw as ExcelJS.CellSharedFormulaValue;
        const out: { f: string; v?: string | number | boolean; t?: number } = { f: '=' + (shared.formula || shared.sharedFormula) };
        const result = shared.result;
        if (typeof result === 'number') { out.v = result; out.t = VALUE_NUMBER; }
        else if (typeof result === 'string') { out.v = result; out.t = VALUE_STRING; }
        else if (typeof result === 'boolean') { out.v = result; out.t = VALUE_BOOLEAN; }
        return out;
    }

    // Hyperlink cell: keep both the visible text and the URL. Univer stores
    // hyperlinks inside cell.p.body.customRanges (CustomRangeType.HYPERLINK).
    // The cell value (v) holds the visible text so existing consumers that
    // only look at v keep working; the link sits beside it on p.
    if (typeof raw === 'object' && 'text' in raw && 'hyperlink' in raw) {
        const text = String((raw as { text: unknown }).text ?? '');
        const url = String((raw as { hyperlink: unknown }).hyperlink ?? '');
        const p = url ? buildHyperlinkCellP(text, url) : undefined;
        return { v: text, t: VALUE_STRING, ...(p ? { p } : {}) };
    }

    // Rich text: flatten to plain text.
    if (typeof raw === 'object' && 'richText' in raw && Array.isArray((raw as { richText: unknown }).richText)) {
        const segments = (raw as { richText: Array<{ text?: string }> }).richText;
        return { v: segments.map((s) => s.text ?? '').join(''), t: VALUE_STRING };
    }

    // Date — convert to Excel's serial-number representation (days since
    // 1899-12-30 UTC, with fractional days for time). Storing as a number
    // lets the cell's numFmt pattern (e.g. "m/d/yy") render the date the
    // way Excel did. Storing as an ISO string would make the formatter
    // produce literal "2025-12-02T00:00:00.000Z" because numFmt only
    // applies to numeric cells.
    if (raw instanceof Date) {
        return { v: dateToExcelSerial(raw), t: VALUE_NUMBER };
    }

    if (typeof raw === 'number') return { v: raw, t: VALUE_NUMBER };
    if (typeof raw === 'boolean') return { v: raw, t: VALUE_BOOLEAN };
    if (typeof raw === 'string') return { v: raw, t: VALUE_STRING };

    // Error cell or other complex types: render the displayed text.
    const text = cell.text;
    if (text) return { v: text, t: VALUE_STRING };
    return {};
}

// Sniff a column's data type from the values in the source xlsx. We only
// need to be in the right ballpark — Univer uses dataType for filter UX, not
// formula resolution. Falls back to 'string' on mixed/empty.
function inferColumnDataType(ws: ExcelJS.Worksheet, colNumber: number, startRow: number, endRow: number): string {
    let nNum = 0, nDate = 0, nBool = 0, nString = 0, total = 0;
    for (let r = startRow; r <= endRow; r++) {
        const v = ws.getCell(r, colNumber).value;
        if (v === null || v === undefined || v === '') continue;
        total++;
        if (v instanceof Date) nDate++;
        else if (typeof v === 'number') nNum++;
        else if (typeof v === 'boolean') nBool++;
        else nString++;
    }
    if (total === 0) return 'string';
    if (nDate / total > 0.5) return 'date';
    if (nNum / total > 0.5) return 'number';
    if (nBool / total > 0.5) return 'bool';
    return 'string';
}

// Parse "B2:E15" or "A1" into 0-based row/col bounds. Returns null on
// malformed input. Single-cell refs ("A1") yield a 1x1 range.
function parseA1Range(ref: string): { startRow: number; endRow: number; startColumn: number; endColumn: number } | null {
    const cellRe = /^([A-Z]+)(\d+)$/;
    const parts = ref.split(':');
    if (parts.length === 0 || parts.length > 2) return null;
    const m1 = cellRe.exec(parts[0]);
    if (!m1) return null;
    const m2 = parts.length === 2 ? cellRe.exec(parts[1]) : m1;
    if (!m2) return null;
    const colToIdx = (s: string): number => {
        let n = 0;
        for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
        return n - 1;
    };
    return {
        startRow: parseInt(m1[2], 10) - 1,
        endRow: parseInt(m2[2], 10) - 1,
        startColumn: colToIdx(m1[1]),
        endColumn: colToIdx(m2[1]),
    };
}

// Raw table parsed directly from xl/tables/*.xml — bypasses exceljs's
// table parser, which has two read bugs in the user's real-world files:
// (1) defaults missing headerRowCount to false (OOXML spec says 1), and
// (2) drops every column after one whose <tableColumn> has nested
// children like <calculatedColumnFormula>. Both problems break round-trip.
interface RawTable {
    name: string;
    ref: string;
    headerRowCount: number;
    totalsRowCount: number;
    columns: string[];
    styleName?: string;
    showRowStripes?: boolean;
}

// Map from 1-based sheet index (matching xl/worksheets/sheetN.xml) to the
// raw tables that sheet references via its _rels file.
type RawTableMap = Map<number, RawTable[]>;

// Strip XML attribute and use the unescaped value (handles &amp; &quot; &#xx; etc).
function decodeXmlAttr(raw: string): string {
    return raw
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
        .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
        .replace(/&amp;/g, '&'); // last to avoid double-decode
}

function getAttr(tag: string, attrName: string): string | null {
    const re = new RegExp(`\\b${attrName}\\s*=\\s*"([^"]*)"`);
    const m = re.exec(tag);
    return m ? decodeXmlAttr(m[1]) : null;
}

function parseTableXml(xml: string): RawTable | null {
    // The <table ...> opening tag carries name, ref, headerRowCount, etc.
    const tableTagMatch = /<table\b[^>]*>/.exec(xml);
    if (!tableTagMatch) return null;
    const tableTag = tableTagMatch[0];
    const name = getAttr(tableTag, 'name');
    const ref = getAttr(tableTag, 'ref');
    if (!name || !ref) return null;

    // OOXML spec: when headerRowCount is omitted, the value is 1.
    // exceljs's read defaults to 0 (the bug we're working around).
    const headerRowCountStr = getAttr(tableTag, 'headerRowCount');
    const headerRowCount = headerRowCountStr === null ? 1 : parseInt(headerRowCountStr, 10);
    const totalsRowCountStr = getAttr(tableTag, 'totalsRowCount');
    const totalsRowCount = totalsRowCountStr === null ? 0 : parseInt(totalsRowCountStr, 10);

    // Pull every <tableColumn ... name="..." ... /> or <tableColumn ...>...</tableColumn>.
    // The opening tag is what carries the name attribute; nested children
    // (<calculatedColumnFormula>, <xmlColumnPr>, etc.) don't matter to us.
    const columns: string[] = [];
    const colRe = /<tableColumn\b[^>]*?(?:\/>|>)/g;
    let m: RegExpExecArray | null;
    while ((m = colRe.exec(xml)) !== null) {
        const cname = getAttr(m[0], 'name');
        if (cname) columns.push(cname);
    }

    const styleTagMatch = /<tableStyleInfo\b[^>]*\/?>/.exec(xml);
    const styleName = styleTagMatch ? getAttr(styleTagMatch[0], 'name') ?? undefined : undefined;
    const stripesAttr = styleTagMatch ? getAttr(styleTagMatch[0], 'showRowStripes') : null;
    // Per OOXML default, showRowStripes is false when omitted. Mirror what
    // the source xlsx asked for so we don't add stripes the user didn't have.
    const showRowStripes = stripesAttr === '1';

    return { name, ref, headerRowCount, totalsRowCount, columns, styleName, showRowStripes };
}

// Walk the zipped xlsx and return tables grouped by 1-based sheet index.
// We map tables to sheets via xl/worksheets/_rels/sheet<N>.xml.rels which
// contains <Relationship Type=".../table" Target="../tables/tableM.xml"/>.
// Read xl/theme/theme1.xml's font scheme. Excel stores the workbook's
// default font here (under <a:fontScheme><a:minorFont><a:latin typeface="..."/>)
// rather than on individual cells, so cells with no explicit font.name
// inherit from this. exceljs doesn't expose the theme XML, hence direct
// zip access. Returns null if the file or attribute is absent.
async function readThemeFont(buffer: ArrayBuffer | Uint8Array | Buffer): Promise<{ minor?: string; major?: string } | null> {
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return null;
    }
    // Theme path is conventionally xl/theme/theme1.xml but tools occasionally
    // use other numbers; pick the first matching file.
    const themePath = Object.keys(zip.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p));
    if (!themePath) return null;
    const xml = await zip.files[themePath].async('string');

    const minor = /<a:minorFont>[^<]*<a:latin\b[^>]*\btypeface="([^"]+)"/.exec(xml)?.[1];
    const major = /<a:majorFont>[^<]*<a:latin\b[^>]*\btypeface="([^"]+)"/.exec(xml)?.[1];
    if (!minor && !major) return null;
    return { minor, major };
}

// Read xl/theme/theme1.xml's <a:clrScheme>. Excel resolves named colors
// like "accent1" through this scheme — and the same TableStyle (e.g.
// TableStyleMedium4) renders different hues depending on which theme is
// active. exceljs ships its own Office-2007 default theme, so an export
// that doesn't preserve the source theme's palette will look noticeably
// different even if the table style name is identical. We capture the raw
// <a:clrScheme>...</a:clrScheme> as a string here and splice it back into
// the exported theme1.xml on save. Returns the captured XML and a
// 12-entry RGB palette indexed by Excel theme color id (0..11) for
// resolving cell-level `{theme: N, tint: T}` color references.
interface ThemePalette {
    raw: string;
    rgb: string[]; // index 0..11 → '#RRGGBB'
}
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
    // Per the OOXML spec, the cell <color theme="N"/> index uses a permuted
    // mapping of the clrScheme elements: 0=lt1, 1=dk1, 2=lt2, 3=dk2,
    // 4=accent1..9=accent6, 10=hlink, 11=folHlink. Note that the cell
    // index swaps lt1↔dk1 and lt2↔dk2 relative to the scheme's element
    // order.
    const ELEMENT_ORDER: Array<'lt1' | 'dk1' | 'lt2' | 'dk2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'hlink' | 'folHlink'> = [
        'lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink',
    ];
    const rgb: string[] = [];
    for (const elName of ELEMENT_ORDER) {
        const elRe = new RegExp(`<a:${elName}\\b[^>]*>([\\s\\S]*?)</a:${elName}>`);
        const elMatch = elRe.exec(raw);
        if (!elMatch) { rgb.push('#000000'); continue; }
        const inner = elMatch[1];
        // Either <a:srgbClr val="RRGGBB"/> or <a:sysClr val="..." lastClr="RRGGBB"/>.
        const srgb = /<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(inner);
        const sys = /<a:sysClr\b[^>]*\blastClr="([0-9A-Fa-f]{6})"/.exec(inner);
        const hex = (srgb?.[1] ?? sys?.[1] ?? '000000').toUpperCase();
        rgb.push('#' + hex);
    }
    return { raw, rgb };
}

// OOXML tint applied in HSL luminance, per Microsoft's spec:
//   tint > 0  →  L = L*(1-tint) + tint     (lighten toward white)
//   tint < 0  →  L = L*(1+tint)            (darken toward black)
// Used to resolve cell-level `{theme: N, tint: T}` color references.
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

// Resolve an exceljs color descriptor. exceljs exposes either:
//   - { argb: 'AARRGGBB' } — explicit RGB; we strip alpha and return.
//   - { theme: N, tint?: T } — references the workbook's clrScheme; needs
//     the import-time palette to resolve.
// Returns null when no resolvable color is present (e.g. {tint: 0.4} alone
// or a missing palette). Callers should treat null as "no color set".
type ExceljsColor = { argb?: string; theme?: number; tint?: number } | null | undefined;
function resolveExceljsColor(color: ExceljsColor, palette: ThemePalette | null): string | undefined {
    if (!color || typeof color !== 'object') return undefined;
    const direct = argbToHex(color.argb);
    if (direct) return direct;
    if (typeof color.theme === 'number' && palette) {
        const base = palette.rgb[color.theme];
        if (!base) return undefined;
        return typeof color.tint === 'number' && color.tint !== 0 ? applyOoxmlTint(base, color.tint) : base;
    }
    return undefined;
}

async function readTablesFromXlsxZip(buffer: ArrayBuffer | Uint8Array | Buffer): Promise<RawTableMap> {
    const result: RawTableMap = new Map();
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return result;
    }

    const tableXmlByName = new Map<string, RawTable>();
    for (const path of Object.keys(zip.files)) {
        const m = /^xl\/tables\/(table\d+\.xml)$/i.exec(path);
        if (!m) continue;
        const xml = await zip.files[path].async('string');
        const parsed = parseTableXml(xml);
        if (parsed) tableXmlByName.set(m[1].toLowerCase(), parsed);
    }
    if (tableXmlByName.size === 0) return result;

    for (const path of Object.keys(zip.files)) {
        const m = /^xl\/worksheets\/_rels\/sheet(\d+)\.xml\.rels$/i.exec(path);
        if (!m) continue;
        const sheetIndex = parseInt(m[1], 10);
        const xml = await zip.files[path].async('string');
        // <Relationship Id="..." Type=".../table" Target="../tables/tableN.xml"/>
        const relRe = /<Relationship\b[^>]*Type="[^"]*\/table"[^>]*Target="[^"]*tables\/(table\d+\.xml)"[^>]*\/?>/gi;
        const tables: RawTable[] = [];
        let rm: RegExpExecArray | null;
        while ((rm = relRe.exec(xml)) !== null) {
            const t = tableXmlByName.get(rm[1].toLowerCase());
            if (t) tables.push(t);
        }
        if (tables.length > 0) result.set(sheetIndex, tables);
    }
    return result;
}

// Build ITableJson entries for one worksheet from its corresponding RawTable
// list. Falls back to exceljs's view if the raw map didn't pick up any
// tables (e.g. files we couldn't unzip — never happens in practice but
// keeps the import resilient).
// When an Excel table uses a built-in style like TableStyleMedium2, the
// styled colors don't live on individual cells — Excel synthesizes them at
// render time from the workbook theme + table style flags. exceljs gives us
// no fill/font records to copy, so cells looked unstyled and Univer
// rendered them with its own table-default-0 theme (light blue), which
// looks nothing like the source's dark teal.
//
// This function walks the table's range and bakes the lookup-table-derived
// header bg/fg + alternating row colors into the corresponding cellData
// entries. Per-cell style records carry the overlay; existing user-set cell
// styles take precedence (we only set fields that aren't already present).
//
// Returns a list of cellStyleAssignments to apply: { row, col, style }.
// The caller is responsible for interning each style and assigning the
// resulting style id to the cell record.
interface CellStyleAssignment {
    row: number;
    column: number;
    style: Record<string, unknown>;
    // List of dotted-path fields we added on top of any pre-existing user
    // style (e.g. `bg`, `cl`, `bl`, `bd.t`, `bd.b`, `bd.l`, `bd.r`). The
    // exporter uses this to subtract our synthesized decoration before
    // writing the cell, so the table style painted by Excel doesn't double
    // up with ours.
    addedFields: string[];
}

function synthesizeTableStyleAssignments(table: RawTable, existingCellStyles: Map<string, Record<string, unknown>>): CellStyleAssignment[] {
    if (!table.styleName) return [];
    const palette: ExcelTableStyle | undefined = EXCEL_TABLE_STYLE_BY_NAME[table.styleName];
    if (!palette) return [];

    const range = parseA1Range(table.ref);
    if (!range) return [];

    const headerRows = table.headerRowCount > 0 ? 1 : 0;
    const totalRows = table.totalsRowCount > 0 ? 1 : 0;
    const dataStartRow = range.startRow + headerRows;
    const dataEndRow = range.endRow - totalRows;

    const out: CellStyleAssignment[] = [];

    // Side keys mirror Univer's bd shape: top/right/bottom/left.
    type BorderSide = 't' | 'r' | 'b' | 'l';
    type BorderEntry = { s: number; cl: { rgb: string } };

    // Helper: build a cell style record additively. Only set fields not
    // already present on the existing cell style — explicit per-cell
    // formatting from the source workbook should win over the synthesized
    // table styling. Borders merge per side (we only add a side if the
    // existing bd doesn't already have that side set).
    const overlay = (
        row: number, col: number,
        addBg?: string, addFg?: string, bold?: boolean,
        addBorders?: Partial<Record<BorderSide, BorderEntry>>,
    ) => {
        const key = `${row}:${col}`;
        const existing = existingCellStyles.get(key) ?? {};
        const next: Record<string, unknown> = { ...existing };
        const addedFields: string[] = [];
        if (addBg && !('bg' in next)) {
            next.bg = { rgb: addBg };
            addedFields.push('bg');
        }
        if (addFg && !('cl' in next)) {
            next.cl = { rgb: addFg };
            addedFields.push('cl');
        }
        if (bold && next.bl !== 1) {
            next.bl = 1;
            addedFields.push('bl');
        }
        if (addBorders) {
            const existingBd = (existing.bd as Partial<Record<BorderSide, BorderEntry>> | undefined) ?? {};
            const mergedBd: Partial<Record<BorderSide, BorderEntry>> = { ...existingBd };
            let added = false;
            for (const side of Object.keys(addBorders) as BorderSide[]) {
                if (!mergedBd[side] && addBorders[side]) {
                    mergedBd[side] = addBorders[side]!;
                    addedFields.push('bd.' + side);
                    added = true;
                }
            }
            if (added) next.bd = mergedBd;
        }
        if (sortedJsonStringify(next) === sortedJsonStringify(existing)) return;
        out.push({ row, column: col, style: next, addedFields });
    };

    // TableStyleMedium2 (and most Medium themes) draw a thin border along
    // the table's outer edges + a thin border under the header in the same
    // accent color used for the header bg. Inner row separators are NOT
    // drawn when showRowStripes is on — the alternating fill is the
    // separator. Skip border synthesis if the catalog has no borderColor.
    const borderRgb = palette.borderColor;
    const thinBorder: BorderEntry | undefined = borderRgb ? { s: BORDER_STYLE_TO_UNIVER.thin, cl: { rgb: borderRgb } } : undefined;

    // Header row.
    if (headerRows === 1) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
            const headerBorders: Partial<Record<BorderSide, BorderEntry>> | undefined = thinBorder ? {
                t: thinBorder,
                b: thinBorder,
                ...(c === range.startColumn ? { l: thinBorder } : {}),
                ...(c === range.endColumn ? { r: thinBorder } : {}),
            } : undefined;
            overlay(range.startRow, c, palette.headerBg, palette.headerFg, /* bold */ true, headerBorders);
        }
    }

    // Data rows with alternating banding (only when showRowStripes is true).
    if (table.showRowStripes) {
        for (let r = dataStartRow; r <= dataEndRow; r++) {
            const isEven = (r - dataStartRow) % 2 === 0;
            const bg = isEven ? palette.bandedRowEvenBg : palette.bandedRowOddBg;
            // Skip white (#FFFFFF) — no-op for default white background.
            const useBg = bg && bg.toUpperCase() !== '#FFFFFF' ? bg : undefined;
            for (let c = range.startColumn; c <= range.endColumn; c++) {
                const rowBorders: Partial<Record<BorderSide, BorderEntry>> | undefined = thinBorder ? {
                    ...(c === range.startColumn ? { l: thinBorder } : {}),
                    ...(c === range.endColumn ? { r: thinBorder } : {}),
                    ...(r === dataEndRow && totalRows === 0 ? { b: thinBorder } : {}),
                } : undefined;
                if (useBg || rowBorders) overlay(r, c, useBg, undefined, false, rowBorders);
            }
        }
    } else if (thinBorder) {
        // Even without banding we still want the table's outer border.
        for (let r = dataStartRow; r <= dataEndRow; r++) {
            for (let c = range.startColumn; c <= range.endColumn; c++) {
                const rowBorders: Partial<Record<BorderSide, BorderEntry>> = {
                    ...(c === range.startColumn ? { l: thinBorder } : {}),
                    ...(c === range.endColumn ? { r: thinBorder } : {}),
                    ...(r === dataEndRow && totalRows === 0 ? { b: thinBorder } : {}),
                };
                if (Object.keys(rowBorders).length > 0) overlay(r, c, undefined, undefined, false, rowBorders);
            }
        }
    }

    // Totals row.
    if (totalRows === 1 && palette.totalsBg) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
            const totalsBorders: Partial<Record<BorderSide, BorderEntry>> | undefined = thinBorder ? {
                t: thinBorder,
                b: thinBorder,
                ...(c === range.startColumn ? { l: thinBorder } : {}),
                ...(c === range.endColumn ? { r: thinBorder } : {}),
            } : undefined;
            overlay(range.endRow, c, palette.totalsBg, palette.totalsFg, /* bold */ true, totalsBorders);
        }
    }

    return out;
}

function buildTableJsonForSheet(ws: ExcelJS.Worksheet, rawTables: RawTable[]): TableJson[] {
    const out: TableJson[] = [];
    for (const t of rawTables) {
        const range = parseA1Range(t.ref);
        if (!range) continue;
        if (t.columns.length === 0) continue;

        const headerRows = t.headerRowCount > 0 ? 1 : 0;
        const totalRows = t.totalsRowCount > 0 ? 1 : 0;
        const dataStartRow = range.startRow + headerRows;
        const dataEndRow = range.endRow - totalRows;

        const columns: TableColumnJson[] = t.columns.map((cname, idx) => ({
            id: `tblcol-${idx}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
            displayName: cname,
            dataType: inferColumnDataType(ws, range.startColumn + idx + 1, dataStartRow + 1, dataEndRow + 1),
            formula: '',
            meta: {},
            style: {},
        }));

        const meta: NotesheetTableMeta = {};
        if (t.styleName) meta.notesheetExcelStyleName = t.styleName;
        if (t.showRowStripes) meta.notesheetShowRowStripes = true;

        out.push({
            id: `tbl-${t.name}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
            name: t.name,
            range,
            options: {
                showHeader: headerRows === 1,
                showFooter: totalRows === 1,
            },
            filters: { tableColumnFilterList: [] },
            columns,
            meta,
        });
    }
    return out;
}

export async function xlsxBufferToSnapshot(buffer: ArrayBuffer | Uint8Array | Buffer): Promise<UniverSnapshot> {
    const wb = new ExcelJS.Workbook();
    // exceljs accepts Buffer in Node and ArrayBuffer in the browser; both are valid
    // at runtime but the .d.ts only types Buffer. Cast away to satisfy TS.
    await wb.xlsx.load(buffer as unknown as Parameters<typeof wb.xlsx.load>[0]);

    // Read tables directly from the xlsx zip — exceljs's table parser drops
    // columns that have nested children and mis-defaults headerRowCount.
    const rawTablesByIndex = await readTablesFromXlsxZip(buffer);

    // Read the workbook's theme font scheme. Excel cells without an explicit
    // font.name inherit from the theme's minorFont (Calibri, Aptos Narrow,
    // etc.). exceljs doesn't surface this, so cells looked font-less; Univer
    // fell back to its own default and the imported sheet showed Arial
    // regardless of source.
    const themeFont = await readThemeFont(buffer);
    // Captured for export-time replay so the workbook's table-style colors
    // resolve against the same accent palette they did originally.
    const themeClrScheme = await readThemeClrScheme(buffer);

    const sheetOrder: string[] = [];
    const sheets: Record<string, SheetRecord> = {};
    const styles: Record<string, Record<string, unknown>> = {};
    const styleIdByKey = new Map<string, string>();
    let nextStyleId = 1;
    // Per-subUnit table state, keyed by our generated sheetId. Filled during
    // the eachSheet walk and serialized into the SHEET_TABLE_PLUGIN resource
    // at the end so Univer's formula engine sees the tables on snapshot load.
    const tableResource: Record<string, { tables: TableJson[]; tableFilteredOutRows: number[] }> = {};

    // Per-cell record of which style fields the M12 table-style synthesizer
    // added on top of the source cell's own style. Keyed `${sheetId}` →
    // `${row}:${col}` → ['bg', 'cl', 'bl', 'bd.t', ...]. Persisted as a
    // snapshot resource so the exporter can subtract our synthesized
    // decoration before writing the cell — Excel re-paints those colors
    // from the table style at render time, and a doubled-up paint reads
    // visually heavier than the original.
    const synthStyleSidecar: Record<string, Record<string, string[]>> = {};

    function internStyle(style: Record<string, unknown> | null): string | undefined {
        if (!style) return undefined;
        const key = styleKey(style);
        const existing = styleIdByKey.get(key);
        if (existing) return existing;
        const id = 'style-' + nextStyleId++;
        styleIdByKey.set(key, id);
        styles[id] = style;
        return id;
    }

    wb.eachSheet((ws) => {
        const sheetId = 'sheet-' + ws.id;
        sheetOrder.push(sheetId);

        const cellData: Record<number, Record<number, CellRecord>> = {};
        let maxRow = 0;
        let maxCol = 0;

        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const r = rowNumber - 1; // exceljs is 1-based
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                const c = colNumber - 1;
                const value = extractCellValue(cell);
                const style = buildStyleFromExcelCell(cell, themeClrScheme);
                const styleId = internStyle(style);
                const record: CellRecord = {};
                if (value.v !== undefined) record.v = value.v;
                if (value.f) record.f = value.f;
                if (value.t !== undefined) record.t = value.t;
                if (value.p) record.p = value.p;
                if (styleId) record.s = styleId;
                if (Object.keys(record).length === 0) return;
                if (!cellData[r]) cellData[r] = {};
                cellData[r][c] = record;
                if (r > maxRow) maxRow = r;
                if (c > maxCol) maxCol = c;
            });
        });

        const mergeData: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }> = [];
        // exceljs exposes _merges as an internal map of "tlAddr:brAddr".
        const merges = (ws as unknown as { _merges?: Record<string, { model: { top: number; left: number; bottom: number; right: number } }> })._merges;
        if (merges) {
            for (const key of Object.keys(merges)) {
                const m = merges[key]?.model;
                if (!m) continue;
                mergeData.push({
                    startRow: m.top - 1,
                    endRow: m.bottom - 1,
                    startColumn: m.left - 1,
                    endColumn: m.right - 1,
                });
            }
        }

        const rowData: Record<number, { h?: number }> = {};
        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (typeof row.height === 'number' && row.height > 0) {
                rowData[rowNumber - 1] = { h: Math.round(row.height * (96 / 72)) };
            }
        });

        const columnData: Record<number, { w?: number }> = {};
        if (Array.isArray(ws.columns)) {
            ws.columns.forEach((col, idx) => {
                if (col && typeof col.width === 'number' && col.width > 0) {
                    // ExcelJS width is "characters of the standard font". A common
                    // approximation is px ≈ width * 7 + 5. Round to integer.
                    columnData[idx] = { w: Math.round(col.width * 7 + 5) };
                }
            });
        }

        // ws.id is the 1-based worksheet index in xl/worksheets/sheetN.xml,
        // which is exactly what readTablesFromXlsxZip keys by.
        const rawTables = rawTablesByIndex.get(ws.id) ?? [];

        // M12: bake table-style colors (TableStyleMedium2 etc.) into per-cell
        // styles so the imported sheet looks like Excel even though Univer's
        // table preset doesn't honor the original style name. The catalog
        // ships every Office-2016+ built-in style. Done BEFORE sheets[sheetId]
        // is sealed so rowCount/columnCount include the synthesized cells.
        for (const t of rawTables) {
            // Build a quick lookup of the existing styles keyed by row:col so
            // the overlay can preserve user-set formatting.
            const existingCellStyles = new Map<string, Record<string, unknown>>();
            for (const rKey of Object.keys(cellData)) {
                const r = Number(rKey);
                const row = cellData[r];
                if (!row) continue;
                for (const cKey of Object.keys(row)) {
                    const c = Number(cKey);
                    const data = row[c];
                    if (!data?.s) continue;
                    const style = styles[data.s];
                    if (style) existingCellStyles.set(`${r}:${c}`, style);
                }
            }
            const assignments = synthesizeTableStyleAssignments(t, existingCellStyles);
            for (const a of assignments) {
                const styleId = internStyle(a.style);
                if (!styleId) continue;
                if (!cellData[a.row]) cellData[a.row] = {};
                if (!cellData[a.row][a.column]) cellData[a.row][a.column] = {};
                cellData[a.row][a.column].s = styleId;
                if (a.row > maxRow) maxRow = a.row;
                if (a.column > maxCol) maxCol = a.column;
                if (a.addedFields.length > 0) {
                    if (!synthStyleSidecar[sheetId]) synthStyleSidecar[sheetId] = {};
                    synthStyleSidecar[sheetId][`${a.row}:${a.column}`] = a.addedFields;
                }
            }
        }

        sheets[sheetId] = {
            id: sheetId,
            name: ws.name,
            cellData,
            rowCount: Math.max(100, maxRow + 1),
            columnCount: Math.max(26, maxCol + 1),
            defaultColumnWidth: 73,
            defaultRowHeight: 19,
            mergeData,
            rowData,
            columnData,
        };

        const tablesForSheet = buildTableJsonForSheet(ws, rawTables);
        if (tablesForSheet.length > 0) {
            tableResource[sheetId] = {
                tables: tablesForSheet,
                tableFilteredOutRows: [],
            };
        }
    });

    if (sheetOrder.length === 0) {
        // Empty workbook fallback — give Univer one empty sheet.
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
    if (Object.keys(tableResource).length > 0) {
        // The plugin requires `data` to be a JSON-stringified string. Univer
        // iterates Object.keys(data) treating each as a subUnitId, so DO NOT
        // mix schema-version markers in here — keep this map subUnit-only.
        resources.push({ name: TABLE_PLUGIN_NAME, data: JSON.stringify(tableResource) });
        // Sibling resource scoped to Notesheet, used to detect 0.23-shaped
        // ITableJson at a later upgrade so migration code can run.
        resources.push({
            name: 'NOTESHEET_TABLE_SCHEMA',
            data: JSON.stringify({ version: TABLE_RESOURCE_SCHEMA }),
        });
    }
    if (Object.keys(synthStyleSidecar).length > 0) {
        resources.push({
            name: NOTESHEET_SYNTH_STYLES_RESOURCE,
            data: JSON.stringify(synthStyleSidecar),
        });
    }
    if (themeClrScheme) {
        resources.push({
            name: NOTESHEET_THEME_CLR_SCHEME_RESOURCE,
            data: themeClrScheme.raw,
        });
    }

    // Workbook-level default style. Univer cells without an explicit `s`
    // inherit from this. We use it to carry the source theme's body font
    // (minorFont) so imported sheets render in Aptos Narrow / Calibri / etc.
    // instead of Univer's hardcoded fallback.
    const defaultStyle = themeFont?.minor ? { ff: themeFont.minor } : undefined;

    return {
        id: 'workbook-' + Date.now(),
        sheetOrder,
        name: 'Spreadsheet',
        appVersion: '0.1.0',
        locale: 'enUS',
        styles,
        sheets,
        ...(defaultStyle ? { defaultStyle } : {}),
        ...(resources.length > 0 ? { resources } : {}),
    } as unknown as UniverSnapshot;
}

function resolveStyle(snapshot: UniverSnapshot, ref: unknown): Record<string, unknown> | null {
    if (!ref) return null;
    if (typeof ref === 'string') {
        const styles = (snapshot as { styles?: Record<string, Record<string, unknown>> }).styles;
        return styles?.[ref] ?? null;
    }
    if (typeof ref === 'object') return ref as Record<string, unknown>;
    return null;
}

// Extract a hyperlink URL from a Univer cell.p (IDocumentData) by finding
// the first customRange of rangeType=0 (HYPERLINK). Returns null when the
// cell has no link. Used on export to feed exceljs's { text, hyperlink }
// cell-value format, which it serializes into <hyperlinks> + rels.
function extractHyperlinkFromCellP(p: unknown): string | null {
    if (!p || typeof p !== 'object') return null;
    const body = (p as { body?: { customRanges?: Array<{ rangeType?: number; properties?: { url?: unknown } }> } }).body;
    const ranges = body?.customRanges;
    if (!Array.isArray(ranges)) return null;
    for (const r of ranges) {
        if (r?.rangeType !== 0) continue; // HYPERLINK
        const url = r?.properties?.url;
        if (typeof url === 'string' && url) return url;
    }
    return null;
}

function applyStyleToCell(
    cell: ExcelJS.Cell,
    style: Record<string, unknown>,
    skipFields?: ReadonlySet<string>,
): void {
    const skip = skipFields ?? EMPTY_SKIP_SET;

    const font: Partial<ExcelJS.Font> = {};
    if (typeof style.ff === 'string') font.name = style.ff;
    if (typeof style.fs === 'number') font.size = style.fs;
    if (style.bl === 1 && !skip.has('bl')) font.bold = true;
    if (style.it === 1) font.italic = true;
    if (style.ul && (style.ul as { s?: number }).s === 1) font.underline = true;
    if (style.st && (style.st as { s?: number }).s === 1) font.strike = true;
    const cl = style.cl as { rgb?: string } | undefined;
    if (cl?.rgb && !skip.has('cl')) {
        const argb = hexToArgb(cl.rgb);
        if (argb) font.color = { argb };
    }
    if (Object.keys(font).length > 0) cell.font = font;

    const bg = style.bg as { rgb?: string } | undefined;
    if (bg?.rgb && !skip.has('bg')) {
        const argb = hexToArgb(bg.rgb);
        if (argb) {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb },
            };
        }
    }

    const align: Partial<ExcelJS.Alignment> = {};
    if (style.ht === 1) align.horizontal = 'left';
    else if (style.ht === 2) align.horizontal = 'center';
    else if (style.ht === 3) align.horizontal = 'right';
    if (style.vt === 1) align.vertical = 'top';
    else if (style.vt === 2) align.vertical = 'middle';
    else if (style.vt === 3) align.vertical = 'bottom';
    if (style.tb === WRAP_STRATEGY_WRAP) align.wrapText = true;
    else if (style.tb === WRAP_STRATEGY_CLIP) align.wrapText = false;
    if (Object.keys(align).length > 0) cell.alignment = align;

    const numFmt = (style.n as { pattern?: string } | undefined)?.pattern;
    if (numFmt) cell.numFmt = numFmt;

    const bd = style.bd as Record<string, { s: number; cl?: { rgb?: string } }> | undefined;
    if (bd && typeof bd === 'object') {
        const out: Partial<ExcelJS.Borders> = {};
        const SIDES: Array<['t' | 'r' | 'b' | 'l', keyof ExcelJS.Borders]> = [
            ['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left'],
        ];
        for (const [univerKey, exceljsKey] of SIDES) {
            if (skip.has('bd.' + univerKey)) continue;
            const side = bd[univerKey];
            if (!side || typeof side.s !== 'number') continue;
            const styleName = BORDER_STYLE_TO_EXCELJS[side.s];
            if (!styleName) continue;
            const border: Partial<ExcelJS.Border> = { style: styleName as ExcelJS.BorderStyle };
            const argb = hexToArgb(side.cl?.rgb);
            if (argb) border.color = { argb };
            (out as Record<string, Partial<ExcelJS.Border>>)[exceljsKey] = border;
        }
        if (Object.keys(out).length > 0) cell.border = out as Partial<ExcelJS.Borders>;
    }
}

const EMPTY_SKIP_SET: ReadonlySet<string> = new Set();

// 0-based column index → A1 column letters (A, B, ..., Z, AA, AB, ...).
function colLetters(idx: number): string {
    let n = idx + 1;
    let s = '';
    while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

// Read the SHEET_TABLE_PLUGIN resource and return a per-subUnitId table map.
// Returns {} if the resource is absent or unparseable; we never throw on
// malformed resources because the rest of the export should still produce a
// valid xlsx (just without table definitions).
function readTableResource(snapshot: UniverSnapshot): Record<string, { tables: TableJson[] }> {
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> }).resources;
    if (!Array.isArray(resources)) return {};
    const entry = resources.find((r) => r?.name === TABLE_PLUGIN_NAME);
    if (!entry || typeof entry.data !== 'string') return {};
    try {
        const parsed = JSON.parse(entry.data);
        if (!parsed || typeof parsed !== 'object') return {};
        return parsed as Record<string, { tables: TableJson[] }>;
    } catch {
        return {};
    }
}

// Read the NOTESHEET_SYNTH_STYLES sidecar that records which per-cell style
// fields the M12 table-style synthesizer added during import. Returns a
// `${sheetId}` → `${row}:${col}` → string[] map; the exporter consults this
// to skip those fields so Excel's TableStyle paint isn't doubled up by ours.
function readSynthStylesSidecar(snapshot: UniverSnapshot): Record<string, Record<string, string[]>> {
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> }).resources;
    if (!Array.isArray(resources)) return {};
    const entry = resources.find((r) => r?.name === NOTESHEET_SYNTH_STYLES_RESOURCE);
    if (!entry || typeof entry.data !== 'string') return {};
    try {
        const parsed = JSON.parse(entry.data);
        if (!parsed || typeof parsed !== 'object') return {};
        return parsed as Record<string, Record<string, string[]>>;
    } catch {
        return {};
    }
}

// Pull the workbook-level default font name out of the snapshot. Set on
// import from the source xlsx's theme1.xml/minorFont; on export every cell
// that doesn't already carry an explicit font.name inherits this so the
// round-trip preserves Aptos Narrow / Calibri / etc.
function readDefaultFontName(snapshot: UniverSnapshot): string | undefined {
    const ds = (snapshot as { defaultStyle?: unknown }).defaultStyle;
    if (!ds || typeof ds !== 'object') return undefined;
    const ff = (ds as { ff?: unknown }).ff;
    return typeof ff === 'string' && ff ? ff : undefined;
}

export async function snapshotToXlsxBuffer(snapshot: UniverSnapshot): Promise<ArrayBuffer> {
    const wb = new ExcelJS.Workbook();
    const sheetOrder = (snapshot as { sheetOrder?: string[] }).sheetOrder ?? [];
    const sheets = (snapshot as { sheets?: Record<string, SheetRecord> }).sheets ?? {};
    const tableResource = readTableResource(snapshot);
    const synthStylesBySheet = readSynthStylesSidecar(snapshot);
    const defaultFontName = readDefaultFontName(snapshot);

    for (const sheetId of sheetOrder) {
        const sheet = sheets[sheetId];
        if (!sheet) continue;
        const ws = wb.addWorksheet(sheet.name || sheetId);

        const synthForSheet = synthStylesBySheet[sheetId] ?? {};

        const cellData = sheet.cellData ?? {};
        for (const rowKey of Object.keys(cellData)) {
            const r = Number(rowKey);
            const row = cellData[r];
            if (!row) continue;
            for (const colKey of Object.keys(row)) {
                const c = Number(colKey);
                const data = row[c];
                if (!data) continue;
                const cell = ws.getCell(r + 1, c + 1);
                const hyperlinkUrl = extractHyperlinkFromCellP(data.p);
                if (data.f) {
                    const formula = data.f.startsWith('=') ? data.f.slice(1) : data.f;
                    const result = data.v;
                    cell.value = { formula, result } as ExcelJS.CellFormulaValue;
                } else if (hyperlinkUrl && data.v !== undefined && data.v !== null) {
                    // Hyperlink-bearing cells use exceljs's { text, hyperlink }
                    // shape — exceljs writes the proper <hyperlinks> block
                    // and the sheet rels entry pointing to the URL.
                    cell.value = { text: String(data.v), hyperlink: hyperlinkUrl };
                } else if (data.v !== undefined && data.v !== null) {
                    cell.value = data.v;
                }
                const style = resolveStyle(snapshot, data.s);
                if (style) {
                    const skipFields = synthForSheet[`${r}:${c}`];
                    const skip = skipFields && skipFields.length > 0 ? new Set(skipFields) : undefined;
                    applyStyleToCell(cell, style, skip);
                }
                // Workbook-default font (Aptos Narrow / Calibri / etc.) is
                // already inherited via theme1.xml's minorFont — see
                // patchThemeFont. Don't write a redundant per-cell
                // `font.name` here: doing so flips applyFont="1" on the
                // cell's xf, which prevents Excel from applying TableStyle
                // dxfs (most visibly: the white-on-color header text of
                // styled tables ends up rendered black because the cell's
                // explicit-font flag overrides the table-style font.)
            }
        }

        if (Array.isArray(sheet.mergeData)) {
            for (const m of sheet.mergeData) {
                ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
            }
        }

        if (sheet.rowData) {
            for (const rowKey of Object.keys(sheet.rowData)) {
                const r = Number(rowKey);
                const h = sheet.rowData[r]?.h;
                if (typeof h === 'number' && h > 0) {
                    ws.getRow(r + 1).height = h * (72 / 96);
                }
            }
        }

        if (sheet.columnData) {
            for (const colKey of Object.keys(sheet.columnData)) {
                const c = Number(colKey);
                const w = sheet.columnData[c]?.w;
                if (typeof w === 'number' && w > 0) {
                    ws.getColumn(c + 1).width = Math.max(1, (w - 5) / 7);
                }
            }
        }

        // Tables go last, after all cell data and styling has been written.
        // exceljs's addTable computes the table's full sheet range from
        // (header? + rows.length + totals?). Passing `rows: []` gives a
        // collapsed range (e.g. A1:E0) that Excel rejects with a "Removed
        // Part: table.xml load error" on open. We instead pass an array of
        // empty arrays sized to the actual data-row count from the
        // snapshot range; exceljs's store() iterates each row's `data`
        // (length 0) and writes nothing into the data cells, leaving the
        // values we already wrote from cellData intact. The header row IS
        // overwritten with column.name (which equals the cellData header
        // we already wrote — idempotent).
        const sheetTables = tableResource[sheetId]?.tables ?? [];
        for (const t of sheetTables) {
            try {
                if (!t.range || !Array.isArray(t.columns) || t.columns.length === 0) continue;
                const headerRow = t.options?.showHeader !== false;
                const totalsRow = !!t.options?.showFooter;
                const totalHeight = t.range.endRow - t.range.startRow + 1;
                const dataRowCount = Math.max(0, totalHeight - (headerRow ? 1 : 0) - (totalsRow ? 1 : 0));
                const tlRef = colLetters(t.range.startColumn) + (t.range.startRow + 1);
                // Reconstruct the original Excel table style from meta we
                // stashed on import. Falls back to exceljs's default
                // (TableStyleMedium2, no stripes) when meta is absent.
                const meta = (t.meta ?? {}) as NotesheetTableMeta;
                const tableStyle: Record<string, unknown> = {};
                if (meta.notesheetExcelStyleName) tableStyle.theme = meta.notesheetExcelStyleName;
                if (meta.notesheetShowRowStripes) tableStyle.showRowStripes = true;
                ws.addTable({
                    name: t.name,
                    ref: tlRef,
                    headerRow,
                    totalsRow,
                    ...(Object.keys(tableStyle).length > 0 ? { style: tableStyle } : {}),
                    columns: t.columns.map((c) => ({ name: c.displayName })),
                    // Sized empty rows so exceljs derives the right tableRef
                    // without writing into our already-populated data cells.
                    rows: Array.from({ length: dataRowCount }, () => []),
                });
            } catch (e) {
                console.warn('[Notesheet] could not export table', t?.name, e);
            }
        }
    }

    if (wb.worksheets.length === 0) {
        wb.addWorksheet('Sheet1');
    }

    const buffer = await wb.xlsx.writeBuffer();
    // Post-process pipeline. Each step is a no-op when its data isn't
    // present and fails soft (returns the input buffer) on error.
    //   1. M12: rewrite theme1.xml's font scheme so the exported workbook's
    //      default font matches what was set on import.
    //   2. M10: inject native OOXML chart parts.
    let out = buffer as ArrayBuffer;
    if (defaultFontName) {
        out = await patchThemeFont(out, defaultFontName);
    }
    const sourceClrScheme = readSourceClrScheme(snapshot);
    if (sourceClrScheme) {
        out = await patchThemeClrScheme(out, sourceClrScheme);
    }
    out = await injectChartsIntoZip(out, snapshot);
    return out;
}

function readSourceClrScheme(snapshot: UniverSnapshot): string | null {
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> }).resources;
    if (!Array.isArray(resources)) return null;
    const entry = resources.find((r) => r?.name === NOTESHEET_THEME_CLR_SCHEME_RESOURCE);
    return entry && typeof entry.data === 'string' ? entry.data : null;
}

async function patchThemeClrScheme(buffer: ArrayBuffer, clrSchemeXml: string): Promise<ArrayBuffer> {
    try {
        const zip = await JSZip.loadAsync(buffer);
        const themePath = Object.keys(zip.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p));
        if (!themePath) return buffer;
        const xml = await zip.files[themePath].async('string');
        // Splice over the existing <a:clrScheme>...</a:clrScheme>, leaving
        // the rest of the theme XML (font scheme, format scheme, etc.)
        // untouched.
        const re = /<a:clrScheme\b[^>]*>[\s\S]*?<\/a:clrScheme>/;
        if (!re.test(xml)) return buffer;
        const patched = xml.replace(re, clrSchemeXml);
        if (patched === xml) return buffer;
        zip.file(themePath, patched);
        return await zip.generateAsync({ type: 'arraybuffer' }) as ArrayBuffer;
    } catch (e) {
        console.warn('[Notesheet] patchThemeClrScheme failed; theme keeps exceljs defaults', e);
        return buffer;
    }
}

// Rewrite the workbook's xl/theme/theme1.xml to use the given font name as
// both major and minor font's latin typeface. exceljs ships a hardcoded
// Calibri/Cambria theme; without this patch, even though we set
// cell.font.name on every cell, the exported file's "default font" picker
// in Excel still says Calibri.
async function patchThemeFont(buffer: ArrayBuffer, fontName: string): Promise<ArrayBuffer> {
    try {
        const zip = await JSZip.loadAsync(buffer);
        const themePath = Object.keys(zip.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p));
        if (!themePath) return buffer;
        const xml = await zip.files[themePath].async('string');
        // Replace the latin typeface attribute inside the major/minor font
        // blocks. We only touch <a:latin typeface="..."/> right after the
        // opening <a:majorFont>/<a:minorFont> tag — the script-specific
        // <a:font script="..."> entries are left as-is.
        const safeName = fontName.replace(/"/g, '&quot;').replace(/&/g, '&amp;');
        const patched = xml
            .replace(/(<a:majorFont>\s*<a:latin\b[^>]*\btypeface=")[^"]*(")/, `$1${safeName}$2`)
            .replace(/(<a:minorFont>\s*<a:latin\b[^>]*\btypeface=")[^"]*(")/, `$1${safeName}$2`);
        if (patched === xml) return buffer;
        zip.file(themePath, patched);
        return await zip.generateAsync({ type: 'arraybuffer' }) as ArrayBuffer;
    } catch (e) {
        console.warn('[Notesheet] patchThemeFont failed; theme keeps Calibri default', e);
        return buffer;
    }
}
