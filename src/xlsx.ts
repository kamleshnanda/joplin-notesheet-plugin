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
import { CHART_PALETTE } from './charts/extractData';
import {
    readChartsFromXlsxZip,
    stripChartPartsFromZip,
    type ImportedChartDrawing,
} from './charts/xlsxChartImport';
import { EXCEL_TABLE_STYLE_BY_NAME, type ExcelTableStyle } from './charts/excelTableStyles';
import {
    EXCEL_TABLE_STYLE_RECIPE_BY_NAME,
    EXCEL_TABLE_STYLE_EMPIRICAL_OVERRIDES,
    resolveColorSlot,
    type ExcelTableStyleRecipe,
    type ColorSlotRecipe,
} from './charts/excelTableStyleRecipes';

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

// Univer's Conditional Formatting plugin reads / writes its rules
// through this resource entry name (M15). Confirmed in
// `@univerjs/sheets-conditional-formatting/lib/types/base/const.d.ts`
// — exported as `SHEET_CONDITIONAL_FORMATTING_PLUGIN`. We DON'T import
// the constant from the package here because that would couple
// `src/xlsx.ts` (a runtime-and-test module) to the CF plugin's
// transitive `lodash-es` ESM, which jest-runtime can't parse without a
// transformer override. The string is stable per Univer's stated API
// surface; a Jest test in `excelReferenceFidelity.test.ts` pins the
// same literal so a Univer-side rename trips a loud test failure.
export const CONDITIONAL_FORMATTING_RESOURCE = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN';

// Typed error surface for .xlsx import failures. exceljs's reconcile pipeline
// has a few known crash sites we can't fix without forking (chart drawings
// that lose their drawing reference, multi-sheet workbooks with multiple
// named tables). When those fire, the caller would otherwise see a raw
// `TypeError: Cannot read properties of undefined (reading 'anchors')`
// stack from deep inside node_modules. Wrap them in this typed error so
// the user-facing dialog (src/index.ts) and editor status bar
// (src/editorView.tsx) can show something a Notesheet user can act on.
//
// `code` values are stable strings; the next agent + docs may key off them.
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
    // Source default column width in Excel "character" units (from
    // <sheetFormatPr defaultColWidth/baseColWidth>). Carried through so
    // export can re-emit it; columns without an explicit <col> inherit
    // this width and would otherwise collapse to exceljs's 8.43 default.
    // Absent when the source didn't specify one.
    defaultColWidthChars?: number;
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
        // Text rotation. exceljs surfaces this as either a signed integer
        // in [-90, 90] (CCW positive) or the literal string 'vertical' for
        // OOXML's stacked-text mode 255. Univer encodes both via
        // ITextRotation: { a: number, v?: 1 }, where v=1 means stacked.
        // We deliberately skip the no-op `textRotation === 0` case so the
        // snapshot doesn't carry a trivial `tr: { a: 0 }` for every cell.
        const rot = align.textRotation;
        if (rot === 'vertical') {
            style.tr = { a: 0, v: 1 };
        } else if (typeof rot === 'number' && rot !== 0) {
            style.tr = { a: rot };
        }
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
            // Border color resolution. Three cases:
            //   1. side.color has an argb (e.g. "FF000000" for black) →
            //      use it directly. argbToHex's "all-zero RGB → undefined"
            //      heuristic is correct for FILLS but WRONG for borders —
            //      a user-selected black border has argb FF000000 and we
            //      must not drop it. We bypass argbToHex here for argb
            //      borders.
            //   2. side.color is a theme reference (e.g. {theme:4, tint:0.4}).
            //      Try the palette resolver; if it can't resolve (no theme
            //      palette imported), drop the border rather than render
            //      bold black — a wrong-color black border looks like a
            //      stray underline on a banded row.
            //   3. side.color is absent or empty → Excel's "automatic" =
            //      black. Use #000000.
            const argb = (side.color as { argb?: string } | undefined)?.argb;
            const isThemeRef = side.color && typeof side.color === 'object' && 'theme' in (side.color as object);
            let rgb: string | null = null;
            if (typeof argb === 'string' && /^[0-9A-Fa-f]{8}$/.test(argb)) {
                rgb = '#' + argb.slice(2).toUpperCase();
            } else if (isThemeRef) {
                rgb = resolveExceljsColor(side.color as ExceljsColor, themePalette) ?? null;
            } else {
                rgb = '#000000';
            }
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

// Build a Univer ITextStyle (ITextRun.ts) from an exceljs run-level Font.
// Mirrors the font-extraction block in buildStyleFromExcelCell but
// scoped to a single rich-text run, and resolves theme/tint colors via
// the same applyOoxmlTint path so a rich-text run that uses theme=10
// resolves consistently with cell-level fonts.
function buildTextStyleFromExceljsFont(
    font: ExcelJS.Font | undefined,
    themePalette: ThemePalette | null,
): Record<string, unknown> {
    if (!font) return {};
    const ts: Record<string, unknown> = {};
    if (font.name) ts.ff = font.name;
    if (typeof font.size === 'number') ts.fs = font.size;
    if (font.bold) ts.bl = 1;
    if (font.italic) ts.it = 1;
    if (font.underline) ts.ul = { s: 1 };
    if (font.strike) ts.st = { s: 1 };
    const color = resolveExceljsColor(font.color as ExceljsColor, themePalette);
    if (color) ts.cl = { rgb: color };
    return ts;
}

// Build a Univer cell.p (IDocumentData) carrying multi-run rich text.
// Same documentSkeleton shape as buildHyperlinkCellP — a finite
// pageSize that survives JSON.stringify, paragraphs/sectionBreaks at
// the right offsets — so Univer's layout pipeline doesn't crash on
// hover. Each input run becomes one textRun with character offsets
// computed from the concatenated text.
function buildRichTextCellP(
    runs: Array<{ text: string; ts: Record<string, unknown> }>,
): Record<string, unknown> {
    const dataStream = runs.map((r) => r.text).join('');
    let pos = 0;
    const textRuns = runs.map((r) => {
        const start = pos;
        pos += r.text.length;
        return { st: start, ed: pos, ts: r.ts };
    });
    return {
        id: '__INTERNAL_EDITOR__DOCS_NORMAL',
        body: {
            dataStream: dataStream + '\r\n',
            textRuns,
            paragraphs: [{ startIndex: dataStream.length, paragraphStyle: {} }],
            sectionBreaks: [{ startIndex: dataStream.length + 1 }],
        },
        documentStyle: {
            pageSize: {
                width: HYPERLINK_PAGE_SIZE,
                height: HYPERLINK_PAGE_SIZE,
            },
        },
    };
}

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

function extractCellValue(
    cell: ExcelJS.Cell,
    themePalette: ThemePalette | null = null,
): { v?: string | number | boolean; f?: string; t?: number; p?: Record<string, unknown> } {
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

    // Rich text: a cell with multiple per-run formatting blocks
    // (e.g. bold word + plain word in one cell). exceljs surfaces this
    // as `cell.value = { richText: [{font, text}, ...] }`. We emit
    // cell.p with one textRun per source run so Univer's editor can
    // render the formatting; cell.v carries the plain-text concat for
    // string-only consumers.
    //
    // Single-run "rich text" (length 1) collapses to a plain string —
    // emitting cell.p just for one uniformly-styled run would bloat
    // every cell exceljs sometimes wraps as a 1-element richText.
    if (typeof raw === 'object' && 'richText' in raw && Array.isArray((raw as { richText: unknown }).richText)) {
        const segments = (raw as { richText: Array<{ text?: string; font?: ExcelJS.Font }> }).richText;
        const plain = segments.map((s) => s.text ?? '').join('');
        if (segments.length <= 1) {
            return { v: plain, t: VALUE_STRING };
        }
        const runs = segments.map((s) => ({
            text: s.text ?? '',
            ts: buildTextStyleFromExceljsFont(s.font, themePalette),
        }));
        const p = buildRichTextCellP(runs);
        return { v: plain, t: VALUE_STRING, p };
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

// Pattern B hyperlink detection. exceljs surfaces cells styled with the
// built-in "Hyperlink" cellStyle (cellStyleXfs builtinId=8) as plain
// strings — `cell.value = "https://..."` with `cell.isHyperlink = false`.
// Excel still renders them as clickable blue-underlined links in the UI
// because the named-style font carries underline + theme=10 color, but
// our import path treats them as ordinary cells.
//
// We resolve Pattern B by reading the raw zip:
//   1. xl/styles.xml's <cellStyles> → find the xfId that points at
//      builtinId=8 (or name="Hyperlink").
//   2. xl/styles.xml's <cellXfs> → collect every cellXfs index whose
//      `xfId` attribute matches the named-Hyperlink xfId.
//   3. Per worksheet, xl/worksheets/sheet*.xml → collect every <c r="A1"
//      s="N"/> where N is in the named-Hyperlink set.
//
// Returns a map: sheetIndex (1-based, matching exceljs's ws.id) → Set
// of A1-format cell refs. The import loop consults this set to decide
// whether to synthesize a hyperlink cell.p for cells exceljs reports as
// plain strings.
//
// IMPORTANT: This is import-only. On export we keep emitting Pattern A
// (`{text, hyperlink}` cell value + <hyperlinks> block), which Excel
// renders identically. Round-tripping the named-style itself would
// require reconstructing builtin cellStyle entries in the exported
// styles.xml, which exceljs doesn't expose cleanly.
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

    // Step 1: find the xfId of <cellStyle name="Hyperlink"> (or
    // builtinId=8). Excel always emits builtinId=8 for Hyperlink, but
    // the name attribute is the safer match because OOXML treats
    // builtinId as a hint, not a contract.
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

    // Step 2: walk <cellXfs> and collect cellXfs indices whose xfId
    // attribute is in the set.
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

    // Step 3: per worksheet, collect <c r="A1" s="N"/> where N is in
    // linkCellXfIndices. We don't try to track which worksheet this is
    // by name — exceljs ws.id is the same 1-based index used in the
    // sheet path "xl/worksheets/sheetN.xml".
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

// ───────── M15: Conditional Formatting translators ─────────────────────
//
// Univer's CF model and exceljs's CF surface differ in shape; each rule
// type gets a per-direction translator pair (xlsx → Univer on import;
// Univer → xlsx on export). All five exceljs rule types we support are
// covered:
//
//   exceljs `colorScale` → Univer `colorScale`
//   exceljs `dataBar`    → Univer `dataBar`
//   exceljs `cellIs`     → Univer `highlightCell` (subType=`number`)
//   exceljs `top10`      → Univer `highlightCell` (subType=`rank`)
//   exceljs `iconSet`    → Univer `iconSet`
//
// Colours: exceljs surfaces argb (`FF...`); Univer uses `#RRGGBB`.
// The existing `argbToHex` / `hexToArgb` helpers are the seam.
//
// Snapshot resource shape (per Univer's CF model service):
//   { [subUnitId]: IConditionFormattingRule[] }
// where each rule has `{cfId, ranges, stopIfTrue, rule: <type-specific>}`.
// We JSON.stringify the entire object as the resource's `data` field —
// exactly what the CF preset's `parseJson` expects on snapshot load.

interface UniverCfRange {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}

interface UniverCfRuleEntry {
    cfId: string;
    ranges: UniverCfRange[];
    stopIfTrue: boolean;
    rule: Record<string, unknown>;
}

// Convert an exceljs CF `ref` (a space-separated list of A1 ranges)
// to Univer's IRange[] shape. exceljs uses tokens like "A2:A11" or
// single cells "A2".
function cfRefToRanges(ref: string | undefined): UniverCfRange[] {
    if (!ref || typeof ref !== 'string') return [];
    const tokens = ref.split(/\s+/).filter(Boolean);
    const ranges: UniverCfRange[] = [];
    for (const token of tokens) {
        const r = parseA1Range(token);
        if (r) ranges.push(r);
    }
    return ranges;
}

// Translate one cfvo entry to Univer's IValueConfig.
// exceljs cfvo: { type: 'min'|'max'|'percentile'|'percent'|'num'|'formula', value?: number|string }
// Univer:       { type, value? }
function translateCfvoToUniver(cfvo: { type?: string; value?: number | string }): { type: string; value?: number | string } {
    const type = cfvo.type ?? 'num';
    const out: { type: string; value?: number | string } = { type };
    if (cfvo.value !== undefined && cfvo.value !== null) {
        // exceljs sometimes surfaces value as a string (formula); we
        // keep numeric percentile values as numbers, formula strings
        // as strings.
        if (type === 'formula') out.value = String(cfvo.value);
        else if (typeof cfvo.value === 'string' && /^-?\d+(?:\.\d+)?$/.test(cfvo.value)) {
            out.value = Number(cfvo.value);
        } else out.value = cfvo.value;
    }
    return out;
}

// Inverse: Univer IValueConfig → exceljs cfvo entry.
function translateCfvoToExceljs(uvo: { type: string; value?: number | string }): { type: string; value?: number | string } {
    const out: { type: string; value?: number | string } = { type: uvo.type };
    if (uvo.value !== undefined && uvo.value !== null) out.value = uvo.value;
    return out;
}

// Generate a stable cfId. Univer expects each rule to carry one;
// uniqueness only matters within a subUnit. Counter-based plus an
// import-tag prefix so collisions with rules added later via the UI
// are extremely unlikely.
let _cfIdCounter = 0;
function nextCfId(): string {
    _cfIdCounter = (_cfIdCounter + 1) | 0;
    return `cf-import-${Date.now().toString(36)}-${_cfIdCounter.toString(36)}`;
}

// 3Arrows / 3TrafficLights1 / etc. — Univer's iconMap lists icon
// indices. Each named iconSet has a fixed length (3, 4, or 5).
// Sources from Univer's source: `iconGroup` in
// @univerjs/sheets-conditional-formatting/.../models/icon-map.
const ICON_SET_LENGTH: Record<string, number> = {
    '3Arrows': 3, '3ArrowsGray': 3,
    '4Arrows': 4, '4ArrowsGray': 4,
    '5Arrows': 5, '5ArrowsGray': 5,
    '3Triangles': 3,
    '3TrafficLights1': 3, '3TrafficLights2': 3,
    '3Signs': 3,
    '4RedToBlack': 4,
    '4TrafficLights': 4,
    '3Symbols': 3, '3Symbols2': 3,
    '3Flags': 3,
    '4Rating': 4,
    '5Rating': 5, '5Quarters': 5, '5Boxes': 5,
    '3Stars': 3,
    '_5Felling': 5,
};

interface ExceljsCfRuleColorScale {
    type: 'colorScale';
    priority?: number;
    cfvo: Array<{ type: string; value?: number | string }>;
    color: Array<{ argb?: string }>;
}
interface ExceljsCfRuleDataBar {
    type: 'dataBar';
    priority?: number;
    cfvo: Array<{ type: string; value?: number | string }>;
    color: { argb?: string };
    gradient?: boolean;
}
interface ExceljsCfRuleCellIs {
    type: 'cellIs';
    priority?: number;
    operator?: string;
    formulae?: string[];
    style?: { fill?: { type?: string; bgColor?: { argb?: string } } };
}
interface ExceljsCfRuleTop10 {
    type: 'top10';
    priority?: number;
    rank?: number;
    bottom?: boolean;
    percent?: boolean;
    style?: { fill?: { type?: string; bgColor?: { argb?: string } } };
}
interface ExceljsCfRuleIconSet {
    type: 'iconSet';
    priority?: number;
    iconSet?: string;
    cfvo: Array<{ type: string; value?: number | string }>;
}
type ExceljsCfRule =
    | ExceljsCfRuleColorScale
    | ExceljsCfRuleDataBar
    | ExceljsCfRuleCellIs
    | ExceljsCfRuleTop10
    | ExceljsCfRuleIconSet
    | { type: string; priority?: number; [k: string]: unknown };

interface ExceljsConditionalFormatting {
    ref?: string;
    rules?: ExceljsCfRule[];
}

// Translate one exceljs CF rule to Univer's IConditionFormattingRule.
// Returns null for rule types we don't (yet) support; the caller drops
// them and the round-trip test pin-down treats unsupported types as
// out of scope.
function translateExceljsCfRuleToUniver(rule: ExceljsCfRule, ranges: UniverCfRange[]): UniverCfRuleEntry | null {
    if (!rule || !rule.type) return null;
    const cfId = nextCfId();
    const stopIfTrue = false;
    switch (rule.type) {
        case 'colorScale': {
            const r = rule as ExceljsCfRuleColorScale;
            if (!Array.isArray(r.cfvo) || !Array.isArray(r.color)) return null;
            // Univer's colorScale shape: config: [{index, color, value}, ...]
            const config: Array<{ index: number; color: string; value: { type: string; value?: number | string } }> = [];
            for (let i = 0; i < r.cfvo.length; i++) {
                const colorArgb = r.color[i]?.argb;
                const hex = argbToHex(colorArgb) ?? '#000000';
                config.push({
                    index: i,
                    color: hex,
                    value: translateCfvoToUniver(r.cfvo[i]),
                });
            }
            return {
                cfId,
                ranges,
                stopIfTrue,
                rule: {
                    type: 'colorScale',
                    config,
                },
            };
        }
        case 'dataBar': {
            const r = rule as ExceljsCfRuleDataBar;
            if (!Array.isArray(r.cfvo) || r.cfvo.length < 2) return null;
            const colorArgb = r.color?.argb;
            const hex = argbToHex(colorArgb) ?? '#638EC6';
            return {
                cfId,
                ranges,
                stopIfTrue,
                rule: {
                    type: 'dataBar',
                    isShowValue: true,
                    config: {
                        min: translateCfvoToUniver(r.cfvo[0]),
                        max: translateCfvoToUniver(r.cfvo[1]),
                        // Excel's <dataBar> defaults gradient=true unless
                        // `<dataBar gradient="0">`. exceljs surfaces the
                        // flag inconsistently across versions; default
                        // to true to match Excel's render.
                        isGradient: r.gradient !== false,
                        positiveColor: hex,
                        nativeColor: hex,
                    },
                },
            };
        }
        case 'cellIs': {
            const r = rule as ExceljsCfRuleCellIs;
            const op = r.operator ?? 'greaterThan';
            const formula = (r.formulae && r.formulae[0]) ?? '0';
            const numValue = /^-?\d+(?:\.\d+)?$/.test(formula) ? Number(formula) : formula;
            const bgArgb = r.style?.fill?.bgColor?.argb;
            const bgHex = argbToHex(bgArgb);
            const style: Record<string, unknown> = {};
            if (bgHex) style.bg = { rgb: bgHex };
            return {
                cfId,
                ranges,
                stopIfTrue,
                rule: {
                    type: 'highlightCell',
                    subType: 'number',
                    operator: op,
                    value: numValue,
                    style,
                },
            };
        }
        case 'top10': {
            const r = rule as ExceljsCfRuleTop10;
            const bgArgb = r.style?.fill?.bgColor?.argb;
            const bgHex = argbToHex(bgArgb);
            const style: Record<string, unknown> = {};
            if (bgHex) style.bg = { rgb: bgHex };
            return {
                cfId,
                ranges,
                stopIfTrue,
                rule: {
                    type: 'highlightCell',
                    subType: 'rank',
                    isBottom: !!r.bottom,
                    isPercent: !!r.percent,
                    value: typeof r.rank === 'number' ? r.rank : 10,
                    style,
                },
            };
        }
        case 'iconSet': {
            const r = rule as ExceljsCfRuleIconSet;
            const iconSet = r.iconSet ?? '3Arrows';
            const length = ICON_SET_LENGTH[iconSet] ?? r.cfvo?.length ?? 3;
            // Univer's IconSetCalculateUnit walks the config from
            // index 0 forward, returning the first item whose
            // value/operator condition the cell value satisfies. The
            // standard layout in Univer's iconMap puts the HIGH icon
            // at index 0 (e.g. 3Arrows = [up-green, right-gold,
            // down-red]). To match Excel's iconSet semantics
            // (lowest icon for the lowest band), we emit config in
            // descending threshold order:
            //   index 0 → highest band, operator >= last cfvo
            //   index 1 → middle band,  operator >= middle cfvo
            //   ...
            //   index N-1 → catch-all, operator <= MAX_SAFE_INTEGER
            //
            // Source cfvo for the M15 fixture is ascending: [0, 33, 67].
            // Mapping: cfvo[N-1]=67 → index 0 (up-green for >=67);
            // cfvo[1]=33 → index 1 (right-gold for >=33); the lowest
            // band gets the catch-all rule (down-red).
            const cfvos = (r.cfvo ?? []).map(translateCfvoToUniver);
            const config: Array<{
                operator: string;
                value: { type: string; value?: number | string };
                iconType: string;
                iconId: string;
            }> = [];
            for (let i = 0; i < length; i++) {
                if (i < length - 1) {
                    // Walk source cfvos in descending order: the highest
                    // threshold goes to index 0, second-highest to
                    // index 1, etc. cfvos.length should equal
                    // `length`, but be defensive against malformed
                    // sources by using a safe cfvo when missing.
                    const cfvoIdx = cfvos.length - 1 - i;
                    const v = cfvos[cfvoIdx >= 0 ? cfvoIdx : 0]
                        ?? { type: 'percent', value: 0 };
                    config.push({
                        operator: 'greaterThanOrEqual',
                        value: v,
                        iconType: iconSet,
                        iconId: String(i),
                    });
                } else {
                    // Last entry is a catch-all that always matches
                    // anything that didn't match a higher-priority
                    // entry. Univer's IconSetCalculateUnit returns the
                    // last entry unconditionally when index ===
                    // length - 1, so the operator+value here are
                    // formally inert; we still write a sensible value
                    // for round-trip purity.
                    config.push({
                        operator: 'lessThanOrEqual',
                        value: { type: 'num', value: Number.MAX_SAFE_INTEGER },
                        iconType: iconSet,
                        iconId: String(i),
                    });
                }
            }
            return {
                cfId,
                ranges,
                stopIfTrue,
                rule: {
                    type: 'iconSet',
                    isShowValue: true,
                    config,
                },
            };
        }
        default:
            // Out-of-scope rule types (text-based, time-period, unique,
            // duplicate, formula, average) get dropped silently. Future
            // milestones may extend this switch; the round-trip test
            // pins the fixture's 5 supported types.
            return null;
    }
}

// Inverse: Univer rule → exceljs CF rule. Returns null for rule shapes
// we don't recognize (e.g. CF rules added via the UI for highlightCell
// subTypes other than `number` / `rank`).
function translateUniverCfRuleToExceljs(entry: UniverCfRuleEntry, priority: number): {
    rule: ExceljsCfRule | null;
    sqref: string;
} | null {
    if (!entry || !entry.rule || !Array.isArray(entry.ranges) || entry.ranges.length === 0) return null;
    const sqref = entry.ranges.map((r) => {
        const tl = colLetters(r.startColumn) + (r.startRow + 1);
        const br = colLetters(r.endColumn) + (r.endRow + 1);
        return tl === br ? tl : `${tl}:${br}`;
    }).join(' ');
    const rule = entry.rule as Record<string, unknown>;
    const type = rule.type as string;
    switch (type) {
        case 'colorScale': {
            const config = rule.config as Array<{ color: string; value: { type: string; value?: number | string } }>;
            if (!Array.isArray(config) || config.length === 0) return null;
            const cfvo = config.map((e) => translateCfvoToExceljs(e.value));
            const color = config.map((e) => ({ argb: hexToArgb(e.color) }));
            return {
                sqref,
                rule: { type: 'colorScale', priority, cfvo, color } as ExceljsCfRuleColorScale,
            };
        }
        case 'dataBar': {
            const config = rule.config as { min: { type: string; value?: number | string }; max: { type: string; value?: number | string }; positiveColor: string };
            if (!config) return null;
            return {
                sqref,
                rule: {
                    type: 'dataBar',
                    priority,
                    cfvo: [translateCfvoToExceljs(config.min), translateCfvoToExceljs(config.max)],
                    color: { argb: hexToArgb(config.positiveColor) },
                } as ExceljsCfRuleDataBar,
            };
        }
        case 'highlightCell': {
            const subType = rule.subType as string;
            const style = rule.style as { bg?: { rgb?: string } } | undefined;
            const bgArgb = hexToArgb(style?.bg?.rgb);
            const dxfStyle = bgArgb
                ? { fill: { type: 'pattern' as const, pattern: 'solid' as const, bgColor: { argb: bgArgb } } }
                : undefined;
            if (subType === 'number') {
                const op = (rule.operator as string) ?? 'greaterThan';
                const value = rule.value;
                const formulae = value !== undefined && value !== null ? [String(value)] : ['0'];
                return {
                    sqref,
                    rule: {
                        type: 'cellIs',
                        priority,
                        operator: op,
                        formulae,
                        ...(dxfStyle ? { style: dxfStyle } : {}),
                    } as ExceljsCfRuleCellIs,
                };
            }
            if (subType === 'rank') {
                const isBottom = !!rule.isBottom;
                const value = rule.value;
                const rank = typeof value === 'number' ? value : 10;
                return {
                    sqref,
                    rule: {
                        type: 'top10',
                        priority,
                        rank,
                        bottom: isBottom,
                        ...(rule.isPercent ? { percent: true } : {}),
                        ...(dxfStyle ? { style: dxfStyle } : {}),
                    } as ExceljsCfRuleTop10,
                };
            }
            // Other highlightCell sub-types (text/timePeriod/formula/
            // average/unique/duplicate) are out of scope for M15.
            return null;
        }
        case 'iconSet': {
            const config = rule.config as Array<{ iconType: string; iconId: string; value: { type: string; value?: number | string } }>;
            if (!Array.isArray(config) || config.length === 0) return null;
            const iconSet = config[0].iconType;
            // Inverse mapping from Univer's descending-threshold layout
            // back to Excel's ascending cfvo. The catch-all at index
            // (length-1) is dropped from the cfvo array; the remaining
            // entries are reversed to ascending order.
            //
            // Univer config (3Arrows): [
            //   { iconId:0, value:{percent,67} },     // highest band
            //   { iconId:1, value:{percent,33} },     // middle band
            //   { iconId:2, value:{num, MAX_SAFE_INTEGER} }, // catch-all
            // ]
            // Excel cfvo (ascending): [
            //   { type:percent, value:0 },             // band 0 (lowest)
            //   { type:percent, value:33 },            // band 1
            //   { type:percent, value:67 },            // band 2
            // ]
            const real = config.slice(0, -1);            // drop the catch-all
            const ascending = real.slice().reverse();    // descend → ascend
            const cfvo: Array<{ type: string; value?: number | string }> = [
                // Excel's first cfvo is always the lowest band — we
                // synthesize one for it (always 0% / min) since Univer
                // doesn't model the lowest-band threshold directly.
                { type: ascending[0]?.value.type ?? 'percent', value: 0 },
                ...ascending.map((c) => translateCfvoToExceljs(c.value)),
            ];
            return {
                sqref,
                rule: {
                    type: 'iconSet',
                    priority,
                    iconSet,
                    cfvo,
                } as ExceljsCfRuleIconSet,
            };
        }
        default:
            return null;
    }
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

// Read each worksheet's effective default column width from its
// <sheetFormatPr>. exceljs surfaces `defaultColWidth` on ws.properties
// but NOT `baseColWidth` (the older attribute Excel actually writes —
// 11-stacked-bar-chart.xlsx ships `baseColWidth="10"` and no
// defaultColWidth). When neither we capture is present the caller keeps
// its own fallback. Returns a map: 1-based sheetIndex -> width in Excel
// "character" units. Excel's effective default when only baseColWidth=N
// is present is N+1 characters (the +1 accounts for cell padding), which
// is what makes a column with no explicit <col> render at ~11.5 chars,
// not 10. Columns WITHOUT an explicit <col> inherit this width; without
// preserving it, such columns (e.g. column B in 11) collapse to exceljs's
// 8.43 default on export and render visibly narrower than the source.
async function readSheetDefaultColWidths(
    buffer: ArrayBuffer | Uint8Array | Buffer,
): Promise<Map<number, number>> {
    const result = new Map<number, number>();
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return result;
    }
    for (const fpath of Object.keys(zip.files)) {
        const m = /^xl\/worksheets\/sheet(\d+)\.xml$/i.exec(fpath);
        if (!m) continue;
        const sheetIndex = parseInt(m[1], 10);
        const xml = await zip.files[fpath].async('string');
        const fmt = xml.match(/<sheetFormatPr\b[^>]*>/i);
        if (!fmt) continue;
        const defW = fmt[0].match(/\bdefaultColWidth="([\d.]+)"/i);
        if (defW) {
            result.set(sheetIndex, parseFloat(defW[1]));
            continue;
        }
        const baseW = fmt[0].match(/\bbaseColWidth="([\d.]+)"/i);
        if (baseW) {
            // Excel's effective default = baseColWidth + 1 character of
            // padding (matches the 11.5-char render observed for a
            // baseColWidth="10" column with no explicit <col>).
            result.set(sheetIndex, parseFloat(baseW[1]) + 1);
        }
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

// Resolve a built-in TableStyle's per-slot colours against a source workbook
// `<a:clrScheme>`. The hardcoded `EXCEL_TABLE_STYLE_BY_NAME` catalog is
// computed for the Aptos accent palette (Office 2016+); when the source
// workbook ships its own clrScheme (e.g. the Classic 2013 palette whose
// accent3 is `#A5A5A5` grey instead of Aptos's `#196B24` green), the same
// `TableStyleMedium4` must paint grey, not green.
//
// We use a parallel recipe table (`EXCEL_TABLE_STYLE_RECIPE_BY_NAME`) that
// names each slot's accent index + tint; at synthesis time we look the
// source clrScheme's accent values up and compute the final RGB via the
// ECMA-376 HSL-L tint formula. Achromatic slots (greys/whites/blacks) keep
// their literal RGB regardless of clrScheme — those are the same in every
// Excel theme.
//
// `themeRgb` is the 12-entry array returned by `readThemeClrScheme`
// (lt1, dk1, lt2, dk2, accent1..accent6, hlink, folHlink). When it's null
// we fall back to the static catalog (Aptos baseline) — the legacy path.
function resolveTableStylePalette(
    styleName: string,
    catalog: ExcelTableStyle | undefined,
    themeRgb: readonly string[] | null,
): ExcelTableStyle | undefined {
    if (!catalog) return undefined;
    if (!themeRgb) return catalog;
    const recipe: ExcelTableStyleRecipe | undefined = EXCEL_TABLE_STYLE_RECIPE_BY_NAME[styleName];
    if (!recipe) return catalog;
    // clrScheme's accent1..accent6 are at indices 4..9 in `themeRgb`
    // (lt1, dk1, lt2, dk2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink).
    const accents: string[] = themeRgb.slice(4, 10);

    // Empirical override lookup: when the source accent for this style's
    // headerBg slot matches a measured RGB in the override table, prefer
    // the measured slot values over the formula. The override map is keyed
    // by `(styleName, accentHex_uppercase)`. Lookup uses the accent that
    // headerBg references (typically the only one a TableStyle uses).
    const overrideTable = EXCEL_TABLE_STYLE_EMPIRICAL_OVERRIDES[styleName];
    const accentIdx = recipe.headerBg.accent;
    let override: ReturnType<typeof Object.values<unknown>>[number] | undefined;
    if (overrideTable && accentIdx) {
        const accentHex = (accents[accentIdx - 1] || '').toUpperCase();
        override = overrideTable[accentHex];
    }

    const resolve = (
        slot: ColorSlotRecipe | undefined,
        fallback: string | undefined,
        overrideValue: string | undefined,
    ): string | undefined => {
        if (overrideValue) return overrideValue.toUpperCase();
        if (!slot) return fallback;
        return resolveColorSlot(slot, accents);
    };

    type OverrideShape = Record<string, string | undefined>;
    const ov = (override || {}) as OverrideShape;

    return {
        styleName: catalog.styleName,
        headerBg: resolve(recipe.headerBg, catalog.headerBg, ov.headerBg)!,
        headerFg: resolve(recipe.headerFg, catalog.headerFg, ov.headerFg)!,
        bandedRowEvenBg: resolve(recipe.bandedRowEvenBg, catalog.bandedRowEvenBg, ov.bandedRowEvenBg),
        bandedRowOddBg: resolve(recipe.bandedRowOddBg, catalog.bandedRowOddBg, ov.bandedRowOddBg),
        totalsBg: resolve(recipe.totalsBg, catalog.totalsBg, ov.totalsBg),
        totalsFg: resolve(recipe.totalsFg, catalog.totalsFg, ov.totalsFg),
        borderColor: resolve(recipe.borderColor, catalog.borderColor, ov.borderColor),
        totalsTopBorder: resolve(recipe.totalsTopBorder, undefined, ov.totalsTopBorder),
        totalsBottomBorder: resolve(recipe.totalsBottomBorder, undefined, ov.totalsBottomBorder),
    };
}

function synthesizeTableStyleAssignments(
    table: RawTable,
    existingCellStyles: Map<string, Record<string, unknown>>,
    themeRgb: readonly string[] | null = null,
): CellStyleAssignment[] {
    if (!table.styleName) return [];
    const catalog: ExcelTableStyle | undefined = EXCEL_TABLE_STYLE_BY_NAME[table.styleName];
    const palette = resolveTableStylePalette(table.styleName, catalog, themeRgb);
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

    // The accent-coloured strip Excel paints at every banded-row
    // boundary (e.g. `#72D068` for Aptos accent3, `#C9C9C9` for Classic
    // accent3). Re-use the recipe's `totalsBottomBorder` slot — it's the
    // same colour Excel uses for this strip in every TableStyleMedium
    // we've measured. (If a future fixture decouples them, split the
    // recipe into a dedicated `interRowStripBorder` slot.) Painted on
    // bd.t of every data row so that, by Univer 0.23's "lower-cell's
    // bd.t wins at a shared edge" rule, the strip shows correctly even
    // on top of the header's bd.b table-outline.
    const interRowStripRgb = palette.totalsBottomBorder;
    const interRowStrip: BorderEntry | undefined = interRowStripRgb
        ? { s: BORDER_STYLE_TO_UNIVER.medium, cl: { rgb: interRowStripRgb } }
        : undefined;

    // Data rows with alternating banding (only when showRowStripes is true).
    if (table.showRowStripes) {
        for (let r = dataStartRow; r <= dataEndRow; r++) {
            const isEven = (r - dataStartRow) % 2 === 0;
            const bg = isEven ? palette.bandedRowEvenBg : palette.bandedRowOddBg;
            // Skip white (#FFFFFF) — no-op for default white background.
            const useBg = bg && bg.toUpperCase() !== '#FFFFFF' ? bg : undefined;
            for (let c = range.startColumn; c <= range.endColumn; c++) {
                const rowBorders: Partial<Record<BorderSide, BorderEntry>> = {};
                if (interRowStrip) rowBorders.t = interRowStrip;
                if (thinBorder) {
                    if (c === range.startColumn) rowBorders.l = thinBorder;
                    if (c === range.endColumn) rowBorders.r = thinBorder;
                    if (r === dataEndRow && totalRows === 0) rowBorders.b = thinBorder;
                }
                const hasBorders = Object.keys(rowBorders).length > 0;
                if (useBg || hasBorders) overlay(r, c, useBg, undefined, false, hasBorders ? rowBorders : undefined);
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
    if (totalRows === 1) {
        // The totals row carries TWO accent-coloured strips with
        // DIFFERENT shapes:
        //   - TOP: header colour (e.g. `#34692E` Aptos accent3 dark
        //     green / `#A5A5A5` Classic accent3 grey), painted as a
        //     DOUBLE-LINE pair (two 2px strips with a 2px white gap,
        //     ~6px total). Pixel-probed at y=826-831 in
        //     `screenshots/excel-reference/FormattingSmorgasboard-Aptos-wide.png`:
        //       y=826-827 #34692E
        //       y=828-829 #FFFFFF
        //       y=830-831 #34692E
        //     and at y=500-505 in the Classic capture:
        //       y=500-501 #A5A5A5
        //       y=502-503 #FFFFFF
        //       y=504-505 #A5A5A5
        //   - BOTTOM: lighter accent (e.g. `#72D068` Aptos / `#C9C9C9`
        //     Classic), single 2px strip — same shade Excel uses for
        //     every banded-row boundary above it. Verified at y=908-909
        //     (Aptos wide) and y=548-549 (Classic).
        //
        // Style choices reflect that asymmetry:
        //   - totals-top → BorderStyleTypes.DOUBLE (s:7). Univer 0.23's
        //     `_renderDoubleBorder` paints two strokes with a gap, which
        //     reproduces Excel's actual double-line shape here.
        //   - totals-bottom → MEDIUM (s:8), single 2px.
        // The earlier choice of MEDIUM for totals-top + `#72D068` was a
        // pixel-probe error — see commit history; corrected here.
        const totalsTopBorderRgb = palette.totalsTopBorder;
        const totalsTopBorder: BorderEntry | undefined = totalsTopBorderRgb
            ? { s: BORDER_STYLE_TO_UNIVER.double, cl: { rgb: totalsTopBorderRgb } }
            : (thinBorder ?? undefined);
        const totalsBottomBorderRgb = palette.totalsBottomBorder;
        const totalsBottomBorder: BorderEntry | undefined = totalsBottomBorderRgb
            ? { s: BORDER_STYLE_TO_UNIVER.medium, cl: { rgb: totalsBottomBorderRgb } }
            : (thinBorder ?? undefined);

        for (let c = range.startColumn; c <= range.endColumn; c++) {
            const totalsBorders: Partial<Record<BorderSide, BorderEntry>> = {};
            if (totalsTopBorder) totalsBorders.t = totalsTopBorder;
            // The totals-row bottom border replaces the table outline's
            // thin frame on the totals row's bottom edge — Excel paints
            // the accent-coloured strip across the full table width, not
            // the outline colour.
            if (totalsBottomBorder) totalsBorders.b = totalsBottomBorder;
            if (thinBorder) {
                if (c === range.startColumn) totalsBorders.l = thinBorder;
                if (c === range.endColumn) totalsBorders.r = thinBorder;
            }
            const useBg = palette.totalsBg && palette.totalsBg.toUpperCase() !== '#FFFFFF' ? palette.totalsBg : undefined;
            overlay(
                range.endRow, c,
                useBg, palette.totalsFg, /* bold */ true,
                Object.keys(totalsBorders).length > 0 ? totalsBorders : undefined,
            );
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
    //
    // M17: BEFORE wb.xlsx.load runs, read chart drawings out of the buffer
    // ourselves and strip them from a copy used for the load. exceljs's
    // chart-side reconcile loop crashes with the legacy `anchors` reference
    // error on workbooks with chart drawings; sidestepping that load by
    // pre-stripping is what unblocks chart-bearing fixtures (MultiSheet.xlsx,
    // LargeWorkbook.xlsx, every fixture under tests/fixtures/charts/).
    //
    // The original buffer is preserved verbatim for the existing post-load
    // zip-direct readers (readTablesFromXlsxZip / readThemeFont /
    // readThemeClrScheme / readNamedHyperlinkCells) — those don't read any
    // chart parts, so the strip step doesn't affect them.
    let importedCharts: ImportedChartDrawing[] = [];
    try {
        importedCharts = await readChartsFromXlsxZip(buffer);
    } catch (e) {
        console.warn('[Notesheet] M17: readChartsFromXlsxZip threw; continuing without chart import', e);
        importedCharts = [];
    }
    // Build a chart-stripped buffer iff there's at least one chart drawing
    // present. For chart-less workbooks we pass the original buffer through
    // unchanged so the existing error-classification path stays intact for
    // non-chart crash classes (e.g. xlsx-multi-table-unsupported).
    let bufferForLoad: ArrayBuffer | Uint8Array | Buffer = buffer;
    if (importedCharts.length > 0) {
        try {
            bufferForLoad = await stripChartPartsFromZip(buffer);
        } catch (e) {
            console.warn('[Notesheet] M17: stripChartPartsFromZip failed; loading original buffer (may crash)', e);
            bufferForLoad = buffer;
        }
    }

    // The try/catch around load() catches three reproducible exceljs reconcile
    // crashes: chart drawings whose drawing reference doesn't resolve (`anchors`
    // crash in lib/xlsx/xlsx.js:100) and multi-sheet workbooks with multiple
    // named tables (`name` crash in lib/doc/worksheet.js:920 inside the tables
    // reduce). We classify by stack frame rather than message alone because
    // "name" is too generic to key off — multiple unrelated exceljs paths
    // can produce a "Cannot read properties of undefined (reading 'name')".
    //
    // The xlsx-charts-unsupported error class stays defined for future
    // drawing-related crash classes M17's strip path doesn't cover (e.g. an
    // image+chart mixed drawing whose strip leaves a dangling ref). The
    // strip pre-empts the common case but isn't a universal cure.
    try {
        await wb.xlsx.load(bufferForLoad as unknown as Parameters<typeof wb.xlsx.load>[0]);
    } catch (err) {
        const e = err as Error;
        const msg = e?.message ?? String(err);
        const stack = e?.stack ?? '';
        if (msg.includes("'anchors'") || msg.includes('anchors')) {
            throw new NotesheetImportError(
                'xlsx-charts-unsupported',
                "This .xlsx contains chart drawings that Notesheet can't import yet. The file imports correctly in Excel but cannot be opened in Notesheet.",
                err,
            );
        }
        if ((msg.includes("'name'") || msg.includes('name')) && /worksheet\.js/.test(stack)) {
            throw new NotesheetImportError(
                'xlsx-multi-table-unsupported',
                "This .xlsx has a structure (multiple sheets each with their own named tables) that Notesheet can't import yet.",
                err,
            );
        }
        throw new NotesheetImportError('xlsx-import-failed', msg, err);
    }

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
    // Pattern B hyperlinks: cells styled with the built-in "Hyperlink"
    // cellStyle (cellStyleXfs builtinId=8). exceljs surfaces these as
    // plain string values — we synthesize cell.p from the string so
    // Univer's hyperlink layer treats them as clickable links.
    const namedHyperlinkCellsBySheet = await readNamedHyperlinkCells(buffer);
    // Per-sheet default column width (Excel "character" units). exceljs
    // drops baseColWidth, so columns with no explicit <col> would lose
    // the source's wider default on round-trip — see issue 11.
    const defaultColWidthBySheet = await readSheetDefaultColWidths(buffer);

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

    // Per-subUnit conditional-formatting rules (M15). Keyed by sheetId,
    // each entry is an array of Univer-shaped IConditionFormattingRule
    // objects. Serialized as the SHEET_CONDITIONAL_FORMATTING_PLUGIN
    // resource so Univer's CF preset paints them on canvas.
    const cfResource: Record<string, UniverCfRuleEntry[]> = {};

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
        const namedHyperlinkCells = namedHyperlinkCellsBySheet.get(ws.id) ?? new Set<string>();

        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const r = rowNumber - 1; // exceljs is 1-based
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                const c = colNumber - 1;
                const value = extractCellValue(cell, themeClrScheme);
                // Pattern B hyperlink synthesis. extractCellValue may already
                // have produced a `p` (Pattern A: cell.value was {text, hyperlink}).
                // For Pattern B (cell.value is a plain URL string AND the cell's
                // xf chain leads to the named "Hyperlink" cellStyle), we
                // construct the same shape of p ourselves so Univer renders
                // the link consistently. extractCellValue's URL-pattern
                // sniffing isn't enough on its own: a cell typed as a plain
                // string value isn't always a link, only when explicitly
                // styled with the named-Hyperlink xf.
                if (
                    !value.p
                    && typeof value.v === 'string'
                    && namedHyperlinkCells.has(cell.address)
                ) {
                    value.p = buildHyperlinkCellP(value.v, value.v);
                }
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
            const assignments = synthesizeTableStyleAssignments(
                t,
                existingCellStyles,
                themeClrScheme?.rgb ?? null,
            );
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

        const defaultColWidthChars = defaultColWidthBySheet.get(ws.id);
        sheets[sheetId] = {
            id: sheetId,
            name: ws.name,
            cellData,
            rowCount: Math.max(100, maxRow + 1),
            columnCount: Math.max(26, maxCol + 1),
            // Convert the source's character-unit default to px with the
            // same w*7+5 approximation columnData uses (keeps the Univer
            // canvas default consistent with explicit per-column widths).
            // Falls back to Univer's 73px default when the source had none.
            defaultColumnWidth: defaultColWidthChars !== undefined
                ? Math.round(defaultColWidthChars * 7 + 5)
                : 73,
            defaultRowHeight: 19,
            mergeData,
            rowData,
            columnData,
            ...(defaultColWidthChars !== undefined ? { defaultColWidthChars } : {}),
        };

        const tablesForSheet = buildTableJsonForSheet(ws, rawTables);
        if (tablesForSheet.length > 0) {
            tableResource[sheetId] = {
                tables: tablesForSheet,
                tableFilteredOutRows: [],
            };
        }

        // M15: read conditional-formatting rules. exceljs surfaces
        // every <conditionalFormatting> block under
        // `ws.conditionalFormattings`. Each block has a `ref` plus a
        // `rules` array; we flatten + translate per-rule into Univer's
        // shape via translateExceljsCfRuleToUniver and accumulate into
        // the per-subUnit resource.
        const wsAny = ws as unknown as { conditionalFormattings?: ExceljsConditionalFormatting[] };
        const cfList = Array.isArray(wsAny.conditionalFormattings) ? wsAny.conditionalFormattings : [];
        if (cfList.length > 0) {
            const ruleEntries: UniverCfRuleEntry[] = [];
            for (const block of cfList) {
                const ranges = cfRefToRanges(block.ref);
                if (ranges.length === 0) continue;
                const rules = Array.isArray(block.rules) ? block.rules : [];
                for (const rule of rules) {
                    const translated = translateExceljsCfRuleToUniver(rule, ranges);
                    if (translated) ruleEntries.push(translated);
                }
            }
            if (ruleEntries.length > 0) cfResource[sheetId] = ruleEntries;
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
    if (Object.keys(cfResource).length > 0) {
        // Univer's CF preset's parseJson does `JSON.parse` on this
        // string and expects `{ [subUnitId]: IConditionFormattingRule[] }`.
        // Emitting the resource even when empty would be harmless but
        // we guard so the snapshot stays minimal.
        resources.push({
            name: CONDITIONAL_FORMATTING_RESOURCE,
            data: JSON.stringify(cfResource),
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

    // M17: chart drawings (read pre-load via readChartsFromXlsxZip). Emit
    // the SHEET_DRAWING_PLUGIN resource in the same shape M10's
    // readChartsFromSnapshot consumes — `componentKey: 'NotesheetChart'`
    // + `data` block + `axisAlignSheetTransform`. subUnitId is the
    // Notesheet-side sheet id (`sheet-<1-based-index>`); we map the
    // importedChart's sheetIndex (xl/worksheets/sheet{N}.xml index) to the
    // matching `sheet-N` here.
    if (importedCharts.length > 0) {
        // Pixel geometry constants — must match the defaults Univer's
        // drawing service uses to compute float-DOM screen positions
        // (DEFAULT_COL_W=73, DEFAULT_ROW_H=19, ROW_HEADER_W=46,
        // COL_HEADER_H=20). Univer recomputes `transform` from
        // `sheetTransform` on load, but supplying it directly lets the
        // drawing service mount the float-DOM at the right position
        // before any layout pass — important for the M17 import path
        // where charts must appear immediately, not after the first
        // resize event.
        const DEFAULT_COL_W = 73;
        const DEFAULT_ROW_H = 19;
        const ROW_HEADER_W = 46;
        const COL_HEADER_H = 20;

        // Helper: when the chart XML didn't ship cached labels/values
        // (older or programmatically-generated workbooks), resolve them
        // from the data sheet's cellData using the chart's sourceRange.
        // Convention mirrors extractDataFromSnapshot: column 0 → labels
        // (skipping the header row), columns 1..N → series.
        const resolveDataFromCells = (
            chart: ImportedChartDrawing,
        ): { labels: string[]; datasets: Array<{ label?: string; data: number[] }> } | null => {
            const sheetName = chart.sourceSheetName;
            if (!sheetName) return null;
            // Find the sheet whose displayed `name` matches sourceSheetName.
            let target: { cellData?: Record<number, Record<number, { v?: unknown }>> } | undefined;
            for (const id of Object.keys(sheets)) {
                if (sheets[id]?.name === sheetName) { target = sheets[id]; break; }
            }
            if (!target?.cellData) return null;
            const sr = chart.sourceRange;
            const headerRow = sr.startRow;
            const dataRowStart = sr.startRow + 1;
            const dataRowEnd = sr.endRow;
            if (dataRowEnd < dataRowStart) return null;
            // Categories: first column.
            const labels: string[] = [];
            for (let r = dataRowStart; r <= dataRowEnd; r++) {
                const v = target.cellData[r]?.[sr.startColumn]?.v;
                labels.push(v == null ? '' : String(v));
            }
            // Series: every column past the first. Tag each with the
            // matching CHART_PALETTE entry so the rendered chart uses
            // Notesheet's recognisable palette (NotesheetChart reads
            // backgroundColor from each dataset). Without this the live
            // Chart.js renderer falls back to its own neutral default
            // and the chart looks generic.
            const datasets: Array<{ label?: string; data: number[]; backgroundColor: string; borderColor: string }> = [];
            for (let c = sr.startColumn + 1; c <= sr.endColumn; c++) {
                const labelV = target.cellData[headerRow]?.[c]?.v;
                const data: number[] = [];
                for (let r = dataRowStart; r <= dataRowEnd; r++) {
                    const v = target.cellData[r]?.[c]?.v;
                    const n = typeof v === 'number' ? v : Number(v);
                    data.push(Number.isFinite(n) ? n : 0);
                }
                const seriesIndex = c - (sr.startColumn + 1);
                const colour = CHART_PALETTE[seriesIndex % CHART_PALETTE.length];
                datasets.push({
                    ...(labelV != null ? { label: String(labelV) } : {}),
                    data,
                    backgroundColor: colour,
                    borderColor: colour,
                });
            }
            return { labels, datasets };
        };

        const drawingResource: Record<string, { data: Record<string, unknown>; order: string[] }> = {};
        for (const chart of importedCharts) {
            const subUnitId = `sheet-${chart.sheetIndex}`;
            // Skip charts whose host sheet didn't survive the import — e.g.
            // a stripped sheet that never reached the eachSheet walk.
            if (!sheets[subUnitId]) continue;

            // Fall back to resolving labels/datasets from the data sheet
            // if the chart XML's <c:cat>/<c:val> caches were empty
            // (programmatic workbooks like MultiSheet.xlsx ship formulas
            // without strCache/numCache).
            //
            // When meta.categoryAxisType === 'index' the source chart had
            // no <c:cat> element at all — Excel shows row index 1..N as
            // the X-axis. Don't synthesize labels from column 0 in that
            // case; that would make column 0's values appear as the
            // X-axis labels (the original 11-stacked-bar bug). Leave
            // labels empty and let NotesheetChart synthesize 1..N at
            // render time.
            let chartLabels = chart.labels;
            let chartDatasets = chart.datasets;
            const isIndexAxis = chart.meta?.categoryAxisType === 'index';
            const cachedLabelsEmpty = chart.labels.length === 0;
            const cachedDataEmpty = chart.datasets.every((ds) => ds.data.length === 0);
            if (cachedDataEmpty) {
                const resolved = resolveDataFromCells(chart);
                if (resolved) chartDatasets = resolved.datasets;
            }
            if (cachedLabelsEmpty && !isIndexAxis) {
                const resolved = resolveDataFromCells(chart);
                if (resolved) chartLabels = resolved.labels;
            }
            if (!drawingResource[subUnitId]) {
                drawingResource[subUnitId] = { data: {}, order: [] };
            }
            const drawingId = chart.chartId;
            const left = ROW_HEADER_W + chart.anchor.fromCol * DEFAULT_COL_W;
            const top = COL_HEADER_H + chart.anchor.fromRow * DEFAULT_ROW_H;
            const right = ROW_HEADER_W + chart.anchor.toCol * DEFAULT_COL_W;
            const bottom = COL_HEADER_H + chart.anchor.toRow * DEFAULT_ROW_H;
            drawingResource[subUnitId].data[drawingId] = {
                unitId: 'workbook',
                subUnitId,
                drawingId,
                // DRAWING_DOM (8). Univer's drawing-DOM service mounts
                // float-DOM components keyed by `componentKey`; setting
                // this to anything else (we used to emit 5 = VIDEO)
                // makes Univer's render pipeline treat the entry as an
                // unsupported drawing and skip it.
                drawingType: 8,
                componentKey: 'NotesheetChart',
                allowTransform: true,
                data: {
                    chartId: chart.chartId,
                    type: chart.type,
                    title: chart.title,
                    sourceRange: chart.sourceRange,
                    ...(chart.sourceSheetName ? { sourceSheetName: chart.sourceSheetName } : {}),
                    labels: chartLabels,
                    datasets: chartDatasets,
                    ...(chart.meta ? { meta: chart.meta } : {}),
                },
                transform: {
                    flipY: false,
                    flipX: false,
                    angle: 0,
                    skewX: 0,
                    skewY: 0,
                    left,
                    top,
                    width: Math.max(50, right - left),
                    height: Math.max(50, bottom - top),
                },
                // sheetTransform's columnOffset/rowOffset are PIXELS in
                // Univer's drawing service, NOT EMUs. The OOXML
                // <xdr:colOff>/<xdr:rowOff> elements ship EMUs (English
                // Metric Units; 9525 EMU = 1 px at 96 DPI). Forwarding
                // raw EMUs here makes Univer recompute transform with
                // nonsense pixel values (~370k px wide, off-screen
                // negative left), which is what produced the
                // "covers the whole sheet" + "blank canvas" symptoms in
                // 06 / 10 — the float-DOM mounted with a transform
                // Chart.js's `responsive` couldn't size against until
                // a manual resize forced a recompute. We convert EMU
                // to pixels here so Univer's drawing service produces
                // the same transform our `transform` block already had.
                sheetTransform: {
                    from: {
                        column: chart.anchor.fromCol,
                        columnOffset: Math.round(chart.anchor.fromColOff / 9525),
                        row: chart.anchor.fromRow,
                        rowOffset: Math.round(chart.anchor.fromRowOff / 9525),
                    },
                    to: {
                        column: chart.anchor.toCol,
                        columnOffset: Math.round(chart.anchor.toColOff / 9525),
                        row: chart.anchor.toRow,
                        rowOffset: Math.round(chart.anchor.toRowOff / 9525),
                    },
                },
                axisAlignSheetTransform: {
                    from: {
                        column: chart.anchor.fromCol,
                        columnOffset: Math.round(chart.anchor.fromColOff / 9525),
                        row: chart.anchor.fromRow,
                        rowOffset: Math.round(chart.anchor.fromRowOff / 9525),
                    },
                    to: {
                        column: chart.anchor.toCol,
                        columnOffset: Math.round(chart.anchor.toColOff / 9525),
                        row: chart.anchor.toRow,
                        rowOffset: Math.round(chart.anchor.toRowOff / 9525),
                    },
                },
            };
            drawingResource[subUnitId].order.push(drawingId);
        }
        if (Object.keys(drawingResource).length > 0) {
            resources.push({
                name: 'SHEET_DRAWING_PLUGIN',
                data: JSON.stringify(drawingResource),
            });
        }
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

// Convert a Univer ITextStyle (the `ts` field of an ITextRun) into an
// exceljs run-level Font. Inverse of buildTextStyleFromExceljsFont.
// Used on export when a multi-run cell.p needs to come out as
// `cell.value = { richText: [{font, text}, ...] }` so Excel renders
// the per-run formatting.
function buildExceljsFontFromTextStyle(ts: Record<string, unknown> | undefined): Partial<ExcelJS.Font> {
    if (!ts) return {};
    const font: Partial<ExcelJS.Font> = {};
    if (typeof ts.ff === 'string') font.name = ts.ff;
    if (typeof ts.fs === 'number') font.size = ts.fs;
    if (ts.bl === 1) font.bold = true;
    if (ts.it === 1) font.italic = true;
    if ((ts.ul as { s?: number } | undefined)?.s === 1) font.underline = true;
    if ((ts.st as { s?: number } | undefined)?.s === 1) font.strike = true;
    const cl = ts.cl as { rgb?: string } | undefined;
    if (cl?.rgb) {
        const argb = hexToArgb(cl.rgb);
        if (argb) font.color = { argb };
    }
    return font;
}

// Extract multi-run rich-text from a Univer cell.p. Returns null when:
//   - cell.p is missing or has < 2 text runs (single-run cells round-trip
//     as plain strings; the M13 import path collapses 1-element richText
//     to plain text and we don't want export to re-promote).
//   - cell.p carries a hyperlink customRange (the hyperlink Pattern A
//     emission wins; multi-run hyperlink text is M13a territory).
//
// The returned array is in exceljs's RichText shape: each element has a
// `text` field and an optional `font` carrying the run's formatting.
function extractRichTextRunsFromCellP(p: unknown): Array<{ text: string; font?: Partial<ExcelJS.Font> }> | null {
    if (!p || typeof p !== 'object') return null;
    const body = (p as { body?: {
        dataStream?: string;
        textRuns?: Array<{ st: number; ed: number; ts?: Record<string, unknown> }>;
        customRanges?: Array<{ rangeType?: number }>;
    } }).body;
    if (!body) return null;
    // Skip when a hyperlink customRange is present — Pattern A handles it.
    const ranges = body.customRanges;
    if (Array.isArray(ranges) && ranges.some((r) => r?.rangeType === 0)) return null;
    const runs = body.textRuns;
    if (!Array.isArray(runs) || runs.length < 2) return null;
    const stream = body.dataStream;
    if (typeof stream !== 'string') return null;
    // The dataStream ends with '\r\n' (paragraph mark + section break);
    // textRun offsets address the text BEFORE that terminator, so it's
    // safe to slice each run's range without trimming the dataStream.
    const out: Array<{ text: string; font?: Partial<ExcelJS.Font> }> = [];
    for (const r of runs) {
        if (typeof r.st !== 'number' || typeof r.ed !== 'number') continue;
        const text = stream.slice(r.st, r.ed);
        const font = buildExceljsFontFromTextStyle(r.ts);
        if (Object.keys(font).length > 0) out.push({ text, font });
        else out.push({ text });
    }
    return out.length >= 2 ? out : null;
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
    // Text rotation. Reverse of the import-side mapping: Univer's
    // ITextRotation { a, v } → exceljs's textRotation (number | 'vertical').
    // Stacked mode (v=1) wins over the angle.
    const tr = style.tr as { a?: number; v?: number } | undefined;
    if (tr) {
        if (tr.v === 1) align.textRotation = 'vertical';
        else if (typeof tr.a === 'number' && tr.a !== 0) align.textRotation = tr.a;
    }
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

// Read the CF resource (M15). Returns a `${sheetId}` → CF rule entries
// map; the exporter walks each entry, translates Univer-shape rules
// back to exceljs CF rule objects, and assigns the array to the
// matching worksheet's `conditionalFormattings` field before
// `writeBuffer()`.
function readCfResource(snapshot: UniverSnapshot): Record<string, UniverCfRuleEntry[]> {
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> }).resources;
    if (!Array.isArray(resources)) return {};
    const entry = resources.find((r) => r?.name === CONDITIONAL_FORMATTING_RESOURCE);
    if (!entry || typeof entry.data !== 'string') return {};
    try {
        const parsed = JSON.parse(entry.data);
        if (!parsed || typeof parsed !== 'object') return {};
        return parsed as Record<string, UniverCfRuleEntry[]>;
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
    const cfBySheet = readCfResource(snapshot);
    const defaultFontName = readDefaultFontName(snapshot);

    for (const sheetId of sheetOrder) {
        const sheet = sheets[sheetId];
        if (!sheet) continue;
        const ws = wb.addWorksheet(sheet.name || sheetId);

        // Re-emit the source's default column width (issue 11). exceljs
        // serializes ws.properties.defaultColWidth into <sheetFormatPr>;
        // columns without an explicit width inherit it instead of
        // collapsing to exceljs's 8.43 default.
        if (typeof sheet.defaultColWidthChars === 'number' && sheet.defaultColWidthChars > 0) {
            (ws.properties as unknown as { defaultColWidth?: number }).defaultColWidth = sheet.defaultColWidthChars;
        }

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
                const richTextRuns = hyperlinkUrl ? null : extractRichTextRunsFromCellP(data.p);
                if (data.f) {
                    const formula = data.f.startsWith('=') ? data.f.slice(1) : data.f;
                    // Drop formulas that reference external workbooks
                    // (`[N]!` prefix or `[filename.xlsx]Sheet!` token).
                    // Notesheet doesn't preserve the source workbook's
                    // `xl/externalLinks/*` parts on round-trip, so emitting
                    // these formulas back makes Excel's open-time validator
                    // strip them and warn the user ("Removed Records:
                    // Formula from /xl/worksheets/sheet1.xml part"). Better
                    // to keep just the cached value — Excel opens the file
                    // cleanly and the cell shows the correct number.
                    // External-link round-trip is a separate cycle.
                    const hasExternalRef = /\[\d+\]!|\[[^\]]+\.xlsx?\]/.test(formula);
                    if (hasExternalRef) {
                        if (data.v !== undefined && data.v !== null) {
                            cell.value = data.v;
                        }
                    } else {
                        const result = data.v;
                        cell.value = { formula, result } as ExcelJS.CellFormulaValue;
                    }
                } else if (hyperlinkUrl && data.v !== undefined && data.v !== null) {
                    // Hyperlink-bearing cells use exceljs's { text, hyperlink }
                    // shape — exceljs writes the proper <hyperlinks> block
                    // and the sheet rels entry pointing to the URL.
                    cell.value = { text: String(data.v), hyperlink: hyperlinkUrl };
                } else if (richTextRuns) {
                    // Multi-run cell with no hyperlink: emit exceljs's
                    // RichText shape so Excel renders the per-run
                    // formatting (bold word + plain word in one cell, etc.).
                    cell.value = { richText: richTextRuns } as unknown as ExcelJS.CellValue;
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

        // M15: write conditional-formatting rules. Univer's snapshot
        // groups rules by subUnitId (sheetId). We translate each rule
        // back to exceljs's CF shape and group by sqref so all rules
        // sharing a ref end up under one block. Priority is preserved
        // by re-numbering 1..N in the order we walk the snapshot.
        const cfEntries = cfBySheet[sheetId];
        if (Array.isArray(cfEntries) && cfEntries.length > 0) {
            const blocksByRef = new Map<string, ExceljsCfRule[]>();
            let priority = 1;
            for (const entry of cfEntries) {
                const translated = translateUniverCfRuleToExceljs(entry, priority);
                if (!translated || !translated.rule) continue;
                const list = blocksByRef.get(translated.sqref) ?? [];
                list.push(translated.rule);
                blocksByRef.set(translated.sqref, list);
                priority++;
            }
            const conditionalFormattings = Array.from(blocksByRef.entries()).map(([ref, rules]) => ({
                ref,
                rules,
            }));
            // exceljs assigns this directly onto ws; the CF blocks are
            // serialized into <conditionalFormatting> elements at
            // writeBuffer time.
            (ws as unknown as { conditionalFormattings: typeof conditionalFormattings }).conditionalFormattings = conditionalFormattings;
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
