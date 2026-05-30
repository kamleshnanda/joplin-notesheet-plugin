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

interface CellRecord {
    v?: string | number | boolean;
    f?: string;
    t?: number;
    s?: string;
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
function styleKey(style: Record<string, unknown>): string {
    return JSON.stringify(style, Object.keys(style).sort());
}

function buildStyleFromExcelCell(cell: ExcelJS.Cell): Record<string, unknown> | null {
    const style: Record<string, unknown> = {};

    const font = cell.font;
    if (font) {
        if (font.name) style.ff = font.name;
        if (typeof font.size === 'number') style.fs = font.size;
        if (font.bold) style.bl = 1;
        if (font.italic) style.it = 1;
        if (font.underline) style.ul = { s: 1 };
        if (font.strike) style.st = { s: 1 };
        const fontColor = argbToHex(font.color?.argb);
        if (fontColor) style.cl = { rgb: fontColor };
    }

    const fill = cell.fill;
    if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
        const bgColor = argbToHex(fill.fgColor?.argb);
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
            const argb = (side.color as { argb?: string } | undefined)?.argb;
            const rgb = argbToHex(argb) ?? '#000000';
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

function extractCellValue(cell: ExcelJS.Cell): { v?: string | number | boolean; f?: string; t?: number } {
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

    // Hyperlink cell: keep the visible text, drop the URL for now (M5 scope).
    if (typeof raw === 'object' && 'text' in raw && 'hyperlink' in raw) {
        return { v: String((raw as { text: unknown }).text ?? ''), t: VALUE_STRING };
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

    const sheetOrder: string[] = [];
    const sheets: Record<string, SheetRecord> = {};
    const styles: Record<string, Record<string, unknown>> = {};
    const styleIdByKey = new Map<string, string>();
    let nextStyleId = 1;
    // Per-subUnit table state, keyed by our generated sheetId. Filled during
    // the eachSheet walk and serialized into the SHEET_TABLE_PLUGIN resource
    // at the end so Univer's formula engine sees the tables on snapshot load.
    const tableResource: Record<string, { tables: TableJson[]; tableFilteredOutRows: number[] }> = {};

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
                const style = buildStyleFromExcelCell(cell);
                const styleId = internStyle(style);
                const record: CellRecord = {};
                if (value.v !== undefined) record.v = value.v;
                if (value.f) record.f = value.f;
                if (value.t !== undefined) record.t = value.t;
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

        // ws.id is the 1-based worksheet index in xl/worksheets/sheetN.xml,
        // which is exactly what readTablesFromXlsxZip keys by.
        const rawTables = rawTablesByIndex.get(ws.id) ?? [];
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

function resolveStyle(snapshot: UniverSnapshot, ref: unknown): Record<string, unknown> | null {
    if (!ref) return null;
    if (typeof ref === 'string') {
        const styles = (snapshot as { styles?: Record<string, Record<string, unknown>> }).styles;
        return styles?.[ref] ?? null;
    }
    if (typeof ref === 'object') return ref as Record<string, unknown>;
    return null;
}

function applyStyleToCell(cell: ExcelJS.Cell, style: Record<string, unknown>): void {
    const font: Partial<ExcelJS.Font> = {};
    if (typeof style.ff === 'string') font.name = style.ff;
    if (typeof style.fs === 'number') font.size = style.fs;
    if (style.bl === 1) font.bold = true;
    if (style.it === 1) font.italic = true;
    if (style.ul && (style.ul as { s?: number }).s === 1) font.underline = true;
    if (style.st && (style.st as { s?: number }).s === 1) font.strike = true;
    const cl = style.cl as { rgb?: string } | undefined;
    if (cl?.rgb) {
        const argb = hexToArgb(cl.rgb);
        if (argb) font.color = { argb };
    }
    if (Object.keys(font).length > 0) cell.font = font;

    const bg = style.bg as { rgb?: string } | undefined;
    if (bg?.rgb) {
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

export async function snapshotToXlsxBuffer(snapshot: UniverSnapshot): Promise<ArrayBuffer> {
    const wb = new ExcelJS.Workbook();
    const sheetOrder = (snapshot as { sheetOrder?: string[] }).sheetOrder ?? [];
    const sheets = (snapshot as { sheets?: Record<string, SheetRecord> }).sheets ?? {};
    const tableResource = readTableResource(snapshot);

    for (const sheetId of sheetOrder) {
        const sheet = sheets[sheetId];
        if (!sheet) continue;
        const ws = wb.addWorksheet(sheet.name || sheetId);

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
                if (data.f) {
                    const formula = data.f.startsWith('=') ? data.f.slice(1) : data.f;
                    const result = data.v;
                    cell.value = { formula, result } as ExcelJS.CellFormulaValue;
                } else if (data.v !== undefined && data.v !== null) {
                    cell.value = data.v;
                }
                const style = resolveStyle(snapshot, data.s);
                if (style) applyStyleToCell(cell, style);
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
    return buffer as ArrayBuffer;
}
