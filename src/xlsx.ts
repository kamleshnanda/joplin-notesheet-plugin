// xlsx ↔ Univer snapshot converters.
//
// exceljs is the underlying engine. Two pure functions are exposed:
//   xlsxBufferToSnapshot(buffer) → IWorkbookData-shaped snapshot
//   snapshotToXlsxBuffer(snapshot) → Promise<ArrayBuffer>
//
// Coverage targets for M5: cell values (string/number/boolean), formulas,
// font (family/size/bold/italic/underline/color), fill (background), alignment
// (horizontal/vertical/wrap), number format, and merged cells. Borders, charts,
// pivots, named tables are out of scope for now.
//
// Universe of enums we care about (mirrored as numeric literals so we don't
// pull @univerjs/core into Jest's node-environment unit tests):
//   HorizontalAlign:  LEFT=1, CENTER=2, RIGHT=3
//   VerticalAlign:    TOP=1, MIDDLE=2, BOTTOM=3
//   WrapStrategy:     OVERFLOW=1, CLIP=2, WRAP=3
//   BooleanNumber:    FALSE=0, TRUE=1
//   CellValueType:    STRING=1, NUMBER=2, BOOLEAN=3

import ExcelJS from 'exceljs';

import type { UniverSnapshot } from './snapshot';

const HORIZONTAL = { left: 1, center: 2, right: 3 } as const;
const VERTICAL = { top: 1, middle: 2, bottom: 3 } as const;
const WRAP_STRATEGY_WRAP = 3;
const WRAP_STRATEGY_CLIP = 2;
const VALUE_STRING = 1;
const VALUE_NUMBER = 2;
const VALUE_BOOLEAN = 3;

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

    return Object.keys(style).length > 0 ? style : null;
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

    // Date — serialize as ISO string for predictability. Numfmt will still render.
    if (raw instanceof Date) {
        return { v: raw.toISOString(), t: VALUE_STRING };
    }

    if (typeof raw === 'number') return { v: raw, t: VALUE_NUMBER };
    if (typeof raw === 'boolean') return { v: raw, t: VALUE_BOOLEAN };
    if (typeof raw === 'string') return { v: raw, t: VALUE_STRING };

    // Error cell or other complex types: render the displayed text.
    const text = cell.text;
    if (text) return { v: text, t: VALUE_STRING };
    return {};
}

export async function xlsxBufferToSnapshot(buffer: ArrayBuffer | Uint8Array | Buffer): Promise<UniverSnapshot> {
    const wb = new ExcelJS.Workbook();
    // exceljs accepts Buffer in Node and ArrayBuffer in the browser; both are valid
    // at runtime but the .d.ts only types Buffer. Cast away to satisfy TS.
    await wb.xlsx.load(buffer as unknown as Parameters<typeof wb.xlsx.load>[0]);

    const sheetOrder: string[] = [];
    const sheets: Record<string, SheetRecord> = {};
    const styles: Record<string, Record<string, unknown>> = {};
    const styleIdByKey = new Map<string, string>();
    let nextStyleId = 1;

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

    return {
        id: 'workbook-' + Date.now(),
        sheetOrder,
        name: 'Spreadsheet',
        appVersion: '0.1.0',
        locale: 'enUS',
        styles,
        sheets,
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
}

export async function snapshotToXlsxBuffer(snapshot: UniverSnapshot): Promise<ArrayBuffer> {
    const wb = new ExcelJS.Workbook();
    const sheetOrder = (snapshot as { sheetOrder?: string[] }).sheetOrder ?? [];
    const sheets = (snapshot as { sheets?: Record<string, SheetRecord> }).sheets ?? {};

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
    }

    if (wb.worksheets.length === 0) {
        wb.addWorksheet('Sheet1');
    }

    const buffer = await wb.xlsx.writeBuffer();
    return buffer as ArrayBuffer;
}
