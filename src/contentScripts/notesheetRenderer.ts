// Notesheet — Markdown-It content-script renderer.
//
// Joplin invokes this content script via `joplin.contentScripts.register(
// ContentScriptType.MarkdownItPlugin, ...)`. It runs inside Joplin's
// renderer worker (the same pipeline that powers the editor preview pane,
// PDF export, and HTML export), NOT in the plugin's main process.
//
// The renderer overrides markdown-it's fence handler. When it sees a
// fenced code block tagged `notesheet v=1`, it parses the JSON inside,
// walks the snapshot's `sheetOrder` / `sheets` / `styles` / `mergeData`,
// and emits an HTML `<table>` per sheet with inline-styled `<td>`s. Any
// other fence falls through to markdown-it's default rendering — we MUST
// not break non-Notesheet notes.
//
// Why we duplicate the fence-parsing logic instead of importing
// `extractSnapshot()` from `src/snapshot.ts`: webpack bundles content
// scripts as standalone files via the `extraScripts` config, but pulling
// in `src/snapshot.ts` would also pull in any of its transitive deps;
// we keep the renderer slim by inlining the small amount of logic we
// need. (The renderer doesn't need import/export helpers from
// `src/xlsx.ts`, which would drag exceljs + JSZip into a renderer
// bundle that runs on every note.)
//
// Conditional formatting: Univer's CF preset paints CF colours at
// canvas-render time. The static HTML renderer doesn't run inside
// Univer; it must evaluate CF rules itself and bake the resulting
// fills into per-cell inline styles. M16 covers cellIs, top10, and
// colorScale; dataBar and iconSet are documented as M16-followup.
//
// CRITICAL invariants:
//   1. The fence-tag check is the FIRST thing — if the tag isn't
//      `notesheet`, return undefined / fall through. Don't break
//      non-Notesheet notes.
//   2. The renderer is read-only on the snapshot. It does NOT modify
//      the snapshot, the note, or any plugin state.
//   3. Output is HTML only. No JavaScript injection.
//   4. Untrusted input — cell values come from .xlsx files of unknown
//      provenance. Every value we interpolate into HTML goes through
//      `escapeHtml()`.

// Markdown-It is supplied by Joplin at runtime; we type it loosely.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MarkdownIt = any;
type FenceToken = {
    info?: string;
    content?: string;
};

// Snapshot shape we walk. Loose typing because Univer's IWorkbookData
// has many optional fields we don't need. We keep it permissive so a
// future schema bump doesn't force a renderer update.
interface SnapshotSheet {
    id: string;
    name: string;
    cellData?: Record<string, Record<string, SnapshotCell>>;
    rowCount?: number;
    columnCount?: number;
    mergeData?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
}
interface SnapshotCell {
    v?: string | number | boolean;
    t?: number;
    s?: string | Record<string, unknown>;
    f?: string;
    p?: { body?: { dataStream?: string } };
}
interface Snapshot {
    sheetOrder?: string[];
    sheets?: Record<string, SnapshotSheet>;
    styles?: Record<string, Record<string, unknown>>;
    resources?: Array<{ name?: string; data?: string }>;
}

// ───────── Fence detection ──────────────────────────────────────────

// Match the same shape `src/snapshot.ts` uses. The fence body is JSON.
// We accept `notesheet v=1` (with optional whitespace), and only `v=1`.
const FENCE_TAG = 'notesheet';

function parseFenceInfo(info: string | undefined): { isNotesheet: boolean; version: number } {
    if (!info) return { isNotesheet: false, version: 0 };
    const trimmed = info.trim();
    if (!trimmed.startsWith(FENCE_TAG)) return { isNotesheet: false, version: 0 };
    const rest = trimmed.slice(FENCE_TAG.length).trim();
    const m = /^v=(\d+)/.exec(rest);
    if (!m) return { isNotesheet: false, version: 0 };
    return { isNotesheet: true, version: Number(m[1]) };
}

function parseSnapshotJson(content: string | undefined): Snapshot | null {
    if (!content) return null;
    try {
        const parsed = JSON.parse(content);
        if (!parsed || typeof parsed !== 'object') return null;
        return parsed as Snapshot;
    } catch {
        return null;
    }
}

// ───────── HTML escaping ────────────────────────────────────────────

const HTML_ESCAPE_RE = /[&<>"']/g;
const HTML_ESCAPES: Record<string, string> = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
};
function escapeHtml(s: string): string {
    return s.replace(HTML_ESCAPE_RE, (c) => HTML_ESCAPES[c]);
}

// ───────── Style → inline CSS ──────────────────────────────────────

// Univer numeric enums for alignment, mirrored from src/xlsx.ts. The
// renderer doesn't depend on @univerjs/core at runtime.
const HORIZONTAL_TO_CSS: Record<number, string> = { 1: 'left', 2: 'center', 3: 'right' };
const VERTICAL_TO_CSS: Record<number, string> = { 1: 'top', 2: 'middle', 3: 'bottom' };

// Univer BorderStyleTypes → CSS border-style. Numeric enum matches
// `BORDER_STYLE_TO_UNIVER` in src/xlsx.ts (mirrored).
const BORDER_STYLE_NUM_TO_CSS: Record<number, string> = {
    1: '1px solid',  // thin
    2: '1px dotted', // hair (approximation)
    3: '1px dotted',
    4: '1px dashed',
    5: '1px dashed',
    6: '1px dashed',
    7: '3px double',
    8: '2px solid',  // medium
    9: '2px dashed',
    10: '2px dashed',
    11: '2px dashed',
    12: '2px dashed',
    13: '3px solid', // thick
};

interface ResolvedStyle {
    bg?: { rgb: string };
    cl?: { rgb: string };
    bl?: number;
    it?: number;
    ul?: { s?: number };
    st?: { s?: number };
    ht?: number;
    vt?: number;
    bd?: {
        t?: { s?: number; cl?: { rgb?: string } };
        r?: { s?: number; cl?: { rgb?: string } };
        b?: { s?: number; cl?: { rgb?: string } };
        l?: { s?: number; cl?: { rgb?: string } };
    };
}

function resolveCellStyle(
    cell: SnapshotCell,
    styles: Record<string, Record<string, unknown>> | undefined,
): ResolvedStyle {
    if (!cell || cell.s === undefined || cell.s === null) return {};
    if (typeof cell.s === 'string') {
        const ref = styles?.[cell.s];
        return (ref ?? {}) as ResolvedStyle;
    }
    if (typeof cell.s === 'object') return cell.s as ResolvedStyle;
    return {};
}

function buildCellInlineStyle(
    base: ResolvedStyle,
    cfFill: string | null,
): string {
    const parts: string[] = [];
    // Background: CF fill (if any) wins over the base style. CF rules
    // explicitly override per-cell formatting in Excel's render order;
    // the M16 renderer mirrors that.
    if (cfFill) {
        parts.push(`background-color: ${cfFill}`);
    } else if (base.bg && base.bg.rgb) {
        parts.push(`background-color: ${base.bg.rgb}`);
    }
    if (base.cl && base.cl.rgb) parts.push(`color: ${base.cl.rgb}`);
    if (base.bl === 1) parts.push('font-weight: bold');
    if (base.it === 1) parts.push('font-style: italic');
    const decorations: string[] = [];
    if (base.ul && base.ul.s === 1) decorations.push('underline');
    if (base.st && base.st.s === 1) decorations.push('line-through');
    if (decorations.length > 0) parts.push(`text-decoration: ${decorations.join(' ')}`);
    if (base.ht !== undefined && HORIZONTAL_TO_CSS[base.ht]) {
        parts.push(`text-align: ${HORIZONTAL_TO_CSS[base.ht]}`);
    }
    if (base.vt !== undefined && VERTICAL_TO_CSS[base.vt]) {
        parts.push(`vertical-align: ${VERTICAL_TO_CSS[base.vt]}`);
    }
    if (base.bd) {
        const sides: Array<['t', 'top'] | ['r', 'right'] | ['b', 'bottom'] | ['l', 'left']> = [
            ['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left'],
        ];
        for (const [k, css] of sides) {
            const side = base.bd[k];
            if (!side || typeof side.s !== 'number') continue;
            const styleRule = BORDER_STYLE_NUM_TO_CSS[side.s] ?? '1px solid';
            const colour = (side.cl && side.cl.rgb) ? side.cl.rgb : '#000000';
            parts.push(`border-${css}: ${styleRule} ${colour}`);
        }
    }
    return parts.join('; ');
}

// ───────── Cell value rendering ─────────────────────────────────────

// Surface the cached value as text. Univer rich-text bodies (cell.p)
// carry a `dataStream` that contains the concatenated text plus a
// trailing `\r\n`; we strip that for display. Per-run formatting is
// out of scope per the M16 plan — we render the plain text.
function renderCellValue(cell: SnapshotCell): string {
    if (cell == null) return '';
    if (cell.p && cell.p.body && typeof cell.p.body.dataStream === 'string') {
        // Univer's dataStream uses \r\n as a logical paragraph break.
        // For HTML, replace with <br/>; the value is escaped first so a
        // cell value that legitimately contains "<br>" stays literal.
        const ds = cell.p.body.dataStream.replace(/\r?\n$/, '');
        return escapeHtml(ds).replace(/\r\n/g, '<br/>').replace(/\r/g, '<br/>').replace(/\n/g, '<br/>');
    }
    if (cell.v === undefined || cell.v === null) return '';
    return escapeHtml(String(cell.v));
}

// ───────── Conditional formatting evaluation ───────────────────────

// CF resource shape (matches CONDITIONAL_FORMATTING_RESOURCE in
// src/xlsx.ts). data is JSON-stringified `{ [subUnitId]: rules[] }`.
const CONDITIONAL_FORMATTING_RESOURCE = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN';

interface CfRule {
    cfId?: string;
    ranges?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
    stopIfTrue?: boolean;
    rule?: {
        type?: string;
        subType?: string;
        operator?: string;
        value?: number | string;
        isBottom?: boolean;
        isPercent?: boolean;
        config?: unknown;
        style?: { bg?: { rgb?: string } };
    };
}

function loadCfRulesFromSnapshot(snapshot: Snapshot): Record<string, CfRule[]> {
    const out: Record<string, CfRule[]> = {};
    if (!Array.isArray(snapshot.resources)) return out;
    const entry = snapshot.resources.find((r) => r && r.name === CONDITIONAL_FORMATTING_RESOURCE);
    if (!entry || typeof entry.data !== 'string') return out;
    try {
        const parsed = JSON.parse(entry.data);
        if (!parsed || typeof parsed !== 'object') return out;
        for (const subUnit of Object.keys(parsed)) {
            const list = (parsed as Record<string, unknown>)[subUnit];
            if (Array.isArray(list)) {
                out[subUnit] = list.filter((r) => !!r) as CfRule[];
            }
        }
    } catch {
        // malformed CF resource — skip silently
    }
    return out;
}

// Linear interpolation between two RGB colours, weighted by `t` in [0,1].
function lerpRgb(a: string, b: string, t: number): string {
    const ah = a.replace('#', '');
    const bh = b.replace('#', '');
    if (ah.length !== 6 || bh.length !== 6) return a;
    const ar = parseInt(ah.slice(0, 2), 16), ag = parseInt(ah.slice(2, 4), 16), ab = parseInt(ah.slice(4, 6), 16);
    const br = parseInt(bh.slice(0, 2), 16), bg = parseInt(bh.slice(2, 4), 16), bb = parseInt(bh.slice(4, 6), 16);
    const cr = Math.round(ar + (br - ar) * t);
    const cg = Math.round(ag + (bg - ag) * t);
    const cb = Math.round(ab + (bb - ab) * t);
    const hex = (n: number) => n.toString(16).padStart(2, '0').toUpperCase();
    return '#' + hex(cr) + hex(cg) + hex(cb);
}

// Resolve a CF value-config (cfvo) to a numeric value over the rule's
// range. Used by colorScale interpolation.
function resolveCfvo(
    config: { type?: string; value?: number | string },
    sortedValues: number[],
): number | null {
    if (!config) return null;
    const type = config.type ?? 'num';
    if (type === 'min') return sortedValues.length > 0 ? sortedValues[0] : 0;
    if (type === 'max') return sortedValues.length > 0 ? sortedValues[sortedValues.length - 1] : 0;
    if (type === 'percentile' || type === 'percent') {
        // Both treat `value` as 0..100 — for percentile we look up the
        // sorted value at that rank, for percent we interpolate min..max.
        const v = typeof config.value === 'number' ? config.value : Number(config.value);
        if (!Number.isFinite(v) || sortedValues.length === 0) return null;
        if (type === 'percent') {
            const lo = sortedValues[0];
            const hi = sortedValues[sortedValues.length - 1];
            return lo + (hi - lo) * (v / 100);
        }
        // percentile
        const idx = Math.min(sortedValues.length - 1, Math.max(0, Math.round(((v / 100) * (sortedValues.length - 1)))));
        return sortedValues[idx];
    }
    if (type === 'num' || type === 'number') {
        const v = typeof config.value === 'number' ? config.value : Number(config.value);
        return Number.isFinite(v) ? v : null;
    }
    // 'formula' — out of scope; return null so the colorScale falls
    // through to the next anchor.
    return null;
}

// Evaluate a colorScale rule for a single cell value. The rule's
// `config` is `[{index, color, value: cfvo}, ...]` (M15 import shape).
function evalColorScale(
    rule: CfRule,
    value: number,
    sortedValues: number[],
): string | null {
    const cfg = rule.rule?.config;
    if (!Array.isArray(cfg) || cfg.length < 2) return null;
    const anchors: Array<{ pos: number; colour: string }> = [];
    for (const a of cfg as Array<{ color?: string; value?: { type?: string; value?: number | string } }>) {
        const pos = a.value ? resolveCfvo(a.value, sortedValues) : null;
        if (pos === null || !Number.isFinite(pos) || !a.color) continue;
        anchors.push({ pos, colour: a.color });
    }
    if (anchors.length < 2) return null;
    anchors.sort((p, q) => p.pos - q.pos);
    if (value <= anchors[0].pos) return anchors[0].colour;
    const last = anchors[anchors.length - 1];
    if (value >= last.pos) return last.colour;
    for (let i = 0; i < anchors.length - 1; i++) {
        const lo = anchors[i], hi = anchors[i + 1];
        if (value >= lo.pos && value <= hi.pos) {
            const span = hi.pos - lo.pos;
            const t = span === 0 ? 0 : (value - lo.pos) / span;
            return lerpRgb(lo.colour, hi.colour, t);
        }
    }
    return null;
}

function compareForCellIs(op: string | undefined, cellValue: number, target: number): boolean {
    switch (op) {
        case 'greaterThan': return cellValue > target;
        case 'greaterThanOrEqual': return cellValue >= target;
        case 'lessThan': return cellValue < target;
        case 'lessThanOrEqual': return cellValue <= target;
        case 'equal': return cellValue === target;
        case 'notEqual': return cellValue !== target;
        case 'between':
            // 'between' uses two formulae; M16 evaluator only sees the
            // first since the import-side translator stores `formulae[0]`
            // in `value`. Treat as no-match for now (M16-followup).
            return false;
        default:
            return cellValue > target;
    }
}

// Collect all numeric values inside a rule's range from the sheet's
// cellData. Used as the input set for colorScale anchors and top10
// rank calculation.
function collectRangeValues(
    sheet: SnapshotSheet,
    rule: CfRule,
): { byRC: Map<string, number>; sorted: number[] } {
    const byRC = new Map<string, number>();
    if (!Array.isArray(rule.ranges)) return { byRC, sorted: [] };
    const cellData = sheet.cellData ?? {};
    for (const range of rule.ranges) {
        for (let r = range.startRow; r <= range.endRow; r++) {
            const rowMap = cellData[String(r)] ?? cellData[r as unknown as string];
            if (!rowMap) continue;
            for (let c = range.startColumn; c <= range.endColumn; c++) {
                const cell = rowMap[String(c)] ?? rowMap[c as unknown as string];
                if (!cell || cell.v === undefined || cell.v === null) continue;
                const v = typeof cell.v === 'number' ? cell.v : Number(cell.v);
                if (!Number.isFinite(v)) continue;
                byRC.set(`${r}:${c}`, v);
            }
        }
    }
    const sorted = Array.from(byRC.values()).sort((a, b) => a - b);
    return { byRC, sorted };
}

// Apply CF rules over a sheet, producing per-cell fill overrides.
// Returns a Map keyed `r:c` to the CSS hex colour to apply.
function evaluateCfFor(sheet: SnapshotSheet, rules: CfRule[]): Map<string, string> {
    const fills = new Map<string, string>();
    if (!Array.isArray(rules) || rules.length === 0) return fills;

    // Pre-compute ranked sets for top10 rules and value sets for
    // colorScale rules. We fold both into one pass per rule.
    for (const rule of rules) {
        const r = rule.rule;
        if (!r) continue;
        switch (r.type) {
            case 'highlightCell': {
                if (r.subType === 'number') {
                    const target = typeof r.value === 'number' ? r.value : Number(r.value);
                    if (!Number.isFinite(target)) break;
                    const fill = r.style?.bg?.rgb;
                    if (!fill) break;
                    const { byRC } = collectRangeValues(sheet, rule);
                    for (const [key, v] of byRC) {
                        if (compareForCellIs(r.operator, v, target)) {
                            fills.set(key, fill);
                        }
                    }
                } else if (r.subType === 'rank') {
                    const fill = r.style?.bg?.rgb;
                    if (!fill) break;
                    const n = typeof r.value === 'number' ? r.value : Number(r.value);
                    if (!Number.isFinite(n) || n <= 0) break;
                    const { byRC, sorted } = collectRangeValues(sheet, rule);
                    if (sorted.length === 0) break;
                    // Top-N: highest n values; bottom: lowest n.
                    const isBottom = !!r.isBottom;
                    const isPercent = !!r.isPercent;
                    let k: number;
                    if (isPercent) {
                        k = Math.max(1, Math.floor((sorted.length * n) / 100));
                    } else {
                        k = Math.max(1, Math.floor(n));
                    }
                    k = Math.min(k, sorted.length);
                    const threshold = isBottom ? sorted[k - 1] : sorted[sorted.length - k];
                    for (const [key, v] of byRC) {
                        if (isBottom ? v <= threshold : v >= threshold) {
                            fills.set(key, fill);
                        }
                    }
                }
                break;
            }
            case 'colorScale': {
                const { byRC, sorted } = collectRangeValues(sheet, rule);
                if (sorted.length === 0) break;
                for (const [key, v] of byRC) {
                    const colour = evalColorScale(rule, v, sorted);
                    if (colour) fills.set(key, colour);
                }
                break;
            }
            // dataBar + iconSet: out of scope per M16. Documented in
            // BUILD_PLAN.md `Out of scope`. The cell value still renders;
            // only the visualisation glyph / bar is dropped.
            default:
                break;
        }
    }
    return fills;
}

// ───────── Sheet → HTML ─────────────────────────────────────────────

interface RenderContext {
    snapshot: Snapshot;
    styles: Record<string, Record<string, unknown>>;
}

function buildMergeIndex(
    sheet: SnapshotSheet,
): { skipKeys: Set<string>; anchors: Map<string, { rowSpan: number; colSpan: number }> } {
    const skipKeys = new Set<string>();
    const anchors = new Map<string, { rowSpan: number; colSpan: number }>();
    if (!Array.isArray(sheet.mergeData)) return { skipKeys, anchors };
    for (const m of sheet.mergeData) {
        if (
            typeof m.startRow !== 'number'
            || typeof m.endRow !== 'number'
            || typeof m.startColumn !== 'number'
            || typeof m.endColumn !== 'number'
        ) continue;
        const rowSpan = m.endRow - m.startRow + 1;
        const colSpan = m.endColumn - m.startColumn + 1;
        anchors.set(`${m.startRow}:${m.startColumn}`, { rowSpan, colSpan });
        for (let r = m.startRow; r <= m.endRow; r++) {
            for (let c = m.startColumn; c <= m.endColumn; c++) {
                if (r === m.startRow && c === m.startColumn) continue;
                skipKeys.add(`${r}:${c}`);
            }
        }
    }
    return { skipKeys, anchors };
}

function computeSheetExtent(sheet: SnapshotSheet): { rows: number; cols: number } {
    let maxRow = -1;
    let maxCol = -1;
    const cellData = sheet.cellData ?? {};
    for (const rKey of Object.keys(cellData)) {
        const r = Number(rKey);
        if (!Number.isFinite(r)) continue;
        const row = cellData[rKey];
        if (!row) continue;
        if (r > maxRow) maxRow = r;
        for (const cKey of Object.keys(row)) {
            const c = Number(cKey);
            if (!Number.isFinite(c)) continue;
            if (c > maxCol) maxCol = c;
        }
    }
    if (Array.isArray(sheet.mergeData)) {
        for (const m of sheet.mergeData) {
            if (typeof m.endRow === 'number' && m.endRow > maxRow) maxRow = m.endRow;
            if (typeof m.endColumn === 'number' && m.endColumn > maxCol) maxCol = m.endColumn;
        }
    }
    // If the sheet is completely empty, render a 1x1 stub so we still
    // emit a `<table>` (the gate "one table per sheet" stays met).
    return { rows: maxRow + 1, cols: maxCol + 1 };
}

function renderSheet(
    sheet: SnapshotSheet,
    ctx: RenderContext,
    cfFills: Map<string, string>,
): string {
    const { skipKeys, anchors } = buildMergeIndex(sheet);
    const { rows, cols } = computeSheetExtent(sheet);
    const sheetName = sheet.name ?? '';

    const out: string[] = [];
    // Sheet name as a heading; the test asserts the name appears within
    // 200 chars before the `<table>` opening tag.
    if (sheetName) {
        out.push(`<h3 class="notesheet-sheet-name">${escapeHtml(sheetName)}</h3>`);
    }
    out.push('<table class="notesheet-table" style="border-collapse: collapse;">');

    if (rows === 0 || cols === 0) {
        out.push('<tbody><tr><td></td></tr></tbody>');
        out.push('</table>');
        return out.join('');
    }

    const cellData = sheet.cellData ?? {};
    out.push('<tbody>');
    for (let r = 0; r < rows; r++) {
        out.push('<tr>');
        const rowMap: Record<string, SnapshotCell> | undefined = (cellData[String(r)] ?? cellData[r as unknown as string]);
        for (let c = 0; c < cols; c++) {
            const key = `${r}:${c}`;
            if (skipKeys.has(key)) continue;
            const anchor = anchors.get(key);
            const cell: SnapshotCell = rowMap?.[String(c)] ?? rowMap?.[c as unknown as string] ?? {};
            const baseStyle = resolveCellStyle(cell, ctx.styles);
            const cfFill = cfFills.get(key) ?? null;
            const inline = buildCellInlineStyle(baseStyle, cfFill);
            const attrs: string[] = [];
            if (anchor && anchor.rowSpan > 1) attrs.push(`rowspan="${anchor.rowSpan}"`);
            if (anchor && anchor.colSpan > 1) attrs.push(`colspan="${anchor.colSpan}"`);
            if (inline) attrs.push(`style="${escapeHtml(inline)}"`);
            const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';
            const valueHtml = renderCellValue(cell);
            out.push(`<td${attrStr}>${valueHtml}</td>`);
        }
        out.push('</tr>');
    }
    out.push('</tbody>');
    out.push('</table>');
    return out.join('');
}

// ───────── Entry: render a fenced notesheet body ────────────────────

// Exposed so Jest tests can call it directly without going through
// markdown-it. Returns the fully-rendered HTML for a fenced notesheet
// body's CONTENT (not the surrounding fence). Returns null if the
// JSON is malformed — the markdown-it integration translates that to
// "fall through to default fence rendering."
export function renderNotesheetSnapshot(content: string): string | null {
    const snapshot = parseSnapshotJson(content);
    if (!snapshot) return null;
    const sheets = snapshot.sheets ?? {};
    const order = Array.isArray(snapshot.sheetOrder) ? snapshot.sheetOrder : Object.keys(sheets);
    const styles = snapshot.styles ?? {};
    const cfBySubUnit = loadCfRulesFromSnapshot(snapshot);
    const ctx: RenderContext = { snapshot, styles };
    const parts: string[] = [];
    parts.push('<div class="notesheet-export">');
    for (const sheetId of order) {
        const sheet = sheets[sheetId];
        if (!sheet) continue;
        // CF rules in the snapshot are keyed by subUnitId (== sheetId in
        // Univer's CF model).
        const rules = cfBySubUnit[sheet.id] ?? cfBySubUnit[sheetId] ?? [];
        const cfFills = evaluateCfFor(sheet, rules);
        parts.push(renderSheet(sheet, ctx, cfFills));
    }
    parts.push('</div>');
    return parts.join('');
}

// Entry point used by the markdown-it fence override. Consults the
// fence info string; when not a notesheet fence, returns null and the
// caller falls through to markdown-it's default renderer.
export function renderFenceToken(token: FenceToken): string | null {
    const info = parseFenceInfo(token.info);
    if (!info.isNotesheet) return null;
    if (info.version !== 1) {
        // Future-version fences fall through too — let the user see the
        // raw JSON instead of pretending to render an unknown shape.
        return null;
    }
    const html = renderNotesheetSnapshot(token.content ?? '');
    return html;
}

// ───────── Markdown-It plugin shape Joplin expects ──────────────────

// Joplin's content script type for ContentScriptType.MarkdownItPlugin
// expects `default function(context)` returning `{ plugin, assets }`.
// `plugin` is invoked with markdown-it's instance + options.
//
// We override the fence renderer rule. When the token's info string
// matches a notesheet fence, we replace the default rendering with our
// HTML; otherwise we delegate to the previous renderer (which handles
// every other code-fence type — generic code blocks, javascript, etc.).
//
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export default function (_context: any) {
    return {
        plugin: function (markdownIt: MarkdownIt, _opts: unknown) {
            if (!markdownIt || !markdownIt.renderer || !markdownIt.renderer.rules) return;
            const defaultFence = markdownIt.renderer.rules.fence
                || function (tokens: FenceToken[], idx: number, options: unknown, _env: unknown, self: { renderToken: (...args: unknown[]) => string }) {
                    return self.renderToken(tokens as unknown as never[], idx as unknown as never, options as never);
                };
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            markdownIt.renderer.rules.fence = function (tokens: FenceToken[], idx: number, options: unknown, env: unknown, self: any) {
                const token = tokens[idx];
                const html = renderFenceToken(token);
                if (html !== null) return html;
                return defaultFence(tokens, idx, options, env, self);
            };
        },
        assets: function () {
            // CSS could be added here later (e.g. table border defaults).
            // Inline styles in renderSheet handle per-cell formatting,
            // so no external CSS is strictly required.
            return [];
        },
    };
}
