// Notesheet editor view — runs inside Joplin's Custom Editor webview.
//
// Mounts a Univer workbook via the @univerjs/presets bootstrap API. The
// host plugin (src/index.ts) sends the snapshot in via webviewApi.postMessage
// and receives save events back the same way.

// Univer CSS must be imported before the JS — they contain the component
// styles, toolbar layouts, and icon font definitions that make the UI render
// properly. Without these, you see scattered icons and unstyled HTML.
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/preset-sheets-sort/lib/index.css';
import '@univerjs/preset-sheets-filter/lib/index.css';
import '@univerjs/preset-sheets-table/lib/index.css';
import '@univerjs/preset-sheets-drawing/lib/index.css';
import '@univerjs/preset-sheets-hyper-link/lib/index.css';
// Conditional Formatting (M15). The preset bundles the CF rendering
// engine (colorScale, dataBar, cellIs/highlightCell, top10/rank,
// iconSet) plus the "Manage Conditional Format Rules" panel users
// reach via the toolbar's Home → Conditional Formatting menu. Without
// this CSS import the panel chrome would render unstyled and the
// iconSet glyph font would be missing.
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';

import {
    createUniver,
    IResourceManagerService,
    LocaleType,
    LogLevel,
    UniverInstanceType,
    defaultTheme,
    merge,
} from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';

import { withFlatTableTheme } from './univerTableTheme';
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import sheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import sheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import sheetsTableEnUS from '@univerjs/preset-sheets-table/locales/en-US';
import sheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import sheetsHyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US';
import sheetsCfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';

import { ColumnChartIcon } from '@univerjs/icons';

import {
    xlsxBufferToSnapshot,
    snapshotToXlsxBuffer,
    NOTESHEET_SYNTH_STYLES_RESOURCE,
    NOTESHEET_THEME_CLR_SCHEME_RESOURCE,
    NOTESHEET_SHAPES_RESOURCE,
} from './xlsx';
import NotesheetChart, { type NotesheetChartType } from './charts/NotesheetChart';
import { extractRangeAsChartData, type RangeAddress } from './charts/extractData';
import { pushChartUpdate } from './charts/dataBus';

declare global {
    interface Window {
        webviewApi?: {
            postMessage: (msg: unknown) => Promise<unknown>;
            onMessage: (cb: (event: { message: unknown }) => void) => void;
        };
    }
}

interface LoadMessage {
    type: 'load';
    snapshot: Record<string, unknown>;
}

const ROOT_ID = 'notesheet-univer-root';
const ACTION_BAR_ID = 'notesheet-action-bar';
const STATUS_ID = 'notesheet-action-status';
const FILE_INPUT_ID = 'notesheet-xlsx-input';
const CHART_MODAL_ID = 'notesheet-chart-modal';
const CHART_COMPONENT_KEY = 'NotesheetChart';
const CHART_ICON_KEY = 'NotesheetChartIcon';
const CHART_MENU_ID = 'notesheet.menu.insert-chart';
const READY_MESSAGE = { type: 'ready' };

let activeUniver: { dispose: () => void } | null = null;
let activeApi: any = null;
let saveTimer: ReturnType<typeof setTimeout> | null = null;
let statusTimer: ReturnType<typeof setTimeout> | null = null;

function ensureRoot(): HTMLElement {
    let el = document.getElementById(ROOT_ID);
    if (!el) {
        el = document.createElement('div');
        el.id = ROOT_ID;
        Object.assign(el.style, {
            position: 'absolute',
            inset: '0',
            width: '100%',
            height: '100%',
        });
        document.body.appendChild(el);
    }
    return el;
}

function setStatus(text: string, isError = false): void {
    const el = document.getElementById(STATUS_ID);
    if (!el) return;
    el.textContent = text;
    el.style.color = isError ? '#b91c1c' : '#374151';
    el.style.opacity = text ? '1' : '0';
    if (statusTimer) clearTimeout(statusTimer);
    if (text) {
        statusTimer = setTimeout(() => {
            el.style.opacity = '0';
            el.textContent = '';
        }, 3000);
    }
}

async function handleImport(file: File): Promise<void> {
    // Replaces the entire workbook — confirm first since this is destructive.
    if (!window.confirm('Import will replace the current sheet contents. Continue?')) return;
    setStatus('Importing…');
    try {
        const buffer = await file.arrayBuffer();
        const snapshot = await xlsxBufferToSnapshot(buffer);
        bootUniver(snapshot);
        // Persist the freshly-loaded workbook so the import survives a reload
        // even if the user makes no edits afterward.
        saveNow();
        setStatus('Imported.');
    } catch (e) {
        console.error('[Notesheet] xlsx import failed', e);
        setStatus('Import failed: ' + (e instanceof Error ? e.message : String(e)), true);
    }
}

async function handleExport(): Promise<void> {
    setStatus('Exporting…');
    try {
        if (!activeApi) throw new Error('workbook not ready');
        const workbook =
            activeApi.getActiveWorkbook?.() || activeApi.getActiveSheet?.()?.getWorkbook?.();
        // save() is the current Univer 0.23 API; getSnapshot is deprecated.
        const snapshot = workbook?.save?.() || workbook?.getSnapshot?.();
        if (!snapshot) throw new Error('no snapshot available');
        const buffer = await snapshotToXlsxBuffer(snapshot);
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'spreadsheet.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        setStatus('Exported.');
    } catch (e) {
        console.error('[Notesheet] xlsx export failed', e);
        setStatus('Export failed: ' + (e instanceof Error ? e.message : String(e)), true);
    }
}

// Disposable returned by fWorkbook.onSelectionChange. Held while the chart
// panel is open so the Range field tracks the user's live selection on the
// sheet, then disposed when the panel closes.
let chartSelectionDisposable: { dispose: () => void } | null = null;

// Charts tracked by the data bus, keyed by chartId. Used by SheetEditEnded
// to figure out which charts need a fresh data push when a cell changes.
// Source-range positions live here; Univer owns the chart's anchor via
// the SHEET_DRAWING_PLUGIN snapshot resource.
import { populateTrackedChartsFromSnapshot, trackedCharts } from './charts/trackedCharts';

function rangeContainsCell(r: RangeAddress, row: number, col: number): boolean {
    return row >= r.startRow && row <= r.endRow && col >= r.startColumn && col <= r.endColumn;
}

// Re-extract chart data and push it through the dataBus when an edit lands
// inside any tracked chart's source range. This never touches Univer's
// drawing service — the chart's float-dom (position, size, drag state) is
// untouched. Only the canvas contents update.
function refreshChartsForEdit(): void {
    if (!activeApi || trackedCharts.size === 0) return;
    try {
        const fWorkbook = activeApi.getActiveWorkbook?.();
        const fSheet = fWorkbook?.getActiveSheet?.();
        if (!fSheet) return;
        const cell = fSheet.getActiveRange?.();
        if (!cell) return;
        const r = cell.getRange?.() ?? cell.getRangeData?.();
        const editedRow = typeof cell.getRow === 'function' ? cell.getRow() : r?.startRow;
        const editedCol = typeof cell.getColumn === 'function' ? cell.getColumn() : r?.startColumn;
        if (typeof editedRow !== 'number' || typeof editedCol !== 'number') return;

        for (const chart of trackedCharts.values()) {
            if (!rangeContainsCell(chart.sourceRange, editedRow, editedCol)) continue;
            // Pass hasHeaderRow so an imported chart (whose sourceRange spans
            // the header row) re-extracts the same way the importer did —
            // skipping row 0 — instead of leaking the header in as a phantom
            // category on edit.
            const fresh = extractRangeAsChartData(fWorkbook, chart.sourceRange, {
                hasHeaderRow: chart.hasHeaderRow === true,
            });
            pushChartUpdate(chart.id, fresh);
        }
    } catch (e) {
        console.warn('[Notesheet] refreshChartsForEdit failed', e);
    }
}

function closeChartModal(): void {
    if (chartSelectionDisposable) {
        try {
            chartSelectionDisposable.dispose();
        } catch {
            /* ignore */
        }
        chartSelectionDisposable = null;
    }
    const modal = document.getElementById(CHART_MODAL_ID);
    if (modal) modal.remove();
}

function insertChart(type: NotesheetChartType, rangeA1: string, title: string): void {
    if (!activeApi) {
        setStatus('Workbook not ready', true);
        return;
    }
    try {
        const fWorkbook = activeApi.getActiveWorkbook?.();
        const fSheet = fWorkbook?.getActiveSheet?.();
        if (!fSheet) throw new Error('no active sheet');

        const fRange = fSheet.getRange?.(rangeA1);
        if (!fRange) throw new Error('range "' + rangeA1 + '" not found');

        // Pull start/end indexes from the FRange so we can later detect edits
        // inside the source range. The facade exposes these directly.
        const r = fRange.getRange?.() ?? fRange.getRangeData?.() ?? null;
        if (!r) throw new Error('could not read range bounds');
        const sourceRange: RangeAddress = {
            startRow: r.startRow,
            endRow: r.endRow,
            startColumn: r.startColumn,
            endColumn: r.endColumn,
            unitId: fWorkbook.getId?.(),
            subUnitId: fSheet.getSheetId?.(),
        };

        const chartData = extractRangeAsChartData(fWorkbook, sourceRange);
        const chartId = 'chart-' + Date.now().toString(36);

        // CRITICAL: use addFloatDomToPosition, NOT addFloatDomToRange.
        // addFloatDomToRange permanently subscribes to the range's screen
        // position and re-anchors the chart on every scroll, which fights
        // user drags. addFloatDomToPosition places the chart at absolute
        // sheet coordinates once and lets the user own the position from
        // there forward (drag/resize/persist via Univer's snapshot).
        //
        // Initial position is computed as "just past the right edge of the
        // source range" using Univer's default cell sizes. Variable
        // row/column sizes will make this approximate, but it's a fine
        // default — Excel/Sheets also drop charts at a default position
        // and let the user drag.
        const DEFAULT_COL_W = 73;
        const DEFAULT_ROW_H = 19;
        const ROW_HEADER_W = 46;
        const COL_HEADER_H = 20;
        const chartW = 480;
        const chartH = 320;
        const startX = ROW_HEADER_W + (sourceRange.endColumn + 2) * DEFAULT_COL_W;
        const startY = COL_HEADER_H + sourceRange.startRow * DEFAULT_ROW_H;

        const handle = fSheet.addFloatDomToPosition?.(
            {
                componentKey: CHART_COMPONENT_KEY,
                // chartId opens a live-update channel — see refreshChartsForEdit.
                data: { chartId, type, sourceRange, title, ...chartData },
                allowTransform: true,
                // Forward pointer events through the chart canvas so Univer's
                // transformer (under the canvas) sees clicks and can attach
                // its drag/resize chrome. Without this the chart swallows all
                // pointer events and the user can't select, move, or delete.
                eventPassThrough: true,
                initPosition: {
                    startX,
                    startY,
                    endX: startX + chartW,
                    endY: startY + chartH,
                },
            },
            chartId,
        );
        if (!handle) throw new Error('addFloatDomToPosition returned no handle');
        trackedCharts.set(chartId, { id: chartId, sourceRange });
        setStatus('Chart inserted.');
        scheduleSave();
    } catch (e) {
        console.error('[Notesheet] insertChart failed', e);
        setStatus('Insert failed: ' + (e instanceof Error ? e.message : String(e)), true);
    }
}

// Open a non-blocking chart-config panel. It docks at the top-right (under the
// action bar) and, while open, mirrors the user's live sheet selection into
// the Range field via fWorkbook.onSelectionChange. This avoids fighting
// Univer's built-in range-selector dialog (whose dismissal model is buggy in
// 0.23) while giving users the same "click on the sheet to pick a range"
// experience Excel/Sheets offers when inserting a chart.
function openChartModal(): void {
    if (document.getElementById(CHART_MODAL_ID)) return;
    if (!activeApi) {
        setStatus('Workbook not ready', true);
        return;
    }

    const fWorkbook = activeApi.getActiveWorkbook?.();
    const fSheet = fWorkbook?.getActiveSheet?.();
    const initialRange = (() => {
        try {
            return fSheet?.getActiveRange?.()?.getA1Notation?.() ?? 'A1:B5';
        } catch {
            return 'A1:B5';
        }
    })();

    const panel = document.createElement('div');
    panel.id = CHART_MODAL_ID;
    Object.assign(panel.style, {
        position: 'absolute',
        top: '52px',
        right: '12px',
        background: '#fff',
        borderRadius: '8px',
        padding: '14px 16px',
        boxShadow: '0 4px 16px rgba(0,0,0,0.18)',
        font: '13px -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif',
        color: '#111827',
        minWidth: '300px',
        display: 'flex',
        flexDirection: 'column',
        gap: '10px',
        zIndex: '10001',
        // Allow clicks on the sheet behind the panel — the panel doesn't
        // overlay the whole editor, just docks in the corner.
        pointerEvents: 'auto',
    });

    const title = document.createElement('div');
    title.textContent = 'Insert Chart';
    Object.assign(title.style, { fontSize: '14px', fontWeight: '600' });

    const hint = document.createElement('div');
    hint.textContent = 'Tip: click or drag on the sheet to pick a range.';
    Object.assign(hint.style, { fontSize: '11px', color: '#6b7280', marginTop: '-4px' });

    const typeRow = document.createElement('label');
    Object.assign(typeRow.style, { display: 'flex', alignItems: 'center', gap: '8px' });
    typeRow.textContent = 'Type';
    const typeSelect = document.createElement('select');
    Object.assign(typeSelect.style, { flex: '1', padding: '4px 6px' });
    for (const opt of ['bar', 'line', 'pie', 'doughnut'] as NotesheetChartType[]) {
        const o = document.createElement('option');
        o.value = opt;
        o.textContent = opt[0].toUpperCase() + opt.slice(1);
        typeSelect.appendChild(o);
    }
    typeRow.appendChild(typeSelect);

    const rangeRow = document.createElement('label');
    Object.assign(rangeRow.style, { display: 'flex', alignItems: 'center', gap: '8px' });
    rangeRow.textContent = 'Range';
    const rangeInput = document.createElement('input');
    rangeInput.type = 'text';
    rangeInput.value = initialRange;
    Object.assign(rangeInput.style, { flex: '1', padding: '4px 6px' });
    rangeRow.appendChild(rangeInput);

    const titleRow = document.createElement('label');
    Object.assign(titleRow.style, { display: 'flex', alignItems: 'center', gap: '8px' });
    titleRow.textContent = 'Title';
    const titleInput = document.createElement('input');
    titleInput.type = 'text';
    titleInput.placeholder = 'optional';
    Object.assign(titleInput.style, { flex: '1', padding: '4px 6px' });
    titleRow.appendChild(titleInput);

    const buttonRow = document.createElement('div');
    Object.assign(buttonRow.style, {
        display: 'flex',
        justifyContent: 'flex-end',
        gap: '8px',
        marginTop: '4px',
    });
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Cancel';
    Object.assign(cancelBtn.style, {
        padding: '4px 12px',
        background: '#fff',
        border: '1px solid #9ca3af',
        borderRadius: '4px',
        cursor: 'pointer',
    });
    const submitBtn = document.createElement('button');
    submitBtn.type = 'button';
    submitBtn.textContent = 'Insert';
    Object.assign(submitBtn.style, {
        padding: '4px 14px',
        background: '#2563eb',
        border: '1px solid #1d4ed8',
        color: '#fff',
        borderRadius: '4px',
        cursor: 'pointer',
    });
    buttonRow.appendChild(cancelBtn);
    buttonRow.appendChild(submitBtn);

    panel.appendChild(title);
    panel.appendChild(hint);
    panel.appendChild(typeRow);
    panel.appendChild(rangeRow);
    panel.appendChild(titleRow);
    panel.appendChild(buttonRow);
    document.body.appendChild(panel);

    // Live-track the user's sheet selection while this panel is open. The
    // Range field updates whenever the selection changes, but only when the
    // user is NOT actively typing in it (so manual edits aren't clobbered).
    try {
        if (typeof fWorkbook?.onSelectionChange === 'function') {
            chartSelectionDisposable = fWorkbook.onSelectionChange(() => {
                if (document.activeElement === rangeInput) return;
                try {
                    const a1 = fWorkbook.getActiveSheet?.()?.getActiveRange?.()?.getA1Notation?.();
                    if (a1) rangeInput.value = a1;
                } catch {
                    /* ignore */
                }
            });
        }
    } catch (e) {
        console.warn('[Notesheet] onSelectionChange subscribe failed', e);
    }

    cancelBtn.addEventListener('click', closeChartModal);
    document.addEventListener('keydown', function escHandler(ev) {
        if (ev.key === 'Escape') {
            closeChartModal();
            document.removeEventListener('keydown', escHandler);
        }
    });
    submitBtn.addEventListener('click', () => {
        const t = typeSelect.value as NotesheetChartType;
        const rng = rangeInput.value.trim() || 'A1:B5';
        const ttl = titleInput.value.trim();
        closeChartModal();
        insertChart(t, rng, ttl);
    });
}

function ensureActionBar(): void {
    if (document.getElementById(ACTION_BAR_ID)) return;

    const bar = document.createElement('div');
    bar.id = ACTION_BAR_ID;
    Object.assign(bar.style, {
        position: 'absolute',
        top: '8px',
        right: '12px',
        // Above Univer's toolbar/menus so the buttons are always reachable.
        zIndex: '10000',
        display: 'flex',
        alignItems: 'center',
        gap: '8px',
        padding: '6px 8px',
        background: 'rgba(255, 255, 255, 0.92)',
        border: '1px solid #d1d5db',
        borderRadius: '6px',
        boxShadow: '0 1px 3px rgba(0, 0, 0, 0.08)',
        font: '12px -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif',
        // The bar accepts clicks within its bounds; everything outside passes
        // through to Univer because we don't put a wrapper over the page.
        pointerEvents: 'auto',
    });

    const buttonStyle: Partial<CSSStyleDeclaration> = {
        padding: '4px 10px',
        background: '#ffffff',
        border: '1px solid #9ca3af',
        borderRadius: '4px',
        cursor: 'pointer',
        font: 'inherit',
        color: '#111827',
    };

    const importBtn = document.createElement('button');
    importBtn.type = 'button';
    importBtn.textContent = 'Import .xlsx';
    Object.assign(importBtn.style, buttonStyle);

    const exportBtn = document.createElement('button');
    exportBtn.type = 'button';
    exportBtn.textContent = 'Export .xlsx';
    Object.assign(exportBtn.style, buttonStyle);

    const status = document.createElement('span');
    status.id = STATUS_ID;
    Object.assign(status.style, {
        marginLeft: '4px',
        opacity: '0',
        transition: 'opacity 200ms ease-out',
        whiteSpace: 'nowrap',
    });

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.id = FILE_INPUT_ID;
    fileInput.accept = '.xlsx,.xls';
    fileInput.style.display = 'none';
    fileInput.addEventListener('change', () => {
        const file = fileInput.files?.[0];
        // Reset so re-selecting the same file fires another change event.
        fileInput.value = '';
        if (file) void handleImport(file);
    });

    importBtn.addEventListener('click', () => fileInput.click());
    exportBtn.addEventListener('click', () => void handleExport());

    bar.appendChild(importBtn);
    bar.appendChild(exportBtn);
    bar.appendChild(status);
    bar.appendChild(fileInput);
    document.body.appendChild(bar);
}

function bootUniver(snapshot: Record<string, unknown>): void {
    if (activeUniver) {
        try {
            activeUniver.dispose();
        } catch {
            /* ignore */
        }
        activeUniver = null;
        activeApi = null;
        const old = document.getElementById(ROOT_ID);
        if (old) old.remove();
    }

    ensureRoot();

    const { univer, univerAPI } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
            [LocaleType.EN_US]: merge(
                {},
                sheetsCoreEnUS,
                sheetsSortEnUS,
                sheetsFilterEnUS,
                sheetsTableEnUS,
                sheetsDrawingEnUS,
                sheetsHyperLinkEnUS,
                sheetsCfEnUS,
                {
                    'sheets-sort': {
                        dialog: {
                            'sort-reminder': 'Sort Warning',
                            'sort-reminder-desc':
                                'Data was found next to your selection. What would you like to do?',
                            'sort-reminder-ext': 'Expand the selection (sort entire rows together)',
                            'sort-reminder-no': 'Continue with the current selection',
                            'first-row-check': 'My data has headers',
                        },
                        general: {
                            'sort-custom': 'Sort...',
                        },
                    },
                },
            ),
        },
        theme: defaultTheme,
        logLevel: LogLevel.WARN,
        presets: [
            UniverSheetsCorePreset({
                container: ROOT_ID,
                // Hide the quick-sort options, keeping only "Sort..." in the
                // dropdown. Same pattern as Excel's Data → Sort & Filter
                // dropdown which shows "Custom Sort..." among other options.
                // Context menu also cleaned up.
                menu: {
                    'sheet.command.sort-range-asc': { hidden: true },
                    'sheet.command.sort-range-desc': { hidden: true },
                    'sheet.command.sort-range-asc-ext': { hidden: true },
                    'sheet.command.sort-range-desc-ext': { hidden: true },
                    'sheet.command.sort-range-asc-ctx': { hidden: true },
                    'sheet.command.sort-range-desc-ctx': { hidden: true },
                    'sheet.command.sort-range-asc-ext-ctx': { hidden: true },
                    'sheet.command.sort-range-desc-ext-ctx': { hidden: true },
                },
            }),
            UniverSheetsSortPreset(),
            UniverSheetsFilterPreset(),
            // Univer's table plugin auto-applies one of 6 default themes
            // (`table-default-0..5`) as a RangeThemeStyle on top of any cell
            // that's part of an ITableJson. Even though our snapshot already
            // synthesizes per-cell `bg`/`cl` values that match the source
            // Excel TableStyleMedium2, the theme overlay shows lavender
            // (#BAC6F8) banding instead of the imported teal (#83CBEB).
            //
            // The fix is to register a no-op "passthrough" theme as
            // userThemes[0] and set defaultThemeIndex: 0. The theme controller
            // resolves the active theme as `userThemes.concat(defaultThemes)`,
            // so ours wins. With every style slot empty, the theme overlay
            // contributes nothing and our synthesized cell colors render
            // unaltered.
            //
            // See sheets-table/lib/es/index.js: SheetsTableThemeController
            // _initUserTableTheme + tableAdd$.subscribe — when fromJSON loads
            // a table from the snapshot the tableAdd event lacks tableStyleId,
            // so the controller falls back to the default index. We use that
            // hook by making index 0 our passthrough theme.
            withFlatTableTheme(UniverSheetsTablePreset()),
            // Drawing preset enables the float-DOM machinery (CanvasFloatDomService),
            // which we use to anchor and drag/resize charts over the grid.
            UniverSheetsDrawingPreset(),
            // Adds the hyperlink layer (cell.p.body.customRanges with
            // CustomRangeType.HYPERLINK). On snapshot load the preset's
            // RichTextRefRangeController walks every cell's `p` and
            // registers the link with RefRangeService — making the URL
            // clickable in the editor automatically.
            UniverSheetsHyperLinkPreset(),
            // Conditional Formatting (M15). Reads CF rules from the
            // snapshot's `SHEET_CONDITIONAL_FORMATTING_PLUGIN` resource
            // and paints them on the canvas in render order: cell fills
            // (cellIs / top10 / highlightCell), then dataBars, then
            // colorScale gradients, then iconSet glyphs. The CF render
            // layer paints OVER per-cell `bg` from the M12 table-style
            // synthesizer; if precedence ever inverts (synthesizer fill
            // masks CF colour), mark CF-bound cells in the synth-styles
            // sidecar so the synthesizer skips them. Users add/edit/
            // delete rules via the "Manage Conditional Format Rules"
            // panel reachable from the toolbar.
            UniverSheetsConditionalFormattingPreset(),
        ],
    });

    // Register Notesheet's snapshot resources with Univer's resource manager.
    // Univer's `loadResources` (called from inside createUnit) iterates the
    // currently-registered hooks and dispatches the matching entries from
    // the input snapshot to each hook's onLoad. Conversely `getResources`
    // (called from workbook.save()) produces output `resources` ONLY from
    // registered hooks. So unregistered resources arrive in the snapshot
    // but get silently dropped on save — which was breaking our theme +
    // synth-styles round-trip when going Joplin save → Joplin reload →
    // Export.
    //
    // Each hook keeps a per-unitId map (so multiple workbooks could
    // coexist in one Univer instance, though Notesheet only uses one).
    // toJson serializes back to the original string we received on import;
    // parseJson does the inverse. Both data types are already strings, so
    // these are identity transforms.
    try {
        const injector = (univer as { __getInjector?: () => unknown }).__getInjector?.();
        const resourceManager = (injector as { get?: (id: unknown) => unknown } | undefined)?.get?.(
            IResourceManagerService,
        ) as { registerPluginResource: (hook: unknown) => unknown } | undefined;
        if (resourceManager?.registerPluginResource) {
            for (const name of [
                NOTESHEET_SYNTH_STYLES_RESOURCE,
                NOTESHEET_THEME_CLR_SCHEME_RESOURCE,
                // M18 A2: preserve-only shape anchors. Univer never renders or
                // touches these; the passthrough hook just carries the string
                // through editor save/reload so shapes survive a round-trip
                // (without it, Univer drops the unregistered resource on save
                // and shapes vanish the moment the user edits the note).
                NOTESHEET_SHAPES_RESOURCE,
            ]) {
                const stash = new Map<string, string>();
                resourceManager.registerPluginResource({
                    pluginName: name,
                    businesses: [UniverInstanceType.UNIVER_SHEET],
                    onLoad: (unitId: string, resource: string) => {
                        if (typeof resource === 'string' && resource) stash.set(unitId, resource);
                    },
                    onUnLoad: (unitId: string) => {
                        stash.delete(unitId);
                    },
                    toJson: (unitId: string) => stash.get(unitId) ?? '',
                    parseJson: (raw: string) => raw,
                });
            }
        } else {
            console.warn(
                "[Notesheet] could not get IResourceManagerService — synth styles + theme palette won't round-trip on save",
            );
        }
    } catch (e) {
        console.warn('[Notesheet] resource hook registration failed', e);
    }

    univer.createUnit(UniverInstanceType.UNIVER_SHEET, snapshot);
    activeUniver = univer;
    activeApi = univerAPI;

    // M17: hydrate the editor's chart-tracking map from the snapshot's
    // SHEET_DRAWING_PLUGIN resource. Charts that arrived via
    // xlsxBufferToSnapshot don't otherwise wire into the data bus —
    // they'd render once with stale data and never refresh on edit.
    populateTrackedChartsFromSnapshot(snapshot);

    // M17 feature-3: imported chart drawings auto-mount via Univer's
    // SHEET_DRAWING_PLUGIN resource hook (drawingType: 8 = DRAWING_DOM
    // + componentKey: 'NotesheetChart' + transform/data block, written
    // by xlsxBufferToSnapshot). The resource hook walks every entry on
    // load and dispatches each to the registered float-DOM component.
    //
    // This is a quiet but load-bearing contract: the M13-style trap
    // (chart subscribed to the bus but never visible) only manifests
    // when the snapshot's drawing entry has the wrong drawingType or
    // is missing the transform block. Univer silently drops malformed
    // entries instead of throwing. The Jest tests in
    // tests/m17ChartImportLiveUpdate.test.ts assert the bus contract
    // independently of the runtime mount; the runtime gate is
    // feature-3's screenshot.

    // Expose the Univer FUniver facade on `window` so the PGE harness's
    // `eval-screenshot.js` (which evaluates JS inside the
    // UserWebviewIndex frame via Playwright `frame.evaluate(...)`) can
    // discover real column widths / row heights at sampling time. The
    // pixel sampler used to hardcode default column widths (73 CSS px)
    // and drifted sideways on any fixture where the operator widened
    // columns to fit content. Reading the live geometry from the
    // workbook is the only reliable fix; structural canvas scanning
    // (e.g. detecting column-header borders by pixel) is brittle on
    // Univer 0.23 because the borders are anti-aliased and faint.
    //
    // Read-only: the harness only calls getter facade methods. Production
    // users have no exposure beyond an extra global property pointer.
    try {
        (window as unknown as { __notesheetUniverAPI?: unknown }).__notesheetUniverAPI = univerAPI;
    } catch {
        // window may be inaccessible in some test runners; harmless.
    }

    // Register the chart component once per Univer instance. componentKey
    // 'NotesheetChart' is what addFloatDomToRange looks up.
    try {
        univerAPI.registerComponent?.(CHART_COMPONENT_KEY, NotesheetChart);
    } catch (e) {
        console.warn('[Notesheet] could not register chart component', e);
    }

    // Add an "Insert Chart" entry to the Insert ribbon next to Univer's
    // built-in image menu. Uses the public createMenu/appendTo Facade so we
    // don't poke private services. Idempotent across bootUniver re-runs:
    // createMenu().appendTo() with the same id is a no-op if already present.
    try {
        univerAPI.registerComponent?.(CHART_ICON_KEY, ColumnChartIcon);
        univerAPI
            .createMenu?.({
                id: CHART_MENU_ID,
                title: 'Insert Chart',
                icon: CHART_ICON_KEY,
                tooltip: 'Insert a Chart.js chart anchored over the grid',
                action: () => openChartModal(),
            })
            .appendTo('ribbon.insert.media');
    } catch (e) {
        console.warn('[Notesheet] could not register chart menu', e);
    }

    // Persist on every workbook change. We debounce so rapid edits don't
    // fire a save per keystroke. The host calls editor.saveNote which is
    // itself debounced by Joplin, but we layer our own gate to keep
    // postMessage traffic light.
    if (typeof univerAPI?.addEvent === 'function' && univerAPI?.Event?.SheetEditEnded) {
        try {
            univerAPI.addEvent(univerAPI.Event.SheetEditEnded, () => {
                scheduleSave();
                refreshChartsForEdit();
            });
        } catch (e) {
            console.warn('[Notesheet] could not subscribe to SheetEditEnded', e);
        }
    }

    // Drawing transforms (drag, resize, anchor changes) don't trigger
    // SheetEditEnded, so without this hook the workbook snapshot picks up
    // the new transform in memory but never gets written back to the note.
    // Hook the generic CommandExecuted stream and trigger a save whenever a
    // drawing-related command fires.
    if (typeof univerAPI?.addEvent === 'function' && univerAPI?.Event?.CommandExecuted) {
        try {
            univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event: { id?: string }) => {
                if (!event?.id) return;
                if (event.id.includes('drawing') || event.id.includes('Drawing')) {
                    scheduleSave();
                }
            });
        } catch (e) {
            console.warn('[Notesheet] could not subscribe to CommandExecuted', e);
        }
    }
}

function scheduleSave(): void {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(saveNow, 400);
}

function saveNow(): void {
    saveTimer = null;
    if (!activeApi) return;
    try {
        const workbook =
            activeApi.getActiveWorkbook?.() || activeApi.getActiveSheet?.()?.getWorkbook?.();
        if (!workbook) return;
        const snapshot = workbook.save?.() || workbook.getSnapshot?.();
        if (!snapshot) return;
        if (window.webviewApi) {
            window.webviewApi.postMessage({ type: 'save', snapshot }).catch((err) => {
                console.error('[Notesheet] save postMessage failed', err);
            });
        }
    } catch (e) {
        console.error('[Notesheet] saveNow failed', e);
    }
}

function init(): void {
    if (!window.webviewApi) {
        console.error('[Notesheet] webviewApi not available; cannot communicate with Joplin host');
        return;
    }
    ensureActionBar();
    window.webviewApi.onMessage((event) => {
        const m = event?.message as LoadMessage | undefined;
        if (m && m.type === 'load' && m.snapshot) {
            bootUniver(m.snapshot);
        }
    });
    window.webviewApi.postMessage(READY_MESSAGE).catch((err) => {
        console.error('[Notesheet] ready postMessage failed', err);
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
