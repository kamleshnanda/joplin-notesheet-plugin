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

import {
    createUniver,
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
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import sheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import sheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import sheetsTableEnUS from '@univerjs/preset-sheets-table/locales/en-US';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from './xlsx';

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
        const workbook = activeApi.getActiveWorkbook?.() || activeApi.getActiveSheet?.()?.getWorkbook?.();
        const snapshot = workbook?.getSnapshot?.() || workbook?.save?.();
        if (!snapshot) throw new Error('no snapshot available');
        const buffer = await snapshotToXlsxBuffer(snapshot);
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
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
        try { activeUniver.dispose(); } catch { /* ignore */ }
        activeUniver = null;
        activeApi = null;
        const old = document.getElementById(ROOT_ID);
        if (old) old.remove();
    }

    ensureRoot();

    const { univer, univerAPI } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
            [LocaleType.EN_US]: merge({}, sheetsCoreEnUS, sheetsSortEnUS, sheetsFilterEnUS, sheetsTableEnUS, {
                'sheets-sort': {
                    dialog: {
                        'sort-reminder': 'Sort Warning',
                        'sort-reminder-desc': 'Data was found next to your selection. What would you like to do?',
                        'sort-reminder-ext': 'Expand the selection (sort entire rows together)',
                        'sort-reminder-no': 'Continue with the current selection',
                        'first-row-check': 'My data has headers',
                    },
                    general: {
                        'sort-custom': 'Sort...',
                    },
                },
            }),
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
            UniverSheetsTablePreset(),
        ],
    });

    univer.createUnit(UniverInstanceType.UNIVER_SHEET, snapshot);
    activeUniver = univer;
    activeApi = univerAPI;

    // Persist on every workbook change. We debounce so rapid edits don't
    // fire a save per keystroke. The host calls editor.saveNote which is
    // itself debounced by Joplin, but we layer our own gate to keep
    // postMessage traffic light.
    if (typeof univerAPI?.addEvent === 'function' && univerAPI?.Event?.SheetEditEnded) {
        try {
            univerAPI.addEvent(univerAPI.Event.SheetEditEnded, scheduleSave);
        } catch (e) {
            console.warn('[Notesheet] could not subscribe to SheetEditEnded', e);
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
        const workbook = activeApi.getActiveWorkbook?.() || activeApi.getActiveSheet?.()?.getWorkbook?.();
        if (!workbook) return;
        const snapshot = workbook.getSnapshot?.() || workbook.save?.();
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
