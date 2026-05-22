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
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import sheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import sheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';

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
const READY_MESSAGE = { type: 'ready' };

let activeUniver: { dispose: () => void } | null = null;
let activeApi: any = null;
let saveTimer: ReturnType<typeof setTimeout> | null = null;

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
            [LocaleType.EN_US]: merge({}, sheetsCoreEnUS, sheetsSortEnUS, sheetsFilterEnUS, {
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
