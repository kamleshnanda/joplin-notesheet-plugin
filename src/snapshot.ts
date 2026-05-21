// Snapshot serialization for Notesheet.
//
// A Notesheet note's body is a markdown fenced code block tagged `notesheet`
// containing a JSON-encoded Univer snapshot. The fence makes the body
// syntactically valid markdown so that toggling Joplin's markdown editor
// shows highlighted code instead of a raw JSON dump, and so that Joplin's
// full-text search still indexes the cell text.
//
// Wire format:
//
//     ```notesheet v=1
//     {"id":"...","sheets":{...},...}
//     ```
//
// `v=1` is a forward-compatibility marker — a future migration can switch
// to a different on-disk shape and dispatch by version.

export const FENCE_TAG = 'notesheet';
export const FENCE_VERSION = 1;

export type UniverSnapshot = Record<string, unknown>;

// Markdown fences are line-anchored: a code block is closed by ``` at the
// start of a line. Anchoring the closing ``` to a line start (with optional
// leading whitespace forbidden by markdown's commonmark spec for our case)
// prevents a stray ``` inside a cell value from terminating the fence early.
const FENCE_RE = /(?:^|\n)```notesheet\s+v=(\d+)\s*\n([\s\S]*?)\n```(?:\n|$)/;

export function isNotesheetBody(body: string | null | undefined): boolean {
    if (!body) return false;
    return FENCE_RE.test(body);
}

export function wrapSnapshot(snapshot: UniverSnapshot): string {
    const json = JSON.stringify(snapshot);
    return '```' + FENCE_TAG + ' v=' + FENCE_VERSION + '\n' + json + '\n```';
}

// Extracts the snapshot from a note body. Returns null if the body is not
// a Notesheet note (no fence) or if the JSON inside the fence is malformed.
// The version is returned alongside so future migrations can branch on it.
export function extractSnapshot(body: string | null | undefined):
    | { ok: true; snapshot: UniverSnapshot; version: number }
    | { ok: false; reason: 'no-fence' | 'bad-json' | 'unsupported-version' } {
    if (!body) return { ok: false, reason: 'no-fence' };
    const m = FENCE_RE.exec(body);
    if (!m) return { ok: false, reason: 'no-fence' };
    const version = Number(m[1]);
    if (!Number.isFinite(version) || version < 1 || version > FENCE_VERSION) {
        return { ok: false, reason: 'unsupported-version' };
    }
    try {
        const snapshot = JSON.parse(m[2]);
        if (!snapshot || typeof snapshot !== 'object') {
            return { ok: false, reason: 'bad-json' };
        }
        return { ok: true, snapshot, version };
    } catch {
        return { ok: false, reason: 'bad-json' };
    }
}

// An empty Univer workbook snapshot. Shape follows Univer's IWorkbookData.
// One sheet, no cells. Everything else gets defaulted by Univer at load.
export function emptySnapshot(): UniverSnapshot {
    return {
        id: 'workbook-' + Date.now(),
        sheetOrder: ['sheet-1'],
        name: 'Spreadsheet',
        appVersion: '0.1.0',
        locale: 'enUS',
        styles: {},
        sheets: {
            'sheet-1': {
                id: 'sheet-1',
                name: 'Sheet1',
                cellData: {},
                rowCount: 100,
                columnCount: 26,
                defaultColumnWidth: 73,
                defaultRowHeight: 19,
            },
        },
    };
}
