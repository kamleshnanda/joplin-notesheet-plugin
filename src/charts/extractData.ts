// Pulls a 2-D range out of a Univer FWorkbook and shapes it for Chart.js.
//
// Convention (Excel-ish): if the range is wider than one column, the FIRST
// column becomes the X-axis labels and each remaining column becomes a series.
// A single-column range is treated as one unlabeled series.
//
// Numeric strings ('42') get coerced; anything that won't coerce becomes NaN
// and Chart.js silently skips it. We never throw — chart rendering should
// degrade gracefully when the data isn't tabular.

export const CHART_PALETTE = [
    '#3b82f6', // blue
    '#ef4444', // red
    '#10b981', // green
    '#f59e0b', // amber
    '#8b5cf6', // violet
    '#ec4899', // pink
    '#06b6d4', // cyan
    '#84cc16', // lime
];

export interface RangeAddress {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
    unitId?: string;
    subUnitId?: string;
}

export interface ChartData {
    labels: string[];
    datasets: Array<{
        label?: string;
        data: number[];
        backgroundColor: string | string[];
        borderColor?: string;
    }>;
}

function toNumber(v: unknown): number {
    if (typeof v === 'number') return v;
    if (typeof v === 'string') {
        const t = v.trim();
        if (t === '') return NaN;
        const n = Number(t);
        return Number.isFinite(n) ? n : NaN;
    }
    if (typeof v === 'boolean') return v ? 1 : 0;
    return NaN;
}

function toLabel(v: unknown): string {
    if (v === null || v === undefined) return '';
    return String(v);
}

// `workbook` is the FWorkbook facade. We accept `unknown` so callers don't
// have to import Univer types just to unit-test this.
export function extractRangeAsChartData(workbook: unknown, range: RangeAddress): ChartData {
    const empty: ChartData = { labels: [], datasets: [] };
    if (!workbook || !range) return empty;

    try {
        const wb = workbook as {
            getActiveSheet?: () => {
                getRange?: (r: RangeAddress) => { getValues?: () => unknown[][] } | null;
            } | null;
            getSheetBySheetId?: (id: string) => {
                getRange?: (r: RangeAddress) => { getValues?: () => unknown[][] } | null;
            } | null;
        };

        // Prefer the sheet the range was anchored on, fall back to active.
        const sheet =
            (range.subUnitId && wb.getSheetBySheetId?.(range.subUnitId)) || wb.getActiveSheet?.();
        if (!sheet) return empty;

        const rangeObj = sheet.getRange?.(range);
        const values = rangeObj?.getValues?.();
        if (!Array.isArray(values) || values.length === 0) return empty;

        const cols = values[0].length;
        if (cols === 1) {
            const data = values.map((row) => toNumber(row[0]));
            return {
                labels: data.map((_, i) => String(i + 1)),
                datasets: [
                    {
                        label: 'Series 1',
                        data,
                        backgroundColor: CHART_PALETTE[0],
                        borderColor: CHART_PALETTE[0],
                    },
                ],
            };
        }

        const labels = values.map((row) => toLabel(row[0]));
        const datasets: ChartData['datasets'] = [];
        for (let c = 1; c < cols; c++) {
            const seriesData = values.map((row) => toNumber(row[c]));
            const color = CHART_PALETTE[(c - 1) % CHART_PALETTE.length];
            datasets.push({
                label: 'Series ' + c,
                data: seriesData,
                backgroundColor: color,
                borderColor: color,
            });
        }
        return { labels, datasets };
    } catch {
        return empty;
    }
}

// Snapshot variant of extractRangeAsChartData. Reads cell values directly
// out of a Univer snapshot's `sheets[id].cellData[row][col].v` map without
// booting Univer. Used by:
//   * src/editorView.tsx populateTrackedChartsFromSnapshot — to rebuild
//     ChartData for an imported chart at snapshot-load time, before any
//     edits land.
//   * src/contentScripts/notesheetRenderer.ts (M17 feature-7) — to drive
//     static SVG emission inside the markdown-it preview pane.
//
// Resolution rules:
//   * If `range.subUnitId` is supplied, it's looked up in `snapshot.sheets`
//     directly. If absent or invalid, fall back to the snapshot's first
//     sheet via sheetOrder[0].
//   * If a target sheet name (string) is supplied via the `sheetName`
//     overload arg, find the sheet whose `name === sheetName`. This is what
//     cross-sheet charts pass — the chart drawing's `sourceSheetName` field
//     holds the target sheet's display name (NOT the subUnitId).
export function extractDataFromSnapshot(
    snapshot: unknown,
    range: RangeAddress,
    sheetName?: string,
): ChartData {
    const empty: ChartData = { labels: [], datasets: [] };
    if (!snapshot || typeof snapshot !== 'object' || !range) return empty;

    const snap = snapshot as {
        sheets?: Record<
            string,
            {
                id?: string;
                name?: string;
                cellData?: Record<string, Record<string, { v?: unknown }>>;
            }
        >;
        sheetOrder?: string[];
    };

    const sheets = snap.sheets ?? {};
    let sheet: { cellData?: Record<string, Record<string, { v?: unknown }>> } | undefined;

    if (sheetName) {
        for (const id of Object.keys(sheets)) {
            if (sheets[id]?.name === sheetName) {
                sheet = sheets[id];
                break;
            }
        }
    }
    if (!sheet && range.subUnitId) sheet = sheets[range.subUnitId];
    if (!sheet) {
        const firstId = (snap.sheetOrder ?? [])[0];
        if (firstId) sheet = sheets[firstId];
    }
    if (!sheet?.cellData) return empty;

    // Materialize values[r - startRow][c - startColumn]; cells absent from the
    // sparse cellData map become null (toNumber → NaN, toLabel → '').
    const values: unknown[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
        const row: unknown[] = [];
        const cellRow = sheet.cellData[String(r)];
        for (let c = range.startColumn; c <= range.endColumn; c++) {
            const cell = cellRow?.[String(c)];
            row.push(cell?.v ?? null);
        }
        values.push(row);
    }

    if (values.length === 0) return empty;
    const cols = values[0].length;
    if (cols === 1) {
        const data = values.map((row) => toNumber(row[0]));
        return {
            labels: data.map((_, i) => String(i + 1)),
            datasets: [
                {
                    label: 'Series 1',
                    data,
                    backgroundColor: CHART_PALETTE[0],
                    borderColor: CHART_PALETTE[0],
                },
            ],
        };
    }

    const labels = values.map((row) => toLabel(row[0]));
    const datasets: ChartData['datasets'] = [];
    for (let c = 1; c < cols; c++) {
        const seriesData = values.map((row) => toNumber(row[c]));
        const color = CHART_PALETTE[(c - 1) % CHART_PALETTE.length];
        datasets.push({
            label: 'Series ' + c,
            data: seriesData,
            backgroundColor: color,
            borderColor: color,
        });
    }
    return { labels, datasets };
}
