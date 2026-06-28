// M17 feature-4: chart-tracking map for the editor view.
//
// `trackedCharts` is consulted on every SheetEditEnded to figure out which
// charts need a fresh data push. The map's contents come from two paths:
//
//   1. Insert flow (insertChart in editorView.tsx) — adds an entry as the
//      user clicks the Insert Chart menu.
//   2. Snapshot-load flow (populateTrackedChartsFromSnapshot below) — reads
//      the SHEET_DRAWING_PLUGIN resource off a freshly-loaded snapshot and
//      registers an entry per chart drawing. WITHOUT this, charts imported
//      from .xlsx render once with stale data and never refresh on edit —
//      exactly the M13 failure mode (snapshot data correct, runtime broken)
//      the harness was built to prevent.
//
// Lives in its own module so tests can import it without booting Univer.

import type { RangeAddress } from './extractData';

export interface TrackedChart {
    id: string;
    sourceRange: RangeAddress;
    sourceSheetName?: string;
    // True when the sourceRange's first row is a header (category-axis title
    // + series names), so the live-edit re-extract must skip it. Derived from
    // the chart's meta.categoryAxisType === 'category'. Without this, editing
    // a cell re-reads the whole range and the header leaks in as a phantom
    // category (the "Quarter" 5th-bar bug).
    hasHeaderRow?: boolean;
}

export const trackedCharts = new Map<string, TrackedChart>();

export function populateTrackedChartsFromSnapshot(snapshot: Record<string, unknown>): void {
    trackedCharts.clear();
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> })
        .resources;
    if (!Array.isArray(resources)) return;
    const entry = resources.find((r) => r?.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry || typeof entry.data !== 'string') return;
    let parsed: Record<string, { data?: Record<string, unknown> }>;
    try {
        parsed = JSON.parse(entry.data);
    } catch {
        return;
    }
    if (!parsed || typeof parsed !== 'object') return;

    for (const subUnitId of Object.keys(parsed)) {
        const subUnit = parsed[subUnitId];
        const drawings = subUnit?.data;
        if (!drawings || typeof drawings !== 'object') continue;
        for (const drawingId of Object.keys(drawings)) {
            const d = (drawings as Record<string, unknown>)[drawingId] as {
                componentKey?: string;
                data?: {
                    chartId?: string;
                    sourceRange?: {
                        startRow?: number;
                        endRow?: number;
                        startColumn?: number;
                        endColumn?: number;
                    };
                    sourceSheetName?: string;
                    meta?: {
                        categoryAxisType?: 'index' | 'category';
                        hasHeaderRow?: boolean;
                    };
                };
            };
            if (d?.componentKey !== 'NotesheetChart') continue;
            const data = d.data;
            const chartId = data?.chartId;
            const sr = data?.sourceRange;
            if (typeof chartId !== 'string' || !chartId) continue;
            if (
                !sr ||
                typeof sr.startRow !== 'number' ||
                typeof sr.endRow !== 'number' ||
                typeof sr.startColumn !== 'number' ||
                typeof sr.endColumn !== 'number'
            )
                continue;
            trackedCharts.set(chartId, {
                id: chartId,
                sourceRange: {
                    startRow: sr.startRow,
                    endRow: sr.endRow,
                    startColumn: sr.startColumn,
                    endColumn: sr.endColumn,
                    subUnitId,
                },
                ...(typeof data?.sourceSheetName === 'string' && data.sourceSheetName
                    ? { sourceSheetName: data.sourceSheetName }
                    : {}),
                // Whether row 0 of sourceRange is a header to skip on
                // live-edit re-extract. Prefer the explicit `hasHeaderRow`
                // signal (set by the importer: true only when the categories
                // start at sheet row ≥ 2, i.e. a real header sits above them).
                // Fall back to the legacy categoryAxisType heuristic for
                // snapshots imported before hasHeaderRow was emitted —
                // imperfect (it over-skips charts whose <c:cat> starts at row
                // 0) but matches the prior behaviour for old snapshots.
                hasHeaderRow:
                    typeof data?.meta?.hasHeaderRow === 'boolean'
                        ? data.meta.hasHeaderRow
                        : data?.meta?.categoryAxisType === 'category',
            });
        }
    }
}
