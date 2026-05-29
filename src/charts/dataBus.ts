// Tiny in-process pubsub for chart live updates.
//
// The editor view emits new chart data via pushChartUpdate(id, data) when a
// source-range cell is edited. The React chart component subscribes by id on
// mount and re-renders Chart.js in place. We deliberately avoid going through
// Univer's drawing service for data updates — that service models a chart as
// an immutable float-dom, so re-anchoring it on every edit clobbers any drag
// or resize the user has applied. The bus is a sidechannel that only carries
// data; positioning stays under Univer's control where it belongs.
//
// Both sides run in the same webview JS context, so a module-level Map is
// sufficient. No serialization, no IPC.

import type { ChartData } from './extractData';

export type ChartUpdateListener = (data: ChartData) => void;

const listeners = new Map<string, Set<ChartUpdateListener>>();

export function subscribeChartUpdate(id: string, listener: ChartUpdateListener): () => void {
    let set = listeners.get(id);
    if (!set) {
        set = new Set();
        listeners.set(id, set);
    }
    set.add(listener);
    return () => {
        const s = listeners.get(id);
        if (!s) return;
        s.delete(listener);
        if (s.size === 0) listeners.delete(id);
    };
}

export function pushChartUpdate(id: string, data: ChartData): void {
    const set = listeners.get(id);
    if (!set || set.size === 0) return;
    for (const fn of set) {
        try { fn(data); } catch (e) { console.error('[Notesheet] chart listener threw', e); }
    }
}

// Test/debug helper — not used in production.
export function _resetChartBus(): void {
    listeners.clear();
}
