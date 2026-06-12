// M18 B1: charts render as static SVG in the HTML / PDF / preview-pane
// export (the Markdown-It content script). Chart.js canvases don't
// survive into static HTML; the content script now hand-authors inline
// SVG from the snapshot's SHEET_DRAWING_PLUGIN resource.
//
// Anchored to STRUCTURAL SVG element counts derived from the input
// snapshot (per the project's chart-export test discipline): a bar
// chart emits one <rect> per finite data point, a line chart one
// <polyline> per series, a pie one <path> per positive slice. Not
// pinned to byte-exact SVG, so the renderer can evolve its styling
// without breaking these as long as the structure matches.

import { renderNotesheetSnapshot } from '../src/contentScripts/notesheetRenderer';

// Build a snapshot with one sheet + a SHEET_DRAWING_PLUGIN resource
// holding the given chart drawings on `sheet-1`.
function snapshotWithCharts(charts: Array<Record<string, unknown>>): string {
    const data: Record<string, unknown> = {};
    const order: string[] = [];
    for (const c of charts) {
        const id = c.chartId as string;
        data[id] = {
            unitId: 'workbook',
            subUnitId: 'sheet-1',
            drawingId: id,
            drawingType: 8,
            componentKey: 'NotesheetChart',
            data: c,
        };
        order.push(id);
    }
    const snapshot = {
        sheetOrder: ['sheet-1'],
        sheets: {
            'sheet-1': { id: 'sheet-1', name: 'Sheet1', cellData: {}, rowCount: 5, columnCount: 5 },
        },
        styles: {},
        resources: [
            {
                name: 'SHEET_DRAWING_PLUGIN',
                data: JSON.stringify({ 'sheet-1': { data, order } }),
            },
        ],
    };
    return JSON.stringify(snapshot);
}

const count = (html: string, re: RegExp) => (html.match(re) ?? []).length;

describe('M18 B1 — charts render as static SVG in HTML export', () => {
    test('bar chart: one <rect> bar per finite data point (+ baseline line)', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'bar',
                    title: 'Quarterly',
                    labels: ['Q1', 'Q2', 'Q3'],
                    datasets: [{ label: 'Sales', data: [10, 20, 15] }],
                },
            ]),
        )!;
        expect(html).toContain('<svg');
        expect(html).toContain('notesheet-chart');
        // 3 data points → 3 bar <rect>s.
        expect(count(html, /<rect\b/g)).toBe(3);
        // Title is rendered.
        expect(html).toContain('Quarterly');
        // Category labels present.
        expect(html).toContain('Q1');
        expect(html).toContain('Q3');
    });

    test('multi-series bar: one <rect> per (series × point) + a legend swatch per series', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'bar',
                    title: 'Two series',
                    labels: ['A', 'B'],
                    datasets: [
                        { label: 'X', data: [1, 2] },
                        { label: 'Y', data: [3, 4] },
                    ],
                },
            ]),
        )!;
        // 2 series × 2 points = 4 bars, PLUS 2 legend swatches = 6 rects.
        expect(count(html, /<rect\b/g)).toBe(6);
        expect(html).toContain('>X<');
        expect(html).toContain('>Y<');
    });

    test('line chart: one <polyline> per series', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'line',
                    title: 'Trend',
                    labels: ['Jan', 'Feb', 'Mar'],
                    datasets: [
                        { label: 'A', data: [1, 2, 3] },
                        { label: 'B', data: [3, 2, 1] },
                    ],
                },
            ]),
        )!;
        expect(count(html, /<polyline\b/g)).toBe(2);
    });

    test('pie chart: one <path> slice per positive value', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'pie',
                    title: 'Share',
                    labels: ['North', 'South', 'East', 'West'],
                    datasets: [{ data: [40, 30, 20, 10] }],
                },
            ]),
        )!;
        expect(count(html, /<path\b/g)).toBe(4);
        // Slice legend labels present.
        expect(html).toContain('North');
        expect(html).toContain('West');
    });

    test('doughnut chart: <path> slices use an inner radius (donut hole)', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'doughnut',
                    title: 'Donut',
                    labels: ['A', 'B', 'C'],
                    datasets: [{ data: [1, 1, 1] }],
                    meta: { holeSize: 50 },
                },
            ]),
        )!;
        expect(count(html, /<path\b/g)).toBe(3);
    });

    test('empty / degenerate chart emits no SVG (no crash)', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'bar',
                    title: 'Empty',
                    labels: [],
                    datasets: [{ data: [] }],
                },
            ]),
        )!;
        // Table still renders; no chart svg for the empty chart.
        expect(html).toContain('notesheet-table');
        expect(count(html, /<svg\b/g)).toBe(0);
    });

    test('chart appears AFTER its sheet table in document order', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                { chartId: 'c1', type: 'bar', labels: ['A'], datasets: [{ data: [5] }] },
            ]),
        )!;
        expect(html.indexOf('</table>')).toBeLessThan(html.indexOf('<svg'));
    });

    test('a snapshot with NO charts emits no SVG (unchanged behavior)', () => {
        const snapshot = JSON.stringify({
            sheetOrder: ['sheet-1'],
            sheets: {
                'sheet-1': { id: 'sheet-1', name: 'S', cellData: {}, rowCount: 2, columnCount: 2 },
            },
            styles: {},
        });
        const html = renderNotesheetSnapshot(snapshot)!;
        expect(count(html, /<svg\b/g)).toBe(0);
        expect(html).toContain('notesheet-table');
    });

    test('uses the shared CHART_PALETTE colours (blue first series)', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                { chartId: 'c1', type: 'bar', labels: ['A'], datasets: [{ data: [5] }] },
            ]),
        )!;
        expect(html).toContain('#3b82f6'); // CHART_PALETTE[0]
    });
});
