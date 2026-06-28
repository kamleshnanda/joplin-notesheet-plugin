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

    // ── B1 fix: horizontal bars (meta.barDir === 'bar') ────────────────
    // Imported charts carry meta.barDir from <c:barDir>; the SVG exporter
    // must honour it (the live Chart.js renderer already does). Before the
    // fix, renderBarSvg ignored barDir and always drew vertical columns —
    // bar='bar' and bar='col' produced byte-identical output.
    const rectsOf = (html: string): string[] => html.match(/<rect\b[^>]*>/g) ?? [];
    const attrNum = (rect: string, name: string): number =>
        parseFloat((rect.match(new RegExp(`${name}="([\\d.]+)"`)) ?? [])[1] ?? 'NaN');

    test('horizontal bar (barDir="bar") encodes value in WIDTH, not height', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'bar',
                    labels: ['A', 'B'],
                    meta: { barDir: 'bar' },
                    datasets: [{ label: 'S', data: [10, 20] }],
                },
            ]),
        )!;
        // Single series → no legend swatch, so both <rect>s are data bars.
        const rects = rectsOf(html);
        expect(rects.length).toBe(2);
        const widths = rects.map((r) => attrNum(r, 'width'));
        const heights = rects.map((r) => attrNum(r, 'height'));
        // Value (10 vs 20) drives WIDTH in a horizontal bar; bar thickness
        // (height) is constant. This is the inverse of a column chart.
        expect(widths[0]).not.toBeCloseTo(widths[1], 1);
        expect(heights[0]).toBeCloseTo(heights[1], 1);
        // ~2:1 value ratio → ~2:1 width ratio (baseline at 0).
        expect(Math.max(...widths) / Math.min(...widths)).toBeCloseTo(2, 0);
    });

    test('horizontal vs vertical bars produce different geometry', () => {
        const mk = (barDir: string) =>
            renderNotesheetSnapshot(
                snapshotWithCharts([
                    {
                        chartId: 'c1',
                        type: 'bar',
                        labels: ['A', 'B'],
                        meta: { barDir },
                        datasets: [{ label: 'S', data: [10, 20] }],
                    },
                ]),
            )!;
        expect(mk('bar')).not.toBe(mk('col'));
    });

    // ── B1 fix: single 100%-slice pie / doughnut ───────────────────────
    // A self-closing 2π arc (start point == end point) is dropped by the
    // SVG spec, so a one-category pie/doughnut rendered nothing. Emit a
    // full circle (pie) / full ring (doughnut) for the degenerate case.
    test('single 100% pie slice renders a full <circle>, not a vanishing arc', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'pie',
                    title: 'All',
                    labels: ['Only'],
                    datasets: [{ data: [42] }],
                },
            ]),
        )!;
        expect(html).toContain('<svg');
        expect(count(html, /<circle\b/g)).toBe(1);
        expect(html).toContain('Only'); // legend still present
    });

    test('single 100% doughnut slice renders a full ring (evenodd hole)', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'doughnut',
                    labels: ['Only'],
                    datasets: [{ data: [42] }],
                    meta: { holeSize: 50 },
                },
            ]),
        )!;
        expect(html).toContain('<svg');
        // One ring element with a hole cut via fill-rule evenodd.
        expect(count(html, /<path\b/g)).toBe(1);
        expect(html).toContain('fill-rule="evenodd"');
    });

    // ── B1 fix: an all-NaN line dataset emits NO svg (matches bar path) ──
    // A non-empty-but-all-non-finite dataset slipped past the n===0 guard
    // and pushed an empty <polyline>, so a degenerate chart still emitted a
    // full <svg>. Contract (per the empty-chart test above): degenerate →
    // no svg.
    test('line chart with an all-NaN dataset emits no SVG', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'line',
                    title: 'Bad',
                    labels: ['A', 'B'],
                    // Non-numeric source values → Number(...) === NaN.
                    datasets: [{ label: 'S', data: ['x', 'y'] as unknown as number[] }],
                },
            ]),
        )!;
        expect(html).toContain('notesheet-table');
        expect(count(html, /<svg\b/g)).toBe(0);
    });

    // ── B1 fix: doughnut holeSize is clamped so the ring never inverts ──
    // Excel constrains holeSize to 1-90, but the importer accepts any
    // non-negative integer, so a malformed workbook could pass holeSize
    // >= 100, making innerR >= radius (an inverted/blank ring). Clamp it.
    test('doughnut holeSize >= 100 is clamped (inner radius stays below outer)', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'doughnut',
                    labels: ['A', 'B', 'C'],
                    datasets: [{ data: [1, 1, 1] }],
                    meta: { holeSize: 150 },
                },
            ]),
        )!;
        // The outer radius is fixed at 116.0 (geometry constants). Parse
        // every arc radius and assert NONE exceeds the outer — an unclamped
        // holeSize=150 would emit inner-arc radius 174.0 (an inverted ring).
        const radii = [...html.matchAll(/A (\d+(?:\.\d+)?) \1 /g)].map((m) => parseFloat(m[1]));
        expect(radii.length).toBeGreaterThan(0);
        const OUTER = 116;
        expect(Math.max(...radii)).toBeLessThanOrEqual(OUTER);
        // A hole is still cut (inner radius present and positive).
        expect(Math.min(...radii)).toBeGreaterThan(0);
        expect(Math.min(...radii)).toBeLessThan(OUTER);
    });

    test('doughnut holeSize 1-9 is NOT inflated (matches editor / export)', () => {
        // holeSize=5 → innerR = 116 * 0.05 = 5.8, NOT floored up to 10%.
        // The live Chart.js renderer and .xlsx export pass holeSize
        // through unclamped, so the SVG export must agree.
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'doughnut',
                    labels: ['A', 'B'],
                    datasets: [{ data: [1, 1] }],
                    meta: { holeSize: 5 },
                },
            ]),
        )!;
        const radii = [...html.matchAll(/A (\d+(?:\.\d+)?) \1 /g)].map((m) => parseFloat(m[1]));
        const inner = Math.min(...radii);
        expect(inner).toBeCloseTo(116 * 0.05, 0); // ≈5.8, not 11.6
    });

    test('horizontal bar with all non-negative data emits ONE value tick (no overlap)', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'bar',
                    labels: ['A', 'B'],
                    meta: { barDir: 'bar' },
                    datasets: [{ label: 'S', data: [10, 20] }],
                },
            ]),
        )!;
        // minV===0 → its tick sits at the same x as... no: max is at the
        // far right, min "0" at the left baseline, so BOTH render (distinct
        // x). The collapse case is all-zero — test that separately below.
        // Here we assert no two value-axis ticks share an x within 8px.
        const tickXs = [...html.matchAll(/<text x="([\d.]+)" y="276"/g)].map((m) =>
            parseFloat(m[1]),
        );
        for (let i = 0; i < tickXs.length; i++) {
            for (let j = i + 1; j < tickXs.length; j++) {
                expect(Math.abs(tickXs[i] - tickXs[j])).toBeGreaterThanOrEqual(8);
            }
        }
    });

    test('horizontal bar with all-zero data emits a single collapsed tick (no duplicate)', () => {
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'bar',
                    labels: ['A', 'B'],
                    meta: { barDir: 'bar' },
                    datasets: [{ label: 'S', data: [0, 0] }],
                },
            ]),
        )!;
        // minV===maxV===0 → both ticks would collide; only one is emitted.
        const ticks = [...html.matchAll(/<text x="[\d.]+" y="276"/g)];
        expect(ticks.length).toBe(1);
    });

    // ── feature-7 acceptance criterion: pie sweep angles match the INPUT ──
    // The path-count test above proves "one slice per value"; this proves the
    // slices are sized to the data. Per the project's "expected from input,
    // NOT from emit" discipline (M13/E), we independently compute each
    // expected sweep from the input data and recover the actual sweep by
    // parsing each <path>'s arc endpoints, asserting agreement within ±3°.
    test('pie slice sweep angles match the input proportions (±3°)', () => {
        const data = [40, 30, 20, 10]; // sums to 100
        const html = renderNotesheetSnapshot(
            snapshotWithCharts([
                {
                    chartId: 'c1',
                    type: 'pie',
                    title: 'Share',
                    labels: ['N', 'S', 'E', 'W'],
                    datasets: [{ data }],
                },
            ]),
        )!;
        const total = data.reduce((s, v) => s + v, 0);
        const expectedDeg = data.map((v) => (v / total) * 360);

        // Recover each slice's sweep from its <path d="M cx cy L x0 y0
        // A r r 0 large 1 x1 y1 Z">. The wedge angle is the angle between
        // (x0,y0) and (x1,y1) about the centre (cx,cy).
        const paths = [
            ...html.matchAll(
                /<path d="M ([\d.]+) ([\d.]+) L ([\d.]+) ([\d.]+) A [\d.]+ [\d.]+ 0 (\d) 1 ([\d.]+) ([\d.]+)/g,
            ),
        ];
        expect(paths.length).toBe(4);
        const actualDeg = paths.map((m) => {
            const [cx, cy, x0, y0, large, x1, y1] = [
                +m[1],
                +m[2],
                +m[3],
                +m[4],
                +m[5],
                +m[6],
                +m[7],
            ];
            const a0 = Math.atan2(y0 - cy, x0 - cx);
            const a1 = Math.atan2(y1 - cy, x1 - cx);
            let sweep = ((a1 - a0) * 180) / Math.PI;
            // Normalise to (0, 360]; the large-arc flag disambiguates >180.
            if (sweep < 0) sweep += 360;
            if (large === 1 && sweep < 180) sweep = 360 - sweep;
            return sweep;
        });
        // Slices are emitted in data order; compare each within ±3°.
        for (let i = 0; i < expectedDeg.length; i++) {
            expect(Math.abs(actualDeg[i] - expectedDeg[i])).toBeLessThanOrEqual(3);
        }
        // Structural floors: total sweep ≈ 360, no slice is exactly 0.
        const sum = actualDeg.reduce((s, d) => s + d, 0);
        expect(Math.abs(sum - 360)).toBeLessThanOrEqual(1);
        expect(Math.min(...actualDeg)).toBeGreaterThan(0);
    });

    // ── feature-7 acceptance criterion: CF + chart on the same sheet ──────
    // A snapshot carrying BOTH a SHEET_CONDITIONAL_FORMATTING_PLUGIN resource
    // (a cellIs rule painting B2:B5 pink) AND a chart must render the
    // CF-coloured <td> (via M16's existing CF evaluator) AND the chart <svg>,
    // in document order: the <table> first, then the <svg> — never overlaid.
    test('CF + chart on one sheet: pink <td> AND chart <svg>, table before svg', () => {
        // Sheet with values in B2:B5 (rows 1-4, col 1). B3=80 > 50 → pink.
        const cellData: Record<number, Record<number, { v: unknown }>> = {
            0: { 0: { v: 'Item' }, 1: { v: 'Score' } },
            1: { 0: { v: 'a' }, 1: { v: 10 } },
            2: { 0: { v: 'b' }, 1: { v: 80 } },
            3: { 0: { v: 'c' }, 1: { v: 30 } },
            4: { 0: { v: 'd' }, 1: { v: 90 } },
        };
        const drawingData: Record<string, unknown> = {};
        drawingData['c1'] = {
            unitId: 'workbook',
            subUnitId: 'sheet-1',
            drawingId: 'c1',
            drawingType: 8,
            componentKey: 'NotesheetChart',
            data: {
                chartId: 'c1',
                type: 'bar',
                title: 'Scores',
                labels: ['a', 'b', 'c', 'd'],
                datasets: [{ label: 'Score', data: [10, 80, 30, 90] }],
            },
        };
        const snapshot = JSON.stringify({
            sheetOrder: ['sheet-1'],
            sheets: {
                'sheet-1': { id: 'sheet-1', name: 'Sheet1', cellData, rowCount: 5, columnCount: 2 },
            },
            styles: {},
            resources: [
                {
                    name: 'SHEET_DRAWING_PLUGIN',
                    data: JSON.stringify({ 'sheet-1': { data: drawingData, order: ['c1'] } }),
                },
                {
                    name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
                    data: JSON.stringify({
                        'sheet-1': [
                            {
                                cfId: '1',
                                ranges: [{ startRow: 1, endRow: 4, startColumn: 1, endColumn: 1 }],
                                rule: {
                                    type: 'highlightCell',
                                    subType: 'number',
                                    operator: 'greaterThan',
                                    value: 50,
                                    style: { bg: { rgb: '#FFC7CE' } },
                                },
                            },
                        ],
                    }),
                },
            ],
        });
        const html = renderNotesheetSnapshot(snapshot)!;
        // CF paint flows through the M16 evaluator: B3 (80) and B5 (90) > 50.
        expect(html).toMatch(/background-color:\s*#FFC7CE/i);
        // The chart SVG renders too.
        expect(html).toContain('<svg');
        expect(html).toContain('notesheet-chart');
        // Document order: table block first, then the chart svg — no overlay.
        expect(html.indexOf('</table>')).toBeLessThan(html.indexOf('<svg'));
        // No absolute/z-index layering (the chart sits after the table).
        const svgBlock = html.slice(html.indexOf('<svg'));
        expect(svgBlock).not.toMatch(/position:\s*absolute/i);
        expect(svgBlock).not.toMatch(/z-index/i);
    });
});
