// React component rendered inside Univer's float-DOM overlay. Receives a
// chart config via the `data` prop (which Univer passes through verbatim from
// addFloatDomToRange's `data` field). Univer's drawing service handles the
// floating div, its drag/resize chrome, and snapshot persistence — we only
// own the canvas inside it.

import * as React from 'react';
import { Chart, registerables, type ChartConfiguration, type ChartType } from 'chart.js';
import type { ChartData, RangeAddress } from './extractData';
import { CHART_PALETTE } from './extractData';
import { subscribeChartUpdate } from './dataBus';
import { notesheetPieLabelsPlugin } from './pieLabelsPlugin';

// registerables covers the core controllers/elements/scales. We register
// our OWN pie-label plugin separately. It paints pie/doughnut slice labels
// the way Excel does — roomy slices get an inside centroid label, small/
// crowded slices get pushed outside with a leader line — driven by
// meta.dLbls (showCatName/showVal/showPercent). It activates only when
// options.plugins.notesheetPieLabels.enabled is set (below), so bar/line
// charts and pies with labels off are untouched. Chart.js core can't draw
// per-slice labels; the EXPORTED .xlsx carries them independently.
Chart.register(...registerables);
Chart.register(notesheetPieLabelsPlugin);

export type NotesheetChartType = 'bar' | 'line' | 'pie' | 'doughnut';

export interface NotesheetChartData {
    chartId?: string;
    type?: NotesheetChartType;
    sourceRange?: RangeAddress;
    title?: string;
    labels?: string[];
    datasets?: ChartData['datasets'];
    // M17: per-chart metadata that doesn't fit ChartType. Today carries:
    //   - barDir: 'bar' (horizontal) or 'col' (vertical, the default).
    //     Source: <c:barDir> in OOXML chart XML, surfaced by xlsxChartImport
    //     for type === 'bar' fixtures. NotesheetChart routes this to
    //     Chart.js's `options.indexAxis = 'y'` for horizontal bars.
    //   - unsupportedSourceType: name of an unsupported chart type
    //     (radar, scatter, area...) that fell back to 'bar'.
    meta?: {
        barDir?: 'bar' | 'col';
        // 'clustered' (default), 'stacked', 'percentStacked', or
        // 'standard' (line only). For bar charts we set the categorical
        // axis stacked + dataset.stack='stack0' so series stack on the
        // same X key. For line charts, 'stacked' makes each dataset
        // fill from the previous dataset's curve.
        barGrouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
        // Excel's gap-between-bar-groups, expressed as a percentage of
        // bar width. 150 = Excel default (gap = 1.5x bar width). Routes
        // to Chart.js's `dataset.categoryPercentage` via the formula
        // 100 / (100 + gapWidth) — so 0 fills the band, 150 makes bars
        // 0.4 of the band, 500 makes bars 0.17 of the band.
        barGapWidth?: number;
        unsupportedSourceType?: string;
        // Where the legend sits. Source: <c:legendPos val="..."/> at
        // import time. Excel values 'r'|'l'|'t'|'b'|'tr' map to Chart.js
        // 'right'|'left'|'top'|'bottom'|'right' (Chart.js has no 'tr').
        legendPos?: 'r' | 'l' | 't' | 'b' | 'tr';
        // 'index' when the source chart had no <c:cat> element. We
        // synthesize 1..N labels to match Excel's "row index as X-axis"
        // behaviour, and stop interpreting column 0 as a label column
        // (the import resolver also skips that fallback).
        categoryAxisType?: 'index' | 'category';
        // Doughnut hole diameter %. Routes to Chart.js's
        // `options.cutout = "${holeSize}%"`.
        holeSize?: number;
        // Line smoothing — applied per-dataset via Chart.js `tension`.
        // 0 = sharp segments, ~0.4 = a typical smooth spline.
        lineSmooth?: boolean;
        // Show data-point markers on lines. Routes to Chart.js's
        // `pointRadius` per dataset (3 = visible, 0 = hidden).
        lineMarkerOn?: boolean;
        // dispBlanksAs ('gap'|'zero'|'span') and crossBetween
        // ('between'|'midCat') round-trip via export but have no
        // Chart.js equivalents that materially change the render.
        // Stored here so M10 export can re-emit them; the runtime
        // ignores them.
        dispBlanksAs?: 'gap' | 'zero' | 'span';
        crossBetween?: 'between' | 'midCat';
        // Tick mark style on bar/line axes; round-tripped only.
        tickMark?: 'none' | 'in' | 'out' | 'cross';
        // Number-format string for value-axis tick labels. Only the
        // small set of patterns we care about are supported at runtime
        // (percentages — "0%", "0.00%"; currency dollars; integers).
        // Anything else falls back to Chart.js's default formatter.
        valAxisNumFmt?: string;
        // Chart-level data-label visibility flags. Pie/doughnut slice
        // labels are the most visible payoff (showVal/showPercent at
        // chart level). Bar/line series may also pick these up.
        dLbls?: {
            showVal?: boolean;
            showCatName?: boolean;
            showPercent?: boolean;
            showSerName?: boolean;
            showLegendKey?: boolean;
            showBubbleSize?: boolean;
        };
    };
}

interface Props {
    data?: NotesheetChartData;
}

const CONTAINER_STYLE: React.CSSProperties = {
    width: '100%',
    height: '100%',
    background: '#fff',
    border: '1px solid #d1d5db',
    boxSizing: 'border-box',
    padding: '8px',
};

// Build a Chart.js tick-callback that formats raw numeric values per
// a small set of Excel-style numFmt strings. Handles only what the
// project's chart fixtures use today; falls back to default toString
// for anything unrecognized so a typo or new pattern doesn't blank
// the axis labels.
function makeNumFmtFormatter(fmt: string): (value: number | string) => string {
    // Percentage: "0%", "0.0%", "0.00%". Multiply raw value by 100,
    // round to N decimals, append "%". Excel convention: a cell with
    // value 0.05 and format "0%" displays as "5%".
    const pctMatch = fmt.match(/^0(?:\.(0+))?%$/);
    if (pctMatch) {
        const decimals = pctMatch[1] ? pctMatch[1].length : 0;
        return (v) => {
            const n = typeof v === 'number' ? v : Number(v);
            if (!Number.isFinite(n)) return String(v);
            return `${(n * 100).toFixed(decimals)}%`;
        };
    }
    // Currency dollars: "$0", "$#,##0", "$#,##0.00".
    const curMatch = fmt.match(/^\$(#,##)?0(?:\.(0+))?$/);
    if (curMatch) {
        const useCommas = !!curMatch[1];
        const decimals = curMatch[2] ? curMatch[2].length : 0;
        return (v) => {
            const n = typeof v === 'number' ? v : Number(v);
            if (!Number.isFinite(n)) return String(v);
            const fixed = n.toFixed(decimals);
            if (!useCommas) return `$${fixed}`;
            const [whole, frac] = fixed.split('.');
            const withCommas = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
            return frac ? `$${withCommas}.${frac}` : `$${withCommas}`;
        };
    }
    // Integer with optional thousands: "0", "#,##0".
    if (fmt === '0' || fmt === '#,##0') {
        const useCommas = fmt === '#,##0';
        return (v) => {
            const n = typeof v === 'number' ? v : Number(v);
            if (!Number.isFinite(n)) return String(v);
            const rounded = Math.round(n).toString();
            return useCommas ? rounded.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : rounded;
        };
    }
    return (v) => String(v);
}

function buildConfig(data: NotesheetChartData | undefined): ChartConfiguration {
    const type = (data?.type ?? 'bar') as ChartType;
    let labels = data?.labels ?? [];
    let datasets = data?.datasets ?? [];

    // No <c:cat> in the source chart: Excel uses row index 1..N as the
    // X-axis. Synthesize matching labels here. When labels are already
    // populated this branch is a no-op.
    if (data?.meta?.categoryAxisType === 'index' && labels.length === 0 && datasets.length > 0) {
        const n = Math.max(...datasets.map((d) => d.data?.length ?? 0));
        labels = Array.from({ length: n }, (_, i) => String(i + 1));
    }

    // Pie/doughnut want one dataset whose backgroundColor is an array (one
    // slice per data point). Bar/line use one color per series.
    if ((type === 'pie' || type === 'doughnut') && datasets.length > 0) {
        const ds = datasets[0];
        datasets = [{
            ...ds,
            backgroundColor: ds.data.map((_, i) => CHART_PALETTE[i % CHART_PALETTE.length]),
        }];
    }

    // Bar orientation. OOXML <c:barDir val="bar"/> = horizontal,
    // val="col" = vertical (Excel's default). Chart.js expresses this
    // via `options.indexAxis`: 'y' = horizontal, 'x' (default) = vertical.
    // Only meaningful when type === 'bar'.
    const indexAxis = (type === 'bar' && data?.meta?.barDir === 'bar') ? 'y' : 'x';

    // Grouping. Excel: 'clustered' | 'stacked' | 'percentStacked' |
    // 'standard'. For bar charts we set the categorical axis stacked
    // and tag each dataset with `stack: 'stack0'` so Chart.js stacks
    // all series on the same X (or Y, for horizontal bars) tick.
    // For line charts, 'stacked' applies via dataset.fill chain.
    const grouping = data?.meta?.barGrouping;
    const isStacked = grouping === 'stacked' || grouping === 'percentStacked';
    if (type === 'bar' && isStacked) {
        datasets = datasets.map((ds) => ({ ...ds, stack: 'stack0' } as ChartData['datasets'][number] & { stack: string }));
    }
    if (type === 'line' && isStacked) {
        // Chart.js stacked-line: each dataset fills from the curve below
        // it. First dataset fills from the X-axis (origin); subsequent
        // datasets fill from the dataset at index-1.
        datasets = datasets.map((ds, i) => ({ ...ds, fill: i === 0 ? 'origin' : ('-1' as const) } as ChartData['datasets'][number] & { fill: string }));
    }

    // Bar width. Excel's <c:gapWidth val="N"/> is "gap between bar
    // groups as a percentage of bar width". Chart.js's
    // categoryPercentage is "what fraction of the category band the
    // bars occupy". gap = N% of barWidth means gap + barWidth =
    // (1 + N/100) * barWidth, so categoryPercentage = 100/(100+N).
    // Excel default 150 → 0.4 (Chart.js default is 0.8 — much wider).
    if (type === 'bar' && typeof data?.meta?.barGapWidth === 'number' && data.meta.barGapWidth >= 0) {
        const cp = 100 / (100 + data.meta.barGapWidth);
        datasets = datasets.map((ds) => ({
            ...ds,
            categoryPercentage: cp,
            barPercentage: 1.0,
        } as ChartData['datasets'][number] & { categoryPercentage: number; barPercentage: number }));
    }

    // Line smoothing — Chart.js `tension`. 0 = polylines, 0.4 = a
    // pleasant Bézier spline matching what Excel's "Smoothed Line"
    // option looks like.
    if (type === 'line' && data?.meta?.lineSmooth) {
        datasets = datasets.map((ds) => ({
            ...ds,
            tension: 0.4,
        } as ChartData['datasets'][number] & { tension: number }));
    }

    // Line markers on/off. Chart.js: pointRadius=0 hides points;
    // a non-zero value shows them. We pick 3 for "on" (matches
    // Chart.js's default visible-marker size).
    if (type === 'line') {
        const radius = data?.meta?.lineMarkerOn ? 3 : 0;
        datasets = datasets.map((ds) => ({
            ...ds,
            pointRadius: radius,
        } as ChartData['datasets'][number] & { pointRadius: number }));
    }

    // Legend position. Excel cf. <c:legendPos>: r/l/t/b/tr. Chart.js
    // accepts 'right' | 'left' | 'top' | 'bottom' | 'chartArea'; 'tr'
    // (Excel's "top-right") collapses to 'right' since Chart.js has no
    // matching position. Default to 'right' (Excel's default and
    // Chart.js's default for Notesheet's component).
    const legendPosMap: Record<string, 'right' | 'left' | 'top' | 'bottom'> = {
        r: 'right', l: 'left', t: 'top', b: 'bottom', tr: 'right',
    };
    const legendPosition = data?.meta?.legendPos
        ? (legendPosMap[data.meta.legendPos] ?? 'right')
        : 'right';

    // Stacked scales. For bar charts: set both axes stacked when
    // grouping is stacked/percentStacked (Chart.js needs both, even
    // though only the value axis actually accumulates). For line:
    // value axis only. percentStacked normalises to 100% via Chart.js's
    // y.max=100 + a custom tooltip — we keep it simple and just stack;
    // a follow-up could refine percent normalisation.
    // Value-axis number format. Excel <c:numFmt formatCode="0%"/> on
    // the valAx (09-bar-percent-axis) means "treat raw 0.05 as 5%."
    // Chart.js doesn't read OOXML formats — we need a tick callback.
    // Support a small set of patterns end-users actually ship:
    //   "0%" / "0.0%" / "0.00%" — percentage, n digits past decimal
    //   "$0" / "$#,##0" / "$#,##0.00" — currency dollars
    //   "0" / "#,##0" — integer with optional thousands separator
    // Anything else falls back to Chart.js's default toString.
    const valNumFmt = data?.meta?.valAxisNumFmt;
    const tickFormatter = valNumFmt ? makeNumFmtFormatter(valNumFmt) : null;

    // Bar orientation matters for which axis carries the values.
    // Vertical bars / lines: y is values. Horizontal bars: x is values.
    const valAxisKey = (type === 'bar' && data?.meta?.barDir === 'bar') ? 'x' : 'y';

    // Horizontal bars (barDir='bar' → indexAxis='y'): Chart.js draws the
    // FIRST category at the TOP of the y-axis and counts down; Excel draws
    // the first category at the BOTTOM and counts up. Left alone, an
    // imported horizontal-bar chart renders its categories in the
    // opposite order from the source (issue 10 / task #24 — the
    // "trend bars totally reversed" report on 10-bar-with-trendline).
    // Reversing the CATEGORY axis (the y-axis for horizontal bars) puts
    // the first category back at the bottom, matching Excel. Vertical
    // bars and line charts are unaffected (their category axis is x and
    // Chart.js already matches Excel's left-to-right order).
    const isHorizontalBar = type === 'bar' && data?.meta?.barDir === 'bar';
    const catAxisKey = isHorizontalBar ? 'y' : 'x';

    const scales: Record<string, Record<string, unknown>> | undefined = (() => {
        const out: Record<string, Record<string, unknown>> = {};
        if (isStacked && (type === 'bar' || type === 'line')) {
            out.x = { ...(out.x ?? {}), stacked: true };
            out.y = { ...(out.y ?? {}), stacked: true };
        }
        if (tickFormatter && (type === 'bar' || type === 'line')) {
            out[valAxisKey] = {
                ...(out[valAxisKey] ?? {}),
                ticks: { callback: tickFormatter },
            };
        }
        if (isHorizontalBar) {
            out[catAxisKey] = { ...(out[catAxisKey] ?? {}), reverse: true };
        }
        return Object.keys(out).length > 0 ? out : undefined;
    })();

    // Doughnut hole. Chart.js's `cutout` accepts pixel ints or `"N%"`
    // strings; we use percent so it scales with the chart. Excel's
    // value is a 1-90 percent. Default to 50 (Excel and Chart.js's
    // typical doughnut). Pie charts ignore this.
    const cutout = type === 'doughnut'
        ? (typeof data?.meta?.holeSize === 'number' ? `${data.meta.holeSize}%` : '50%')
        : undefined;

    // Slice data labels (pie/doughnut). Excel keeps these flags on
    // meta.dLbls (showCatName / showPercent / showVal). Our custom
    // notesheetPieLabelsPlugin reads these options and paints inside /
    // pushed-out-with-leader-line labels like Excel. Enabled only for
    // pie/doughnut when at least one label flag is set, so bar/line
    // charts (and pies with labels off) are never touched.
    const dl = data?.meta?.dLbls;
    const wantsSliceLabels = (type === 'pie' || type === 'doughnut')
        && !!dl && !!(dl.showCatName || dl.showPercent || dl.showVal);
    const pieLabelsConfig = {
        enabled: wantsSliceLabels,
        flags: {
            showCatName: !!dl?.showCatName,
            showVal: !!dl?.showVal,
            showPercent: !!dl?.showPercent,
        },
    };

    return {
        type,
        data: { labels, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            ...(type === 'bar' ? { indexAxis } : {}),
            ...(scales ? { scales } : {}),
            ...(cutout !== undefined ? { cutout } : {}),
            // When slice labels are on, reserve horizontal padding so the
            // pie shrinks and there's a gutter on each side for pushed-out
            // small-slice labels + their leader lines (they'd otherwise be
            // clipped at the chart's left/right edge).
            ...(wantsSliceLabels ? { layout: { padding: { left: 96, right: 96, top: 8 } } } : {}),
            plugins: {
                legend: {
                    display: datasets.length > 1 || type === 'pie' || type === 'doughnut',
                    position: legendPosition,
                },
                title: data?.title
                    ? { display: true, text: data.title, font: { size: 14 } }
                    : { display: false },
                notesheetPieLabels: pieLabelsConfig,
            },
        },
    };
}

const NotesheetChart: React.FC<Props> = ({ data }) => {
    const canvasRef = React.useRef<HTMLCanvasElement | null>(null);
    const chartRef = React.useRef<Chart | null>(null);
    const lastTypeRef = React.useRef<ChartType | null>(null);

    // Live data overrides whatever was baked into props.data at insert time.
    // Set by the dataBus subscription on every source-range edit.
    const [liveData, setLiveData] = React.useState<ChartData | null>(null);

    React.useEffect(() => {
        const id = data?.chartId;
        if (!id) return;
        const off = subscribeChartUpdate(id, (next) => {
            setLiveData(next);
        });
        return off;
    }, [data?.chartId]);

    const merged: NotesheetChartData = React.useMemo(() => {
        if (!liveData) return data ?? {};
        return { ...(data ?? {}), labels: liveData.labels, datasets: liveData.datasets };
    }, [data, liveData]);

    React.useEffect(() => {
        const canvas = canvasRef.current;
        if (!canvas) return;
        const config = buildConfig(merged);

        // Type change requires recreate; everything else can update in place.
        if (chartRef.current && lastTypeRef.current === config.type) {
            try {
                chartRef.current.data = config.data;
                chartRef.current.options = config.options ?? {};
                chartRef.current.update();
                return;
            } catch (e) {
                console.error('[Notesheet] chart update failed', e);
            }
        }

        if (chartRef.current) {
            try { chartRef.current.destroy(); } catch { /* ignore */ }
            chartRef.current = null;
        }

        try {
            chartRef.current = new Chart(canvas, config);
            lastTypeRef.current = config.type;
        } catch (e) {
            console.error('[Notesheet] chart create failed', e);
        }
    }, [merged]);

    // M17 feature-3 followup: imported charts mount inside Univer's
    // float-DOM whose container size is 0×0 at the moment React's
    // create-Chart effect runs (the float-DOM transform settles a beat
    // later). Chart.js's `responsive: true` only resizes on window
    // events, NOT on container-only changes — so the chart paints to
    // a 0×0 surface and stays blank until the user resizes the window
    // or the float-DOM. Watch the container with a ResizeObserver and
    // call chart.resize() on every dimension change. Cheap (one-shot
    // per layout pass) and idempotent.
    React.useEffect(() => {
        const canvas = canvasRef.current;
        const container = canvas?.parentElement ?? null;
        if (!container) return;
        if (typeof ResizeObserver === 'undefined') return;
        const ro = new ResizeObserver(() => {
            if (chartRef.current) {
                try { chartRef.current.resize(); } catch { /* ignore */ }
            }
        });
        ro.observe(container);
        return () => { ro.disconnect(); };
    }, []);

    React.useEffect(() => {
        return () => {
            if (chartRef.current) {
                try { chartRef.current.destroy(); } catch { /* ignore */ }
                chartRef.current = null;
            }
        };
    }, []);

    return (
        <div style={CONTAINER_STYLE}>
            <canvas ref={canvasRef} style={{ width: '100%', height: '100%' }} />
        </div>
    );
};

export default NotesheetChart;
