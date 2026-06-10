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

Chart.register(...registerables);

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
    const scales: Record<string, { stacked?: boolean }> | undefined =
        (isStacked && (type === 'bar' || type === 'line'))
            ? { x: { stacked: true }, y: { stacked: true } }
            : undefined;

    return {
        type,
        data: { labels, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            ...(type === 'bar' ? { indexAxis } : {}),
            ...(scales ? { scales } : {}),
            plugins: {
                legend: {
                    display: datasets.length > 1 || type === 'pie' || type === 'doughnut',
                    position: legendPosition,
                },
                title: data?.title
                    ? { display: true, text: data.title, font: { size: 14 } }
                    : { display: false },
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
