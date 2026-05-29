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
    const labels = data?.labels ?? [];
    let datasets = data?.datasets ?? [];

    // Pie/doughnut want one dataset whose backgroundColor is an array (one
    // slice per data point). Bar/line use one color per series.
    if ((type === 'pie' || type === 'doughnut') && datasets.length > 0) {
        const ds = datasets[0];
        datasets = [{
            ...ds,
            backgroundColor: ds.data.map((_, i) => CHART_PALETTE[i % CHART_PALETTE.length]),
        }];
    }

    return {
        type,
        data: { labels, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            plugins: {
                legend: { display: datasets.length > 1 || type === 'pie' || type === 'doughnut' },
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
