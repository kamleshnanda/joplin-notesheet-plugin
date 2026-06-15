// M18 C3: percentStacked charts normalise each category's series to 100%.
//
// Before C3, a percentStacked chart routed to Chart.js as a plain stacked
// chart — series accumulated to their raw totals, not normalised to 100%, so
// the chart looked like a regular stacked chart instead of Excel's
// "100% Stacked". C3 normalises each category column to percentages and caps
// the value axis at 100. Plain 'stacked' is unaffected.

import { buildConfig } from '../src/charts/NotesheetChart';

type ScaleAxis = { stacked?: boolean; max?: number };
function valAxis(cfg: ReturnType<typeof buildConfig>, key: 'x' | 'y'): ScaleAxis {
    const scales = (cfg.options as { scales?: Record<string, ScaleAxis> })?.scales ?? {};
    return scales[key] ?? {};
}

describe('M18 C3 — percentStacked normalisation', () => {
    const base = {
        chartId: 'c1',
        type: 'bar' as const,
        labels: ['Q1', 'Q2'],
        datasets: [
            { label: 'A', data: [30, 10], backgroundColor: '#3b82f6' },
            { label: 'B', data: [10, 30], backgroundColor: '#ef4444' },
        ],
    };

    test('percentStacked normalises each category column to sum 100', () => {
        const cfg = buildConfig({ ...base, meta: { barGrouping: 'percentStacked' } });
        const ds = cfg.data.datasets as Array<{ data: number[] }>;
        // Q1: 30 + 10 = 40 → 75% + 25%. Q2: 10 + 30 = 40 → 25% + 75%.
        expect(ds[0].data[0]).toBeCloseTo(75, 5);
        expect(ds[1].data[0]).toBeCloseTo(25, 5);
        expect(ds[0].data[1]).toBeCloseTo(25, 5);
        expect(ds[1].data[1]).toBeCloseTo(75, 5);
        // Each category column sums to 100.
        expect(ds[0].data[0] + ds[1].data[0]).toBeCloseTo(100, 5);
    });

    test('percentStacked caps the value axis at 100 and stacks', () => {
        const cfg = buildConfig({ ...base, meta: { barGrouping: 'percentStacked' } });
        const y = valAxis(cfg, 'y');
        expect(y.stacked).toBe(true);
        expect(y.max).toBe(100);
    });

    test('plain stacked is NOT normalised (raw values kept, no max cap)', () => {
        const cfg = buildConfig({ ...base, meta: { barGrouping: 'stacked' } });
        const ds = cfg.data.datasets as Array<{ data: number[] }>;
        expect(ds[0].data[0]).toBe(30); // unchanged
        expect(ds[1].data[1]).toBe(30);
        expect(valAxis(cfg, 'y').max).toBeUndefined();
    });

    test('a zero-total category leaves its (zero) values alone, no NaN', () => {
        const cfg = buildConfig({
            ...base,
            datasets: [
                { label: 'A', data: [0, 10], backgroundColor: '#3b82f6' },
                { label: 'B', data: [0, 30], backgroundColor: '#ef4444' },
            ],
            meta: { barGrouping: 'percentStacked' as const },
        });
        const ds = cfg.data.datasets as Array<{ data: number[] }>;
        // Q1 total 0 → both stay 0 (no division by zero).
        expect(ds[0].data[0]).toBe(0);
        expect(ds[1].data[0]).toBe(0);
        // Q2 total 40 → 25% / 75%.
        expect(ds[0].data[1]).toBeCloseTo(25, 5);
    });
});
