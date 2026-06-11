// Custom Chart.js v4 plugin: Excel-style pie / doughnut slice labels.
//
// WHY THIS EXISTS. chartjs-plugin-datalabels places every label at the
// slice centroid. For thin adjacent slices the centroids nearly coincide
// near the pie center, so the labels smear into each other (observed on
// 06-two-charts-one-sheet: "Latin America 10%" and "Middle East & Africa
// 5%" overlapped illegibly). Excel instead keeps roomy slices' labels
// INSIDE at the centroid and pushes small/crowded slices' labels OUTSIDE
// the pie with a leader line connecting label → slice. No off-the-shelf
// Chart.js v4 plugin does leader lines (the only `outlabels` plugin on
// npm is locked to Chart.js 2.x), so we hand-roll one.
//
// BEHAVIOUR (mirrors Excel "best fit" label placement):
//   - Compute each label's multi-line text from the dLbls flags.
//   - Try to place it INSIDE at the slice centroid. If the text box fits
//     within the slice's angular + radial bounds, draw it there in white.
//   - Otherwise place it OUTSIDE: anchor a point just past the slice's
//     mid-angle on the pie rim, extend radially to an outer label point,
//     and draw a leader line (rim → elbow → label) in the slice colour's
//     muted grey, with the label text in dark grey near the chart edge.
//   - De-overlap the OUTSIDE labels per side (left/right) by pushing them
//     apart vertically so stacked callouts (the 10% + 5% case) separate,
//     exactly like Excel's stacked "Latin America" / "Middle East" labels.
//
// This is a pure render-time plugin (afterDatasetsDraw) — it reads slice
// geometry from the arc elements Chart.js already laid out and paints on
// the same canvas context. It owns NO state and mutates nothing.

import type { Chart, ChartType, Plugin } from 'chart.js';

export interface PieLabelFlags {
    showCatName?: boolean;
    showVal?: boolean;
    showPercent?: boolean;
}

// Per-chart options the renderer attaches under options.plugins.notesheetPieLabels.
export interface NotesheetPieLabelsOptions {
    enabled: boolean;
    flags: PieLabelFlags;
    // Optional trendline equation/R² text (bar/line charts). When set, the
    // plugin draws it as a small floating label near the top of the plot,
    // matching how Excel shows the trendline equation as a chart label
    // rather than a legend entry.
    trendlineText?: string | null;
}

// Teach Chart.js's typed plugin-options map about our custom plugin so
// `options.plugins.notesheetPieLabels` typechecks at the call site.
declare module 'chart.js' {
    interface PluginOptionsByType<TType extends ChartType> {
        notesheetPieLabels?: NotesheetPieLabelsOptions;
    }
}

const PLUGIN_ID = 'notesheetPieLabels';

// Build the label text for one slice from the dLbls flags. Mirrors the
// part order Excel/our datalabels formatter used: category, value, percent.
function labelLines(
    flags: PieLabelFlags,
    label: unknown,
    value: number,
    total: number,
): string[] {
    const lines: string[] = [];
    if (flags.showCatName && label != null) lines.push(String(label));
    if (flags.showVal) lines.push(String(value));
    if (flags.showPercent && total > 0) lines.push(`${Math.round((value / total) * 100)}%`);
    return lines;
}

// Word-wrap each line to a max pixel width (the chart's float-DOM canvas
// is narrow — ~461px — so a long category like "Middle East & Africa"
// must wrap, exactly as Excel wraps it in a constrained box rather than
// run off the edge).
function wrapLines(ctx: CanvasRenderingContext2D, lines: string[], maxW: number): string[] {
    const out: string[] = [];
    for (const line of lines) {
        const words = line.split(/\s+/);
        let cur = '';
        for (const w of words) {
            const trial = cur ? `${cur} ${w}` : w;
            if (ctx.measureText(trial).width <= maxW || !cur) {
                cur = trial;
            } else {
                out.push(cur);
                cur = w;
            }
        }
        if (cur) out.push(cur);
    }
    return out;
}

// Measure the widest line so we can fit-test / position the text box.
function measure(ctx: CanvasRenderingContext2D, lines: string[]): { w: number; h: number } {
    let w = 0;
    for (const ln of lines) {
        const m = ctx.measureText(ln);
        if (m.width > w) w = m.width;
    }
    const lineH = 14; // matches the 11px bold font set below, with leading
    return { w, h: lines.length * lineH };
}

interface PlacedOutside {
    index: number;
    lines: string[];
    side: 'left' | 'right';
    // anchor on the pie rim (where the leader line starts)
    rimX: number;
    rimY: number;
    // label box center (mutated during de-overlap)
    labelX: number;
    labelY: number;
    color: string;
    boxW: number;
    boxH: number;
}

export const notesheetPieLabelsPlugin: Plugin = {
    id: PLUGIN_ID,

    afterDatasetsDraw(chart: Chart) {
        const opts = (chart.options.plugins as Record<string, unknown> | undefined)?.[PLUGIN_ID] as
            | NotesheetPieLabelsOptions
            | undefined;
        if (!opts) return;

        // Trendline equation/R² annotation (bar/line). Drawn as a small
        // grey floating label near the top-right of the plot, independent
        // of the pie-label path. Matches Excel showing the trendline
        // equation as a chart label, not a legend entry.
        if (opts.trendlineText) {
            const ctx = chart.ctx;
            const { chartArea } = chart;
            ctx.save();
            ctx.font = '11px sans-serif';
            ctx.fillStyle = '#404040';
            ctx.textAlign = 'right';
            ctx.textBaseline = 'top';
            ctx.fillText(opts.trendlineText, chartArea.right - 6, chartArea.top + 4);
            ctx.restore();
        }

        if (!opts.enabled) return;

        const meta = chart.getDatasetMeta(0);
        if (!meta || !meta.data || meta.data.length === 0) return;
        const dataset = chart.data.datasets[0];
        const values = (dataset?.data ?? []) as number[];
        const total = values.reduce((s, n) => s + (typeof n === 'number' ? n : 0), 0);
        if (total <= 0) return;

        const ctx = chart.ctx;
        const { chartArea } = chart;
        const cx = (chartArea.left + chartArea.right) / 2;
        const cy = (chartArea.top + chartArea.bottom) / 2;

        ctx.save();
        ctx.font = 'bold 11px sans-serif';
        ctx.textBaseline = 'middle';

        const outside: PlacedOutside[] = [];

        for (let i = 0; i < meta.data.length; i++) {
            const arc = meta.data[i] as unknown as {
                startAngle: number;
                endAngle: number;
                outerRadius: number;
                innerRadius: number;
                options?: { backgroundColor?: string };
            };
            const value = typeof values[i] === 'number' ? values[i] : 0;
            if (value <= 0) continue;

            const lines = labelLines(opts.flags, chart.data.labels?.[i], value, total);
            if (lines.length === 0) continue;

            const mid = (arc.startAngle + arc.endAngle) / 2;
            const sweep = arc.endAngle - arc.startAngle;
            const outerR = arc.outerRadius;
            const innerR = arc.innerRadius;
            const box = measure(ctx, lines);

            // Centroid radius (mid between inner and outer for doughnuts, ~0.6R
            // for pies — matches Chart.js datalabels' default anchor).
            const centroidR = innerR + (outerR - innerR) * 0.6;
            const centroidX = cx + Math.cos(mid) * centroidR;
            const centroidY = cy + Math.sin(mid) * centroidR;

            // Fit test: does the text box fit inside the slice? Approximate
            // the available tangential width at the centroid radius by the
            // arc chord (2*R*sin(sweep/2)), and require some radial room too.
            const chord = 2 * centroidR * Math.sin(Math.min(sweep, Math.PI) / 2);
            const radialRoom = outerR - innerR;
            const fitsInside = box.w + 8 <= chord && box.h + 4 <= radialRoom && sweep > 0.45;

            if (fitsInside) {
                // INSIDE: white text centred at the centroid.
                ctx.fillStyle = '#ffffff';
                ctx.textAlign = 'center';
                drawLines(ctx, lines, centroidX, centroidY);
            } else {
                // OUTSIDE: collect for a second pass (so we can de-overlap
                // per side before drawing leader lines + text).
                const side: 'left' | 'right' = Math.cos(mid) >= 0 ? 'right' : 'left';
                const rimX = cx + Math.cos(mid) * outerR;
                const rimY = cy + Math.sin(mid) * outerR;
                // Initial label point: a bit beyond the rim along the radius.
                const labelR = outerR + 24;
                const labelX = cx + Math.cos(mid) * labelR;
                const labelY = cy + Math.sin(mid) * labelR;
                const bg = arc.options?.backgroundColor;
                // Wrap to the gutter width available between the plot edge
                // and the canvas edge so long categories don't run off.
                const gutterW = (side === 'left' ? chartArea.left : chart.width - chartArea.right) - 12;
                const wrapped = wrapLines(ctx, lines, Math.max(40, gutterW));
                const wbox = measure(ctx, wrapped);
                outside.push({
                    index: i,
                    lines: wrapped,
                    side,
                    rimX,
                    rimY,
                    labelX,
                    labelY,
                    color: typeof bg === 'string' ? bg : '#888888',
                    boxW: wbox.w,
                    boxH: wbox.h,
                });
            }
        }

        // Place + de-overlap the outside labels in a per-side gutter near
        // the chart edge, then draw each with a rim → kink → gutter leader
        // line. Routing them to a side gutter (rather than straight out
        // along the radius) keeps top-sector small slices — the 10% + 5%
        // case, whose mid-angles both point near the top — from stacking up
        // INTO the chart title. This mirrors how Excel lays its callouts
        // along the side. Y is clamped within the plot area, leaving room
        // at the top for the title.
        const titleRoom = 22; // px reserved at the top for the chart title
        const yMin = chartArea.top + titleRoom;
        const yMax = chartArea.bottom - 4;
        for (const side of ['left', 'right'] as const) {
            const group = outside.filter((o) => o.side === side).sort((a, b) => a.labelY - b.labelY);
            if (group.length === 0) continue;
            const dir = side === 'right' ? 1 : -1;
            // Text anchors JUST OUTSIDE the plot edge and grows away from the
            // pie (right-aligned on the left, left-aligned on the right), so
            // it stays close to the slice and never jams against the canvas
            // edge. The renderer reserved layout padding on each side so this
            // sits within the canvas. The leader-line kink is just inside the
            // plot edge; the line runs rim → kink → text anchor.
            const textAnchorX = side === 'right' ? chartArea.right + 8 : chartArea.left - 8;
            const kinkX = side === 'right' ? chartArea.right - 6 : chartArea.left + 6;

            // De-overlap: clamp into [yMin, yMax] then push successive
            // labels apart by at least their half-heights + a gap.
            const gap = 6;
            for (let i = 0; i < group.length; i++) {
                group[i].labelY = Math.max(yMin, Math.min(yMax, group[i].labelY));
            }
            for (let i = 1; i < group.length; i++) {
                const prev = group[i - 1];
                const cur = group[i];
                const minDist = prev.boxH / 2 + cur.boxH / 2 + gap;
                if (cur.labelY - prev.labelY < minDist) cur.labelY = prev.labelY + minDist;
            }
            // If we ran past the bottom, shift the whole stack up to fit.
            const last = group[group.length - 1];
            const overflow = last.labelY + last.boxH / 2 - yMax;
            if (overflow > 0) for (const o of group) o.labelY -= overflow;

            for (const o of group) {
                // Clamp the text anchor so the (right-aligned left / left-
                // aligned right) text box stays fully on-canvas.
                const anchor = side === 'right'
                    ? Math.min(textAnchorX, chart.width - 4 - o.boxW)
                    : Math.max(textAnchorX, 4 + o.boxW);
                // Leader line: slice rim → radial kink → text anchor. The
                // kink (just inside the plot edge) at the label's
                // de-overlapped y gives the classic Excel elbow.
                ctx.beginPath();
                ctx.moveTo(o.rimX, o.rimY);
                ctx.lineTo(kinkX, o.labelY);
                ctx.lineTo(anchor, o.labelY);
                ctx.strokeStyle = o.color;
                ctx.lineWidth = 1.25;
                ctx.stroke();

                // Label text grows AWAY from the pie: right-aligned on the
                // left side (text ends at the anchor, extends left), left-
                // aligned on the right side. Small inset past the anchor so
                // the glyphs don't sit on the leader line.
                ctx.fillStyle = '#404040';
                ctx.textAlign = side === 'right' ? 'left' : 'right';
                drawLines(ctx, o.lines, anchor + dir * 4, o.labelY);
            }
        }

        ctx.restore();
    },
};

// Draw a vertically-centred stack of text lines at (x, yCenter).
function drawLines(ctx: CanvasRenderingContext2D, lines: string[], x: number, yCenter: number): void {
    const lineH = 14;
    const startY = yCenter - ((lines.length - 1) * lineH) / 2;
    for (let i = 0; i < lines.length; i++) {
        ctx.fillText(lines[i], x, startY + i * lineH);
    }
}
