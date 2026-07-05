// M18 bugfix: live-edit chart re-extraction must skip the header row.
//
// REPRO: import 01-bar-simple (range A1:B5, where row 0 = "Quarter"/"Sales"
// header). The initial render uses the chart's cached labels (Q1..Q4, header
// excluded) and is correct. But editing a value fires refreshChartsForEdit,
// which re-extracts via extractRangeAsChartData over the WHOLE sourceRange —
// including row 0 — so the header "Quarter" becomes a 5th category with value
// 0 (a phantom empty bar labelled "Quarter").
//
// FIX: extractRangeAsChartData takes a hasHeaderRow flag; when set it drops
// row 0 from labels + series data, matching what the importer does when it
// builds cached labels from <c:cat> (A2:A5, not A1:A5). The flag is derived
// from the chart's meta.categoryAxisType === 'category' and threaded through
// the tracked chart to refreshChartsForEdit.

import { readFileSync } from 'fs';
import path from 'path';
import { extractRangeAsChartData, type RangeAddress } from '../src/charts/extractData';
import { populateTrackedChartsFromSnapshot, trackedCharts } from '../src/charts/trackedCharts';
import { xlsxBufferToSnapshot } from '../src/xlsx';

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

// Minimal FWorkbook facade returning a fixed 5-row range (A1:B5) where row 0
// is the header — exactly the 01-bar-simple shape after a B4 edit to 195.
function fakeWorkbook(values: unknown[][]): unknown {
    return {
        getActiveSheet: () => ({
            getRange: (_r: RangeAddress) => ({ getValues: () => values }),
        }),
        getSheetBySheetId: () => null,
    };
}

const A1_B5: unknown[][] = [
    ['Quarter', 'Sales'],
    ['Q1', 100],
    ['Q2', 120],
    ['Q3', 195],
    ['Q4', 140],
];
const range: RangeAddress = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 };

describe('M18 — live-edit chart re-extraction skips the header row', () => {
    test('REPRO: without the header flag, the header row leaks in as a category', () => {
        // Documents the buggy default (back-compat path for header-less ranges).
        const data = extractRangeAsChartData(fakeWorkbook(A1_B5), range);
        // 5 labels incl. the "Quarter" header — this is the phantom bar.
        expect(data.labels).toEqual(['Quarter', 'Q1', 'Q2', 'Q3', 'Q4']);
    });

    test('with hasHeaderRow, row 0 is dropped — labels are Q1..Q4 and data is 4 points', () => {
        const data = extractRangeAsChartData(fakeWorkbook(A1_B5), range, {
            hasHeaderRow: true,
        });
        expect(data.labels).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
        expect(data.datasets.length).toBe(1);
        expect(data.datasets[0].data).toEqual([100, 120, 195, 140]);
        // The series label comes from the header cell of the value column.
        expect(data.datasets[0].label).toBe('Sales');
        // No phantom "Quarter" category, no leading 0 datapoint.
        expect(data.labels).not.toContain('Quarter');
    });

    test('single-column range with hasHeaderRow drops the header too', () => {
        const oneCol: unknown[][] = [['Sales'], [100], [120], [195], [140]];
        const data = extractRangeAsChartData(
            fakeWorkbook(oneCol),
            { startRow: 0, endRow: 4, startColumn: 1, endColumn: 1 },
            { hasHeaderRow: true },
        );
        expect(data.datasets[0].data).toEqual([100, 120, 195, 140]);
    });

    test('integration: 01-bar-simple tracks hasHeaderRow=true (category axis)', async () => {
        const buf = readFileSync(path.join(__dirname, 'fixtures', 'charts', '01-bar-simple.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        populateTrackedChartsFromSnapshot(snap as unknown as Record<string, unknown>);
        const chart = [...trackedCharts.values()][0];
        expect(chart).toBeDefined();
        // sourceRange spans the header row (startRow 0)...
        expect(chart.sourceRange.startRow).toBe(0);
        // ...so the tracked chart must flag it for the live-edit skip.
        expect(chart.hasHeaderRow).toBe(true);
    });
});
