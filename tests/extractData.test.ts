import { extractRangeAsChartData } from '../src/charts/extractData';

function fakeWorkbook(values: unknown[][] | null) {
    return {
        getActiveSheet: () => ({
            getRange: () => (values === null ? null : { getValues: () => values }),
        }),
    };
}

describe('extractRangeAsChartData', () => {
    test('labels + 1 series', () => {
        const wb = fakeWorkbook([
            ['Q1', 100],
            ['Q2', 150],
            ['Q3', 200],
        ]);
        const { labels, datasets } = extractRangeAsChartData(wb, {
            startRow: 0,
            endRow: 2,
            startColumn: 0,
            endColumn: 1,
        });
        expect(labels).toEqual(['Q1', 'Q2', 'Q3']);
        expect(datasets).toHaveLength(1);
        expect(datasets[0].data).toEqual([100, 150, 200]);
    });

    test('labels + 2 series', () => {
        const wb = fakeWorkbook([
            ['Jan', 10, 20],
            ['Feb', 15, 25],
        ]);
        const { labels, datasets } = extractRangeAsChartData(wb, {
            startRow: 0,
            endRow: 1,
            startColumn: 0,
            endColumn: 2,
        });
        expect(labels).toEqual(['Jan', 'Feb']);
        expect(datasets).toHaveLength(2);
        expect(datasets[0].data).toEqual([10, 15]);
        expect(datasets[1].data).toEqual([20, 25]);
    });

    test('single column → one unlabeled series', () => {
        const wb = fakeWorkbook([[1], [2], [3]]);
        const { labels, datasets } = extractRangeAsChartData(wb, {
            startRow: 0,
            endRow: 2,
            startColumn: 0,
            endColumn: 0,
        });
        expect(labels).toEqual(['1', '2', '3']);
        expect(datasets).toHaveLength(1);
        expect(datasets[0].data).toEqual([1, 2, 3]);
    });

    test('numeric strings get coerced; text is NaN', () => {
        const wb = fakeWorkbook([
            ['A', '42'],
            ['B', 'oops'],
            ['C', '7.5'],
        ]);
        const { datasets } = extractRangeAsChartData(wb, {
            startRow: 0,
            endRow: 2,
            startColumn: 0,
            endColumn: 1,
        });
        expect(datasets[0].data[0]).toBe(42);
        expect(Number.isNaN(datasets[0].data[1])).toBe(true);
        expect(datasets[0].data[2]).toBe(7.5);
    });

    test('missing workbook returns empty', () => {
        const result = extractRangeAsChartData(null as unknown, {
            startRow: 0,
            endRow: 0,
            startColumn: 0,
            endColumn: 0,
        });
        expect(result).toEqual({ labels: [], datasets: [] });
    });

    test('range that the sheet cannot resolve returns empty', () => {
        const wb = fakeWorkbook(null);
        const result = extractRangeAsChartData(wb, {
            startRow: 0,
            endRow: 0,
            startColumn: 0,
            endColumn: 0,
        });
        expect(result).toEqual({ labels: [], datasets: [] });
    });

    test('palette assigns distinct colors to multi-series', () => {
        const wb = fakeWorkbook([
            ['Jan', 10, 20, 30],
            ['Feb', 15, 25, 35],
        ]);
        const { datasets } = extractRangeAsChartData(wb, {
            startRow: 0,
            endRow: 1,
            startColumn: 0,
            endColumn: 3,
        });
        const colors = datasets.map((d) => d.backgroundColor);
        expect(new Set(colors).size).toBe(3);
    });
});
