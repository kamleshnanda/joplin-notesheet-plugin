import { _resetChartBus, pushChartUpdate, subscribeChartUpdate } from '../src/charts/dataBus';

const sample = { labels: ['a'], datasets: [{ data: [1], backgroundColor: '#000' }] };

describe('chart data bus', () => {
    beforeEach(() => _resetChartBus());

    test('subscribers receive updates for their id', () => {
        const seen: number[][] = [];
        subscribeChartUpdate('chart-1', (d) => seen.push(d.datasets[0].data));
        pushChartUpdate('chart-1', sample);
        expect(seen).toEqual([[1]]);
    });

    test('push to an id with no subscribers is a no-op', () => {
        // Should not throw.
        expect(() => pushChartUpdate('nobody-home', sample)).not.toThrow();
    });

    test('different ids are isolated', () => {
        let aCount = 0;
        let bCount = 0;
        subscribeChartUpdate('a', () => aCount++);
        subscribeChartUpdate('b', () => bCount++);
        pushChartUpdate('a', sample);
        pushChartUpdate('a', sample);
        pushChartUpdate('b', sample);
        expect(aCount).toBe(2);
        expect(bCount).toBe(1);
    });

    test('unsubscribe stops further deliveries', () => {
        let count = 0;
        const off = subscribeChartUpdate('x', () => count++);
        pushChartUpdate('x', sample);
        off();
        pushChartUpdate('x', sample);
        expect(count).toBe(1);
    });

    test('multiple subscribers per id all receive', () => {
        let a = 0;
        let b = 0;
        subscribeChartUpdate('shared', () => a++);
        subscribeChartUpdate('shared', () => b++);
        pushChartUpdate('shared', sample);
        expect(a).toBe(1);
        expect(b).toBe(1);
    });

    test('throwing listener does not block others', () => {
        let secondCalled = false;
        subscribeChartUpdate('y', () => { throw new Error('boom'); });
        subscribeChartUpdate('y', () => { secondCalled = true; });
        // Mute console.error during this test only.
        const orig = console.error;
        console.error = () => undefined;
        try {
            pushChartUpdate('y', sample);
        } finally {
            console.error = orig;
        }
        expect(secondCalled).toBe(true);
    });
});
