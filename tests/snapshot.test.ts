import {
    FENCE_VERSION,
    emptySnapshot,
    extractSnapshot,
    isNotesheetBody,
    wrapSnapshot,
} from '../src/snapshot';

describe('snapshot helpers', () => {
    test('wrap then extract round-trips', () => {
        const snap = { id: 'x', sheets: { a: { cellData: { 0: { 0: { v: 1 } } } } } };
        const body = wrapSnapshot(snap);
        const result = extractSnapshot(body);
        expect(result.ok).toBe(true);
        if (result.ok) {
            expect(result.version).toBe(FENCE_VERSION);
            expect(result.snapshot).toEqual(snap);
        }
    });

    test('isNotesheetBody detects fence', () => {
        expect(isNotesheetBody(wrapSnapshot({}))).toBe(true);
        expect(isNotesheetBody('# Just a markdown note')).toBe(false);
        expect(isNotesheetBody('')).toBe(false);
        expect(isNotesheetBody(null)).toBe(false);
        expect(isNotesheetBody(undefined)).toBe(false);
    });

    test('detects body even when surrounded by other markdown', () => {
        const body = '# Title\n\nSome prose.\n\n' + wrapSnapshot({ id: 'x' }) + '\n\nMore prose.\n';
        expect(isNotesheetBody(body)).toBe(true);
        const result = extractSnapshot(body);
        expect(result.ok).toBe(true);
    });

    test('rejects body with no fence', () => {
        const result = extractSnapshot('# Hello world');
        expect(result.ok).toBe(false);
        expect(result.ok === false && result.reason).toBe('no-fence');
    });

    test('rejects body with malformed JSON inside the fence', () => {
        const body = '```notesheet v=1\nnot-json\n```';
        const result = extractSnapshot(body);
        expect(result.ok).toBe(false);
        expect(result.ok === false && result.reason).toBe('bad-json');
    });

    test('rejects body with non-object JSON inside the fence', () => {
        const body = '```notesheet v=1\n42\n```';
        const result = extractSnapshot(body);
        expect(result.ok).toBe(false);
        expect(result.ok === false && result.reason).toBe('bad-json');
    });

    test('rejects unsupported future version', () => {
        const body = '```notesheet v=999\n{}\n```';
        const result = extractSnapshot(body);
        expect(result.ok).toBe(false);
        expect(result.ok === false && result.reason).toBe('unsupported-version');
    });

    test('emptySnapshot is a valid round-trippable workbook', () => {
        const snap = emptySnapshot();
        const body = wrapSnapshot(snap);
        const result = extractSnapshot(body);
        expect(result.ok).toBe(true);
        if (result.ok) {
            expect(result.snapshot).toMatchObject({
                sheetOrder: ['sheet-1'],
                sheets: { 'sheet-1': { id: 'sheet-1', name: 'Sheet1' } },
            });
        }
    });

    test('wrapped snapshot can contain backticks safely', () => {
        // Univer cell values may contain backticks; the JSON.stringify escapes
        // them to \", so the markdown fence terminator (```) doesn't get
        // confused. Verify a sample.
        const snap = { sheets: { a: { cellData: { 0: { 0: { v: '```hello```' } } } } } };
        const body = wrapSnapshot(snap);
        const result = extractSnapshot(body);
        expect(result.ok).toBe(true);
        if (result.ok) expect(result.snapshot).toEqual(snap);
    });
});
