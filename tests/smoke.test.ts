// Placeholder smoke test so Jest exits 0 in CI before M1 introduces real
// helpers. M1+ tests should land in tests/<feature>.test.ts files.

describe('Jest setup', () => {
    test('runs', () => {
        expect(true).toBe(true);
    });
});
