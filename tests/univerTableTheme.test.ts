// Regression test for the passthrough-theme wiring (M12 Bug 2 fix).
//
// Univer's sheets-table plugin auto-applies one of its 6 default themes
// (`table-default-0..5`) on top of any imported table, masking the per-cell
// `bg`/`cl` values we synthesize from the source Excel TableStyleMedium2.
// `withFlatTableTheme` rewrites the preset so the data plugin is constructed
// with `userThemes[0] = notesheet-flat` and `defaultThemeIndex = 0`. Then
// the controller's resolved-theme list (`userThemes.concat(defaults)`) puts
// our empty theme at index 0, and tableAdd$ events without a tableStyleId
// (the case for fromJSON-loaded snapshots) pick up that empty theme.

// Mock the data plugin import so we don't drag the whole Univer ESM graph
// (lodash-es etc.) through jest's CJS transform. The helper's correctness is
// purely about identity-matching plugin references in the preset's plugin
// array, which a sentinel exercises just as well.
jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { UniverSheetsTablePlugin } from '@univerjs/sheets-table';

import {
    FLAT_TABLE_THEME_CONFIG,
    FLAT_TABLE_THEME_NAME,
    withFlatTableTheme,
} from '../src/univerTableTheme';

describe('withFlatTableTheme', () => {
    test('replaces UniverSheetsTablePlugin with [plugin, config] tuple', () => {
        const fakeUiPlugin = function FakeUi() { /* no-op */ };
        const preset = {
            plugins: [UniverSheetsTablePlugin, fakeUiPlugin],
            locales: { 'en-US': {} },
        };

        const result = withFlatTableTheme(preset);

        expect(result.plugins[0]).toEqual([UniverSheetsTablePlugin, FLAT_TABLE_THEME_CONFIG]);
        // Other plugin entries pass through untouched.
        expect(result.plugins[1]).toBe(fakeUiPlugin);
        // Locales preserved.
        expect(result.locales).toEqual({ 'en-US': {} });
    });

    test('config has the passthrough theme at userThemes[0] and defaultThemeIndex 0', () => {
        // The controller resolves themes as `userThemes.concat(builtins)` and
        // falls back to `_allThemes[defaultThemeIndex]` when tableAdd$ fires
        // without a tableStyleId. Both invariants must hold for the imported
        // colors to survive.
        expect(FLAT_TABLE_THEME_CONFIG.defaultThemeIndex).toBe(0);
        expect(FLAT_TABLE_THEME_CONFIG.userThemes).toHaveLength(1);
        expect(FLAT_TABLE_THEME_CONFIG.userThemes[0]).toEqual({
            name: FLAT_TABLE_THEME_NAME,
            style: {},
        });
    });

    test('does not mutate the input preset', () => {
        const fakeUiPlugin = function FakeUi() { /* no-op */ };
        const preset = { plugins: [UniverSheetsTablePlugin, fakeUiPlugin] };
        const original = preset.plugins.slice();

        withFlatTableTheme(preset);

        expect(preset.plugins).toEqual(original);
    });
});
