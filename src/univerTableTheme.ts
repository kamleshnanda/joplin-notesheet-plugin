// Univer's table plugin auto-applies one of 6 default themes
// (`table-default-0..5`) as a RangeThemeStyle on top of any cell that's part
// of an ITableJson. Even though our snapshot already synthesizes per-cell
// `bg`/`cl` values that match the source Excel TableStyleMedium2, the theme
// overlay shows lavender (#BAC6F8) banding instead of the imported teal
// (#83CBEB).
//
// We register a no-op "passthrough" theme as userThemes[0] and set
// defaultThemeIndex: 0. SheetsTableThemeController resolves the active
// theme list as `userThemes.concat(defaultThemes)`, so ours wins. With every
// style slot empty, the theme overlay contributes nothing and our
// synthesized cell colors render unaltered.
//
// fromJSON-loaded tables fire tableAdd$ without a tableStyleId (see
// sheets-table/lib/es/index.js fromJSON), so they always pick up our index-0
// theme. Tables created from scratch will also get this passthrough — that's
// fine; user-applied formatting via the toolbar lands on cells directly.

import { UniverSheetsTablePlugin } from '@univerjs/sheets-table';

export const FLAT_TABLE_THEME_NAME = 'notesheet-flat';

export const FLAT_TABLE_THEME_CONFIG = {
    userThemes: [{ name: FLAT_TABLE_THEME_NAME, style: {} }],
    defaultThemeIndex: 0,
} as const;

export interface UniverPreset {
    plugins: unknown[];
    locales?: unknown;
}

// Replace the data plugin entry of a sheets-table preset with the tuple form
// `[plugin, config]`, leaving every other entry (UI plugin, etc.) untouched.
export function withFlatTableTheme(preset: UniverPreset): UniverPreset {
    const plugins = preset.plugins.map((p) =>
        p === UniverSheetsTablePlugin ? [UniverSheetsTablePlugin, FLAT_TABLE_THEME_CONFIG] : p,
    );
    return { ...preset, plugins };
}
