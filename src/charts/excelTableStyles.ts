// =============================================================================
// Excel built-in TableStyle catalog for M12 ("synthesize per-cell styling").
//
// PURPOSE
//   When importing an .xlsx whose <table>/<tableStyleInfo> references a built-in
//   style name like "TableStyleMedium2", we have no styled <fill>/<font> records
//   to copy from the workbook -- Excel synthesizes the look at render time from
//   the workbook's theme accent colors. This catalog precomputes the rendered
//   colors so the import path can bake them into per-cell fills/fonts.
//
// PALETTE ASSUMPTION (READ ME)
//   This catalog is computed against the **Office 2016+ "Aptos" default theme
//   palette** (the one shipped with Microsoft 365 / Office 2019+). Verified by
//   inspecting `xl/theme/theme1.xml` in tests/fixtures/charts/01..10:
//       accent1 = #156082   (dark teal)
//       accent2 = #E97132   (orange)
//       accent3 = #196B24   (green)
//       accent4 = #0F9ED5   (cyan/sky)
//       accent5 = #A02B93   (magenta)
//       accent6 = #4EA72E   (lime green)
//   The legacy Office 2007-2013 palette (accent1=#5B9BD5, accent2=#ED7D31, ...)
//   is NOT used here. Workbooks authored in Excel 2013 or earlier will render
//   different colors for the same TableStyle name -- if M12 needs to support
//   those, the importer must read xl/theme/theme1.xml and recompute, OR pick a
//   different catalog. For the 99% case (modern Excel, modern Sheets export),
//   this catalog is the authoritative answer.
//
// DERIVATION
//   Tints/shades are applied in HSL L-space per ECMA-376 §18.8.19 ("color"
//   element @tint). Approximations used:
//     - Light header   = full accent;        rowStripe = tint(+0.80)  (very pale)
//     - Medium header  = full accent;        rowStripe = tint(+0.60)  (pastel)
//     - Dark header    = tint(-0.25) accent; rowStripe = #404040/#595959 greys
//     - "None" column (indices 1, 8, 15, 22): neutral grey/black variants
//   These match Excel's rendered output to within a few RGB units for the
//   modern palette; small drift is acceptable since user perception of
//   "dark teal banded table" doesn't depend on exact #83CBEB vs #84CBEC.
//
// COVERAGE / GAPS
//   - TableStyleLight 1..21        (21 entries -- all)
//   - TableStyleMedium 1..28       (28 entries -- all)
//   - TableStyleDark 1..11         (11 entries -- all)
//   Total: 60 entries. No styles deliberately omitted.
//
//   NOT MODELED (out of scope for M12 minimum):
//     - First-column / last-column emphasis (showFirstColumn, showLastColumn).
//     - Column stripes (showColumnStripes; rare in practice).
//     - Custom <tableStyle> entries authored in the workbook -- those have
//       real <dxf> records the importer can read directly.
//
// USAGE
//   const style = EXCEL_TABLE_STYLES.find(s => s.styleName === 'TableStyleMedium2');
//   if (style) {
//     applyFill(headerRow, style.headerBg);
//     applyFontColor(headerRow, style.headerFg);
//     for (let r = dataStart; r <= dataEnd; r++) {
//       applyFill(r, (r - dataStart) % 2 === 0 ? style.bandedRowEvenBg : style.bandedRowOddBg);
//     }
//   }
// =============================================================================

export interface ExcelTableStyle {
  styleName: string;
  headerBg: string;       // #RRGGBB
  headerFg: string;       // #RRGGBB
  bandedRowEvenBg?: string;
  bandedRowOddBg?: string;
  totalsBg?: string;
  totalsFg?: string;
  borderColor?: string;
}

export const EXCEL_TABLE_STYLES: ExcelTableStyle[] = [
  // --- Light (21) -----------------------------------------------------------
  { styleName: 'TableStyleLight1',  headerBg: '#000000', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9D9D9', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#000000', borderColor: '#000000' },
  { styleName: 'TableStyleLight2',  headerBg: '#156082', headerFg: '#FFFFFF', bandedRowEvenBg: '#C1E5F5', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#156082', borderColor: '#156082' },
  { styleName: 'TableStyleLight3',  headerBg: '#E97132', headerFg: '#FFFFFF', bandedRowEvenBg: '#FBE3D6', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#E97132', borderColor: '#E97132' },
  { styleName: 'TableStyleLight4',  headerBg: '#196B24', headerFg: '#FFFFFF', bandedRowEvenBg: '#C2F1C8', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#196B24', borderColor: '#196B24' },
  { styleName: 'TableStyleLight5',  headerBg: '#0F9ED5', headerFg: '#FFFFFF', bandedRowEvenBg: '#CAEEFB', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#0F9ED5', borderColor: '#0F9ED5' },
  { styleName: 'TableStyleLight6',  headerBg: '#A02B93', headerFg: '#FFFFFF', bandedRowEvenBg: '#F2CFEE', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#A02B93', borderColor: '#A02B93' },
  { styleName: 'TableStyleLight7',  headerBg: '#4EA72E', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9F2D0', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#4EA72E', borderColor: '#4EA72E' },
  { styleName: 'TableStyleLight8',  headerBg: '#000000', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9D9D9', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#000000', borderColor: '#000000' },
  { styleName: 'TableStyleLight9',  headerBg: '#156082', headerFg: '#FFFFFF', bandedRowEvenBg: '#C1E5F5', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#156082', borderColor: '#156082' },
  { styleName: 'TableStyleLight10', headerBg: '#E97132', headerFg: '#FFFFFF', bandedRowEvenBg: '#FBE3D6', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#E97132', borderColor: '#E97132' },
  { styleName: 'TableStyleLight11', headerBg: '#196B24', headerFg: '#FFFFFF', bandedRowEvenBg: '#C2F1C8', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#196B24', borderColor: '#196B24' },
  { styleName: 'TableStyleLight12', headerBg: '#0F9ED5', headerFg: '#FFFFFF', bandedRowEvenBg: '#CAEEFB', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#0F9ED5', borderColor: '#0F9ED5' },
  { styleName: 'TableStyleLight13', headerBg: '#A02B93', headerFg: '#FFFFFF', bandedRowEvenBg: '#F2CFEE', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#A02B93', borderColor: '#A02B93' },
  { styleName: 'TableStyleLight14', headerBg: '#4EA72E', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9F2D0', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#4EA72E', borderColor: '#4EA72E' },
  { styleName: 'TableStyleLight15', headerBg: '#000000', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9D9D9', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#000000', borderColor: '#000000' },
  { styleName: 'TableStyleLight16', headerBg: '#156082', headerFg: '#FFFFFF', bandedRowEvenBg: '#C1E5F5', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#156082', borderColor: '#156082' },
  { styleName: 'TableStyleLight17', headerBg: '#E97132', headerFg: '#FFFFFF', bandedRowEvenBg: '#FBE3D6', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#E97132', borderColor: '#E97132' },
  { styleName: 'TableStyleLight18', headerBg: '#196B24', headerFg: '#FFFFFF', bandedRowEvenBg: '#C2F1C8', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#196B24', borderColor: '#196B24' },
  { styleName: 'TableStyleLight19', headerBg: '#0F9ED5', headerFg: '#FFFFFF', bandedRowEvenBg: '#CAEEFB', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#0F9ED5', borderColor: '#0F9ED5' },
  { styleName: 'TableStyleLight20', headerBg: '#A02B93', headerFg: '#FFFFFF', bandedRowEvenBg: '#F2CFEE', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#A02B93', borderColor: '#A02B93' },
  { styleName: 'TableStyleLight21', headerBg: '#4EA72E', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9F2D0', bandedRowOddBg: '#FFFFFF', totalsBg: '#FFFFFF', totalsFg: '#4EA72E', borderColor: '#4EA72E' },

  // --- Medium (28) ----------------------------------------------------------
  { styleName: 'TableStyleMedium1',  headerBg: '#A6A6A6', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9D9D9', bandedRowOddBg: '#FFFFFF', totalsBg: '#A6A6A6', totalsFg: '#000000', borderColor: '#A6A6A6' },
  { styleName: 'TableStyleMedium2',  headerBg: '#156082', headerFg: '#FFFFFF', bandedRowEvenBg: '#83CBEB', bandedRowOddBg: '#FFFFFF', totalsBg: '#83CBEB', totalsFg: '#000000', borderColor: '#156082' },
  { styleName: 'TableStyleMedium3',  headerBg: '#E97132', headerFg: '#FFFFFF', bandedRowEvenBg: '#F6C6AD', bandedRowOddBg: '#FFFFFF', totalsBg: '#F6C6AD', totalsFg: '#000000', borderColor: '#E97132' },
  { styleName: 'TableStyleMedium4',  headerBg: '#196B24', headerFg: '#FFFFFF', bandedRowEvenBg: '#84E291', bandedRowOddBg: '#FFFFFF', totalsBg: '#84E291', totalsFg: '#000000', borderColor: '#196B24' },
  { styleName: 'TableStyleMedium5',  headerBg: '#0F9ED5', headerFg: '#FFFFFF', bandedRowEvenBg: '#96DCF8', bandedRowOddBg: '#FFFFFF', totalsBg: '#96DCF8', totalsFg: '#000000', borderColor: '#0F9ED5' },
  { styleName: 'TableStyleMedium6',  headerBg: '#A02B93', headerFg: '#FFFFFF', bandedRowEvenBg: '#E59EDD', bandedRowOddBg: '#FFFFFF', totalsBg: '#E59EDD', totalsFg: '#000000', borderColor: '#A02B93' },
  { styleName: 'TableStyleMedium7',  headerBg: '#4EA72E', headerFg: '#FFFFFF', bandedRowEvenBg: '#B4E5A2', bandedRowOddBg: '#FFFFFF', totalsBg: '#B4E5A2', totalsFg: '#000000', borderColor: '#4EA72E' },
  { styleName: 'TableStyleMedium8',  headerBg: '#A6A6A6', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9D9D9', bandedRowOddBg: '#FFFFFF', totalsBg: '#A6A6A6', totalsFg: '#000000', borderColor: '#A6A6A6' },
  { styleName: 'TableStyleMedium9',  headerBg: '#156082', headerFg: '#FFFFFF', bandedRowEvenBg: '#83CBEB', bandedRowOddBg: '#FFFFFF', totalsBg: '#83CBEB', totalsFg: '#000000', borderColor: '#156082' },
  { styleName: 'TableStyleMedium10', headerBg: '#E97132', headerFg: '#FFFFFF', bandedRowEvenBg: '#F6C6AD', bandedRowOddBg: '#FFFFFF', totalsBg: '#F6C6AD', totalsFg: '#000000', borderColor: '#E97132' },
  { styleName: 'TableStyleMedium11', headerBg: '#196B24', headerFg: '#FFFFFF', bandedRowEvenBg: '#84E291', bandedRowOddBg: '#FFFFFF', totalsBg: '#84E291', totalsFg: '#000000', borderColor: '#196B24' },
  { styleName: 'TableStyleMedium12', headerBg: '#0F9ED5', headerFg: '#FFFFFF', bandedRowEvenBg: '#96DCF8', bandedRowOddBg: '#FFFFFF', totalsBg: '#96DCF8', totalsFg: '#000000', borderColor: '#0F9ED5' },
  { styleName: 'TableStyleMedium13', headerBg: '#A02B93', headerFg: '#FFFFFF', bandedRowEvenBg: '#E59EDD', bandedRowOddBg: '#FFFFFF', totalsBg: '#E59EDD', totalsFg: '#000000', borderColor: '#A02B93' },
  { styleName: 'TableStyleMedium14', headerBg: '#4EA72E', headerFg: '#FFFFFF', bandedRowEvenBg: '#B4E5A2', bandedRowOddBg: '#FFFFFF', totalsBg: '#B4E5A2', totalsFg: '#000000', borderColor: '#4EA72E' },
  { styleName: 'TableStyleMedium15', headerBg: '#A6A6A6', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9D9D9', bandedRowOddBg: '#FFFFFF', totalsBg: '#A6A6A6', totalsFg: '#000000', borderColor: '#A6A6A6' },
  { styleName: 'TableStyleMedium16', headerBg: '#156082', headerFg: '#FFFFFF', bandedRowEvenBg: '#83CBEB', bandedRowOddBg: '#FFFFFF', totalsBg: '#83CBEB', totalsFg: '#000000', borderColor: '#156082' },
  { styleName: 'TableStyleMedium17', headerBg: '#E97132', headerFg: '#FFFFFF', bandedRowEvenBg: '#F6C6AD', bandedRowOddBg: '#FFFFFF', totalsBg: '#F6C6AD', totalsFg: '#000000', borderColor: '#E97132' },
  { styleName: 'TableStyleMedium18', headerBg: '#196B24', headerFg: '#FFFFFF', bandedRowEvenBg: '#84E291', bandedRowOddBg: '#FFFFFF', totalsBg: '#84E291', totalsFg: '#000000', borderColor: '#196B24' },
  { styleName: 'TableStyleMedium19', headerBg: '#0F9ED5', headerFg: '#FFFFFF', bandedRowEvenBg: '#96DCF8', bandedRowOddBg: '#FFFFFF', totalsBg: '#96DCF8', totalsFg: '#000000', borderColor: '#0F9ED5' },
  { styleName: 'TableStyleMedium20', headerBg: '#A02B93', headerFg: '#FFFFFF', bandedRowEvenBg: '#E59EDD', bandedRowOddBg: '#FFFFFF', totalsBg: '#E59EDD', totalsFg: '#000000', borderColor: '#A02B93' },
  { styleName: 'TableStyleMedium21', headerBg: '#4EA72E', headerFg: '#FFFFFF', bandedRowEvenBg: '#B4E5A2', bandedRowOddBg: '#FFFFFF', totalsBg: '#B4E5A2', totalsFg: '#000000', borderColor: '#4EA72E' },
  { styleName: 'TableStyleMedium22', headerBg: '#A6A6A6', headerFg: '#FFFFFF', bandedRowEvenBg: '#D9D9D9', bandedRowOddBg: '#FFFFFF', totalsBg: '#A6A6A6', totalsFg: '#000000', borderColor: '#A6A6A6' },
  { styleName: 'TableStyleMedium23', headerBg: '#156082', headerFg: '#FFFFFF', bandedRowEvenBg: '#83CBEB', bandedRowOddBg: '#FFFFFF', totalsBg: '#83CBEB', totalsFg: '#000000', borderColor: '#156082' },
  { styleName: 'TableStyleMedium24', headerBg: '#E97132', headerFg: '#FFFFFF', bandedRowEvenBg: '#F6C6AD', bandedRowOddBg: '#FFFFFF', totalsBg: '#F6C6AD', totalsFg: '#000000', borderColor: '#E97132' },
  { styleName: 'TableStyleMedium25', headerBg: '#196B24', headerFg: '#FFFFFF', bandedRowEvenBg: '#84E291', bandedRowOddBg: '#FFFFFF', totalsBg: '#84E291', totalsFg: '#000000', borderColor: '#196B24' },
  { styleName: 'TableStyleMedium26', headerBg: '#0F9ED5', headerFg: '#FFFFFF', bandedRowEvenBg: '#96DCF8', bandedRowOddBg: '#FFFFFF', totalsBg: '#96DCF8', totalsFg: '#000000', borderColor: '#0F9ED5' },
  { styleName: 'TableStyleMedium27', headerBg: '#A02B93', headerFg: '#FFFFFF', bandedRowEvenBg: '#E59EDD', bandedRowOddBg: '#FFFFFF', totalsBg: '#E59EDD', totalsFg: '#000000', borderColor: '#A02B93' },
  { styleName: 'TableStyleMedium28', headerBg: '#4EA72E', headerFg: '#FFFFFF', bandedRowEvenBg: '#B4E5A2', bandedRowOddBg: '#FFFFFF', totalsBg: '#B4E5A2', totalsFg: '#000000', borderColor: '#4EA72E' },

  // --- Dark (11) ------------------------------------------------------------
  { styleName: 'TableStyleDark1',  headerBg: '#000000', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark2',  headerBg: '#104862', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark3',  headerBg: '#C04F15', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark4',  headerBg: '#13501B', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark5',  headerBg: '#0B76A0', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark6',  headerBg: '#78206E', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark7',  headerBg: '#3B7D23', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark8',  headerBg: '#000000', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark9',  headerBg: '#104862', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark10', headerBg: '#C04F15', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
  { styleName: 'TableStyleDark11', headerBg: '#13501B', headerFg: '#FFFFFF', bandedRowEvenBg: '#404040', bandedRowOddBg: '#595959', totalsBg: '#000000', totalsFg: '#FFFFFF', borderColor: '#FFFFFF' },
];

// Convenience lookup map.
export const EXCEL_TABLE_STYLE_BY_NAME: Record<string, ExcelTableStyle> =
  Object.fromEntries(EXCEL_TABLE_STYLES.map(s => [s.styleName, s]));

// =============================================================================
// Theme-aware resolution (M13 — workstream A).
//
// The catalog above is computed against a single fixed theme palette
// (Aptos). Workbooks authored against any other theme — Office 2007
// (exceljs's writer default), Office 2013-2022 Classic, custom themes —
// would produce visibly wrong colors in Joplin even though the exported
// xlsx round-trips the source clrScheme correctly.
//
// `resolveTableStylePalette()` takes a TableStyle name + the workbook's
// own 12-entry theme palette (RGB values indexed by Excel theme color
// id 0..11; same shape as ThemePalette.rgb in src/xlsx.ts) and returns
// an ExcelTableStyle computed against THAT palette. When the palette
// is null (workbook ships no theme1.xml), fall back to the hardcoded
// catalog.
//
// Mapping cycle for "Light"/"Medium"/"Dark" styles:
//   Each variant has 21+/28+/11+ entries. The first cycle entry
//   (TableStyleMediumN where (N-1) mod 7 === 0: M1, M8, M15, M22) uses
//   a neutral grey palette and ignores theme accents. The remaining
//   entries cycle through accent1..accent6:
//
//     (N-1) mod 7 === 0  →  grey palette (no theme reference)
//     (N-1) mod 7 === 1  →  accent1
//     (N-1) mod 7 === 2  →  accent2
//     (N-1) mod 7 === 3  →  accent3
//     (N-1) mod 7 === 4  →  accent4
//     (N-1) mod 7 === 5  →  accent5
//     (N-1) mod 7 === 6  →  accent6
//
// "Light" styles use the same accent for headerBg + a much lighter
// banded row (tint +0.80). "Medium" uses tint +0.60. "Dark" uses
// inverted greys for bands and a slightly darkened accent (-0.25) for
// the header.

// Excel theme color indices (cell-level <color theme="N"/> map):
//   0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2,
//   4 = accent1, 5 = accent2, 6 = accent3, 7 = accent4,
//   8 = accent5, 9 = accent6, 10 = hlink, 11 = folHlink.
const ACCENT_INDICES_FROM_CYCLE = [
    -1, // (N-1) mod 7 === 0 → no accent
    4,  // accent1
    5,  // accent2
    6,  // accent3
    7,  // accent4
    8,  // accent5
    9,  // accent6
];

// Apply OOXML's HSL-luminance tint, identical to applyOoxmlTint() in
// src/xlsx.ts. Duplicated here to avoid a cross-package import; the
// math is small and rarely changes.
function applyTint(hex: string, tint: number): string {
    const r = parseInt(hex.slice(1, 3), 16) / 255;
    const g = parseInt(hex.slice(3, 5), 16) / 255;
    const b = parseInt(hex.slice(5, 7), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    let l = (max + min) / 2;
    if (max !== min) {
        const d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
        else if (max === g) h = (b - r) / d + 2;
        else h = (r - g) / d + 4;
        h /= 6;
    }
    if (tint < 0) l = l * (1 + tint);
    else l = l * (1 - tint) + tint;
    const hue2rgb = (p: number, q: number, t: number): number => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
    };
    let r2: number, g2: number, b2: number;
    if (s === 0) { r2 = g2 = b2 = l; }
    else {
        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        r2 = hue2rgb(p, q, h + 1 / 3);
        g2 = hue2rgb(p, q, h);
        b2 = hue2rgb(p, q, h - 1 / 3);
    }
    const toHex = (x: number) => Math.round(x * 255).toString(16).padStart(2, '0');
    return ('#' + toHex(r2) + toHex(g2) + toHex(b2)).toUpperCase();
}

// Parse "TableStyleLightN" / "TableStyleMediumN" / "TableStyleDarkN" →
// ('Light' | 'Medium' | 'Dark', N). Returns null on unrecognized input.
function parseTableStyleName(name: string): { variant: 'Light' | 'Medium' | 'Dark'; n: number } | null {
    const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
    if (!m) return null;
    return { variant: m[1] as 'Light' | 'Medium' | 'Dark', n: parseInt(m[2], 10) };
}

// Resolve a TableStyle name against the workbook's theme palette.
// Returns null when:
//   - The style name isn't recognized (caller should try the catalog).
//   - The palette is missing AND the catalog also has no entry.
// Returns the catalog entry when the palette is missing but the
// catalog has a fallback. Returns a freshly-computed entry when the
// palette is present.
export function resolveTableStylePalette(
    styleName: string,
    palette: string[] | null | undefined,
): ExcelTableStyle | undefined {
    // Fall back to the catalog when there's no workbook theme to resolve
    // against, OR when the style is one of the cycle-0 "neutral grey"
    // entries that intentionally ignores theme accents.
    const parsed = parseTableStyleName(styleName);
    const cycle = parsed ? (parsed.n - 1) % 7 : -1;
    if (!palette || cycle === 0) return EXCEL_TABLE_STYLE_BY_NAME[styleName];
    if (!parsed) return EXCEL_TABLE_STYLE_BY_NAME[styleName];

    const accentIdx = ACCENT_INDICES_FROM_CYCLE[cycle];
    const accent = palette[accentIdx];
    if (!accent) return EXCEL_TABLE_STYLE_BY_NAME[styleName];

    if (parsed.variant === 'Medium') {
        // Header = full accent, banded even = tint(+0.6) of accent,
        // banded odd = white, totals = same as banded even, border =
        // accent. Matches the structure of EXCEL_TABLE_STYLES Medium
        // entries; only the RGBs change.
        const band = applyTint(accent, 0.6);
        return {
            styleName,
            headerBg: accent,
            headerFg: '#FFFFFF',
            bandedRowEvenBg: band,
            bandedRowOddBg: '#FFFFFF',
            totalsBg: band,
            totalsFg: '#000000',
            borderColor: accent,
        };
    }
    if (parsed.variant === 'Light') {
        // Header = full accent, banded = tint(+0.8) of accent (paler),
        // borders use the accent.
        const band = applyTint(accent, 0.8);
        return {
            styleName,
            headerBg: accent,
            headerFg: '#FFFFFF',
            bandedRowEvenBg: band,
            bandedRowOddBg: '#FFFFFF',
            totalsBg: '#FFFFFF',
            totalsFg: accent,
            borderColor: accent,
        };
    }
    // Dark variant: header = darkened accent (-0.25), bands = neutral
    // greys, totals = black/white inverted, border = white.
    const headerBg = applyTint(accent, -0.25);
    return {
        styleName,
        headerBg,
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    };
}
