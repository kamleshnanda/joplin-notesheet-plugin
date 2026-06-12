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
    headerBg: string; // #RRGGBB
    headerFg: string; // #RRGGBB
    bandedRowEvenBg?: string;
    bandedRowOddBg?: string;
    totalsBg?: string;
    totalsFg?: string;
    borderColor?: string;
    /**
     * The accent-coloured top border on the totals row. Distinct from
     * `borderColor` (which models the table outline). Populated at
     * synthesis time when the recipe carries a `totalsTopBorder` slot —
     * the static catalog leaves it undefined.
     */
    totalsTopBorder?: string;
    /**
     * The accent-coloured bottom border on the totals row. Excel paints
     * this in addition to (and the same colour as) `totalsTopBorder` —
     * see `screenshots/excel-reference/FormattingSmorgasboard-Aptos.png`,
     * which carries `#72D068` strips at both the top and bottom of the
     * totals row body. Populated by the recipe layer; replaces the
     * table-outline `borderColor` on the totals row's bottom edge.
     */
    totalsBottomBorder?: string;
}

export const EXCEL_TABLE_STYLES: ExcelTableStyle[] = [
    // --- Light (21) -----------------------------------------------------------
    {
        styleName: 'TableStyleLight1',
        headerBg: '#000000',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9D9D9',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#000000',
        borderColor: '#000000',
    },
    {
        styleName: 'TableStyleLight2',
        headerBg: '#156082',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#C1E5F5',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#156082',
        borderColor: '#156082',
    },
    {
        styleName: 'TableStyleLight3',
        headerBg: '#E97132',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#FBE3D6',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#E97132',
        borderColor: '#E97132',
    },
    {
        styleName: 'TableStyleLight4',
        headerBg: '#196B24',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#C2F1C8',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#196B24',
        borderColor: '#196B24',
    },
    {
        styleName: 'TableStyleLight5',
        headerBg: '#0F9ED5',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#CAEEFB',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#0F9ED5',
        borderColor: '#0F9ED5',
    },
    {
        styleName: 'TableStyleLight6',
        headerBg: '#A02B93',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#F2CFEE',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#A02B93',
        borderColor: '#A02B93',
    },
    {
        styleName: 'TableStyleLight7',
        headerBg: '#4EA72E',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9F2D0',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#4EA72E',
        borderColor: '#4EA72E',
    },
    {
        styleName: 'TableStyleLight8',
        headerBg: '#000000',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9D9D9',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#000000',
        borderColor: '#000000',
    },
    {
        styleName: 'TableStyleLight9',
        headerBg: '#156082',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#C1E5F5',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#156082',
        borderColor: '#156082',
    },
    {
        styleName: 'TableStyleLight10',
        headerBg: '#E97132',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#FBE3D6',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#E97132',
        borderColor: '#E97132',
    },
    {
        styleName: 'TableStyleLight11',
        headerBg: '#196B24',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#C2F1C8',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#196B24',
        borderColor: '#196B24',
    },
    {
        styleName: 'TableStyleLight12',
        headerBg: '#0F9ED5',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#CAEEFB',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#0F9ED5',
        borderColor: '#0F9ED5',
    },
    {
        styleName: 'TableStyleLight13',
        headerBg: '#A02B93',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#F2CFEE',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#A02B93',
        borderColor: '#A02B93',
    },
    {
        styleName: 'TableStyleLight14',
        headerBg: '#4EA72E',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9F2D0',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#4EA72E',
        borderColor: '#4EA72E',
    },
    {
        styleName: 'TableStyleLight15',
        headerBg: '#000000',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9D9D9',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#000000',
        borderColor: '#000000',
    },
    {
        styleName: 'TableStyleLight16',
        headerBg: '#156082',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#C1E5F5',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#156082',
        borderColor: '#156082',
    },
    {
        styleName: 'TableStyleLight17',
        headerBg: '#E97132',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#FBE3D6',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#E97132',
        borderColor: '#E97132',
    },
    {
        styleName: 'TableStyleLight18',
        headerBg: '#196B24',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#C2F1C8',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#196B24',
        borderColor: '#196B24',
    },
    {
        styleName: 'TableStyleLight19',
        headerBg: '#0F9ED5',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#CAEEFB',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#0F9ED5',
        borderColor: '#0F9ED5',
    },
    {
        styleName: 'TableStyleLight20',
        headerBg: '#A02B93',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#F2CFEE',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#A02B93',
        borderColor: '#A02B93',
    },
    {
        styleName: 'TableStyleLight21',
        headerBg: '#4EA72E',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9F2D0',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#FFFFFF',
        totalsFg: '#4EA72E',
        borderColor: '#4EA72E',
    },

    // --- Medium (28) ----------------------------------------------------------
    {
        styleName: 'TableStyleMedium1',
        headerBg: '#A6A6A6',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9D9D9',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#A6A6A6',
        totalsFg: '#000000',
        borderColor: '#A6A6A6',
    },
    {
        styleName: 'TableStyleMedium2',
        headerBg: '#156082',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#83CBEB',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#83CBEB',
        totalsFg: '#000000',
        borderColor: '#156082',
    },
    {
        styleName: 'TableStyleMedium3',
        headerBg: '#E97132',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#F6C6AD',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#F6C6AD',
        totalsFg: '#000000',
        borderColor: '#E97132',
    },
    {
        styleName: 'TableStyleMedium4',
        headerBg: '#196B24',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#84E291',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#84E291',
        totalsFg: '#000000',
        borderColor: '#196B24',
    },
    {
        styleName: 'TableStyleMedium5',
        headerBg: '#0F9ED5',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#96DCF8',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#96DCF8',
        totalsFg: '#000000',
        borderColor: '#0F9ED5',
    },
    {
        styleName: 'TableStyleMedium6',
        headerBg: '#A02B93',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#E59EDD',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#E59EDD',
        totalsFg: '#000000',
        borderColor: '#A02B93',
    },
    {
        styleName: 'TableStyleMedium7',
        headerBg: '#4EA72E',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#B4E5A2',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#B4E5A2',
        totalsFg: '#000000',
        borderColor: '#4EA72E',
    },
    {
        styleName: 'TableStyleMedium8',
        headerBg: '#A6A6A6',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9D9D9',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#A6A6A6',
        totalsFg: '#000000',
        borderColor: '#A6A6A6',
    },
    {
        styleName: 'TableStyleMedium9',
        headerBg: '#156082',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#83CBEB',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#83CBEB',
        totalsFg: '#000000',
        borderColor: '#156082',
    },
    {
        styleName: 'TableStyleMedium10',
        headerBg: '#E97132',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#F6C6AD',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#F6C6AD',
        totalsFg: '#000000',
        borderColor: '#E97132',
    },
    {
        styleName: 'TableStyleMedium11',
        headerBg: '#196B24',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#84E291',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#84E291',
        totalsFg: '#000000',
        borderColor: '#196B24',
    },
    {
        styleName: 'TableStyleMedium12',
        headerBg: '#0F9ED5',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#96DCF8',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#96DCF8',
        totalsFg: '#000000',
        borderColor: '#0F9ED5',
    },
    {
        styleName: 'TableStyleMedium13',
        headerBg: '#A02B93',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#E59EDD',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#E59EDD',
        totalsFg: '#000000',
        borderColor: '#A02B93',
    },
    {
        styleName: 'TableStyleMedium14',
        headerBg: '#4EA72E',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#B4E5A2',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#B4E5A2',
        totalsFg: '#000000',
        borderColor: '#4EA72E',
    },
    {
        styleName: 'TableStyleMedium15',
        headerBg: '#A6A6A6',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9D9D9',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#A6A6A6',
        totalsFg: '#000000',
        borderColor: '#A6A6A6',
    },
    {
        styleName: 'TableStyleMedium16',
        headerBg: '#156082',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#83CBEB',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#83CBEB',
        totalsFg: '#000000',
        borderColor: '#156082',
    },
    {
        styleName: 'TableStyleMedium17',
        headerBg: '#E97132',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#F6C6AD',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#F6C6AD',
        totalsFg: '#000000',
        borderColor: '#E97132',
    },
    {
        styleName: 'TableStyleMedium18',
        headerBg: '#196B24',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#84E291',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#84E291',
        totalsFg: '#000000',
        borderColor: '#196B24',
    },
    {
        styleName: 'TableStyleMedium19',
        headerBg: '#0F9ED5',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#96DCF8',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#96DCF8',
        totalsFg: '#000000',
        borderColor: '#0F9ED5',
    },
    {
        styleName: 'TableStyleMedium20',
        headerBg: '#A02B93',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#E59EDD',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#E59EDD',
        totalsFg: '#000000',
        borderColor: '#A02B93',
    },
    {
        styleName: 'TableStyleMedium21',
        headerBg: '#4EA72E',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#B4E5A2',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#B4E5A2',
        totalsFg: '#000000',
        borderColor: '#4EA72E',
    },
    {
        styleName: 'TableStyleMedium22',
        headerBg: '#A6A6A6',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#D9D9D9',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#A6A6A6',
        totalsFg: '#000000',
        borderColor: '#A6A6A6',
    },
    {
        styleName: 'TableStyleMedium23',
        headerBg: '#156082',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#83CBEB',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#83CBEB',
        totalsFg: '#000000',
        borderColor: '#156082',
    },
    {
        styleName: 'TableStyleMedium24',
        headerBg: '#E97132',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#F6C6AD',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#F6C6AD',
        totalsFg: '#000000',
        borderColor: '#E97132',
    },
    {
        styleName: 'TableStyleMedium25',
        headerBg: '#196B24',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#84E291',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#84E291',
        totalsFg: '#000000',
        borderColor: '#196B24',
    },
    {
        styleName: 'TableStyleMedium26',
        headerBg: '#0F9ED5',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#96DCF8',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#96DCF8',
        totalsFg: '#000000',
        borderColor: '#0F9ED5',
    },
    {
        styleName: 'TableStyleMedium27',
        headerBg: '#A02B93',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#E59EDD',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#E59EDD',
        totalsFg: '#000000',
        borderColor: '#A02B93',
    },
    {
        styleName: 'TableStyleMedium28',
        headerBg: '#4EA72E',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#B4E5A2',
        bandedRowOddBg: '#FFFFFF',
        totalsBg: '#B4E5A2',
        totalsFg: '#000000',
        borderColor: '#4EA72E',
    },

    // --- Dark (11) ------------------------------------------------------------
    {
        styleName: 'TableStyleDark1',
        headerBg: '#000000',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark2',
        headerBg: '#104862',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark3',
        headerBg: '#C04F15',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark4',
        headerBg: '#13501B',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark5',
        headerBg: '#0B76A0',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark6',
        headerBg: '#78206E',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark7',
        headerBg: '#3B7D23',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark8',
        headerBg: '#000000',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark9',
        headerBg: '#104862',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark10',
        headerBg: '#C04F15',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
    {
        styleName: 'TableStyleDark11',
        headerBg: '#13501B',
        headerFg: '#FFFFFF',
        bandedRowEvenBg: '#404040',
        bandedRowOddBg: '#595959',
        totalsBg: '#000000',
        totalsFg: '#FFFFFF',
        borderColor: '#FFFFFF',
    },
];

// Convenience lookup map.
export const EXCEL_TABLE_STYLE_BY_NAME: Record<string, ExcelTableStyle> = Object.fromEntries(
    EXCEL_TABLE_STYLES.map((s) => [s.styleName, s]),
);
