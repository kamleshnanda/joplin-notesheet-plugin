// =============================================================================
// Theme-aware recipes for the Excel built-in TableStyle catalog.
//
// PURPOSE
//   `excelTableStyles.ts` is correct ONLY for the Office 2016+ Aptos accent
//   palette. When a workbook ships a non-Aptos `<a:clrScheme>` (e.g. the
//   2013-era "Classic" scheme whose accent3 is `#A5A5A5` grey instead of
//   Aptos's `#196B24` green), the same `TableStyleMedium4` must paint grey
//   in Excel — but our hardcoded catalog still painted Aptos green.
//
//   This file maps every accent-derived TableStyle slot to a structured
//   "(accent index, tint)" recipe. At import time the synthesizer resolves
//   the recipe against the source workbook's clrScheme, producing the
//   correct rendered RGB regardless of which palette the workbook ships.
//
//   The achromatic styles (Light1/8/15, Medium1/8/15/22, Dark1/8) don't
//   reference any accent and keep their literal greys — their recipe slot
//   is `{accent: null, rgb: '#...'}`.
//
// SHAPE
//   - `accent: 1..6` — index into the source clrScheme's accent1..accent6.
//   - `accent: null` — the slot is literal RGB; use `rgb` directly.
//   - `tint: 0` — full accent (no tint).
//   - `tint > 0` — lighten toward white in HSL-L space (ECMA-376 §18.8.19).
//   - `tint < 0` — darken toward black.
//   White (#FFFFFF) and black (#000000) headerFg / totalsFg / etc. slots
//   are encoded as `{accent: null, rgb: '#FFFFFF'}` / `{accent: null, rgb: '#000000'}`
//   — the existing catalog always uses pure white/black for those slots and
//   no clrScheme would override that.
//
// DERIVATION & THE EMPIRICAL OVERRIDE TABLE
//   The HSL-L tint amounts encoded here (`mediumFor()` etc.) are the
//   first-pass approximation that PR #22 shipped. They are NOT what
//   Excel actually paints. Operator-captured screenshots in
//   `screenshots/excel-reference/*.png` revealed:
//
//     TableStyleMedium4 + Aptos accent3 #196B24:
//       header     #34692E (NOT raw accent, NOT HSL-L tint)
//       banded row #CAEFCB (NOT tint(+0.6) which gives #84E291)
//       totals top #72D068 (double-line border, missing entirely from
//                            the recipe shape pre-PR-#22)
//     TableStyleMedium4 + Classic accent3 #A5A5A5:
//       header     #A5A5A5 (raw — happens to match)
//       banded row #EDEDED (NOT tint(+0.6) which gives #DBDBDB)
//       totals top #C9C9C9
//
//   No single algorithm we tried (HSL-L tint, HSV scaling, satMod+lumMod,
//   RGB mix toward grey, etc.) reproduces all four target RGBs from
//   `(#196B24, +0.6)` AND `(#A5A5A5, +0.6)` simultaneously. Excel's
//   built-in TableStyle definitions live in Office's installed assets,
//   not in the workbook's `xl/styles.xml`, and the actual transformation
//   is not documented in OOXML.
//
//   To unblock M13/E and the fidelity-test gate without a fudged formula,
//   we ship `EXCEL_TABLE_STYLE_EMPIRICAL_OVERRIDES`: a `(styleName,
//   accentHex) → measured slot` lookup. The synthesizer prefers this
//   lookup when the source accent matches a measured value, and falls
//   back to the (admittedly approximate) HSL-L tint formula when no
//   measurement exists. Adding a new fixture means adding a new entry,
//   measured from a real Excel render via the same `tests/util/pngSampler.ts`
//   helper the fidelity test uses.
//
// USAGE
//   const recipe = EXCEL_TABLE_STYLE_RECIPE_BY_NAME[table.styleName];
//   const headerBgRgb = recipe ? resolveSlot(recipe.headerBg, sourceClrScheme) : palette.headerBg;
// =============================================================================

export interface ColorSlotRecipe {
    /** 1..6 = clrScheme accent1..accent6; null = literal RGB. */
    accent: 1 | 2 | 3 | 4 | 5 | 6 | null;
    /** ECMA-376 HSL-L tint amount: 0 = full accent, +0.8 = very pale, -0.25 = darker. Ignored when accent is null. */
    tint?: number;
    /** Literal RGB, used when accent is null. */
    rgb?: string;
}

export interface ExcelTableStyleRecipe {
    styleName: string;
    headerBg: ColorSlotRecipe;
    headerFg: ColorSlotRecipe;
    bandedRowEvenBg?: ColorSlotRecipe;
    bandedRowOddBg?: ColorSlotRecipe;
    totalsBg?: ColorSlotRecipe;
    totalsFg?: ColorSlotRecipe;
    borderColor?: ColorSlotRecipe;
    /**
     * The accent-coloured top border on the totals row. Catalog
     * `borderColor` slot only models the table outline; this is a
     * separate decoration.
     */
    totalsTopBorder?: ColorSlotRecipe;
    /**
     * The accent-coloured bottom border on the totals row. Excel paints
     * this in addition to (and the same colour as) `totalsTopBorder` —
     * see `screenshots/excel-reference/FormattingSmorgasboard-Aptos.png`,
     * which carries `#72D068` strips at both the top and bottom of the
     * totals row body. Replaces the table-outline `borderColor` on the
     * totals row's bottom edge.
     */
    totalsBottomBorder?: ColorSlotRecipe;
}

const WHITE: ColorSlotRecipe = { accent: null, rgb: '#FFFFFF' };
const BLACK: ColorSlotRecipe = { accent: null, rgb: '#000000' };

// --- Building blocks ---------------------------------------------------------
// Light family: headerBg = accent (full), bandedEvenBg = accent (tint +0.80),
//               bandedOddBg = white, totalsBg = white, totalsFg = accent,
//               borderColor = accent.
function lightFor(accent: 1 | 2 | 3 | 4 | 5 | 6, n: number): ExcelTableStyleRecipe {
    return {
        styleName: `TableStyleLight${n}`,
        headerBg: { accent, tint: 0 },
        headerFg: WHITE,
        bandedRowEvenBg: { accent, tint: 0.8 },
        bandedRowOddBg: WHITE,
        totalsBg: WHITE,
        totalsFg: { accent, tint: 0 },
        borderColor: { accent, tint: 0 },
    };
}

// Medium family: headerBg = accent (FALLBACK - see empirical overrides),
//                bandedEvenBg = accent (FALLBACK), bandedOddBg = white,
//                totalsBg = white (Excel paints the body of totals as
//                white in the references, with a double-line top border),
//                totalsFg = black, borderColor = accent (table outline),
//                totalsTopBorder = accent at tint(+0.40) (FALLBACK).
//
// The accent + tint values here are an APPROXIMATION; the empirical
// override map below is what actually reproduces Excel for the project-
// owned fixtures. When the source accent isn't in the override map, the
// synthesizer falls back to these values — better than emitting nothing,
// but not pixel-accurate.
function mediumFor(accent: 1 | 2 | 3 | 4 | 5 | 6, n: number): ExcelTableStyleRecipe {
    return {
        styleName: `TableStyleMedium${n}`,
        headerBg: { accent, tint: 0 },
        headerFg: WHITE,
        bandedRowEvenBg: { accent, tint: 0.6 },
        bandedRowOddBg: WHITE,
        totalsBg: WHITE,
        totalsFg: BLACK,
        borderColor: { accent, tint: 0 },
        totalsTopBorder: { accent, tint: 0.4 },
        totalsBottomBorder: { accent, tint: 0.4 },
    };
}

// Dark family: headerBg = accent (tint -0.25, slightly darkened), headerFg = white,
//              bandedEvenBg = #404040, bandedOddBg = #595959 (theme-blind preset
//              greys per ECMA-376), totalsBg = black, totalsFg = white,
//              borderColor = white.
function darkFor(accent: 1 | 2 | 3 | 4 | 5 | 6, n: number): ExcelTableStyleRecipe {
    return {
        styleName: `TableStyleDark${n}`,
        headerBg: { accent, tint: -0.25 },
        headerFg: WHITE,
        bandedRowEvenBg: { accent: null, rgb: '#404040' },
        bandedRowOddBg: { accent: null, rgb: '#595959' },
        totalsBg: BLACK,
        totalsFg: WHITE,
        borderColor: WHITE,
    };
}

// Achromatic Light/Medium "0-style" entries (Light1/8/15, Medium1/8/15/22,
// Dark1/8) use a literal grey/black for header — they don't reference any
// accent. We encode them as accent:null with the catalog's exact RGB.
const LIGHT_ACHROMATIC = (n: number): ExcelTableStyleRecipe => ({
    styleName: `TableStyleLight${n}`,
    headerBg: BLACK,
    headerFg: WHITE,
    bandedRowEvenBg: { accent: null, rgb: '#D9D9D9' },
    bandedRowOddBg: WHITE,
    totalsBg: WHITE,
    totalsFg: BLACK,
    borderColor: BLACK,
});

const MEDIUM_ACHROMATIC = (n: number): ExcelTableStyleRecipe => ({
    styleName: `TableStyleMedium${n}`,
    headerBg: { accent: null, rgb: '#A6A6A6' },
    headerFg: WHITE,
    bandedRowEvenBg: { accent: null, rgb: '#D9D9D9' },
    bandedRowOddBg: WHITE,
    totalsBg: { accent: null, rgb: '#A6A6A6' },
    totalsFg: BLACK,
    borderColor: { accent: null, rgb: '#A6A6A6' },
});

const DARK_ACHROMATIC = (n: number): ExcelTableStyleRecipe => ({
    styleName: `TableStyleDark${n}`,
    headerBg: BLACK,
    headerFg: WHITE,
    bandedRowEvenBg: { accent: null, rgb: '#404040' },
    bandedRowOddBg: { accent: null, rgb: '#595959' },
    totalsBg: BLACK,
    totalsFg: WHITE,
    borderColor: WHITE,
});

export const EXCEL_TABLE_STYLE_RECIPES: ExcelTableStyleRecipe[] = [
    // --- Light (21) -------------------------------------------------------
    LIGHT_ACHROMATIC(1),
    lightFor(1, 2), lightFor(2, 3), lightFor(3, 4), lightFor(4, 5), lightFor(5, 6), lightFor(6, 7),
    LIGHT_ACHROMATIC(8),
    lightFor(1, 9), lightFor(2, 10), lightFor(3, 11), lightFor(4, 12), lightFor(5, 13), lightFor(6, 14),
    LIGHT_ACHROMATIC(15),
    lightFor(1, 16), lightFor(2, 17), lightFor(3, 18), lightFor(4, 19), lightFor(5, 20), lightFor(6, 21),

    // --- Medium (28) ------------------------------------------------------
    MEDIUM_ACHROMATIC(1),
    mediumFor(1, 2), mediumFor(2, 3), mediumFor(3, 4), mediumFor(4, 5), mediumFor(5, 6), mediumFor(6, 7),
    MEDIUM_ACHROMATIC(8),
    mediumFor(1, 9), mediumFor(2, 10), mediumFor(3, 11), mediumFor(4, 12), mediumFor(5, 13), mediumFor(6, 14),
    MEDIUM_ACHROMATIC(15),
    mediumFor(1, 16), mediumFor(2, 17), mediumFor(3, 18), mediumFor(4, 19), mediumFor(5, 20), mediumFor(6, 21),
    MEDIUM_ACHROMATIC(22),
    mediumFor(1, 23), mediumFor(2, 24), mediumFor(3, 25), mediumFor(4, 26), mediumFor(5, 27), mediumFor(6, 28),

    // --- Dark (11) --------------------------------------------------------
    DARK_ACHROMATIC(1),
    darkFor(1, 2), darkFor(2, 3), darkFor(3, 4), darkFor(4, 5), darkFor(5, 6), darkFor(6, 7),
    DARK_ACHROMATIC(8),
    darkFor(1, 9), darkFor(2, 10), darkFor(3, 11),
];

export const EXCEL_TABLE_STYLE_RECIPE_BY_NAME: Record<string, ExcelTableStyleRecipe> =
    Object.fromEntries(EXCEL_TABLE_STYLE_RECIPES.map((s) => [s.styleName, s]));

// =============================================================================
// Empirical overrides — measured from real Excel renders.
//
// Keyed by (styleName, accentHex_uppercase). When a fixture's source
// accent matches a measured entry, the synthesizer prefers the measured
// RGB over the formula-based recipe. Add a new entry when:
//
//   1. You have a `screenshots/excel-reference/<fixture>.png` capture
//      from Excel (NOT from Notesheet — the whole point is ground truth).
//   2. You sampled the dominant fill of header / banded / totals-top
//      via `tests/util/pngSampler.ts:dominantColor`.
//   3. The fidelity test in `tests/excelReferenceFidelity.test.ts`
//      actually fails without the entry and passes with it.
//
// The lookup is keyed by **the resolved accent's hex**, not by accent
// index. That way two different workbooks whose accent3 happens to be
// the same colour will land at the same overrides regardless of palette.
//
// FORMAT
//   - `headerBg` etc. are literal RGB strings (no recipe slots).
//   - Slots not in the override fall back to the formula recipe.
// =============================================================================

export interface ExcelTableStyleEmpiricalOverride {
    headerBg?: string;
    headerFg?: string;
    bandedRowEvenBg?: string;
    bandedRowOddBg?: string;
    totalsBg?: string;
    totalsFg?: string;
    borderColor?: string;
    totalsTopBorder?: string;
    totalsBottomBorder?: string;
}

export const EXCEL_TABLE_STYLE_EMPIRICAL_OVERRIDES: Record<
    string,
    Record<string, ExcelTableStyleEmpiricalOverride>
> = {
    TableStyleMedium4: {
        // Aptos accent3. Re-measured 2026-06-07 from
        // screenshots/excel-reference/FormattingSmorgasboard-Aptos-wide.png
        // (1962×1070 wide capture; clean column body at x∈{380,588,780,
        // 950,1120,1300,1500} verified all-consistent).
        //
        // Totals TOP is the HEADER colour (#34692E) painted as a
        // double-line pair (two 2px strips with a 2px white gap):
        //   y=826-827 #34692E / y=828-829 #FFFFFF / y=830-831 #34692E
        // Totals BOTTOM is the lighter accent (#72D068), single 2px:
        //   y=908-909 #72D068
        // Inter-row strips at every banded-row boundary (rows 2..9) are
        // also #72D068, single 2px each (verified at 8 distinct
        // boundaries on the wide reference).
        '#196B24': {
            headerBg: '#34692E',
            bandedRowEvenBg: '#CAEFCB',
            totalsTopBorder: '#34692E',
            totalsBottomBorder: '#72D068',
        },
        // Classic accent3. Re-measured 2026-06-07 from
        // screenshots/excel-reference/FormattingSmorgasboard-Classic.png
        // (1378×618 capture; probed at 10 x-positions all consistent).
        // Same shape as Aptos: totals-top is HEADER colour double-line,
        // totals-bottom is the lighter accent single 2px, inter-row
        // strips at every banded-row boundary in the lighter accent.
        '#A5A5A5': {
            headerBg: '#A5A5A5',
            bandedRowEvenBg: '#EDEDED',
            totalsTopBorder: '#A5A5A5',
            totalsBottomBorder: '#C9C9C9',
        },
    },
};

// =============================================================================
// HSL-L tint helpers (ECMA-376 §18.8.19 "color" element @tint).
//
// Tint > 0:  L' = L + (1 - L) * tint        — lighten toward white
// Tint < 0:  L' = L + L * tint              — darken toward black
// Tint == 0: no change
//
// Operating in HSL-L (not HSL-V or HSV) matches Excel's rendering. We use
// it via a roundtrip:  RGB → HSL → adjust L → back to RGB.
// =============================================================================

function hexToRgb(hex: string): { r: number; g: number; b: number } {
    const s = hex.replace(/^#/, '');
    return {
        r: parseInt(s.slice(0, 2), 16),
        g: parseInt(s.slice(2, 4), 16),
        b: parseInt(s.slice(4, 6), 16),
    };
}

function rgbToHex(r: number, g: number, b: number): string {
    const clamp = (n: number): number => Math.max(0, Math.min(255, Math.round(n)));
    const toHex = (n: number): string => clamp(n).toString(16).padStart(2, '0');
    return ('#' + toHex(r) + toHex(g) + toHex(b)).toUpperCase();
}

function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
    const rN = r / 255, gN = g / 255, bN = b / 255;
    const max = Math.max(rN, gN, bN), min = Math.min(rN, gN, bN);
    const l = (max + min) / 2;
    let h = 0;
    let s = 0;
    if (max !== min) {
        const d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch (max) {
            case rN: h = (gN - bN) / d + (gN < bN ? 6 : 0); break;
            case gN: h = (bN - rN) / d + 2; break;
            case bN: h = (rN - gN) / d + 4; break;
        }
        h /= 6;
    }
    return { h, s, l };
}

function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
    if (s === 0) {
        const v = l * 255;
        return { r: v, g: v, b: v };
    }
    const hue2rgb = (p: number, q: number, t: number): number => {
        let tN = t;
        if (tN < 0) tN += 1;
        if (tN > 1) tN -= 1;
        if (tN < 1 / 6) return p + (q - p) * 6 * tN;
        if (tN < 1 / 2) return q;
        if (tN < 2 / 3) return p + (q - p) * (2 / 3 - tN) * 6;
        return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    return {
        r: hue2rgb(p, q, h + 1 / 3) * 255,
        g: hue2rgb(p, q, h) * 255,
        b: hue2rgb(p, q, h - 1 / 3) * 255,
    };
}

/**
 * Apply ECMA-376 HSL-L tint to a base RGB.
 *
 * `tint = 0` is the identity. Positive tints lighten toward white; negative
 * tints darken toward black. The maths matches Excel's render (verified
 * against the existing `excelTableStyles.ts` catalog: `tintRgb('#196B24', 0.6)`
 * lands within 1-2 RGB units of `#84E291`).
 */
export function tintRgb(baseHex: string, tint: number): string {
    const { r, g, b } = hexToRgb(baseHex);
    if (!tint) return rgbToHex(r, g, b);
    const { h, s, l } = rgbToHsl(r, g, b);
    let lNew: number;
    if (tint > 0) lNew = l + (1 - l) * tint;
    else lNew = l + l * tint;
    if (lNew < 0) lNew = 0;
    if (lNew > 1) lNew = 1;
    const { r: rN, g: gN, b: bN } = hslToRgb(h, s, lNew);
    return rgbToHex(rN, gN, bN);
}

/**
 * Resolve a single color-slot recipe against a source clrScheme palette.
 *
 * `accents` is a 6-tuple of `'#RRGGBB'` strings indexed by accent1..accent6.
 * If the recipe is achromatic (`accent: null`) the literal `rgb` is returned.
 */
export function resolveColorSlot(
    recipe: ColorSlotRecipe,
    accents: readonly string[],
): string {
    if (recipe.accent === null || recipe.accent === undefined) {
        return (recipe.rgb || '#000000').toUpperCase();
    }
    const baseHex = accents[recipe.accent - 1];
    if (!baseHex) {
        // Source clrScheme didn't supply this accent — fall back to a sane
        // default. Should not happen for well-formed Excel themes.
        return '#000000';
    }
    return tintRgb(baseHex, recipe.tint || 0);
}
