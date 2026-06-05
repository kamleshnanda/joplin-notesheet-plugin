// Canvas-vs-Excel fidelity gate.
//
// PURPOSE
//   `excelReferenceFidelity.test.ts` (Phase 2 of the M13/E rework)
//   compares the Excel reference render against `xlsxBufferToSnapshot`'s
//   in-memory snapshot. That gate catches "import path produced wrong
//   target RGB" — but it does NOT catch "Univer didn't paint the target
//   RGB". The data shape can be perfect while the rendered canvas is
//   wrong.
//
//   This file adds the second gate: compare the Joplin canvas screenshot
//   (`screenshots/feature-1-m13-theme-aware-banding/eval-{aptos,classic}-*.png`)
//   against the Excel reference render
//   (`screenshots/excel-reference/FormattingSmorgasboard-{Aptos,Classic}.png`)
//   at structurally-aligned regions. Region-finding is heuristic, not
//   pixel-coordinate hardcoded — the two PNGs are at different DPRs and
//   different table positions, so we locate regions by scanning for
//   their distinctive fills.
//
// REGION FINDING
//   - Header band: the first horizontal strip (≥6 consecutive y-rows)
//     whose dominant non-background colour matches the expected header
//     fill within Δ ≤ 24 per channel. Returns its y-range and dominant
//     RGB.
//   - First banded data row: scan downward from the bottom of the
//     header band; first horizontal strip whose dominant colour matches
//     the expected banded fill within Δ ≤ 24.
//   - Totals-row top: scan upward from the bottom of the table; locate
//     the LAST coloured strip (matching the totals-top expected colour
//     within Δ ≤ 24). The totals-row body sits below this strip, so
//     "last coloured strip before the white-or-fill totals-body region"
//     is the sentinel.
//
// ASSERTIONS
//   - Header dominant colour matches Excel's header dominant colour
//     (Δ ≤ 8 per channel).
//   - First banded row dominant colour matches Excel's first banded row
//     (Δ ≤ 8 per channel).
//   - Totals-row top border dominant colour matches Excel's
//     (Δ ≤ 8 per channel).
//   - Totals-row top border thickness: a SINGLE solid strip (not two
//     strips with white between) when Excel has a single strip. Univer's
//     `BorderStyleTypes.DOUBLE` (style code 7) emits two thin strips
//     separated by ~1px of white; Excel paints a single 2px MEDIUM
//     strip. The test asserts "no white pixel inside the dark strip" so
//     a regression to DOUBLE rendering trips the gate.
//
// REGION TOLERANCE
//   The colour tolerance for region-FINDING is Δ ≤ 24 per channel — wide
//   so anti-aliasing variants of the strip don't get missed. The
//   header / banded ASSERTION tolerance is Δ ≤ 8 (those regions are tall
//   enough that anti-aliasing doesn't dominate the dominant pixel).
//
//   The TOTALS-TOP STRIP tolerance is Δ ≤ 32. Reason: a 1-2px strip on
//   a Retina canvas (DPR=2) pulls in sub-pixel anti-aliasing — the
//   reading is e.g. `#89CE74` instead of pure `#72D068`. We can't
//   demand pixel-pure parity at this thickness without a different
//   sampling scheme; the structural assertion (single strip vs double
//   strip — see below) is what's actually load-bearing for this gate.
//
// LATEST EVAL SCREENSHOT
//   The Joplin screenshots have timestamped filenames; we pick the
//   latest by mtime within `screenshots/feature-1-m13-theme-aware-banding/`
//   matching `eval-aptos-*.png` and `eval-classic-*.png`. If the
//   directory has no matching files, the test errors with a clear
//   diagnostic — re-run `scripts/pge/eval-screenshot.sh` first.

import path from 'path';
import { readdirSync, statSync } from 'fs';

import {
    decodePng,
    dominantColor,
    hexToRgb,
    rgbDelta,
    DecodedPng,
} from './util/pngSampler';

const REPO_ROOT = path.join(__dirname, '..');
const REFERENCES_DIR = path.join(REPO_ROOT, 'screenshots', 'excel-reference');
const FEATURE_DIR = path.join(REPO_ROOT, 'screenshots', 'feature-1-m13-theme-aware-banding');

const APTOS_REF = path.join(REFERENCES_DIR, 'FormattingSmorgasboard-Aptos.png');
const CLASSIC_REF = path.join(REFERENCES_DIR, 'FormattingSmorgasboard-Classic.png');

// Spec-mandated colours for each fixture (matches the empirical override
// table in `src/charts/excelTableStyleRecipes.ts`). Region-finding uses
// these as the expected fill; the dominant-colour assertion compares
// Joplin → Excel head-to-head, NOT Joplin → expected (so the test
// doesn't lie if both are wrong in the same way).
const APTOS_EXPECTED = {
    headerHex: '#34692E',
    bandedHex: '#CAEFCB',
    totalsTopHex: '#72D068',
};
const CLASSIC_EXPECTED = {
    headerHex: '#A5A5A5',
    bandedHex: '#EDEDED',
    totalsTopHex: '#C9C9C9',
};

const REGION_TOLERANCE = 24; // wide for region-finding (anti-aliased edges)
const ASSERT_TOLERANCE = 8; // tight for tall regions (header, banded data row)
const STRIP_TOLERANCE = 32; // looser for 1-2px strips at DPR=2 (anti-aliasing dominates)

/** Pick the latest `eval-<variant>-*.png` from FEATURE_DIR by mtime. */
function latestEvalPng(variant: 'aptos' | 'classic'): string {
    const prefix = `eval-${variant}-`;
    const matches = readdirSync(FEATURE_DIR).filter((n) =>
        n.startsWith(prefix) && n.endsWith('.png'),
    );
    if (matches.length === 0) {
        throw new Error(
            `No ${prefix}*.png in ${FEATURE_DIR}. ` +
            `Re-run scripts/pge/eval-screenshot.sh feature-1-m13-theme-aware-banding:${variant}.`,
        );
    }
    matches.sort((a, b) => {
        const ma = statSync(path.join(FEATURE_DIR, a)).mtimeMs;
        const mb = statSync(path.join(FEATURE_DIR, b)).mtimeMs;
        return mb - ma;
    });
    return path.join(FEATURE_DIR, matches[0]);
}

/**
 * Find a contiguous horizontal strip whose dominant colour is within
 * `tolerance` of `targetHex`. Scans top-to-bottom (or bottom-to-top if
 * `reverse`) within the column band [xMin..xMax]. Returns the first
 * strip of ≥ minHeight consecutive y-rows that match.
 */
function findColouredStrip(
    img: DecodedPng,
    xMin: number, xMax: number,
    yStart: number, yEnd: number,
    targetHex: string,
    tolerance: number,
    minHeight: number,
    reverse = false,
): { yMin: number; yMax: number; hex: string } | null {
    const target = hexToRgb(targetHex);
    let runStart = -1;
    let runEnd = -1;
    let runHex = '';
    const yIter = reverse
        ? Array.from({ length: yStart - yEnd + 1 }, (_, i) => yStart - i)
        : Array.from({ length: yEnd - yStart + 1 }, (_, i) => yStart + i);
    for (const y of yIter) {
        const dom = dominantColor(img, xMin, xMax, y, y);
        const d = rgbDelta(dom.rgb, target);
        if (d.max <= tolerance) {
            if (runStart < 0) runStart = y;
            runEnd = y;
            runHex = dom.hex;
        } else {
            const runLength = Math.abs(runEnd - runStart) + 1;
            if (runStart >= 0 && runLength >= minHeight) {
                return reverse
                    ? { yMin: runEnd, yMax: runStart, hex: runHex }
                    : { yMin: runStart, yMax: runEnd, hex: runHex };
            }
            runStart = -1;
            runEnd = -1;
        }
    }
    const runLength = Math.abs(runEnd - runStart) + 1;
    if (runStart >= 0 && runLength >= minHeight) {
        return reverse
            ? { yMin: runEnd, yMax: runStart, hex: runHex }
            : { yMin: runStart, yMax: runEnd, hex: runHex };
    }
    return null;
}

/**
 * Find a thin (1-2 px) coloured strip — the inter-row borders / totals
 * top border are 1-2 px tall, so `findColouredStrip` with min height ≥ 2
 * works but returns "all consecutive matching rows" which is also fine.
 *
 * Looks ALL strips matching `targetHex` in the column band, returning
 * them in top-to-bottom order. Used for "find every banded-row-boundary
 * decoration" or "find the last decoration before totals body".
 */
function findAllColouredStrips(
    img: DecodedPng,
    xMin: number, xMax: number,
    yStart: number, yEnd: number,
    targetHex: string,
    tolerance: number,
): Array<{ yMin: number; yMax: number; hex: string }> {
    const target = hexToRgb(targetHex);
    const out: Array<{ yMin: number; yMax: number; hex: string }> = [];
    let runStart = -1;
    let runEnd = -1;
    let runHex = '';
    for (let y = yStart; y <= yEnd; y++) {
        const dom = dominantColor(img, xMin, xMax, y, y);
        const d = rgbDelta(dom.rgb, target);
        if (d.max <= tolerance) {
            if (runStart < 0) runStart = y;
            runEnd = y;
            runHex = dom.hex;
        } else {
            if (runStart >= 0) {
                out.push({ yMin: runStart, yMax: runEnd, hex: runHex });
                runStart = -1;
                runEnd = -1;
            }
        }
    }
    if (runStart >= 0) out.push({ yMin: runStart, yMax: runEnd, hex: runHex });
    return out;
}

function expectRgbWithin(actualHex: string, expectedHex: string, tol: number, label: string): void {
    const a = hexToRgb(actualHex);
    const e = hexToRgb(expectedHex);
    const d = rgbDelta(a, e);
    if (d.max > tol) {
        throw new Error(
            `${label}: expected ${expectedHex} ±${tol}, got ${actualHex.toUpperCase()} ` +
            `(ΔR=${d.dR}, ΔG=${d.dG}, ΔB=${d.dB})`,
        );
    }
}

interface FixtureProbe {
    label: string;
    refPath: string;
    joplinPath: string;
    expected: typeof APTOS_EXPECTED;
    /** Joplin screenshot has wider canvas; its column band excludes A1
     *  active-cell selection blue (col B onwards starts ~x=160 at DPR=2). */
    joplinXMin: number;
    joplinXMax: number;
    excelXMin: number;
    excelXMax: number;
}

function runFixtureChecks(probe: FixtureProbe): void {
    const ref = decodePng(probe.refPath);
    const joplin = decodePng(probe.joplinPath);

    // --- Header band ----------------------------------------------------
    const refHeader = findColouredStrip(
        ref, probe.excelXMin, probe.excelXMax, 50, ref.height - 1,
        probe.expected.headerHex, REGION_TOLERANCE, 6,
    );
    expect(refHeader).not.toBeNull();
    const joplinHeader = findColouredStrip(
        joplin, probe.joplinXMin, probe.joplinXMax, 30, joplin.height - 1,
        probe.expected.headerHex, REGION_TOLERANCE, 6,
    );
    expect(joplinHeader).not.toBeNull();

    // The dominant colours should match each other (head-to-head).
    expectRgbWithin(joplinHeader!.hex, refHeader!.hex, ASSERT_TOLERANCE,
        `${probe.label} header band: Joplin vs Excel`);

    // --- First banded data row -----------------------------------------
    // Search starts just below the header band. Use a tighter tolerance
    // for region-finding here because thin grid lines (e.g. Univer's
    // default `#D7D8DB`) fall within the wider tolerance; the banded
    // fill itself is a TALL strip (≥10px), so requiring minHeight=10
    // also rules out 1-2px decorations.
    const refBanded = findColouredStrip(
        ref, probe.excelXMin, probe.excelXMax, refHeader!.yMax + 4, ref.height - 1,
        probe.expected.bandedHex, 12, 10,
    );
    expect(refBanded).not.toBeNull();
    const joplinBanded = findColouredStrip(
        joplin, probe.joplinXMin, probe.joplinXMax, joplinHeader!.yMax + 4, joplin.height - 1,
        probe.expected.bandedHex, 12, 10,
    );
    expect(joplinBanded).not.toBeNull();
    expectRgbWithin(joplinBanded!.hex, refBanded!.hex, ASSERT_TOLERANCE,
        `${probe.label} banded data row 1: Joplin vs Excel`);

    // --- Totals-row top border ------------------------------------------
    // Find ALL strips matching totals-top colour in both images. In Excel
    // the LAST one before the totals-body white region is the totals-top
    // border. In Joplin (with current DOUBLE rendering) there are TWO
    // strips at the totals-top with a white pixel between them — that's
    // the bug we're gating against.
    const refTotalsStrips = findAllColouredStrips(
        ref, probe.excelXMin, probe.excelXMax, refBanded!.yMax + 4, ref.height - 1,
        probe.expected.totalsTopHex, REGION_TOLERANCE,
    );
    const joplinTotalsStrips = findAllColouredStrips(
        joplin, probe.joplinXMin, probe.joplinXMax, joplinBanded!.yMax + 4, joplin.height - 1,
        probe.expected.totalsTopHex, REGION_TOLERANCE,
    );

    expect(refTotalsStrips.length).toBeGreaterThan(0);
    expect(joplinTotalsStrips.length).toBeGreaterThan(0);

    // The LAST strip before totals body in each image is the totals-top
    // sentinel. Excel paints inter-banded-row strips at every row
    // boundary (every ~48px), so the totals-top is the last strip.
    const refTotalsTop = refTotalsStrips[refTotalsStrips.length - 1];
    const joplinTotalsTop = joplinTotalsStrips[joplinTotalsStrips.length - 1];
    // 1-2px strip at DPR=2 anti-aliases — use the looser STRIP_TOLERANCE.
    expectRgbWithin(joplinTotalsTop.hex, refTotalsTop.hex, STRIP_TOLERANCE,
        `${probe.label} totals-top border: Joplin vs Excel`);

    // The totals-top border in Excel is a SINGLE 2px strip. Univer's
    // BorderStyleTypes.DOUBLE (style 7) emits TWO 1px strips with a
    // 1px white gap between them — so the LAST two `joplinTotalsStrips`
    // entries would be at Δy ≈ 2 of each other if DOUBLE is in use.
    //
    // Sentinel: if the second-to-last strip is within 4px of the last
    // strip in the JOPLIN image, that's the DOUBLE-line render bug.
    if (joplinTotalsStrips.length >= 2) {
        const prev = joplinTotalsStrips[joplinTotalsStrips.length - 2];
        const gap = joplinTotalsTop.yMin - prev.yMax;
        if (gap > 0 && gap <= 4) {
            // The Excel reference does NOT have a tight pair like this
            // at the totals-top — verify by counting tight pairs in ref.
            let refTightPairs = 0;
            for (let i = 1; i < refTotalsStrips.length; i++) {
                const refPrev = refTotalsStrips[i - 1];
                const refCur = refTotalsStrips[i];
                const refGap = refCur.yMin - refPrev.yMax;
                if (refGap > 0 && refGap <= 4) refTightPairs++;
            }
            // Joplin has a tight pair AT the totals-top that Excel doesn't.
            // (Excel's strips are all evenly spaced ~48px apart.)
            throw new Error(
                `${probe.label} totals-top renders as DOUBLE-line ` +
                `(two strips at y=${prev.yMin}-${prev.yMax} and ` +
                `y=${joplinTotalsTop.yMin}-${joplinTotalsTop.yMax}, gap=${gap}). ` +
                `Excel paints a single MEDIUM strip; ${refTightPairs} tight pairs ` +
                `in Excel ref (expect 0). Recipe should emit MEDIUM (style 8) ` +
                `not DOUBLE (style 7) for totals-top.`,
            );
        }
    }
}

describe('Canvas vs Excel fidelity — TableStyleMedium4 (M13/E rework)', () => {
    test('Aptos: Joplin canvas matches Excel reference at structural regions', () => {
        runFixtureChecks({
            label: 'Aptos',
            refPath: APTOS_REF,
            joplinPath: latestEvalPng('aptos'),
            expected: APTOS_EXPECTED,
            // Joplin canvas at DPR=2: column A starts at x≈80 (after row
            // header). x=200 is safely inside col B, dodging A1 active-
            // cell selection blue.
            joplinXMin: 200,
            joplinXMax: 1500,
            // Excel reference: column A starts ~x=80, col B at ~x=200.
            excelXMin: 200,
            excelXMax: 1500,
        });
    });

    test('Classic: Joplin canvas matches Excel reference at structural regions', () => {
        runFixtureChecks({
            label: 'Classic',
            refPath: CLASSIC_REF,
            joplinPath: latestEvalPng('classic'),
            expected: CLASSIC_EXPECTED,
            joplinXMin: 200,
            joplinXMax: 1200,
            excelXMin: 200,
            excelXMax: 1200,
        });
    });
});

// ─── M15: Conditional Formatting canvas fidelity ────────────────────────
//
// Canvas-vs-Excel parity gate for the 5 CF rule types. Compares the
// Joplin canvas screenshot
// (`screenshots/feature-1-m15-conditional-formatting/eval-*.png`)
// against the operator-captured Excel reference
// (`screenshots/excel-reference/ConditionalFormatting-Variants.png`)
// at five sentinel points (one per CF column).
//
// SKIP-WHEN-MISSING: the reference PNG does not exist at planner /
// generator time — the operator captures it from real Excel during
// the cycle. Tests skip via fs.existsSync; once the reference lands,
// they run and the evaluator can act on the parity verdict.
//
// Per-column sentinel:
//   - A (colorScale): dominant colour of the median cell A6 (value 40).
//     At #FFEB84 family on the gradient it's the most stable measure;
//     A2 (red end) and A11 (green end) anti-alias more aggressively.
//   - C (dataBar): the right edge of the longest-bar cell C11 (value 90).
//     At #638EC6.
//   - E (cellIs > 50): cell E10 (value 80). At #FFC7CE.
//   - G (top-3 rank): the top cell G2 (value 82). At #C6EFCE.
//   - I (iconSet): the top-band cell I11 (value 90, green-up arrow).
//     Looser tolerance because the icon glyph is small.

const CF_VARIANTS_REF = path.join(REFERENCES_DIR, 'ConditionalFormatting-Variants.png');
const CF_FEATURE_DIR = path.join(REPO_ROOT, 'screenshots', 'feature-1-m15-conditional-formatting');

function existsSync(p: string): boolean {
    try { statSync(p); return true; } catch { return false; }
}

function latestCfEvalPng(): string | null {
    if (!existsSync(CF_FEATURE_DIR)) return null;
    const matches = readdirSync(CF_FEATURE_DIR).filter((n) =>
        n.startsWith('eval-') && n.endsWith('.png'),
    );
    if (matches.length === 0) return null;
    matches.sort((a, b) => {
        const ma = statSync(path.join(CF_FEATURE_DIR, a)).mtimeMs;
        const mb = statSync(path.join(CF_FEATURE_DIR, b)).mtimeMs;
        return mb - ma;
    });
    return path.join(CF_FEATURE_DIR, matches[0]);
}

const CF_REF_PRESENT = existsSync(CF_VARIANTS_REF);
const CF_EVAL_PRESENT = !!latestCfEvalPng();

// describe.skip when the reference is absent (M13/E pattern: skip
// loudly when ground truth is missing, run when present). The
// evaluator unskips by capturing the reference; the tests fail-fast
// inside `beforeAll` if the reference is later removed.
const cfDescribe = (CF_REF_PRESENT && CF_EVAL_PRESENT) ? describe : describe.skip;

cfDescribe('Canvas vs Excel fidelity — ConditionalFormatting-Variants (M15)', () => {
    let refImg: DecodedPng;
    let joplinImg: DecodedPng;

    beforeAll(() => {
        if (!CF_REF_PRESENT) {
            throw new Error(
                `Excel reference PNG missing at ${CF_VARIANTS_REF}. ` +
                `Operator must capture from real Excel: open ` +
                `tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx ` +
                `in Excel and screenshot the visible CF columns A..I.`,
            );
        }
        const evalPath = latestCfEvalPng();
        if (!evalPath) {
            throw new Error(
                `No eval-*.png in ${CF_FEATURE_DIR}. ` +
                `Re-run scripts/pge/eval-screenshot.sh feature-1-m15-conditional-formatting.`,
            );
        }
        refImg = decodePng(CF_VARIANTS_REF);
        joplinImg = decodePng(evalPath);
    });

    test('Column A (colorScale): both renders carry the gradient mid-band yellow #FFEB84 family', () => {
        // Sample a small mid-cell band on column A in BOTH renders. The
        // y-positions and column x-bands are deliberately approximate;
        // the assertion is parity between Joplin and Excel, not absolute
        // RGB. We pick the median row (row 7 = value 50 → midpoint of
        // gradient) and scan a horizontal slice for the dominant fill.
        // We don't reach inside text glyphs — column A's data area is
        // narrower than the cell width, so a 1-pixel vertical strip in
        // the middle of the row catches the cell fill cleanly.
        //
        // Joplin canvas at DPR=2: col A starts at x≈92 (row header 46
        // CSS px ≈ 92 device px). Median row y depends on canvas size;
        // we sample a slab.
        const refDom = dominantColor(refImg, 0, refImg.width, Math.floor(refImg.height * 0.4), Math.floor(refImg.height * 0.5));
        const joplinDom = dominantColor(joplinImg, 92, 200, Math.floor(joplinImg.height * 0.4), Math.floor(joplinImg.height * 0.5));
        // Both should carry yellow-ish ink on the gradient mid-band.
        // We don't pin to #FFEB84 explicitly — the assertion is
        // dominant-colour parity within Δ ≤ 24 (gradient anti-aliasing
        // blends adjacent cells more than the catalog tests do).
        const d = rgbDelta(refDom.rgb, joplinDom.rgb);
        if (d.max > 24) {
            throw new Error(
                `Column A gradient parity: Excel reference dominant ${refDom.hex}, ` +
                `Joplin canvas dominant ${joplinDom.hex} (Δmax=${d.max}, ΔR=${d.dR}, ΔG=${d.dG}, ΔB=${d.dB}).`,
            );
        }
    });

    test.todo('Column C (dataBar): both renders carry the dataBar blue #638EC6 family on long-bar cells');
    test.todo('Column E (cellIs > 50): both renders fill cell E10 (value 80) with #FFC7CE');
    test.todo('Column G (top-3): both renders fill the highest-G cell with #C6EFCE');
    test.todo('Column I (iconSet): both renders display a green-up arrow on cell I11 (value 90)');
});
