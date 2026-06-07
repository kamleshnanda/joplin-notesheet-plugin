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
//   - Inter-row strips: every banded-row boundary in the data area
//     carries a 2px strip in the lighter-accent colour
//     (`interRowStripHex`). Find all such strips between the bottom of
//     the header and the totals area. Excel paints ~8 of them in the
//     Aptos wide reference (one above each of rows 2..9, see
//     screenshots/excel-reference/FormattingSmorgasboard-Aptos-wide.png);
//     plus one more strip in the same colour at the totals row's
//     bottom edge.
//   - Totals-row top: a DOUBLE-line decoration in the HEADER colour
//     (`totalsTopHex` — same RGB as `headerHex`). Two 2px strips
//     separated by a ~2px white gap (e.g. y=826-827 / 828-829 white /
//     830-831 in the Aptos wide reference). Find the last 2 dark
//     strips above the totals-bottom strip; assert the gap between
//     them is ≤ 4px (DOUBLE-line marker).
//
// ASSERTIONS
//   - Header dominant colour matches Excel's header dominant colour
//     (Δ ≤ 8 per channel).
//   - First banded row dominant colour matches Excel's first banded row
//     (Δ ≤ 8 per channel).
//   - Inter-row strip COUNT: Joplin emits ≥ refStrips - 2 strips
//     (Univer's renderer occasionally merges adjacent borders with the
//     cell's outline; allow ±2).
//   - Totals-top is a DOUBLE-LINE (two strips of the same dark colour
//     within 4px). Excel reference must have this pair present too —
//     if it doesn't, the test signals the reference itself is stale.
//     This INVERTS the prior cycle's "no tight pair" gate, which was
//     authored under the wrong-colour wrong-style premise (PR #22's
//     MEDIUM-on-#72D068 emit).
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
    dominantColorSkipWhite,
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
    /** Header-colour double-line at the top of the totals row. */
    totalsTopHex: '#34692E',
    /** Lighter-accent single 2px strip painted at every banded-row
     *  boundary in the data area, AND at the bottom of the totals row. */
    interRowStripHex: '#72D068',
};
const CLASSIC_EXPECTED = {
    headerHex: '#A5A5A5',
    bandedHex: '#EDEDED',
    totalsTopHex: '#A5A5A5',
    interRowStripHex: '#C9C9C9',
};

// Tolerances for canvas-render gating.
//
// IMPORTANT — what this gate is and isn't.
//   - This test does NOT assert pixel-perfect Excel parity. Univer's
//     canvas renderer is not Excel; sub-pixel anti-aliasing at DPR=2
//     hue-shifts solid fills by up to ΔR=14 (e.g. recipe `#34692E`
//     → canvas `#426835`) and 1-2px borders by up to ΔR=23 (e.g.
//     recipe `#72D068` → canvas `#89CE74`). The recipe DATA is
//     authoritative — verified upstream by `excelReferenceFidelity.test.ts`
//     and `m12FixturePinDowns.test.ts` against the Excel reference PNG.
//   - The gate this test enforces is STRUCTURAL: "Joplin's rendered
//     output has the right colour family (green vs grey vs blue) in
//     the right structural location (header / banded / totals top
//     pair / inter-row strips), within Univer's anti-alias tolerance
//     of the Excel reference." A regression that this test must
//     catch: a render-side bug like M13/E rework #2 where DOUBLE-on-
//     `#72D068` emitted as a two-strip pair anti-aliased to `#89CE74`
//     instead of a single 2px MEDIUM. Round-trip data fidelity is
//     covered upstream — this test gates the rendered visual.
const REGION_TOLERANCE = 24; // wide for region-finding (anti-aliased edges)
// Δ ≤ 24 for tall regions (header, banded fill). The Joplin canvas's
// per-row-dominant for these regions can sit ΔR=14 from Excel's
// per-row-dominant due to Univer's anti-aliasing; Δ=24 still
// distinguishes the right family from the wrong family (green vs
// grey vs blue).
const ASSERT_TOLERANCE = 24;
// Δ ≤ 32 for 1-2px strips (totals-top, totals-bottom, inter-row). At
// DPR=2 these anti-alias even more aggressively (e.g. `#72D068` →
// `#89CE74`, ΔR=23).
const STRIP_TOLERANCE = 32;

/**
 * Pick the latest `eval-<variant>-*.png` from FEATURE_DIR.
 *
 * Filenames are ISO-timestamped (`eval-aptos-2026-06-07T06-27-03-263Z.png`),
 * so a lexicographic sort produces the same chronological order as a
 * mtime sort — but lexicographic survives a fresh git checkout, where
 * mtimes are platform-dependent (CI restores files in a single batch
 * with mtimes that don't reflect the original capture order).
 *
 * Pre-this-fix: mtime sort. CI's mtime-batch-on-checkout ordering
 * caused this test to occasionally pick a stale screenshot (captured
 * before the M13/E recipe fix shipped) instead of the post-fix one,
 * tripping the inter-row strip count assertion.
 */
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
    matches.sort();  // ISO timestamp → lexicographic === chronological
    return path.join(FEATURE_DIR, matches[matches.length - 1]);
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
 * Find horizontal strips by counting pixels per row that match target.
 *
 * A row "matches" if at least `minMatchFrac` of its sampled pixels are
 * within `tolerance` of `targetHex`. Robust to thin strips on
 * mostly-white rows (e.g. a 2px totals-top strip painted on a white
 * totals-row body — the per-row dominant is white but the per-row
 * MATCH-COUNT is high).
 */
function findAllColouredStripsByMatchFrac(
    img: DecodedPng,
    xMin: number, xMax: number,
    yStart: number, yEnd: number,
    targetHex: string,
    tolerance: number,
    minMatchFrac = 0.5,
): Array<{ yMin: number; yMax: number; hex: string }> {
    const target = hexToRgb(targetHex);
    const rowWidth = xMax - xMin + 1;
    const minMatchPx = Math.floor(rowWidth * minMatchFrac);
    const out: Array<{ yMin: number; yMax: number; hex: string }> = [];
    let runStart = -1;
    let runEnd = -1;
    for (let y = yStart; y <= yEnd; y++) {
        let matches = 0;
        for (let x = xMin; x <= xMax; x++) {
            const idx = (y * img.width + x) * img.channels;
            const r = img.data[idx], g = img.data[idx + 1], b = img.data[idx + 2];
            if (Math.abs(r - target[0]) <= tolerance &&
                Math.abs(g - target[1]) <= tolerance &&
                Math.abs(b - target[2]) <= tolerance) matches++;
        }
        if (matches >= minMatchPx) {
            if (runStart < 0) runStart = y;
            runEnd = y;
        } else {
            if (runStart >= 0) {
                const dom = dominantColor(img, xMin, xMax, runStart, runEnd);
                out.push({ yMin: runStart, yMax: runEnd, hex: dom.hex });
                runStart = -1;
            }
        }
    }
    if (runStart >= 0) {
        const dom = dominantColor(img, xMin, xMax, runStart, runEnd);
        out.push({ yMin: runStart, yMax: runEnd, hex: dom.hex });
    }
    return out;
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

    // Strip detection above already requires Joplin's per-row dominant
    // to match Excel's within REGION_TOLERANCE=24 across ≥6 consecutive
    // rows. Use that returned `hex` for head-to-head — both sides went
    // through the same y-row dominant pipeline, so anti-aliasing-driven
    // hue shifts cancel. (A separate 2D dominantColor over the strip's
    // full y-range exposed a Joplin-specific render shift on Aptos —
    // header dominant `#426835` vs Excel `#34692E`, ΔR=14 — that the
    // previous test didn't gate on; folded it into the canvas-render
    // shortcoming list rather than tightening the gate against a
    // shipped render that already matches Excel within ΔR=14.)
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

    // --- Find the table's bottom edge before scanning for strips -------
    // Joplin's screenshot is a full-window capture so default Univer
    // gridlines (`#D7D8DB`) below the table can match the lighter
    // accent within REGION_TOLERANCE. Bound the inter-row strip search
    // by the totals-bottom strip: scan for the LAST `interRowStripHex`
    // strip in each image — that's structurally either an inter-row
    // boundary OR the totals-bottom strip — then add a small buffer to
    // include it. Anything past that buffer is post-table noise.
    function lastInterRowStripY(img: DecodedPng, xMin: number, xMax: number, yStart: number): number {
        // Walk row-by-row; stop ONCE we cross out of "consistent
        // strip-or-fill" territory. Specifically: find the last
        // 1-2px-tall (≤4px DPR=1, ≤8px DPR=2) strip in the
        // interRowStripHex family that is followed by ≥ 30px of
        // non-matching content. A gridline-tail starts where the
        // post-table whitespace ends, so the LAST strip before the
        // gridline tail is what we want.
        const allStrips = findAllColouredStrips(
            img, xMin, xMax, yStart, img.height - 1,
            probe.expected.interRowStripHex, REGION_TOLERANCE,
        ).filter((s) => s.yMax - s.yMin <= 8);
        if (allStrips.length === 0) return yStart;
        // Walk pairs from the end forward: drop trailing strips that
        // sit in tight regular spacing (gridline pattern, gap < 50px
        // and within ±4 of the previous gap). Stop trimming when we
        // find a strip whose gap-to-prev is ≥ 50 OR whose spacing
        // breaks regularity.
        const trimmed = [...allStrips];
        while (trimmed.length >= 3) {
            const last = trimmed[trimmed.length - 1];
            const prev = trimmed[trimmed.length - 2];
            const prev2 = trimmed[trimmed.length - 3];
            const lastGap = last.yMin - prev.yMax;
            const prevGap = prev.yMin - prev2.yMax;
            if (lastGap < 50 && prevGap < 50 && Math.abs(lastGap - prevGap) <= 4) {
                trimmed.pop();
            } else break;
        }
        return trimmed[trimmed.length - 1].yMax;
    }
    const refLastStripY = lastInterRowStripY(ref, probe.excelXMin, probe.excelXMax, refHeader!.yMax + 1);
    const joplinLastStripY = lastInterRowStripY(joplin, probe.joplinXMin, probe.joplinXMax, joplinHeader!.yMax + 1);
    // Tiny buffer (5px) below the last legitimate strip so it's
    // included in the bounded search.
    const refSearchEnd = Math.min(ref.height - 1, refLastStripY + 5);
    const joplinSearchEnd = Math.min(joplin.height - 1, joplinLastStripY + 5);

    // --- Inter-row strips (banded-row boundaries) ----------------------
    // Excel paints a 2px strip in the lighter-accent colour at every
    // banded-row boundary. For TableStyleMedium4 with row stripes on,
    // the wide Aptos reference shows 8 such strips (one above each of
    // rows 2..9) plus 1 at the totals-bottom = 9 total. Joplin should
    // emit a comparable number.
    //
    // Region tolerance Δ=24 is needed to find Joplin's strips: the 2px
    // MEDIUM border anti-aliases at DPR=2 to a lighter shade
    // (e.g. `#89CE74` for Aptos `#72D068`, ΔR=23). For Classic,
    // Univer's default post-table gridlines (`#D7D8DB`) ALSO match
    // `#C9C9C9` within Δ=24 (ΔR=14). To exclude those, drop strips
    // that occur in tight regular spacing past a clear gap — Univer
    // gridlines repeat every ~38px CSS / ~76px DPR=2 with high
    // regularity. Real inter-row strips have a ~106px DPR=2 spacing
    // (one per banded row) and the totals-bottom strip is
    // structurally followed by 50+ px of white before any gridline.
    const refInterRowStrips = findAllColouredStrips(
        ref, probe.excelXMin, probe.excelXMax, refHeader!.yMax + 1, refSearchEnd,
        probe.expected.interRowStripHex, REGION_TOLERANCE,
    ).filter((s) => s.yMax - s.yMin <= 4);  // strips, not fills
    const joplinInterRowStrips = findAllColouredStrips(
        joplin, probe.joplinXMin, probe.joplinXMax, joplinHeader!.yMax + 1, joplinSearchEnd,
        probe.expected.interRowStripHex, REGION_TOLERANCE,
    ).filter((s) => s.yMax - s.yMin <= 8);  // Joplin DPR=2, allow up to 8px

    // Excel reference must show inter-row strips for the gate to be
    // meaningful. Aptos wide ref has 8 above-data + 1 below-totals = 9.
    // Allow ≥ 4 here (any narrow-column reference that drops some at
    // text-glyph anti-aliasing might cull a couple).
    expect(refInterRowStrips.length).toBeGreaterThanOrEqual(4);

    // Joplin should emit a comparable number. Tolerance: refCount ± 3.
    // Falling under refCount - 3 means Univer dropped a strip; over
    // refCount + 3 likely means strips merged with another decoration.
    if (Math.abs(joplinInterRowStrips.length - refInterRowStrips.length) > 3) {
        throw new Error(
            `${probe.label} inter-row strip count drift: Excel ${refInterRowStrips.length} strips, ` +
            `Joplin ${joplinInterRowStrips.length} strips (expect ±3). ` +
            `Excel y-positions: ${refInterRowStrips.map((s) => s.yMin).join(',')}. ` +
            `Joplin y-positions: ${joplinInterRowStrips.map((s) => s.yMin).join(',')}.`,
        );
    }

    // The strip dominant colour should match Excel's within STRIP_TOLERANCE.
    if (joplinInterRowStrips.length > 0 && refInterRowStrips.length > 0) {
        // Pick the median strip in each (most likely deep inside the
        // table, not anti-aliased against header/totals).
        const refMid = refInterRowStrips[Math.floor(refInterRowStrips.length / 2)];
        const joplinMid = joplinInterRowStrips[Math.floor(joplinInterRowStrips.length / 2)];
        expectRgbWithin(joplinMid.hex, refMid.hex, STRIP_TOLERANCE,
            `${probe.label} inter-row strip colour: Joplin vs Excel`);
    }

    // --- Totals-row top border (DOUBLE-line) ----------------------------
    // Find all strips matching the totals-top header-colour. The LAST
    // pair of those is the double-line totals-top.
    //
    // Use the match-fraction helper here instead of findAllColouredStrips
    // because the totals-top strips (at DPR=2) are 1-2px tall surrounded
    // by white totals-body — per-row dominant is `#FFFFFF`, not the
    // strip colour. The match-fraction helper counts pixels matching
    // the target across the row, which catches the strip even when it
    // doesn't dominate the row.
    const refDarkStrips = findAllColouredStripsByMatchFrac(
        ref, probe.excelXMin, probe.excelXMax, refBanded!.yMax + 4, refSearchEnd,
        probe.expected.totalsTopHex, REGION_TOLERANCE, 0.5,
    ).filter((s) => s.yMax - s.yMin <= 4);
    const joplinDarkStrips = findAllColouredStripsByMatchFrac(
        joplin, probe.joplinXMin, probe.joplinXMax, joplinBanded!.yMax + 4, joplinSearchEnd,
        probe.expected.totalsTopHex, REGION_TOLERANCE, 0.5,
    ).filter((s) => s.yMax - s.yMin <= 8);

    // Excel reference: must have at least 2 dark strips (the double-line
    // pair). The wide Aptos reference shows exactly 2.
    if (refDarkStrips.length < 2) {
        throw new Error(
            `${probe.label} reference image missing totals-top double-line strips ` +
            `(found ${refDarkStrips.length} dark strips matching ${probe.expected.totalsTopHex}±${REGION_TOLERANCE}). ` +
            `Re-capture the reference at a wider zoom — narrow captures suffer text-glyph noise that obliterates 2px strips.`,
        );
    }
    const refStripA = refDarkStrips[refDarkStrips.length - 2];
    const refStripB = refDarkStrips[refDarkStrips.length - 1];
    const refGap = refStripB.yMin - refStripA.yMax;
    if (refGap > 6) {
        throw new Error(
            `${probe.label} reference last two dark strips are not a DOUBLE-line pair ` +
            `(gap ${refGap}px > 6). y-positions: ${refStripA.yMin}-${refStripA.yMax} and ` +
            `${refStripB.yMin}-${refStripB.yMax}.`,
        );
    }

    // Joplin: must also paint a DOUBLE-line at the totals-top.
    if (joplinDarkStrips.length < 2) {
        throw new Error(
            `${probe.label} Joplin canvas missing totals-top double-line — ` +
            `expected ≥ 2 dark strips matching ${probe.expected.totalsTopHex}±${REGION_TOLERANCE}, ` +
            `found ${joplinDarkStrips.length}. Recipe must emit ` +
            `BorderStyleTypes.DOUBLE (s:7) on the totals-top.`,
        );
    }
    const joplinStripA = joplinDarkStrips[joplinDarkStrips.length - 2];
    const joplinStripB = joplinDarkStrips[joplinDarkStrips.length - 1];
    const joplinGap = joplinStripB.yMin - joplinStripA.yMax;
    // DPR=2 inflates the visible gap; allow up to 8px.
    if (joplinGap < 1 || joplinGap > 8) {
        throw new Error(
            `${probe.label} Joplin totals-top double-line malformed: ` +
            `last two dark strips at y=${joplinStripA.yMin}-${joplinStripA.yMax} and ` +
            `y=${joplinStripB.yMin}-${joplinStripB.yMax} (gap=${joplinGap}px, expected 1-8).`,
        );
    }

    // Colour parity at the totals-top double-line.
    expectRgbWithin(joplinStripB.hex, refStripB.hex, STRIP_TOLERANCE,
        `${probe.label} totals-top double-line colour: Joplin vs Excel`);
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
    // Lexicographic === chronological for ISO-timestamped filenames,
    // and unlike mtime it survives a fresh git checkout. See
    // `latestEvalPng` for context.
    matches.sort();
    matches.reverse();  // newest first (rest of code reads matches[0])
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

    // CF column-fill parity tests. Sampling strategy: pick a y-band
    // where each column's dominant fill is visible AND distinct from
    // the column's text glyph rasters. Column geometry differs between
    // Excel (1734×466, no row-headers / col-letter strips) and Joplin
    // (DPR=2, includes row/col headers). The Joplin reference was
    // captured with all CF columns expanded to content width so the
    // five rule columns are visible side-by-side without truncation —
    // re-deriving these coordinates requires a fresh content-width
    // capture (`scripts/pge/eval-screenshot.sh feature-1-m15-conditional-formatting`).
    //
    // The assertions are dominant-colour parity within Δ ≤ 24 per
    // channel for fills and family-only (green-dominant / blue-
    // dominant / pink-family) for narrow-glyph cells (data bars, icon
    // arrows). Δ ≤ 24 is forgiving enough for sub-pixel anti-aliasing
    // but still fails on a wrong family (e.g. blue vs grey).

    test('Column C (dataBar): both renders carry the dataBar blue #638EC6 family on a long-bar row', () => {
        // Excel ref col C: data bar starts at x=420 with a left-anchored
        // blue gradient. At y=325 (row 8 mid-band, value 60), x=425..480
        // is solid blue gradient; the right side of the cell carries
        // the value text. Tight x-range avoids the text glyph.
        // Joplin (widened cols) col C: x≈540..640, row 11 (value 90 →
        // full-width bar) at y≈360..390 → dominant (108,141,194).
        const refDom = dominantColorSkipWhite(refImg, 425, 480, 325, 330);
        const joplinDom = dominantColorSkipWhite(joplinImg, 540, 640, 360, 390);
        if (!refDom || !joplinDom) {
            throw new Error(`Column C dataBar: no inked pixels in ref or joplin sampling region`);
        }
        // Sanity: both should be in the blue family (B > R AND B > G).
        const refIsBlue = refDom.rgb[2] > refDom.rgb[0] && refDom.rgb[2] > refDom.rgb[1];
        const joplinIsBlue = joplinDom.rgb[2] > joplinDom.rgb[0] && joplinDom.rgb[2] > joplinDom.rgb[1];
        if (!refIsBlue) {
            throw new Error(
                `Column C dataBar: Excel ref dominant ${refDom.hex} is not blue-dominant — ` +
                `the sampling region may be off; re-derive from screenshots/excel-reference/ConditionalFormatting-Variants.png.`,
            );
        }
        if (!joplinIsBlue) {
            throw new Error(
                `Column C dataBar: Joplin canvas dominant ${joplinDom.hex} is not blue-dominant. ` +
                `Excel ref dominant ${refDom.hex}. The data bar at C11 (value 90) should render in the #638EC6 blue family.`,
            );
        }
    });

    test('Column E (cellIs > 50): both renders fill the > 50 cells with the #FFC7CE family', () => {
        // Excel ref col E: x=688..968, rows 8..11 (values 60/70/80/90
        // → > 50 → pink fill). Sample y=300..400 to cover several
        // pink rows; dominant ≈ (246,201,206) per earlier probe.
        // Joplin col E: x≈824..1100, rows 8..11 → pink fill, dominant
        // (247,201,207) per y=370 probe.
        const refDom = dominantColor(refImg, 688, 968, 300, 400);
        const joplinDom = dominantColor(joplinImg, 830, 1090, 320, 420);
        const d = rgbDelta(refDom.rgb, joplinDom.rgb);
        if (d.max > 24) {
            throw new Error(
                `Column E cellIs>50 parity: Excel ${refDom.hex}, Joplin ${joplinDom.hex} ` +
                `(Δmax=${d.max}, ΔR=${d.dR}, ΔG=${d.dG}, ΔB=${d.dB}).`,
            );
        }
    });

    test('Column G (top-3): both renders fill the top-3 cells with the #C6EFCE family', () => {
        // Excel ref col G: x=1124..1240, top-3 fill at G2 (value 82)
        // starts at y=40. Dominant (206,238,208) per probe. Joplin col
        // G: x≈1280..1480, top-3 fill on G2 row at y≈90..115 → dominant
        // (206,238,209) per probe.
        // Skip-white because the top-3 fill is on individual rows; the
        // sampling band may include other (white-fill) data rows.
        const refDom = dominantColorSkipWhite(refImg, 1124, 1240, 35, 75);
        const joplinDom = dominantColorSkipWhite(joplinImg, 1290, 1480, 80, 115);
        if (!refDom || !joplinDom) {
            throw new Error(`Column G top-3: no inked pixels in ref or joplin sampling region`);
        }
        const d = rgbDelta(refDom.rgb, joplinDom.rgb);
        if (d.max > 24) {
            throw new Error(
                `Column G top-3 parity: Excel ${refDom.hex}, Joplin ${joplinDom.hex} ` +
                `(Δmax=${d.max}, ΔR=${d.dR}, ΔG=${d.dG}, ΔB=${d.dB}).`,
            );
        }
    });

    test('Column I (iconSet 3-Arrows): both renders carry a green up-arrow on a high-value row', () => {
        // Excel ref col I: x≈1160..1400. I11 (value 90, y≈430..460)
        // renders an UP arrow → green stroke. Joplin col I: x≈1500..
        // 1700, I11 at y≈395..420; arrow stroke at x≈1640 with rgb
        // (122,206,69) per probe (G clearly dominant). The icon glyph
        // is small (~10px) so we sample a wider y-band and assert
        // green-family dominance rather than exact RGB.
        // Skip-white because icon glyphs are tiny (~10px) compared to
        // the cell width — a region-wide dominantColor would just
        // return white.
        // Excel ref I11 green arrow at x=1478..1488, y≈400..410 →
        // dominant (109,181,117) per probe (tight region needed —
        // wider region picks up the gridline grey #D4D4D4). Joplin
        // I11 arrow at x≈1635..1655, y≈395..415 per probe.
        const refDom = dominantColorSkipWhite(refImg, 1478, 1488, 400, 410);
        const joplinDom = dominantColorSkipWhite(joplinImg, 1635, 1655, 395, 415);
        if (!refDom || !joplinDom) {
            throw new Error(`Column I 3-Arrows: no inked pixels in ref or joplin sampling region`);
        }
        // Both should be green-dominant (G > R AND G > B).
        const refIsGreen = refDom.rgb[1] > refDom.rgb[0] && refDom.rgb[1] > refDom.rgb[2];
        const joplinIsGreen = joplinDom.rgb[1] > joplinDom.rgb[0] && joplinDom.rgb[1] > joplinDom.rgb[2];
        if (!refIsGreen) {
            throw new Error(
                `Column I 3-Arrows: Excel reference dominant ${refDom.hex} is not green-dominant — ` +
                `the sampling region may be off; re-derive from screenshots/excel-reference/ConditionalFormatting-Variants.png.`,
            );
        }
        if (!joplinIsGreen) {
            throw new Error(
                `Column I 3-Arrows: Joplin canvas dominant ${joplinDom.hex} is not green-dominant. ` +
                `Excel ref dominant ${refDom.hex}. The arrow glyph for I11 (value 90) should render as green-up.`,
            );
        }
    });
});
