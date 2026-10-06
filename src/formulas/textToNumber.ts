// M18 manual-test finding #3: a generic "Text to Number" conversion.
//
// Univer's built-in `sheet.command.text-to-number` only converts a cell whose
// stored string parses as a bare number (its `isRealNum` gate rejects anything
// `Number(...)` turns into NaN — so "0%", "1,234", "$5.00" all no-op). Excel's
// "Convert to Number" is more generous: it strips the format decoration, stores
// the real numeric value, and applies the matching number format.
//
// This module is the PURE parsing core (no Univer imports) so it can be unit
// tested directly. The Univer command wiring lives in src/editorView.tsx and
// calls parseTextToNumber() per cell.
//
// Deliberately conservative: we only convert text that is UNAMBIGUOUSLY a
// formatted number. Anything else (words, dates, mixed text) returns null and
// the cell is left untouched — we must not clobber cells that are genuinely
// text (Excel marks these with quotePrefix, and the user chose text on import).

export interface TextToNumberResult {
    // The numeric value to store (Univer CellValueType.NUMBER).
    value: number;
    // The number-format pattern to apply, or null to leave as General.
    // e.g. '0%' for a percent, '"$"#,##0.00' for USD. null → no numfmt (a plain
    // integer/float displays via General, matching Excel's Convert-to-Number).
    pattern: string | null;
}

// Parse a numeric-looking string into { value, pattern }, or null if the text
// isn't unambiguously a number. Handles:
//   "0" / "-12" / "1234"          → integer, General (pattern null)
//   "0.00" / "1234.5678"          → float, General (pattern null)
//   "1,234" / "1,234.56"          → grouped number → value 1234[.56], '#,##0[.00]'
//   "50%" / "0%" / "12.5%"        → value/100, '0%' or '0.00%'
//   "$1,234.56" / "1234.56 €"     → currency → value, currency-ish pattern
//   "(1,234.00)"                  → negative accounting → -1234
// Leading/trailing whitespace is tolerated. Returns null for empty, NaN, or
// text with disallowed characters.
export function parseTextToNumber(raw: string): TextToNumberResult | null {
    if (typeof raw !== 'string') return null;
    const s = raw.trim();
    if (s === '') return null;

    // ---- percent: trailing % (optionally after the number) ----
    const pct = s.match(/^([-+]?[\d,]*\.?\d+)\s*%$/);
    if (pct) {
        const numStr = pct[1].replace(/,/g, '');
        const n = Number(numStr);
        if (!Number.isFinite(n)) return null;
        // Preserve decimal precision in the pattern: "12.5%" → '0.00%'.
        const decimals = decimalCount(numStr);
        const pattern = decimals > 0 ? `0.${'0'.repeat(decimals)}%` : '0%';
        return { value: n / 100, pattern };
    }

    // ---- accounting negatives in parens: "(1,234.00)" → -1234 ----
    let working = s;
    let parenNegative = false;
    const paren = working.match(/^\((.*)\)$/);
    if (paren) {
        parenNegative = true;
        working = paren[1].trim();
    }

    // ---- currency: a leading or trailing currency symbol/code ----
    // Capture the symbol so the applied pattern shows it. Support the common
    // set ($ € £ ¥ and a trailing 3-letter code like "kr"/"USD" preceded by a
    // space). Kept intentionally small; unusual formats fall through to null.
    let currencyPattern: string | null = null;
    const leadSym = working.match(/^([$€£¥])\s?(.*)$/);
    const trailSym = working.match(/^(.*?)\s?([$€£¥])$/);
    const trailCode = working.match(/^(.*?)\s+([A-Za-z]{2,3})$/);
    if (leadSym) {
        currencyPattern = `"${leadSym[1]}"#,##0.00`;
        working = leadSym[2].trim();
    } else if (trailSym) {
        currencyPattern = `#,##0.00 "${trailSym[2]}"`;
        working = trailSym[1].trim();
    } else if (trailCode) {
        currencyPattern = `#,##0.00 "${trailCode[2]}"`;
        working = trailCode[1].trim();
    }

    // ---- plain / grouped number ----
    // Allow an optional sign, digit groups (commas), and a decimal part.
    const numMatch = working.match(/^([-+]?)(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?$/);
    if (!numMatch) return null;
    const grouped = /,/.test(numMatch[2]);
    const digits = numMatch[2].replace(/,/g, '');
    const decimalPart = numMatch[3] ?? '';
    const magnitude = Number(`${digits}${decimalPart}`);
    if (!Number.isFinite(magnitude)) return null;
    const signedValue = applySign(magnitude, numMatch[1], parenNegative);

    if (currencyPattern) {
        return { value: signedValue, pattern: currencyPattern };
    }
    if (grouped) {
        const decimals = decimalPart ? decimalPart.length - 1 : 0;
        const pattern = decimals > 0 ? `#,##0.${'0'.repeat(decimals)}` : '#,##0';
        return { value: signedValue, pattern };
    }
    // Plain integer/float: General format (Excel's Convert-to-Number leaves the
    // format General and lets the value display as typed).
    return { value: signedValue, pattern: null };
}

// Count digits after the decimal point in a numeric string ("12.50" → 2).
function decimalCount(numStr: string): number {
    const dot = numStr.indexOf('.');
    return dot < 0 ? 0 : numStr.length - dot - 1;
}

// Combine the magnitude with an explicit leading sign and/or accounting parens.
function applySign(magnitude: number, explicitSign: string, parenNegative: boolean): number {
    let negative = explicitSign === '-';
    if (parenNegative) negative = !negative; // parens flip; "-(5)" is unusual but consistent
    return negative ? -magnitude : magnitude;
}
