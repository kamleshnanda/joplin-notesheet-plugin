// Unit tests for the generic Text-to-Number parser (finding #3). Covers the
// numeric-text shapes from NumberFormats.xlsx and Excel's Convert-to-Number
// semantics: derive the numeric value AND the display format from the text.

import { parseTextToNumber } from '../src/formulas/textToNumber';

describe('parseTextToNumber — plain integers and floats', () => {
    test('"0" → 0, General', () => {
        expect(parseTextToNumber('0')).toEqual({ value: 0, pattern: null });
    });
    test('"1234" → 1234, General', () => {
        expect(parseTextToNumber('1234')).toEqual({ value: 1234, pattern: null });
    });
    test('"0.00" → 0 (float), General', () => {
        expect(parseTextToNumber('0.00')).toEqual({ value: 0, pattern: null });
    });
    test('"1234.5678" → float, General', () => {
        expect(parseTextToNumber('1234.5678')).toEqual({ value: 1234.5678, pattern: null });
    });
    test('"-1234.5" → negative float', () => {
        expect(parseTextToNumber('-1234.5')).toEqual({ value: -1234.5, pattern: null });
    });
    test('leading/trailing whitespace tolerated', () => {
        expect(parseTextToNumber('  42  ')).toEqual({ value: 42, pattern: null });
    });
});

describe('parseTextToNumber — percent', () => {
    test('"0%" → 0 with 0% pattern', () => {
        expect(parseTextToNumber('0%')).toEqual({ value: 0, pattern: '0%' });
    });
    test('"50%" → 0.5 with 0% pattern', () => {
        expect(parseTextToNumber('50%')).toEqual({ value: 0.5, pattern: '0%' });
    });
    test('"12.5%" → 0.125 with 0.0% pattern (preserves decimals)', () => {
        expect(parseTextToNumber('12.5%')).toEqual({ value: 0.125, pattern: '0.0%' });
    });
    test('"100 %" (space before %) → 1', () => {
        expect(parseTextToNumber('100 %')).toEqual({ value: 1, pattern: '0%' });
    });
});

describe('parseTextToNumber — grouped numbers', () => {
    test('"1,234" → 1234 with #,##0 pattern', () => {
        expect(parseTextToNumber('1,234')).toEqual({ value: 1234, pattern: '#,##0' });
    });
    test('"1,234.56" → 1234.56 with #,##0.00 pattern', () => {
        expect(parseTextToNumber('1,234.56')).toEqual({ value: 1234.56, pattern: '#,##0.00' });
    });
});

describe('parseTextToNumber — currency', () => {
    test('"$1,234.56" → 1234.56 with USD pattern', () => {
        expect(parseTextToNumber('$1,234.56')).toEqual({
            value: 1234.56,
            pattern: '"$"#,##0.00',
        });
    });
    test('"1234.56 €" (trailing symbol) → 1234.56 with EUR pattern', () => {
        expect(parseTextToNumber('1234.56 €')).toEqual({
            value: 1234.56,
            pattern: '#,##0.00 "€"',
        });
    });
});

describe('parseTextToNumber — accounting negatives', () => {
    test('"(1,234.00)" → -1234', () => {
        expect(parseTextToNumber('(1,234.00)')).toEqual({ value: -1234, pattern: '#,##0.00' });
    });
});

describe('parseTextToNumber — rejects non-numbers (must not clobber real text)', () => {
    test.each([
        '',
        '   ',
        'Hello',
        'General',
        '#,##0.00', // a format code string, not a number
        'm/d/yy',
        'N/A',
        '12abc',
        '1.2.3',
        '$', // symbol alone
    ])('%p → null', (input) => {
        expect(parseTextToNumber(input)).toBeNull();
    });
});
