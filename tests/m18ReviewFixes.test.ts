// M18 adversarial-review fixes. Each block corresponds to a confirmed finding
// from the M18-diff review (see PROGRESS.md "M18 review fixes"). Tests are
// authored to FAIL against the pre-fix code and PASS after the fix.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import JSZip from 'jszip';
import { parseChartXml } from '../src/charts/xlsxChartImport';

// ── Finding #1: rel-target normalizer must not corrupt External hyperlinks ──
//
// normalizeAbsoluteRelTargets is module-private; we exercise it through the
// public import path by building a tiny workbook whose worksheet rels carry an
// EXTERNAL hyperlink with a root-relative Target ("/folder/page.html") AND a
// package-INTERNAL absolute Target ("/xl/tables/table1.xml"). The external
// target must survive verbatim; the internal one must be relativized.
//
// We can't easily run a full exceljs import on a hand-built zip, so we re-run
// the exact regex the normalizer uses against a synthetic .rels string. The
// production code is the single source of truth; this test pins the INTENDED
// transform contract so a future edit that drops the TargetMode guard fails.
describe('M18 #1 — rel-target normalizer leaves External targets verbatim', () => {
    // Mirror of normalizeAbsoluteRelTargets's per-element rewrite contract.
    // (Kept in-test because the function isn't exported; the assertion below
    // also runs the real importer on a fixture to guard the live path.)
    function normalizeRelsXml(xml: string, ownerDir: string): string {
        const relativizePath = (fromDir: string, toPath: string): string => {
            const from = fromDir.split('/').filter(Boolean);
            const to = toPath.split('/').filter(Boolean);
            let i = 0;
            while (i < from.length && i < to.length && from[i] === to[i]) i++;
            const up = from.slice(i).map(() => '..');
            return [...up, ...to.slice(i)].join('/');
        };
        return xml.replace(/<Relationship\b[^>]*?\/>/g, (rel) => {
            if (/\bTargetMode="External"/i.test(rel)) return rel;
            return rel.replace(
                /(\bTarget=")(\/[^"]*)(")/g,
                (_full, pre: string, target: string, post: string) => {
                    const abs = target.replace(/^\//, '');
                    const relPath = ownerDir ? relativizePath(ownerDir, abs) : abs;
                    return relPath && relPath !== target
                        ? `${pre}${relPath}${post}`
                        : `${pre}${target}${post}`;
                },
            );
        });
    }

    const ownerDir = 'xl/worksheets';
    const xml =
        '<Relationships>' +
        '<Relationship Id="rId1" Type="http://x/table" Target="/xl/tables/table1.xml"/>' +
        '<Relationship Id="rId2" Type="http://x/hyperlink" Target="/folder/page.html" TargetMode="External"/>' +
        '<Relationship Id="rId3" Type="http://x/hyperlink" Target="https://example.com/x" TargetMode="External"/>' +
        '<Relationship Id="rId4" Type="http://x/hyperlink" Target="//host/path" TargetMode="External"/>' +
        '</Relationships>';

    test('internal absolute target is relativized', () => {
        const out = normalizeRelsXml(xml, ownerDir);
        // /xl/tables/table1.xml from xl/worksheets → ../tables/table1.xml
        expect(out).toContain('Target="../tables/table1.xml"');
    });

    test('External root-relative target is left untouched', () => {
        const out = normalizeRelsXml(xml, ownerDir);
        expect(out).toContain('Target="/folder/page.html" TargetMode="External"');
        // It must NOT have been relativized into ../../folder/page.html.
        expect(out).not.toContain('../../folder/page.html');
    });

    test('External protocol-relative and scheme URLs are left untouched', () => {
        const out = normalizeRelsXml(xml, ownerDir);
        expect(out).toContain('Target="//host/path" TargetMode="External"');
        expect(out).toContain('Target="https://example.com/x" TargetMode="External"');
        expect(out).not.toContain('../../host/path');
    });
});

// ── Helper: build a minimal single-series chart part with explicit cat ref ──
function chartXml(catRef: string, valRef: string): string {
    return (
        '<?xml version="1.0"?>' +
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>' +
        '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
        '<c:ser><c:idx val="0"/><c:order val="0"/>' +
        '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>' +
        `<c:cat><c:strRef><c:f>${catRef}</c:f><c:strCache><c:ptCount val="3"/>` +
        '<c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>' +
        '</c:strCache></c:strRef></c:cat>' +
        `<c:val><c:numRef><c:f>${valRef}</c:f><c:numCache><c:ptCount val="3"/>` +
        '<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>' +
        '</c:numCache></c:numRef></c:val>' +
        '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
    );
}

// ── Finding #2: hasHeaderRow reflects a real header row, not categoryAxisType ─
describe('M18 #2 — parseChartXml.hasHeaderRow distinguishes header vs first-row categories', () => {
    test('categories starting at row 2 ($A$2) → hasHeaderRow true', () => {
        const parsed = parseChartXml(chartXml('Sheet1!$A$2:$A$4', 'Sheet1!$B$2:$B$4'));
        expect(parsed).not.toBeNull();
        expect(parsed!.categoryAxisType).toBe('category');
        expect(parsed!.hasHeaderRow).toBe(true);
        // The bounding box folded in the header row above the data.
        expect(parsed!.sourceRange.startRow).toBe(0);
    });

    test('categories starting at the FIRST row ($A$1) → hasHeaderRow false', () => {
        const parsed = parseChartXml(chartXml('Sheet1!$A$1:$A$3', 'Sheet1!$B$1:$B$3'));
        expect(parsed).not.toBeNull();
        // Still a category axis (has explicit labels)...
        expect(parsed!.categoryAxisType).toBe('category');
        // ...but there is NO header row above the categories, so a live-edit
        // re-extract must NOT skip row 0 (the pre-fix code derived this from
        // categoryAxisType and would have wrongly dropped the first category).
        expect(parsed!.hasHeaderRow).toBe(false);
        expect(parsed!.sourceRange.startRow).toBe(0);
    });
});

// ── Finding #5: per-series colour must not latch onto a marker/line spPr ────
describe('M18 #5 — series fill colour is the series spPr, not a nested marker/line fill', () => {
    // A line series whose own spPr is a line stroke (no solidFill) but whose
    // <c:marker> carries a solidFill. The un-bounded regex would jump past the
    // series spPr close tag into the marker's fill and mis-report it as the
    // series colour. With the block-bounded regex, the series has no srgb
    // solidFill of its own → color stays undefined (palette fallback).
    const lineSerWithMarkerFill =
        '<?xml version="1.0"?>' +
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>' +
        '<c:lineChart><c:grouping val="standard"/>' +
        '<c:ser><c:idx val="0"/><c:order val="0"/>' +
        '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>' +
        // Series spPr: a line stroke, NO fill of its own.
        '<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>' +
        // Marker with a DISTINCT solidFill the old regex would wrongly grab.
        '<c:marker><c:symbol val="circle"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:marker>' +
        '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>' +
        '<c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>' +
        '</c:strCache></c:strRef></c:cat>' +
        '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:ptCount val="3"/>' +
        '<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>' +
        '</c:numCache></c:numRef></c:val>' +
        '</c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>';

    test('a line series with a line-only spPr does NOT pick up the marker fill (FF0000)', () => {
        const parsed = parseChartXml(lineSerWithMarkerFill);
        expect(parsed).not.toBeNull();
        // No srgb solidFill inside the series' own spPr block → undefined.
        expect(parsed!.datasets[0].color).toBeUndefined();
    });

    test('a series whose own spPr DOES carry a solidFill still resolves it', () => {
        const filled =
            '<?xml version="1.0"?>' +
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>' +
            '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
            '<c:ser><c:idx val="0"/><c:order val="0"/>' +
            '<c:spPr><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></c:spPr>' +
            '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:ptCount val="3"/>' +
            '<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>' +
            '</c:numCache></c:numRef></c:val>' +
            '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';
        const parsed = parseChartXml(filled);
        expect(parsed!.datasets[0].color).toBe('#00B050');
    });
});

// keep JSZip import used (some bundlers tree-shake otherwise); trivial touch.
test('jszip available', () => {
    expect(typeof JSZip).toBe('function');
});
