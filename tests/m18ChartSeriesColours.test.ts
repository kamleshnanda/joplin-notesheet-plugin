// M18 C1: per-series chart colours from Excel <c:ser><c:spPr><a:solidFill>.
//
// Before C1, imported charts always used CHART_PALETTE — an Excel chart with
// custom series colours imported looking like a generic Notesheet chart, and
// re-export emitted palette colours, not the source ones. C1 reads each
// series' solidFill on import, carries it on the dataset (`color`), uses it
// in the live render + static SVG, and re-emits it on export. Falls back to
// CHART_PALETTE when the source series has no explicit fill.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import JSZip from 'jszip';
import { readChartsFromXlsxZip } from '../src/charts/xlsxChartImport';
import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// A bar chart with TWO series, each carrying an explicit <c:spPr> solidFill.
function chartXml(): string {
    const ser = (idx: number, name: string, rgb: string, vals: number[]) =>
        `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>` +
        `<c:tx><c:strRef><c:f>Sheet1!$${String.fromCharCode(66 + idx)}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
        `<c:spPr><a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill></c:spPr>` +
        `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
        `<c:val><c:numRef><c:f>Sheet1!$${String.fromCharCode(66 + idx)}$2:$${String.fromCharCode(66 + idx)}$3</c:f><c:numCache><c:ptCount val="2"/>` +
        vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
        `</c:numCache></c:numRef></c:val></c:ser>`;
    return (
        `<?xml version="1.0"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
        ser(0, 'Alpha', 'FF0000', [10, 20]) +
        ser(1, 'Beta', '00B050', [15, 25]) +
        `<c:axId val="1"/><c:axId val="2"/></c:barChart></c:plotArea></c:chart></c:chartSpace>`
    );
}

async function buildChartXlsx(): Promise<Buffer> {
    const zip = new JSZip();
    zip.file(
        '[Content_Types].xml',
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`,
    );
    zip.file(
        '_rels/.rels',
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    );
    zip.file(
        'xl/workbook.xml',
        `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    );
    zip.file(
        'xl/_rels/workbook.xml.rels',
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    );
    zip.file(
        'xl/worksheets/sheet1.xml',
        `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheetData/><drawing r:id="rId1"/></worksheet>`,
    );
    zip.file(
        'xl/worksheets/_rels/sheet1.xml.rels',
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    );
    zip.file(
        'xl/drawings/drawing1.xml',
        `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
    );
    zip.file(
        'xl/drawings/_rels/drawing1.xml.rels',
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`,
    );
    zip.file('xl/charts/chart1.xml', chartXml());
    return zip.generateAsync({ type: 'nodebuffer' });
}

describe('M18 C1 — per-series chart colours', () => {
    test('import reads each series solidFill into dataset.color', async () => {
        const charts = await readChartsFromXlsxZip(await buildChartXlsx());
        expect(charts.length).toBe(1);
        const ds = charts[0].datasets;
        expect(ds.length).toBe(2);
        // Colours are surfaced as #RRGGBB on each dataset.
        expect((ds[0] as { color?: string }).color).toBe('#FF0000');
        expect((ds[1] as { color?: string }).color).toBe('#00B050');
    });

    test('series without an explicit fill leaves color undefined (palette fallback later)', async () => {
        // Re-use the chart but strip the second series' spPr.
        const buf = await buildChartXlsx();
        const zip = await JSZip.loadAsync(buf);
        let chartXmlStr = await zip.file('xl/charts/chart1.xml')!.async('string');
        // Remove the 00B050 fill only.
        chartXmlStr = chartXmlStr.replace(
            '<c:spPr><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></c:spPr>',
            '',
        );
        zip.file('xl/charts/chart1.xml', chartXmlStr);
        const charts = await readChartsFromXlsxZip(await zip.generateAsync({ type: 'nodebuffer' }));
        const ds = charts[0].datasets;
        expect((ds[0] as { color?: string }).color).toBe('#FF0000');
        expect((ds[1] as { color?: string }).color).toBeUndefined();
    });

    test('round-trip: source series colours survive import -> export', async () => {
        const snap = await xlsxBufferToSnapshot((await buildChartXlsx()) as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(exported);
        const cpath = Object.keys(zip.files).find((p) => /xl\/charts\/chart\d+\.xml$/.test(p))!;
        const xml = await zip.files[cpath].async('string');
        // Both source series fills are re-emitted (not replaced by palette).
        expect(xml).toContain('<a:srgbClr val="FF0000"/>');
        expect(xml).toContain('<a:srgbClr val="00B050"/>');
        // The default palette blue (3B82F6) must NOT appear for series 0 —
        // the source red won.
        expect(xml).not.toContain('3B82F6');
    });

    test('palette fallback: a series with no source fill exports a palette colour', async () => {
        const buf = await buildChartXlsx();
        const zip0 = await JSZip.loadAsync(buf);
        let chartXmlStr = await zip0.file('xl/charts/chart1.xml')!.async('string');
        chartXmlStr = chartXmlStr.replace(
            '<c:spPr><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></c:spPr>',
            '',
        );
        zip0.file('xl/charts/chart1.xml', chartXmlStr);
        const snap = await xlsxBufferToSnapshot(
            (await zip0.generateAsync({ type: 'nodebuffer' })) as unknown as Buffer,
        );
        const exported = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(exported);
        const cpath = Object.keys(zip.files).find((p) => /xl\/charts\/chart\d+\.xml$/.test(p))!;
        const xml = await zip.files[cpath].async('string');
        // Series 0 keeps its source red; series 1 (no fill) gets palette[1] (EF4444).
        expect(xml).toContain('<a:srgbClr val="FF0000"/>');
        expect(xml).toContain('<a:srgbClr val="EF4444"/>');
    });
});
