// M18 C2: rich-text chart titles. A chart title can carry per-run formatting
// (bold / italic / size / colour) across multiple <a:r> runs. Before C2 we
// flattened the title to plain concatenated text, so re-export lost the bold
// segment, the colour, the size — the title came back as one uniform run.
//
// C2 captures the runs (text + rPr) on import into `titleRuns`, carries them
// through the snapshot, and re-emits each run with its <a:rPr> on export. The
// flat `title` string is kept for the live Chart.js render (which shows a
// single-font title — documented limitation) and as a fallback.

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

// Title with TWO runs: "Sales " (bold, red, 14pt) + "2026" (italic, 11pt).
const RICH_TITLE =
    `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>` +
    `<a:r><a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Sales </a:t></a:r>` +
    `<a:r><a:rPr lang="en-US" sz="1100" i="1"/><a:t>2026</a:t></a:r>` +
    `</a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;

function chartXml(): string {
    return (
        `<?xml version="1.0"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart>${RICH_TITLE}<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
        `<c:ser><c:idx val="0"/><c:order val="0"/>` +
        `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
        `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>` +
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

interface TitleRun {
    text: string;
    bold?: boolean;
    italic?: boolean;
    size?: number; // OOXML hundredths of a point (sz)
    color?: string; // #RRGGBB
}

describe('M18 C2 — rich-text chart titles', () => {
    test('import captures per-run title formatting into titleRuns', async () => {
        const charts = await readChartsFromXlsxZip(await buildChartXlsx());
        const runs = (charts[0] as { titleRuns?: TitleRun[] }).titleRuns;
        expect(runs).toBeDefined();
        expect(runs!.length).toBe(2);
        expect(runs![0]).toMatchObject({
            text: 'Sales ',
            bold: true,
            size: 1400,
            color: '#FF0000',
        });
        expect(runs![1]).toMatchObject({ text: '2026', italic: true, size: 1100 });
        // The flat title is still the concatenation (live render + fallback).
        expect(charts[0].title).toBe('Sales 2026');
    });

    test('export re-emits each run with its rPr (bold/colour/size preserved)', async () => {
        const snap = await xlsxBufferToSnapshot((await buildChartXlsx()) as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(exported);
        const cpath = Object.keys(zip.files).find((p) => /xl\/charts\/chart\d+\.xml$/.test(p))!;
        const xml = await zip.files[cpath].async('string');
        // Two title runs, with their formatting.
        const titleBlock = xml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
        const runCount = (titleBlock.match(/<a:r>/g) ?? []).length;
        expect(runCount).toBe(2);
        expect(titleBlock).toContain('<a:t>Sales </a:t>');
        expect(titleBlock).toContain('<a:t>2026</a:t>');
        expect(titleBlock).toMatch(/b="1"/); // bold run
        expect(titleBlock).toMatch(/i="1"/); // italic run
        expect(titleBlock).toContain('<a:srgbClr val="FF0000"/>'); // red run
        expect(titleBlock).toMatch(/sz="1400"/);
    });

    test('a plain (single-run) title still round-trips as one run', async () => {
        const buf = await buildChartXlsx();
        const zip0 = await JSZip.loadAsync(buf);
        let cx = await zip0.file('xl/charts/chart1.xml')!.async('string');
        // Replace the rich title with a plain one-run title.
        cx = cx.replace(
            /<c:title>[\s\S]*?<c:autoTitleDeleted val="0"\/>/,
            `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Plain Title</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val="0"/>`,
        );
        zip0.file('xl/charts/chart1.xml', cx);
        const snap = await xlsxBufferToSnapshot(
            (await zip0.generateAsync({ type: 'nodebuffer' })) as unknown as Buffer,
        );
        const exported = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(exported);
        const cpath = Object.keys(zip.files).find((p) => /xl\/charts\/chart\d+\.xml$/.test(p))!;
        const xml = await zip.files[cpath].async('string');
        const titleBlock = xml.match(/<c:title>[\s\S]*?<\/c:title>/)![0];
        expect((titleBlock.match(/<a:r>/g) ?? []).length).toBe(1);
        expect(titleBlock).toContain('<a:t>Plain Title</a:t>');
    });
});
