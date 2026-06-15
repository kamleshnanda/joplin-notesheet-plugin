// M18 A1 stabilization: parser-robustness for image import OOXML shapes that
// the checked-in fixtures don't exercise. Each test builds a minimal .xlsx zip
// (the parts readImagesFromXlsxZip actually reads) and asserts the parser
// surfaces every image. Written failing-first against the 5 review findings:
//   - multiple <xdr:pic> in one anchor (grouped images)
//   - absoluteAnchor images
//   - non-xdr: namespace prefix
//   - multiple drawing parts referenced by one sheet
//   - imageId/drawingId uniqueness when two anchors share one media part
import JSZip from 'jszip';
import { readImagesFromXlsxZip } from '../src/drawings/xlsxImageImport';

const PNG = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64',
);

// Build a one-sheet workbook zip with the given drawing XML + rels mapping
// rIds → media filenames. `sheetRelsExtra` lets a test add a second drawing.
async function buildZip(opts: {
    drawings: Record<string, string>; // drawingN.xml path -> xml
    drawingRels: Record<string, Array<{ id: string; target: string }>>; // drawingN.xml -> rels
    sheetDrawingRels: Array<{ id: string; target: string }>; // sheet1.xml.rels drawing rels
    media: Record<string, Buffer>; // xl/media/imageN.ext -> bytes
}): Promise<Buffer> {
    const zip = new JSZip();
    zip.file(
        'xl/workbook.xml',
        `<?xml version="1.0"?><workbook xmlns:r="http://x"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    );
    zip.file(
        'xl/_rels/workbook.xml.rels',
        `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    );
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet/>`);
    const sheetRelXml = opts.sheetDrawingRels
        .map(
            (r) =>
                `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="${r.target}"/>`,
        )
        .join('');
    zip.file(
        'xl/worksheets/_rels/sheet1.xml.rels',
        `<?xml version="1.0"?><Relationships>${sheetRelXml}</Relationships>`,
    );
    for (const [path, xml] of Object.entries(opts.drawings)) zip.file(path, xml);
    for (const [drawingPath, rels] of Object.entries(opts.drawingRels)) {
        const folder = drawingPath.split('/').slice(0, -1).join('/');
        const file = drawingPath.split('/').pop();
        const relXml = rels
            .map(
                (r) =>
                    `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${r.target}"/>`,
            )
            .join('');
        zip.file(
            `${folder}/_rels/${file}.rels`,
            `<?xml version="1.0"?><Relationships>${relXml}</Relationships>`,
        );
    }
    for (const [path, bytes] of Object.entries(opts.media)) zip.file(path, bytes);
    return zip.generateAsync({ type: 'nodebuffer' });
}

const picXml = (rEmbed: string) =>
    `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip r:embed="${rEmbed}"/></xdr:blipFill><xdr:spPr/></xdr:pic>`;

const fromTo =
    `<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;

describe('M18 A1 — image import parser robustness', () => {
    test('two <xdr:pic> in ONE anchor → both images imported', async () => {
        const drawing =
            `<xdr:wsDr xmlns:xdr="x" xmlns:a="a" xmlns:r="r">` +
            `<xdr:twoCellAnchor>${fromTo}${picXml('rId1')}${picXml('rId2')}</xdr:twoCellAnchor>` +
            `</xdr:wsDr>`;
        const buf = await buildZip({
            drawings: { 'xl/drawings/drawing1.xml': drawing },
            drawingRels: {
                'xl/drawings/drawing1.xml': [
                    { id: 'rId1', target: '../media/image1.png' },
                    { id: 'rId2', target: '../media/image2.png' },
                ],
            },
            sheetDrawingRels: [{ id: 'rId1', target: '../drawings/drawing1.xml' }],
            media: { 'xl/media/image1.png': PNG, 'xl/media/image2.png': PNG },
        });
        const images = await readImagesFromXlsxZip(buf);
        expect(images.length).toBe(2);
    });

    test('absoluteAnchor image → imported (from-cell synthesized)', async () => {
        const drawing =
            `<xdr:wsDr xmlns:xdr="x" xmlns:a="a" xmlns:r="r">` +
            `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>${picXml('rId1')}</xdr:absoluteAnchor>` +
            `</xdr:wsDr>`;
        const buf = await buildZip({
            drawings: { 'xl/drawings/drawing1.xml': drawing },
            drawingRels: {
                'xl/drawings/drawing1.xml': [{ id: 'rId1', target: '../media/image1.png' }],
            },
            sheetDrawingRels: [{ id: 'rId1', target: '../drawings/drawing1.xml' }],
            media: { 'xl/media/image1.png': PNG },
        });
        const images = await readImagesFromXlsxZip(buf);
        expect(images.length).toBe(1);
        expect(images[0].ext).toEqual({ width: 100, height: 100 }); // 952500/9525
    });

    test('non-xdr namespace prefix → still imported', async () => {
        // A producer using prefix `a1:` for the drawing-ml elements.
        const drawing =
            `<a1:wsDr xmlns:a1="x" xmlns:a="a" xmlns:r="r"><a1:twoCellAnchor>` +
            `<a1:from><a1:col>1</a1:col><a1:colOff>0</a1:colOff><a1:row>1</a1:row><a1:rowOff>0</a1:rowOff></a1:from>` +
            `<a1:to><a1:col>4</a1:col><a1:colOff>0</a1:colOff><a1:row>10</a1:row><a1:rowOff>0</a1:rowOff></a1:to>` +
            `<a1:pic><a1:blipFill><a:blip r:embed="rId1"/></a1:blipFill></a1:pic>` +
            `</a1:twoCellAnchor></a1:wsDr>`;
        const buf = await buildZip({
            drawings: { 'xl/drawings/drawing1.xml': drawing },
            drawingRels: {
                'xl/drawings/drawing1.xml': [{ id: 'rId1', target: '../media/image1.png' }],
            },
            sheetDrawingRels: [{ id: 'rId1', target: '../drawings/drawing1.xml' }],
            media: { 'xl/media/image1.png': PNG },
        });
        const images = await readImagesFromXlsxZip(buf);
        expect(images.length).toBe(1);
    });

    test('sheet referencing TWO drawing parts → both images imported', async () => {
        const mk = (rEmbed: string) =>
            `<xdr:wsDr xmlns:xdr="x" xmlns:a="a" xmlns:r="r"><xdr:twoCellAnchor>${fromTo}${picXml(rEmbed)}</xdr:twoCellAnchor></xdr:wsDr>`;
        const buf = await buildZip({
            drawings: {
                'xl/drawings/drawing1.xml': mk('rId1'),
                'xl/drawings/drawing2.xml': mk('rId1'),
            },
            drawingRels: {
                'xl/drawings/drawing1.xml': [{ id: 'rId1', target: '../media/image1.png' }],
                'xl/drawings/drawing2.xml': [{ id: 'rId1', target: '../media/image2.png' }],
            },
            sheetDrawingRels: [
                { id: 'rId1', target: '../drawings/drawing1.xml' },
                { id: 'rId2', target: '../drawings/drawing2.xml' },
            ],
            media: { 'xl/media/image1.png': PNG, 'xl/media/image2.png': PNG },
        });
        const images = await readImagesFromXlsxZip(buf);
        expect(images.length).toBe(2);
    });

    test('two anchors sharing one media part → unique imageIds (no overwrite)', async () => {
        const drawing =
            `<xdr:wsDr xmlns:xdr="x" xmlns:a="a" xmlns:r="r">` +
            `<xdr:twoCellAnchor>${fromTo}${picXml('rId1')}</xdr:twoCellAnchor>` +
            `<xdr:twoCellAnchor>${fromTo}${picXml('rId1')}</xdr:twoCellAnchor>` +
            `</xdr:wsDr>`;
        const buf = await buildZip({
            drawings: { 'xl/drawings/drawing1.xml': drawing },
            drawingRels: {
                'xl/drawings/drawing1.xml': [{ id: 'rId1', target: '../media/image1.png' }],
            },
            sheetDrawingRels: [{ id: 'rId1', target: '../drawings/drawing1.xml' }],
            media: { 'xl/media/image1.png': PNG },
        });
        const images = await readImagesFromXlsxZip(buf);
        expect(images.length).toBe(2);
        // The two must be distinguishable so the snapshot drawingId doesn't collide.
        const ids = new Set(images.map((i) => `${i.sheetIndex}-${i.imageId}`));
        expect(ids.size).toBe(2);
    });
});
