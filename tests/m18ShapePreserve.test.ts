// M18 A2: shape drawings (non-chart, non-image <xdr:sp>) are PRESERVE-ONLY.
// Univer 0.23 has no working native shape renderer (DRAWING_SHAPE is a stub:
// sheets-drawing-ui mounts only DRAWING_IMAGE + DRAWING_DOM), so we do NOT
// render shapes in the editor or in Joplin HTML/PDF export. Instead we
// round-trip them through .xlsx losslessly: import stashes each <xdr:sp>
// anchor (verbatim) into a passthrough resource keyed by sheet, the editor
// resource hook carries it through save/reload untouched, and export injects
// the anchors back into the worksheet's drawing part so Excel renders them.
//
// Tests written failing-first against:
//   - readShapesFromXlsxZip: extracts shape anchors verbatim, per sheet
//   - the snapshot carries a SHEET_NOTESHEET_SHAPES_PLUGIN resource
//   - export injects the shapes back; re-import recovers them byte-for-byte
//   - shapes coexist with an image on the same sheet (one drawing part)

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import JSZip from 'jszip';
import { readShapesFromXlsxZip } from '../src/drawings/xlsxShapeImport';
import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

// A standalone shape anchor (rect text box). No <xdr:pic>, no chart.
function shapeAnchor(name: string, fill: string): string {
    return (
        `<xdr:twoCellAnchor>` +
        `<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
        `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="${name}"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
        `<xdr:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="900" cy="500"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></xdr:spPr>` +
        `<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>${name}</a:t></a:r></a:p></xdr:txBody></xdr:sp>` +
        `<xdr:clientData/></xdr:twoCellAnchor>`
    );
}

// Build a minimal valid .xlsx (loadable by exceljs) with the given drawing
// body (the inner content of <xdr:wsDr>) on sheet1.
function buildXlsx(drawingBody: string): Promise<Buffer> {
    const zip = new JSZip();
    zip.file(
        '[Content_Types].xml',
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`,
    );
    zip.file(
        '_rels/.rels',
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    );
    zip.file(
        'xl/workbook.xml',
        `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    );
    zip.file(
        'xl/_rels/workbook.xml.rels',
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    );
    zip.file(
        'xl/worksheets/sheet1.xml',
        `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>`,
    );
    zip.file(
        'xl/worksheets/_rels/sheet1.xml.rels',
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    );
    zip.file(
        'xl/drawings/drawing1.xml',
        `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}">${drawingBody}</xdr:wsDr>`,
    );
    return zip.generateAsync({ type: 'nodebuffer' });
}

function shapeResource(snap: unknown): Record<string, string[]> | null {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_NOTESHEET_SHAPES_PLUGIN');
    return entry ? JSON.parse(entry.data) : null;
}

describe('M18 A2 — shape preserve-only round-trip', () => {
    test('readShapesFromXlsxZip extracts the shape anchor verbatim, keyed by sheet', async () => {
        const buf = await buildXlsx(shapeAnchor('Box1', 'FFFF00'));
        const shapes = await readShapesFromXlsxZip(buf);
        expect(shapes.length).toBe(1);
        expect(shapes[0].sheetIndex).toBe(1);
        // The stashed anchor XML must contain the shape verbatim.
        expect(shapes[0].anchorXml).toContain('<xdr:sp');
        expect(shapes[0].anchorXml).toContain('Box1');
        expect(shapes[0].anchorXml).toContain('FFFF00');
    });

    test('a pic-only drawing yields NO shapes (shapes != images)', async () => {
        // An <xdr:pic> with a <xdr:spPr> inside must NOT be misread as a shape.
        const picBody =
            `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
            `<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
            `<xdr:pic><xdr:blipFill><a:blip r:embed="rId9"/></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
        const buf = await buildXlsx(picBody);
        const shapes = await readShapesFromXlsxZip(buf);
        expect(shapes.length).toBe(0);
    });

    test('import surfaces a SHEET_NOTESHEET_SHAPES_PLUGIN resource', async () => {
        const buf = await buildXlsx(shapeAnchor('Box1', 'FFFF00'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const res = shapeResource(snap);
        expect(res).not.toBeNull();
        // Keyed by subUnitId (sheet-1); value is an array of anchor XML strings.
        expect(res!['sheet-1']).toBeDefined();
        expect(res!['sheet-1'].length).toBe(1);
        expect(res!['sheet-1'][0]).toContain('Box1');
    });

    test('shape survives import -> export -> re-import (bytes preserved)', async () => {
        const buf = await buildXlsx(shapeAnchor('Box1', 'FFFF00'));
        const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap1);

        // Exported zip carries the <xdr:sp> in a drawing part.
        const zip = await JSZip.loadAsync(exported);
        const drawings = Object.keys(zip.files).filter((p) =>
            /^xl\/drawings\/drawing\d+\.xml$/.test(p),
        );
        expect(drawings.length).toBeGreaterThanOrEqual(1);
        let foundShape = false;
        for (const d of drawings) {
            const xml = await zip.files[d].async('string');
            if (xml.includes('<xdr:sp') && xml.includes('Box1')) foundShape = true;
        }
        expect(foundShape).toBe(true);

        // Re-import recovers the shape.
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(exported) as unknown as Buffer);
        const res2 = shapeResource(snap2);
        expect(res2?.['sheet-1']?.length).toBe(1);
        expect(res2!['sheet-1'][0]).toContain('FFFF00');
    });

    test('two shapes on one sheet both round-trip', async () => {
        const buf = await buildXlsx(shapeAnchor('Box1', 'FFFF00') + shapeAnchor('Box2', '00FF00'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const res = shapeResource(snap);
        expect(res?.['sheet-1']?.length).toBe(2);
        const exported = await snapshotToXlsxBuffer(snap);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(exported) as unknown as Buffer);
        expect(shapeResource(snap2)?.['sheet-1']?.length).toBe(2);
    });

    test('shape anchor content is preserved faithfully (geometry, fill, text)', async () => {
        // Preserve-only is worthless if it silently corrupts the shape, so
        // assert the SUBSTANTIVE content survives import (the cNvPr id may be
        // renumbered on export-merge; everything else must round-trip).
        const buf = await buildXlsx(shapeAnchor('Callout', 'A1B2C3'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const anchor = shapeResource(snap)!['sheet-1'][0];
        // Geometry, fill, preset, and text all present verbatim.
        expect(anchor).toContain('<a:ext cx="900" cy="500"/>');
        expect(anchor).toContain('<a:prstGeom prst="rect">');
        expect(anchor).toContain('<a:srgbClr val="A1B2C3"/>');
        expect(anchor).toContain('<a:t>Callout</a:t>');
        // The from/to anchor cells survive.
        expect(anchor).toMatch(/<xdr:from>[\s\S]*<xdr:col>1<\/xdr:col>[\s\S]*<\/xdr:from>/);
    });

    test('an image-only fixture surfaces NO shapes (no false positives)', async () => {
        // Regression guard: every <xdr:pic> carries an <xdr:spPr>; the shape
        // reader must not mistake that for a standalone <xdr:sp>.
        const { readFileSync } = await import('fs');
        const pathMod = await import('path');
        const imgBuf = readFileSync(
            pathMod.join(__dirname, 'fixtures', 'images', 'HumanImage-SingleSheet.xlsx'),
        );
        const snap = await xlsxBufferToSnapshot(imgBuf as unknown as Buffer);
        // No shape resource emitted for a pure-image workbook.
        expect(shapeResource(snap)).toBeNull();
    });

    test('shape coexists with an image on the same sheet (one drawing part)', async () => {
        // Build a drawing with BOTH an <xdr:pic> (image) and a standalone
        // <xdr:sp> (shape) so the image importer (A1) and shape importer (A2)
        // both run on the same sheet. After export, both must end up in ONE
        // drawing part (Excel honors one <drawing> per sheet) and re-import
        // must recover both independently.
        const PNG = Buffer.from(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
            'base64',
        );
        const zip = new JSZip();
        zip.file(
            '[Content_Types].xml',
            `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`,
        );
        zip.file(
            '_rels/.rels',
            `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
        );
        zip.file(
            'xl/workbook.xml',
            `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        );
        zip.file(
            'xl/_rels/workbook.xml.rels',
            `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
        );
        zip.file(
            'xl/worksheets/sheet1.xml',
            `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>`,
        );
        zip.file(
            'xl/worksheets/_rels/sheet1.xml.rels',
            `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
        );
        const picAnchor =
            `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
            `<xdr:ext cx="952500" cy="952500"/>` +
            `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Pic"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
        zip.file(
            'xl/drawings/drawing1.xml',
            `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${picAnchor}${shapeAnchor('Box1', 'FFFF00')}</xdr:wsDr>`,
        );
        zip.file(
            'xl/drawings/_rels/drawing1.xml.rels',
            `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
        );
        zip.file('xl/media/image1.png', PNG);
        const buf = await zip.generateAsync({ type: 'nodebuffer' });

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        // Image went to SHEET_DRAWING_PLUGIN; shape went to its own resource.
        expect(shapeResource(snap)?.['sheet-1']?.length).toBe(1);

        const exported = await snapshotToXlsxBuffer(snap);
        const outZip = await JSZip.loadAsync(exported);
        const drawings = Object.keys(outZip.files).filter((p) =>
            /^xl\/drawings\/drawing\d+\.xml$/.test(p),
        );
        // Exactly ONE drawing part on the sheet, carrying BOTH pic and sp.
        let picCount = 0;
        let spCount = 0;
        for (const d of drawings) {
            const xml = await outZip.files[d].async('string');
            picCount += (xml.match(/<xdr:pic\b/g) ?? []).length;
            spCount += (xml.match(/<xdr:sp\b/g) ?? []).length;
        }
        expect(picCount).toBe(1);
        expect(spCount).toBe(1);

        // Re-import recovers both.
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(exported) as unknown as Buffer);
        expect(shapeResource(snap2)?.['sheet-1']?.length).toBe(1);
    });

    test('real fixture ShapeTextBox-SingleSheet.xlsx round-trips', async () => {
        // Regression guard against a real .xlsx file (not synthetic inline
        // XML) — the same fixture used for the live passthrough-survival
        // verification (screenshots/m18-a2/).
        const { readFileSync } = await import('fs');
        const pathMod = await import('path');
        const buf = readFileSync(
            pathMod.join(__dirname, 'fixtures', 'shapes', 'ShapeTextBox-SingleSheet.xlsx'),
        );
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const anchor = shapeResource(snap)?.['sheet-1']?.[0] ?? '';
        expect(anchor).toContain('Preserved Shape');
        expect(anchor).toContain('roundRect');
        const exported = await snapshotToXlsxBuffer(snap);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(exported) as unknown as Buffer);
        expect(shapeResource(snap2)?.['sheet-1']?.[0]).toContain('Preserved Shape');
    });
});
