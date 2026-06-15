// M18 A1: image drawing export to .xlsx via post-processing the zip exceljs
// writes — the mirror of injectChartsIntoZip.
//
// Architecture (parallels src/charts/xlsxChart.ts):
//   snapshotToXlsxBuffer()  -> exceljs writes cells/styles/tables -> buffer
//                          -> injectChartsIntoZip(buffer, snapshot)
//                          -> injectImagesIntoZip(buffer, snapshot)   (here)
//                          -> JSZip surgery: add xl/media/* parts +
//                            xl/drawings <xdr:pic> anchors + rels +
//                            [Content_Types] Default-by-extension
//                          -> return new buffer
//
// CRITICAL coexistence requirement (the plan's #1 risk): when a sheet
// already has a drawing part (because injectChartsIntoZip ran first and added
// a chart drawing), images for that sheet MUST be MERGED into the existing
// drawingN.xml + its _rels — NOT emitted as a second drawing part. Excel only
// honors ONE <drawing r:id> per worksheet, so a second drawing part would be
// silently dropped (and the chart or the image would vanish). We therefore
// scan the worksheet's rels for an existing drawing target, reuse it when
// present, and only create a fresh drawing part when the sheet has none.
//
// exceljs writes ZERO media for our exported workbooks (we build a FRESH
// workbook from snapshot cellData; exceljs only preserves media on workbooks
// it LOADED with images). So we own the media parts entirely via zip
// injection; we never call ws.addImage.

import JSZip from 'jszip';

import type { UniverSnapshot } from '../snapshot';
import { escapeXml, maxExistingRId, upsertRelationship } from '../charts/xlsxChart';

// EMU per pixel at 96 DPI (mirrors the import path).
const EMU_PER_PX = 9525;

export type ImageExtension = 'png' | 'jpeg' | 'gif' | 'bmp' | 'webp';

// One image, normalized for emission. Read out of the snapshot's
// SHEET_DRAWING_PLUGIN resource, native Univer image-drawing entries only.
export interface ImageDrawing {
    drawingId: string;
    sheetId: string;
    extension: ImageExtension;
    // Raw image bytes (decoded from the base64 data: URI).
    bytes: Uint8Array;
    anchor: {
        fromCol: number;
        fromColOff: number; // EMU
        fromRow: number;
        fromRowOff: number; // EMU
        toCol: number;
        toColOff: number; // EMU
        toRow: number;
        toRowOff: number; // EMU
    };
}

// --- MIME / extension helpers ---

function mimeToExtension(mime: string): ImageExtension | null {
    switch (mime.toLowerCase()) {
        case 'image/png':
            return 'png';
        case 'image/jpeg':
        case 'image/jpg':
            return 'jpeg';
        case 'image/gif':
            return 'gif';
        case 'image/bmp':
            return 'bmp';
        case 'image/webp':
            return 'webp';
        default:
            return null;
    }
}

function extensionToContentType(ext: ImageExtension): string {
    switch (ext) {
        case 'png':
            return 'image/png';
        case 'jpeg':
            return 'image/jpeg';
        case 'gif':
            return 'image/gif';
        case 'bmp':
            return 'image/bmp';
        case 'webp':
            return 'image/webp';
    }
}

// Decode a data:<mime>;base64,<payload> URI into mime + raw bytes.
function decodeDataUri(source: string): { mime: string; bytes: Uint8Array } | null {
    const m = source.match(/^data:([^;,]+);base64,([\s\S]*)$/);
    if (!m) return null;
    const mime = m[1];
    try {
        const buf = Buffer.from(m[2], 'base64');
        return { mime, bytes: new Uint8Array(buf) };
    } catch {
        return null;
    }
}

// --- Read images from snapshot ---

// The SHEET_DRAWING_PLUGIN resource's data field is a JSON-stringified map:
//   { [subUnitId]: { data: { [drawingId]: ISheetDrawing | ISheetImage }, order } }
// We filter to NATIVE image drawings: drawingType === 0, an imageSourceType,
// a base64 data: URI source, and NO 'NotesheetChart' componentKey (charts use
// drawingType 8 + componentKey; they must be left to injectChartsIntoZip).
export function readImagesFromSnapshot(snapshot: UniverSnapshot): ImageDrawing[] {
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> })
        .resources;
    if (!Array.isArray(resources)) return [];
    const entry = resources.find((r) => r?.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry || typeof entry.data !== 'string') return [];

    let parsed: Record<string, { data?: Record<string, unknown> }>;
    try {
        parsed = JSON.parse(entry.data);
    } catch {
        return [];
    }
    if (!parsed || typeof parsed !== 'object') return [];

    const out: ImageDrawing[] = [];
    for (const subUnitId of Object.keys(parsed)) {
        const subUnit = parsed[subUnitId];
        const drawings = subUnit?.data;
        if (!drawings || typeof drawings !== 'object') continue;

        for (const drawingId of Object.keys(drawings)) {
            const d = drawings[drawingId] as {
                componentKey?: string;
                drawingType?: number;
                imageSourceType?: string;
                source?: string;
                sheetTransform?: {
                    from?: {
                        column?: number;
                        columnOffset?: number;
                        row?: number;
                        rowOffset?: number;
                    };
                    to?: {
                        column?: number;
                        columnOffset?: number;
                        row?: number;
                        rowOffset?: number;
                    };
                };
                axisAlignSheetTransform?: {
                    from?: {
                        column?: number;
                        columnOffset?: number;
                        row?: number;
                        rowOffset?: number;
                    };
                    to?: {
                        column?: number;
                        columnOffset?: number;
                        row?: number;
                        rowOffset?: number;
                    };
                };
            };

            // Charts are drawingType 8 + componentKey; never treat as image.
            if (d?.componentKey === 'NotesheetChart') continue;
            if (d?.drawingType !== 0) continue;
            if (!d.imageSourceType || typeof d.source !== 'string') continue;

            const decoded = decodeDataUri(d.source);
            if (!decoded) continue;
            const ext = mimeToExtension(decoded.mime);
            if (!ext) continue;

            const tx = d.axisAlignSheetTransform ?? d.sheetTransform;
            if (!tx?.from || !tx?.to) continue;

            // sheetTransform offsets are PIXELS in Univer; OOXML anchor offsets
            // are EMUs. Convert px -> EMU here (inverse of the import path).
            out.push({
                drawingId,
                sheetId: subUnitId,
                extension: ext,
                bytes: decoded.bytes,
                anchor: {
                    fromCol: tx.from.column ?? 0,
                    fromColOff: Math.round((tx.from.columnOffset ?? 0) * EMU_PER_PX),
                    fromRow: tx.from.row ?? 0,
                    fromRowOff: Math.round((tx.from.rowOffset ?? 0) * EMU_PER_PX),
                    toCol: tx.to.column ?? 0,
                    toColOff: Math.round((tx.to.columnOffset ?? 0) * EMU_PER_PX),
                    toRow: tx.to.row ?? 0,
                    toRowOff: Math.round((tx.to.rowOffset ?? 0) * EMU_PER_PX),
                },
            });
        }
    }
    return out;
}

// --- <xdr:pic> builder ---

// Build a single <xdr:twoCellAnchor editAs="oneCell"> ... <xdr:pic> block.
// `picId` is the cNvPr id (must be unique within the drawing); `rId` is the
// drawing-rels relationship that points at the media part; `name` is a
// human-readable picture name.
export function buildImagePicXml(
    image: ImageDrawing,
    picId: number,
    rId: number,
    name: string,
): string {
    const a = image.anchor;
    return (
        `<xdr:twoCellAnchor editAs="oneCell">` +
        `<xdr:from><xdr:col>${a.fromCol}</xdr:col><xdr:colOff>${a.fromColOff}</xdr:colOff><xdr:row>${a.fromRow}</xdr:row><xdr:rowOff>${a.fromRowOff}</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>${a.toCol}</xdr:col><xdr:colOff>${a.toColOff}</xdr:colOff><xdr:row>${a.toRow}</xdr:row><xdr:rowOff>${a.toRowOff}</xdr:rowOff></xdr:to>` +
        `<xdr:pic>` +
        `<xdr:nvPicPr>` +
        `<xdr:cNvPr id="${picId}" name="${escapeXml(name)}"/>` +
        `<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>` +
        `</xdr:nvPicPr>` +
        `<xdr:blipFill>` +
        `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${rId}"/>` +
        `<a:stretch><a:fillRect/></a:stretch>` +
        `</xdr:blipFill>` +
        `<xdr:spPr>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `</xdr:spPr>` +
        `</xdr:pic>` +
        `<xdr:clientData/>` +
        `</xdr:twoCellAnchor>`
    );
}

// Wrap a set of anchor blocks in a fresh wsDr document.
function buildDrawingDoc(anchorsXml: string): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchorsXml}</xdr:wsDr>`;
}

// --- zip helpers ---

// Resolve subUnitId -> 1-based worksheet index. Mirrors the chart export
// path's lookupSheet (kept local to avoid coupling).
function lookupSheet(snapshot: UniverSnapshot, sheetId: string): { index: number } | null {
    const sheetOrder = (snapshot as { sheetOrder?: string[] }).sheetOrder ?? [];
    const idx0 = sheetOrder.indexOf(sheetId);
    if (idx0 < 0) return null;
    return { index: idx0 + 1 };
}

// Scan existing xl/media/imageN.<ext> parts for the highest N, so new media
// parts don't collide with any exceljs OR chart-injected media already there.
function maxExistingMediaNum(zip: JSZip): number {
    let max = 0;
    for (const p of Object.keys(zip.files)) {
        const m = p.match(/^xl\/media\/image(\d+)\.[A-Za-z0-9]+$/);
        if (m) {
            const n = parseInt(m[1], 10);
            if (n > max) max = n;
        }
    }
    return max;
}

// Highest cNvPr id used by graphicFrame/pic elements in a drawing doc, so a
// merged <xdr:pic> gets a non-colliding id.
function maxExistingPicId(drawingXml: string): number {
    let max = 1; // Excel reserves id=1 for the drawing root conceptually
    const re = /<xdr:cNvPr\s+id="(\d+)"/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(drawingXml)) !== null) {
        const n = parseInt(m[1], 10);
        if (n > max) max = n;
    }
    return max;
}

// Find the existing drawing target (e.g. "drawing1.xml") referenced by a
// worksheet's rels, if any. Returns the absolute zip key of the drawing part.
function findExistingDrawingForSheet(relsXml: string | null): string | null {
    if (!relsXml) return null;
    // Type ends in /drawing; Target is a relative path like ../drawings/drawingN.xml.
    const re =
        /<Relationship\b[^>]*Type="[^"]*\/drawing"[^>]*Target="([^"]+)"|<Relationship\b[^>]*Target="([^"]+)"[^>]*Type="[^"]*\/drawing"/;
    const m = relsXml.match(re);
    const target = m ? (m[1] ?? m[2]) : null;
    if (!target) return null;
    // Targets are relative to xl/worksheets; normalize "../drawings/x.xml".
    const cleaned = target.replace(/^\.\.\//, '').replace(/^\//, '');
    return cleaned.startsWith('xl/') ? cleaned : `xl/${cleaned}`;
}

// Insert <drawing r:id="..."/> into worksheet XML at its schema-correct slot
// (after pageSetup/headerFooter, before tableParts). Mirrors the chart path.
function insertDrawingRef(sheetXml: string, rId: number): string {
    if (sheetXml.includes(`<drawing r:id="rId${rId}"/>`)) return sheetXml;
    const drawingTag = `<drawing r:id="rId${rId}"/>`;
    if (sheetXml.includes('<tableParts')) {
        return sheetXml.replace(/<tableParts\b/, `${drawingTag}<tableParts`);
    }
    return sheetXml.replace(/<\/worksheet>\s*$/, `${drawingTag}</worksheet>`);
}

// Patch [Content_Types].xml with Default-by-extension entries for each image
// extension used. Images use <Default Extension=.../> (idempotent — Excel
// dedupes by extension), NOT per-part <Override> like charts.
function patchContentTypesForImages(
    contentTypesXml: string,
    extensions: Set<ImageExtension>,
): string {
    const toAdd: string[] = [];
    for (const ext of extensions) {
        // Skip if a Default for this extension already exists.
        const re = new RegExp(`<Default\\s+Extension="${ext}"`, 'i');
        if (re.test(contentTypesXml)) continue;
        toAdd.push(`<Default Extension="${ext}" ContentType="${extensionToContentType(ext)}"/>`);
    }
    if (toAdd.length === 0) return contentTypesXml;
    // Defaults conventionally precede Overrides; inserting right after <Types ...>
    // keeps that ordering and is valid regardless.
    return contentTypesXml.replace(/(<Types\b[^>]*>)/, `$1${toAdd.join('')}`);
}

// --- Public entry point ---

// Post-process the xlsx buffer (already chart-injected) to add image parts
// for every native image drawing in the snapshot. Runs AFTER
// injectChartsIntoZip and coexists with it: images on a sheet that already
// has a chart drawing are merged into that drawing rather than creating a
// second one. Fail-soft: returns the input buffer on any throw.
export async function injectImagesIntoZip(
    buffer: ArrayBuffer,
    snapshot: UniverSnapshot,
): Promise<ArrayBuffer> {
    let images: ImageDrawing[];
    try {
        images = readImagesFromSnapshot(snapshot);
    } catch (e) {
        console.warn('[Notesheet] M18: readImagesFromSnapshot threw; skipping image export', e);
        return buffer;
    }
    if (images.length === 0) return buffer;

    try {
        const zip = await JSZip.loadAsync(buffer);

        // Group images by sheet — one drawing part per sheet (merged with any
        // existing chart drawing on that sheet).
        const bySheet = new Map<string, ImageDrawing[]>();
        for (const img of images) {
            const arr = bySheet.get(img.sheetId);
            if (arr) arr.push(img);
            else bySheet.set(img.sheetId, [img]);
        }

        let mediaCounter = maxExistingMediaNum(zip);
        // drawing{N} numbering for NEW drawing parts: continue past any
        // existing drawingN.xml in the zip (chart-injected or otherwise).
        let drawingCounter = 0;
        for (const p of Object.keys(zip.files)) {
            const m = p.match(/^xl\/drawings\/drawing(\d+)\.xml$/);
            if (m) {
                const n = parseInt(m[1], 10);
                if (n > drawingCounter) drawingCounter = n;
            }
        }

        const extensionsUsed = new Set<ImageExtension>();

        for (const [sheetId, sheetImages] of bySheet) {
            const sheet = lookupSheet(snapshot, sheetId);
            if (!sheet) continue;

            const sheetRelsPath = `xl/worksheets/_rels/sheet${sheet.index}.xml.rels`;
            const sheetXmlPath = `xl/worksheets/sheet${sheet.index}.xml`;
            const existingRelsFile = zip.file(sheetRelsPath);
            const existingRelsXml = existingRelsFile
                ? await existingRelsFile.async('string')
                : null;

            // Does this sheet already have a drawing part (chart-injected)?
            const existingDrawingPath = findExistingDrawingForSheet(existingRelsXml);

            if (existingDrawingPath && zip.file(existingDrawingPath)) {
                // MERGE path: add <xdr:pic> anchors + image rels into the
                // existing drawing part (the coexistence case).
                const drawingFile = zip.file(existingDrawingPath)!;
                const drawingXml = await drawingFile.async('string');
                const drawingNum = (() => {
                    const m = existingDrawingPath.match(/drawing(\d+)\.xml$/);
                    return m ? parseInt(m[1], 10) : drawingCounter;
                })();
                const drawingRelsPath = `xl/drawings/_rels/drawing${drawingNum}.xml.rels`;
                const drawingRelsFile = zip.file(drawingRelsPath);
                let drawingRelsXml = drawingRelsFile ? await drawingRelsFile.async('string') : null;

                let picId = maxExistingPicId(drawingXml);
                let newAnchors = '';
                for (let i = 0; i < sheetImages.length; i++) {
                    const img = sheetImages[i];
                    extensionsUsed.add(img.extension);
                    mediaCounter += 1;
                    const mediaName = `image${mediaCounter}.${img.extension}`;
                    zip.file(`xl/media/${mediaName}`, img.bytes);

                    const newRId = maxExistingRId(drawingRelsXml) + 1;
                    const rel = `<Relationship Id="rId${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`;
                    drawingRelsXml = upsertRelationship(drawingRelsXml, rel);

                    picId += 1;
                    newAnchors += buildImagePicXml(img, picId, newRId, `Picture ${picId}`);
                }
                const mergedDrawing = drawingXml.replace(
                    /<\/xdr:wsDr>\s*$/,
                    `${newAnchors}</xdr:wsDr>`,
                );
                zip.file(existingDrawingPath, mergedDrawing);
                if (drawingRelsXml) zip.file(drawingRelsPath, drawingRelsXml);
            } else {
                // NEW path: create a fresh drawing part + sheet rel +
                // <drawing r:id> + drawing rels.
                const drawingNum = ++drawingCounter;
                let drawingRelsXml: string | null = null;
                let picId = 1;
                let anchorsXml = '';
                for (let i = 0; i < sheetImages.length; i++) {
                    const img = sheetImages[i];
                    extensionsUsed.add(img.extension);
                    mediaCounter += 1;
                    const mediaName = `image${mediaCounter}.${img.extension}`;
                    zip.file(`xl/media/${mediaName}`, img.bytes);

                    const newRId = maxExistingRId(drawingRelsXml) + 1;
                    const rel = `<Relationship Id="rId${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`;
                    drawingRelsXml = upsertRelationship(drawingRelsXml, rel);

                    picId += 1;
                    anchorsXml += buildImagePicXml(img, picId, newRId, `Picture ${picId}`);
                }
                zip.file(`xl/drawings/drawing${drawingNum}.xml`, buildDrawingDoc(anchorsXml));
                if (drawingRelsXml) {
                    zip.file(`xl/drawings/_rels/drawing${drawingNum}.xml.rels`, drawingRelsXml);
                }

                // Sheet rel -> the new drawing part.
                const newRId = maxExistingRId(existingRelsXml) + 1;
                const sheetRel = `<Relationship Id="rId${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingNum}.xml"/>`;
                zip.file(sheetRelsPath, upsertRelationship(existingRelsXml, sheetRel));

                // Worksheet xml -> <drawing r:id="...">.
                const sheetXmlFile = zip.file(sheetXmlPath);
                if (sheetXmlFile) {
                    const sheetXml = await sheetXmlFile.async('string');
                    zip.file(sheetXmlPath, insertDrawingRef(sheetXml, newRId));
                }

                // Patch [Content_Types] with a drawing Override for the new part.
                const ctFileEarly = zip.file('[Content_Types].xml');
                if (ctFileEarly) {
                    const ctXml = await ctFileEarly.async('string');
                    if (!ctXml.includes(`/xl/drawings/drawing${drawingNum}.xml`)) {
                        const drawOverride = `<Override PartName="/xl/drawings/drawing${drawingNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
                        zip.file(
                            '[Content_Types].xml',
                            ctXml.replace(/<\/Types>\s*$/, `${drawOverride}</Types>`),
                        );
                    }
                }
            }
        }

        // Patch [Content_Types].xml with Default-by-extension for every image
        // extension used (idempotent — Excel dedupes by extension).
        const ctFile = zip.file('[Content_Types].xml');
        if (ctFile) {
            const ctXml = await ctFile.async('string');
            zip.file('[Content_Types].xml', patchContentTypesForImages(ctXml, extensionsUsed));
        }

        const out = await zip.generateAsync({ type: 'arraybuffer' });
        return out as ArrayBuffer;
    } catch (e) {
        console.error(
            '[Notesheet] M18: injectImagesIntoZip failed; falling back to image-less xlsx',
            e,
        );
        return buffer;
    }
}
