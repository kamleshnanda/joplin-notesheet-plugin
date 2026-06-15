// M18 A2: shape drawing export to .xlsx — PRESERVE-ONLY, the mirror of
// injectImagesIntoZip but simpler.
//
// Shapes were captured verbatim on import (full <xdr:*Anchor> XML) into the
// SHEET_NOTESHEET_SHAPES_PLUGIN resource. Here we splice those anchors back
// into each worksheet's drawing part. Unlike images, a shape anchor is
// SELF-CONTAINED: no media parts, no drawing rels, no [Content_Types] Default
// — the <xdr:sp> carries its own geometry/fill/text inline. So we only need
// to (a) ensure the sheet has a drawing part, and (b) append the anchors into
// it before </xdr:wsDr>.
//
// Runs AFTER injectChartsIntoZip + injectImagesIntoZip. Coexistence is the
// same one-drawing-per-sheet rule: if the sheet already has a drawing part
// (chart- or image-injected), MERGE the shape anchors into it; only create a
// fresh drawing part when the sheet has none. Fail-soft: returns the input
// buffer on any throw.

import JSZip from 'jszip';

import type { UniverSnapshot } from '../snapshot';
import { maxExistingRId, upsertRelationship } from '../charts/xlsxChart';
import { NOTESHEET_SHAPES_RESOURCE } from './sheetIdResolver';

interface ShapeGroup {
    sheetId: string;
    anchors: string[];
}

// Pull preserve-only shape anchors out of the snapshot resource.
function readShapesFromSnapshot(snapshot: UniverSnapshot): ShapeGroup[] {
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> })
        .resources;
    if (!Array.isArray(resources)) return [];
    const entry = resources.find((r) => r && r.name === NOTESHEET_SHAPES_RESOURCE);
    if (!entry || typeof entry.data !== 'string') return [];
    let parsed: Record<string, unknown>;
    try {
        parsed = JSON.parse(entry.data);
    } catch {
        return [];
    }
    const out: ShapeGroup[] = [];
    for (const sheetId of Object.keys(parsed)) {
        const anchors = parsed[sheetId];
        if (Array.isArray(anchors) && anchors.length > 0) {
            out.push({
                sheetId,
                anchors: anchors.filter((a) => typeof a === 'string') as string[],
            });
        }
    }
    return out;
}

function buildDrawingDoc(anchorsXml: string): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchorsXml}</xdr:wsDr>`;
}

function lookupSheetIndex(snapshot: UniverSnapshot, sheetId: string): number | null {
    const sheetOrder = (snapshot as { sheetOrder?: string[] }).sheetOrder ?? [];
    const idx0 = sheetOrder.indexOf(sheetId);
    return idx0 < 0 ? null : idx0 + 1;
}

function findExistingDrawingForSheet(relsXml: string | null): string | null {
    if (!relsXml) return null;
    const re =
        /<Relationship\b[^>]*Type="[^"]*\/drawing"[^>]*Target="([^"]+)"|<Relationship\b[^>]*Target="([^"]+)"[^>]*Type="[^"]*\/drawing"/;
    const m = relsXml.match(re);
    const target = m ? (m[1] ?? m[2]) : null;
    if (!target) return null;
    const cleaned = target.replace(/^\.\.\//, '').replace(/^\//, '');
    return cleaned.startsWith('xl/') ? cleaned : `xl/${cleaned}`;
}

function insertDrawingRef(sheetXml: string, rId: number): string {
    if (sheetXml.includes(`<drawing r:id="rId${rId}"/>`)) return sheetXml;
    const drawingTag = `<drawing r:id="rId${rId}"/>`;
    if (sheetXml.includes('<tableParts')) {
        return sheetXml.replace(/<tableParts\b/, `${drawingTag}<tableParts`);
    }
    return sheetXml.replace(/<\/worksheet>\s*$/, `${drawingTag}</worksheet>`);
}

// Highest <xdr:cNvPr id="N"> in a drawing doc, so merged anchors get
// non-colliding ids.
function maxExistingCnvPrId(drawingXml: string): number {
    let max = 1;
    const re = /<xdr:cNvPr\s+id="(\d+)"/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(drawingXml)) !== null) {
        const n = parseInt(m[1], 10);
        if (n > max) max = n;
    }
    return max;
}

// Renumber every <xdr:cNvPr id="N"> in the given anchors starting past
// `startId`, so they don't collide with ids already in the drawing part.
function renumberCnvPrIds(anchorsXml: string, startId: number): string {
    let id = startId;
    return anchorsXml.replace(/(<xdr:cNvPr\s+id=")(\d+)(")/g, (_full, pre, _n, post) => {
        id += 1;
        return `${pre}${id}${post}`;
    });
}

export async function injectShapesIntoZip(
    buffer: ArrayBuffer,
    snapshot: UniverSnapshot,
): Promise<ArrayBuffer> {
    let groups: ShapeGroup[];
    try {
        groups = readShapesFromSnapshot(snapshot);
    } catch (e) {
        console.warn('[Notesheet] M18 A2: readShapesFromSnapshot threw; skipping shape export', e);
        return buffer;
    }
    if (groups.length === 0) return buffer;

    try {
        const zip = await JSZip.loadAsync(buffer);

        let drawingCounter = 0;
        for (const p of Object.keys(zip.files)) {
            const m = p.match(/^xl\/drawings\/drawing(\d+)\.xml$/);
            if (m) {
                const n = parseInt(m[1], 10);
                if (n > drawingCounter) drawingCounter = n;
            }
        }

        for (const group of groups) {
            const index = lookupSheetIndex(snapshot, group.sheetId);
            if (index === null) continue;

            const sheetRelsPath = `xl/worksheets/_rels/sheet${index}.xml.rels`;
            const sheetXmlPath = `xl/worksheets/sheet${index}.xml`;
            const existingRelsFile = zip.file(sheetRelsPath);
            const existingRelsXml = existingRelsFile
                ? await existingRelsFile.async('string')
                : null;
            const existingDrawingPath = findExistingDrawingForSheet(existingRelsXml);

            if (existingDrawingPath && zip.file(existingDrawingPath)) {
                // MERGE: append shape anchors into the existing drawing part,
                // renumbering cNvPr ids past whatever's already there.
                const drawingFile = zip.file(existingDrawingPath)!;
                const drawingXml = await drawingFile.async('string');
                const renumbered = renumberCnvPrIds(
                    group.anchors.join(''),
                    maxExistingCnvPrId(drawingXml),
                );
                const merged = drawingXml.replace(/<\/xdr:wsDr>\s*$/, `${renumbered}</xdr:wsDr>`);
                zip.file(existingDrawingPath, merged);
            } else {
                // NEW: create a fresh drawing part + sheet rel + <drawing r:id>
                // + [Content_Types] Override.
                const drawingNum = ++drawingCounter;
                const anchorsXml = renumberCnvPrIds(group.anchors.join(''), 1);
                zip.file(`xl/drawings/drawing${drawingNum}.xml`, buildDrawingDoc(anchorsXml));

                const newRId = maxExistingRId(existingRelsXml) + 1;
                const sheetRel = `<Relationship Id="rId${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingNum}.xml"/>`;
                zip.file(sheetRelsPath, upsertRelationship(existingRelsXml, sheetRel));

                const sheetXmlFile = zip.file(sheetXmlPath);
                if (sheetXmlFile) {
                    const sheetXml = await sheetXmlFile.async('string');
                    zip.file(sheetXmlPath, insertDrawingRef(sheetXml, newRId));
                }

                const ctFile = zip.file('[Content_Types].xml');
                if (ctFile) {
                    const ctXml = await ctFile.async('string');
                    if (!ctXml.includes(`/xl/drawings/drawing${drawingNum}.xml`)) {
                        const override = `<Override PartName="/xl/drawings/drawing${drawingNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
                        zip.file(
                            '[Content_Types].xml',
                            ctXml.replace(/<\/Types>\s*$/, `${override}</Types>`),
                        );
                    }
                }
            }
        }

        const out = await zip.generateAsync({ type: 'arraybuffer' });
        return out as ArrayBuffer;
    } catch (e) {
        console.error('[Notesheet] M18 A2: injectShapesIntoZip failed; shapes dropped', e);
        return buffer;
    }
}
