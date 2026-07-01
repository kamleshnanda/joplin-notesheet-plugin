// M18 A1 regression + class guard: image export MUST work in the Joplin
// editor webview, where Node's global `Buffer` is UNDEFINED (only `atob`/
// `btoa` exist).
//
// The live bug: decodeDataUri() in src/drawings/xlsxImage.ts used
// `Buffer.from(base64, 'base64')`. In Jest (Node) that works, so every
// prior round-trip test passed — but in the real webview `Buffer` is not
// defined, decodeDataUri threw, the try/catch swallowed it to null, and
// the image was SILENTLY dropped from the exported .xlsx. Proven live via
// CDP: workbook.save() carried the image, yet the exported zip had no
// xl/media.
//
// This file has three layers:
//   1. Unit: readImagesFromSnapshot (decode) + bytesToBase64 (encode) run
//      correctly with global.Buffer deleted — reproduces the webview for OUR
//      code, isolated from JSZip/exceljs (whose Node builds call Buffer
//      internally and are webpack-swapped for browser builds in production).
//   2. Node smoke: injectImagesIntoZip end-to-end emits xl/media.
//   3. CLASS GUARD: a static scan of every webview-reachable module for an
//      unguarded bare-global `Buffer` — so the NEXT such bug (in charts,
//      shapes, or a future codec) fails at test time, not in a user's export.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import {
    readImagesFromSnapshot,
    bytesToBase64,
    injectImagesIntoZip,
} from '../src/drawings/xlsxImage';

// A 1x1 red PNG as a data URI (base64 body is what decodeDataUri must decode).
const RED_PNG_B64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function snapshotWithImage() {
    return {
        resources: [
            {
                name: 'SHEET_DRAWING_PLUGIN',
                data: JSON.stringify({
                    'sheet-1': {
                        data: {
                            img1: {
                                drawingType: 0,
                                imageSourceType: 'BASE64',
                                source: `data:image/png;base64,${RED_PNG_B64}`,
                                sheetTransform: {
                                    from: { column: 0, columnOffset: 0, row: 0, rowOffset: 0 },
                                    to: { column: 2, columnOffset: 0, row: 2, rowOffset: 0 },
                                },
                            },
                        },
                        order: ['img1'],
                    },
                }),
            },
        ],
    };
}

describe('M18 A1: decodeDataUri survives a Buffer-less (webview) runtime', () => {
    test('readImagesFromSnapshot decodes base64 with global Buffer undefined', () => {
        const snap = snapshotWithImage();
        const expectedBytes = Buffer.from(RED_PNG_B64, 'base64').length;

        const savedBuffer = (global as { Buffer?: unknown }).Buffer;

        delete (global as any).Buffer;
        let images: ReturnType<typeof readImagesFromSnapshot>;
        try {
            expect(typeof (global as { Buffer?: unknown }).Buffer).toBe('undefined');
            images = readImagesFromSnapshot(snap as never);
        } finally {
            (global as { Buffer?: unknown }).Buffer = savedBuffer;
        }

        // Before the fix: 0 images (decodeDataUri threw -> null -> skipped).
        expect(images).toHaveLength(1);
        expect(images[0].extension).toBe('png');
        expect(images[0].bytes.length).toBe(expectedBytes);
    });

    // The IMPORT path (xlsx.ts:2741) built a data: URI via
    // Buffer.from(bytes).toString('base64'), which throws in the webview the
    // same way. bytesToBase64 must round-trip with global Buffer undefined.
    test('bytesToBase64 encodes with global Buffer undefined', () => {
        const bytes = Buffer.from(RED_PNG_B64, 'base64');
        const expected = bytes.toString('base64');
        const input = new Uint8Array(bytes);

        const savedBuffer = (global as { Buffer?: unknown }).Buffer;

        delete (global as any).Buffer;
        let encoded: string;
        try {
            expect(typeof (global as { Buffer?: unknown }).Buffer).toBe('undefined');
            encoded = bytesToBase64(input);
        } finally {
            (global as { Buffer?: unknown }).Buffer = savedBuffer;
        }

        expect(encoded).toBe(expected);
    });

    // Smoke: injectImagesIntoZip runs and emits xl/media under Node (Buffer
    // present). We do NOT delete global.Buffer around the full path because
    // JSZip's Node build itself calls Buffer.isBuffer internally — in the real
    // webview webpack swaps in JSZip's browser bundle, a resolution Jest can't
    // reproduce, so a Buffer-less run here fails for a PRODUCTION-IMPOSSIBLE
    // reason (JSZip, not our code). The class of bug this feature had —
    // OUR OWN code touching the bare global Buffer — is caught statically by
    // the "no unguarded bare Buffer" guard below, which covers every
    // webview-reachable codec module, not just this one.
    test('injectImagesIntoZip emits xl/media (Node smoke)', async () => {
        const base = new JSZip();
        base.file(
            '[Content_Types].xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                '<Default Extension="xml" ContentType="application/xml"/>' +
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                '</Types>',
        );
        base.file(
            'xl/worksheets/sheet1.xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                '<sheetData/></worksheet>',
        );
        const baseBuffer = (await base.generateAsync({ type: 'arraybuffer' })) as ArrayBuffer;

        const snap = snapshotWithImage();
        (snap as { sheetOrder?: string[] }).sheetOrder = ['sheet-1'];
        const expectedBytes = Buffer.from(RED_PNG_B64, 'base64').length;

        const out = await injectImagesIntoZip(baseBuffer, snap as never);

        const zip = await JSZip.loadAsync(out);
        const mediaKeys = Object.keys(zip.files).filter((p) =>
            /^xl\/media\/image\d+\.png$/.test(p),
        );
        expect(mediaKeys).toHaveLength(1);
        const mediaBytes = await zip.files[mediaKeys[0]].async('uint8array');
        expect(mediaBytes.length).toBe(expectedBytes);
        expect(Array.from(mediaBytes.slice(0, 4))).toEqual([0x89, 0x50, 0x4e, 0x47]);
    });
});

// The fix above (atob/btoa in xlsxImage.ts) closes ONE site. This guard closes
// the CLASS: it fails if any webview-reachable module references the bare
// global `Buffer` without an explicit runtime guard. The Joplin editor
// webview has no global `Buffer`; webpack polyfills third-party libs'
// module-scoped `require('buffer')` but NOT a bare global, so our own code
// must never assume it. Node's Jest env masks this — hence a static guard.
//
// Intentional, guarded Node-only fallbacks are exempt via a `webview-buffer-ok`
// marker comment on the same line (see base64ToBytes/bytesToBase64), which
// only run when atob/btoa are absent (i.e. never in the webview).
describe('M18 A1: no webview-reachable module uses the bare global Buffer', () => {
    // Transitive closure of what editorView.tsx (webpack target:'web') loads.
    // index.ts is the Node HOST entry (target:'node') and is intentionally
    // excluded — it may use fs/Buffer freely.
    const WEBVIEW_MODULES = [
        'src/xlsx.ts',
        'src/snapshot.ts',
        'src/univerTableTheme.ts',
        'src/drawings/xlsxImage.ts',
        'src/drawings/xlsxImageImport.ts',
        'src/drawings/xlsxShape.ts',
        'src/drawings/xlsxShapeImport.ts',
        'src/drawings/sheetIdResolver.ts',
        'src/charts/xlsxChart.ts',
        'src/charts/xlsxChartImport.ts',
        'src/charts/extractData.ts',
    ];

    // Matches a runtime `Buffer` reference (Buffer.from, new Buffer, Buffer(),
    // Buffer.isBuffer, Buffer.alloc, Buffer.concat). Deliberately NOT matching
    // `: Buffer` / `| Buffer` type positions or `ArrayBuffer`.
    const BARE_BUFFER = /(?<![A-Za-z0-9_])Buffer\s*(?:\.\s*(?:from|alloc|concat|isBuffer)|\()/;

    test.each(WEBVIEW_MODULES)('%s has no unguarded bare Buffer', (relPath) => {
        const abs = path.join(__dirname, '..', relPath);
        const src = readFileSync(abs, 'utf8');
        const offenders: string[] = [];
        src.split('\n').forEach((line, i) => {
            const trimmed = line.trim();
            // Skip comment-only lines — prose mentioning Buffer isn't a
            // runtime reference (and can't throw).
            if (trimmed.startsWith('//') || trimmed.startsWith('*')) return;
            if (!BARE_BUFFER.test(line)) return;
            if (line.includes('webview-buffer-ok')) return; // explicit exemption
            offenders.push(`${relPath}:${i + 1}  ${trimmed}`);
        });
        expect(offenders).toEqual([]);
    });
});
