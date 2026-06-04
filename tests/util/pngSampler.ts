// Pure-stdlib PNG reader for tests.
//
// Exists so `tests/excelReferenceFidelity.test.ts` can sample dominant
// fill colours from operator-captured Excel renders without adding a
// runtime dependency (`pngjs`, `sharp`, etc). Supports the only PNG
// shapes our reference screenshots actually use:
//
//   - bit depth 8
//   - color type 2 (RGB) or 6 (RGBA)
//   - non-interlaced
//
// Filter unfilter (None/Sub/Up/Average/Paeth) follows the PNG spec
// (RFC 2083 §6 / W3C PNG (Second Edition) §9.2).
//
// **Don't promote this helper to a runtime dependency** — it's
// test-only. If a runtime path needs PNG decoding, that's a bigger
// design question.

import { readFileSync } from 'fs';
import { inflateSync } from 'zlib';

export interface DecodedPng {
    width: number;
    height: number;
    /** 3 (RGB) or 4 (RGBA). Test paths only sample the first 3. */
    channels: 3 | 4;
    /** Row-major raw bytes, length = width * height * channels. */
    data: Buffer;
}

const PNG_SIGNATURE_HEX = '89504e470d0a1a0a';

export function decodePng(filePath: string): DecodedPng {
    const buf = readFileSync(filePath);
    if (buf.subarray(0, 8).toString('hex') !== PNG_SIGNATURE_HEX) {
        throw new Error(`Not a PNG: ${filePath}`);
    }
    let i = 8;
    let width = 0; let height = 0; let bitDepth = 0; let colorType = 0; let interlace = 0;
    const idat: Buffer[] = [];
    while (i < buf.length) {
        const len = buf.readUInt32BE(i); i += 4;
        const type = buf.subarray(i, i + 4).toString('ascii'); i += 4;
        const data = buf.subarray(i, i + len); i += len;
        i += 4; // CRC, ignored
        if (type === 'IHDR') {
            width = data.readUInt32BE(0);
            height = data.readUInt32BE(4);
            bitDepth = data[8];
            colorType = data[9];
            interlace = data[12];
        } else if (type === 'IDAT') {
            idat.push(data);
        } else if (type === 'IEND') {
            break;
        }
    }
    if (bitDepth !== 8) throw new Error(`Unsupported bit depth ${bitDepth} (expected 8)`);
    if (interlace !== 0) throw new Error('Interlaced PNGs are not supported');
    const channels: 3 | 4 = colorType === 2 ? 3 : colorType === 6 ? 4 : (() => {
        throw new Error(`Unsupported colour type ${colorType} (expected 2 or 6)`);
    })();

    const raw = inflateSync(Buffer.concat(idat));
    const stride = width * channels + 1;
    const out = Buffer.alloc(width * height * channels);
    let prev = Buffer.alloc(width * channels);
    for (let y = 0; y < height; y++) {
        const filterByte = raw[y * stride];
        const scanline = raw.subarray(y * stride + 1, y * stride + stride);
        const cur = Buffer.alloc(width * channels);
        for (let x = 0; x < scanline.length; x++) {
            const a = x >= channels ? cur[x - channels] : 0;
            const b = prev[x];
            const c = x >= channels ? prev[x - channels] : 0;
            let v = 0;
            switch (filterByte) {
                case 0: v = scanline[x]; break;
                case 1: v = (scanline[x] + a) & 0xff; break;
                case 2: v = (scanline[x] + b) & 0xff; break;
                case 3: v = (scanline[x] + Math.floor((a + b) / 2)) & 0xff; break;
                case 4: {
                    const p = a + b - c;
                    const pa = Math.abs(p - a);
                    const pb = Math.abs(p - b);
                    const pc = Math.abs(p - c);
                    let nearest: number;
                    if (pa <= pb && pa <= pc) nearest = a;
                    else if (pb <= pc) nearest = b;
                    else nearest = c;
                    v = (scanline[x] + nearest) & 0xff;
                    break;
                }
                default: throw new Error(`Unknown PNG filter type ${filterByte}`);
            }
            cur[x] = v;
        }
        cur.copy(out, y * width * channels);
        prev = cur;
    }
    return { width, height, channels, data: out };
}

export function pixelAt(img: DecodedPng, x: number, y: number): [number, number, number] {
    const idx = (y * img.width + x) * img.channels;
    return [img.data[idx], img.data[idx + 1], img.data[idx + 2]];
}

/**
 * Find the most-frequent RGB tuple in a rectangular region.
 *
 * Used to identify the dominant fill colour of a header band / banded
 * row / etc., shrugging off anti-aliasing pixels at the cell edges.
 * Excel's screenshots are stable enough that the banded fill ends up
 * with thousands of identical-RGB pixels in the body of the row, so
 * "most-frequent" reliably picks out the band's fill.
 */
export function dominantColor(
    img: DecodedPng,
    xMin: number, xMax: number,
    yMin: number, yMax: number,
): { rgb: [number, number, number]; hex: string; hits: number; total: number } {
    const counts = new Map<number, number>();
    let total = 0;
    for (let y = yMin; y <= yMax; y++) {
        for (let x = xMin; x <= xMax; x++) {
            const idx = (y * img.width + x) * img.channels;
            const key = (img.data[idx] << 16) | (img.data[idx + 1] << 8) | img.data[idx + 2];
            counts.set(key, (counts.get(key) ?? 0) + 1);
            total += 1;
        }
    }
    let bestKey = 0; let bestHits = 0;
    for (const [key, hits] of counts) {
        if (hits > bestHits) { bestKey = key; bestHits = hits; }
    }
    const r = (bestKey >> 16) & 0xff;
    const g = (bestKey >> 8) & 0xff;
    const b = bestKey & 0xff;
    const hex = '#' + [r, g, b].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
    return { rgb: [r, g, b], hex, hits: bestHits, total };
}

/** Manhattan-style per-channel deviation between two RGB tuples. */
export function rgbDelta(a: readonly [number, number, number], b: readonly [number, number, number]): { dR: number; dG: number; dB: number; max: number } {
    const dR = Math.abs(a[0] - b[0]);
    const dG = Math.abs(a[1] - b[1]);
    const dB = Math.abs(a[2] - b[2]);
    return { dR, dG, dB, max: Math.max(dR, dG, dB) };
}

export function hexToRgb(hex: string): [number, number, number] {
    const s = hex.replace(/^#/, '');
    return [parseInt(s.slice(0, 2), 16), parseInt(s.slice(2, 4), 16), parseInt(s.slice(4, 6), 16)];
}
