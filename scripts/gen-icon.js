// Generates a Notesheet icon: 96x96 PNG, white background, dark grid,
// one cell filled in Notesheet green (#34692E — the Aptos Medium4
// header colour we ship for the table style).
//
// Pure stdlib zlib + manual PNG encoder (no pngjs dep).

const { writeFileSync } = require('fs');
const { deflateSync } = require('zlib');

const W = 96;
const H = 96;
const BG = [255, 255, 255]; // white
const GRID = [80, 80, 80]; // medium-dark grey grid lines
const FILL = [52, 105, 46]; // #34692E Notesheet green

// Cell layout: 4x4 cells with 1px grid between. Margin: 4px. Each cell
// is (96 - 8 margin - 5 grid lines) / 4 ≈ 20 px wide. Use exact maths.
const MARGIN = 4;
const ROWS = 4,
    COLS = 4;
const GRID_W = 1;
const CELL_W = Math.floor((W - 2 * MARGIN - (COLS + 1) * GRID_W) / COLS);
const CELL_H = Math.floor((H - 2 * MARGIN - (ROWS + 1) * GRID_W) / ROWS);
// Actual painted region: MARGIN + 5 grid lines + 4 cells
// = MARGIN*2 + (ROWS+1)*GRID_W + ROWS*CELL_H = 4*2 + 5*1 + 4*20 = 93
// — leaves 3px of margin at the bottom-right; centre by recomputing.
const PAINT_W = MARGIN * 2 + (COLS + 1) * GRID_W + COLS * CELL_W;
const PAINT_H = MARGIN * 2 + (ROWS + 1) * GRID_W + ROWS * CELL_H;
const X_OFFSET = Math.floor((W - PAINT_W) / 2);
const Y_OFFSET = Math.floor((H - PAINT_H) / 2);

const filledCellRow = 2; // 0-indexed (3rd row from top)
const filledCellCol = 1; // 0-indexed (2nd col from left)

// rgba buffer — use RGBA so the PNG encoder is simple.
const buf = Buffer.alloc(W * H * 4);
function setPx(x, y, [r, g, b]) {
    if (x < 0 || x >= W || y < 0 || y >= H) return;
    const i = (y * W + x) * 4;
    buf[i] = r;
    buf[i + 1] = g;
    buf[i + 2] = b;
    buf[i + 3] = 255;
}

// Background.
for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) setPx(x, y, BG);

// Compute cell origin.
function cellRect(row, col) {
    const x0 = X_OFFSET + MARGIN + GRID_W + col * (CELL_W + GRID_W);
    const y0 = Y_OFFSET + MARGIN + GRID_W + row * (CELL_H + GRID_W);
    return { x0, y0, x1: x0 + CELL_W - 1, y1: y0 + CELL_H - 1 };
}

// Fill the highlight cell first so the grid lines paint on top.
{
    const { x0, y0, x1, y1 } = cellRect(filledCellRow, filledCellCol);
    for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) setPx(x, y, FILL);
}

// Grid lines: ROWS+1 horizontal + COLS+1 vertical inside the paint area.
const gridX0 = X_OFFSET + MARGIN;
const gridY0 = Y_OFFSET + MARGIN;
const gridX1 = X_OFFSET + PAINT_W - MARGIN - 1;
const gridY1 = Y_OFFSET + PAINT_H - MARGIN - 1;

// Horizontal lines.
for (let r = 0; r <= ROWS; r++) {
    const y = gridY0 + r * (CELL_H + GRID_W);
    for (let x = gridX0; x <= gridX1; x++) setPx(x, y, GRID);
}
// Vertical lines.
for (let c = 0; c <= COLS; c++) {
    const x = gridX0 + c * (CELL_W + GRID_W);
    for (let y = gridY0; y <= gridY1; y++) setPx(x, y, GRID);
}

// PNG encoding ---------------------------------------------------------
function chunk(type, data) {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const typeBuf = Buffer.from(type, 'ascii');
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])) >>> 0);
    return Buffer.concat([len, typeBuf, data, crc]);
}

function crc32(buf) {
    let c = 0xffffffff;
    for (const b of buf) {
        c ^= b;
        for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
    }
    return ~c >>> 0;
}

const SIG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const ihdr = Buffer.alloc(13);
ihdr.writeUInt32BE(W, 0);
ihdr.writeUInt32BE(H, 4);
ihdr[8] = 8; // bit depth
ihdr[9] = 6; // color type RGBA
ihdr[10] = 0;
ihdr[11] = 0;
ihdr[12] = 0; // not interlaced

// Add filter byte (0 = None) per scanline.
const filtered = Buffer.alloc((W * 4 + 1) * H);
for (let y = 0; y < H; y++) {
    filtered[y * (W * 4 + 1)] = 0;
    buf.copy(filtered, y * (W * 4 + 1) + 1, y * W * 4, (y + 1) * W * 4);
}
const idat = deflateSync(filtered);

const png = Buffer.concat([
    SIG,
    chunk('IHDR', ihdr),
    chunk('IDAT', idat),
    chunk('IEND', Buffer.alloc(0)),
]);

const outPath = process.argv[2] || 'src/icon.png';
writeFileSync(outPath, png);
console.log('Wrote', outPath, 'size:', png.length, 'bytes');
