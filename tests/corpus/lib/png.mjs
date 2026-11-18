// Deterministic PNG writer for builder-generated pictures. The zlib stream
// uses stored (uncompressed) deflate blocks, so the bytes do not depend on the
// zlib build. Pictures are palette images drawn from rectangles.

import { COLOR } from './grid.mjs';
import { crc32 } from './zip.mjs';

function chunk(type, data) {
  const out = Buffer.alloc(12 + data.length);
  out.writeUInt32BE(data.length, 0);
  out.write(type, 4, 'ascii');
  data.copy(out, 8);
  out.writeUInt32BE(crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
  return out;
}

function zlibStored(raw) {
  const blocks = [Buffer.from([0x78, 0x01])];
  for (let p = 0; p < raw.length || p === 0; p += 65535) {
    const part = raw.subarray(p, Math.min(p + 65535, raw.length));
    const head = Buffer.alloc(5);
    head[0] = p + 65535 >= raw.length ? 1 : 0;
    head.writeUInt16LE(part.length, 1);
    head.writeUInt16LE(part.length ^ 0xffff, 3);
    blocks.push(head, part);
    if (raw.length === 0) break;
  }
  let a = 1;
  let b = 0;
  for (const byte of raw) { a = (a + byte) % 65521; b = (b + a) % 65521; }
  const adler = Buffer.alloc(4);
  adler.writeUInt32BE(((b << 16) | a) >>> 0);
  blocks.push(adler);
  return Buffer.concat(blocks);
}

/** palette: ['RRGGBB', ...]; rects: [[x, y, w, h, paletteIndex], ...] drawn in
 * order over index 0. Returns PNG bytes (8-bit indexed). */
export function png(width, height, palette, rects) {
  const pixels = new Uint8Array(width * height);
  for (const [x, y, w, h, c] of rects) {
    for (let j = Math.max(0, y); j < Math.min(height, y + h); j += 1) {
      pixels.fill(c, j * width + Math.max(0, x), j * width + Math.min(width, x + w));
    }
  }
  const raw = Buffer.alloc((width + 1) * height);
  for (let j = 0; j < height; j += 1) {
    raw[j * (width + 1)] = 0;
    Buffer.from(pixels.subarray(j * width, (j + 1) * width)).copy(raw, j * (width + 1) + 1);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 3; // indexed colour
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('PLTE', Buffer.from(palette.join(''), 'hex')),
    chunk('IDAT', zlibStored(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

/** A picture of a grid: `cols` × `rows` fields separated by gutters, drawn as
 * 2 px outlines (or solid fields) in ink on paper. Sizes in pixels. */
export function gridPicture({ cols, rows, field, gutterX, fieldH, gutterY, solid = false, accent = [] }) {
  const width = cols * field + (cols - 1) * gutterX;
  const height = rows * fieldH + (rows - 1) * gutterY;
  const rects = [];
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const x = c * (field + gutterX);
      const y = r * (fieldH + gutterY);
      const ink = accent.some(([ac, ar]) => ac === c && ar === r) ? 2 : 1;
      if (solid) rects.push([x, y, field, fieldH, ink]);
      else rects.push([x, y, field, 2, ink], [x, y + fieldH - 2, field, 2, ink], [x, y, 2, fieldH, ink], [x + field - 2, y, 2, fieldH, ink]);
    }
  }
  return { width, height, bytes: png(width, height, [COLOR.paper, COLOR.ink, COLOR.accent], rects) };
}
