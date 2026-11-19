// Enhanced Metafile writer ([MS-EMF]) for the specimens. Records are written
// one by one, little-endian, so a specimen states exactly what a renderer
// must play back. Coordinates are device pixels of a reference device with
// PX_PER_MM pixels per millimetre (MM_TEXT unless a case sets a map mode);
// the picture frame is the whole drawing surface.

export const PX_PER_MM = 10;
const REF_MM = [210, 297]; // reference device: A4 at PX_PER_MM

const rec = (type, fill) => (size) => {
  const b = Buffer.alloc(size);
  b.writeUInt32LE(type, 0);
  b.writeUInt32LE(size, 4);
  fill?.(b);
  return b;
};
const rectAt = (b, o, [l, t, r, bt]) => { b.writeInt32LE(l, o); b.writeInt32LE(t, o + 4); b.writeInt32LE(r, o + 8); b.writeInt32LE(bt, o + 12); };
const colorAt = (b, o, hex) => { b.writeUInt8(parseInt(hex.slice(0, 2), 16), o); b.writeUInt8(parseInt(hex.slice(2, 4), 16), o + 1); b.writeUInt8(parseInt(hex.slice(4, 6), 16), o + 2); };

export const RGN_AND = 1;
export const RGN_DIFF = 4;
export const RGN_COPY = 5;
export const MM_ANISOTROPIC = 8;
export const NULL_BRUSH = 0x80000005;
export const NULL_PEN = 0x80000008;

/** Record builders; each returns a Buffer. Rectangles are [left, top, right, bottom]. */
export const R = {
  setMapMode: (mode) => rec(17, (b) => b.writeUInt32LE(mode, 8))(12),
  setWindowExt: (cx, cy) => rec(9, (b) => { b.writeInt32LE(cx, 8); b.writeInt32LE(cy, 12); })(16),
  setViewportExt: (cx, cy) => rec(11, (b) => { b.writeInt32LE(cx, 8); b.writeInt32LE(cy, 12); })(16),
  saveDC: () => rec(33)(8),
  restoreDC: (rel = -1) => rec(34, (b) => b.writeInt32LE(rel, 8))(12),
  /** Solid brush (LogBrush BS_SOLID) in object slot ih. */
  brush: (ih, hex) => rec(39, (b) => { b.writeUInt32LE(ih, 8); b.writeUInt32LE(0, 12); colorAt(b, 16, hex); b.writeUInt32LE(0, 20); })(24),
  /** Cosmetic solid pen (LogPen PS_SOLID) of the given width in object slot ih. */
  pen: (ih, width, hex) => rec(38, (b) => { b.writeUInt32LE(ih, 8); b.writeUInt32LE(0, 12); b.writeInt32LE(width, 16); b.writeInt32LE(0, 20); colorAt(b, 24, hex); })(28),
  select: (ih) => rec(37, (b) => b.writeUInt32LE(ih, 8))(12),
  remove: (ih) => rec(40, (b) => b.writeUInt32LE(ih, 8))(12),
  rectangle: (box) => rec(43, (b) => rectAt(b, 8, box))(24),
  ellipse: (box) => rec(42, (b) => rectAt(b, 8, box))(24),
  moveTo: (x, y) => rec(27, (b) => { b.writeInt32LE(x, 8); b.writeInt32LE(y, 12); })(16),
  lineTo: (x, y) => rec(54, (b) => { b.writeInt32LE(x, 8); b.writeInt32LE(y, 12); })(16),
  intersectClipRect: (box) => rec(30, (b) => rectAt(b, 8, box))(24),
  excludeClipRect: (box) => rec(29, (b) => rectAt(b, 8, box))(24),
  offsetClipRgn: (dx, dy) => rec(26, (b) => { b.writeInt32LE(dx, 8); b.writeInt32LE(dy, 12); })(16),
  /** EMR_EXTSELECTCLIPRGN with a region of rectangles (device units). */
  extSelectClipRgn: (mode, rects) => {
    const data = 32 + 16 * rects.length;
    return rec(75, (b) => {
      b.writeUInt32LE(data, 8);
      b.writeUInt32LE(mode, 12);
      b.writeUInt32LE(32, 16); b.writeUInt32LE(1, 20); b.writeUInt32LE(rects.length, 24); b.writeUInt32LE(16 * rects.length, 28);
      rectAt(b, 32, [Math.min(...rects.map((r) => r[0])), Math.min(...rects.map((r) => r[1])), Math.max(...rects.map((r) => r[2])), Math.max(...rects.map((r) => r[3]))]);
      rects.forEach((r, i) => rectAt(b, 48 + 16 * i, r));
    })(16 + data);
  },
  beginPath: () => rec(59)(8),
  endPath: () => rec(60)(8),
  selectClipPath: (mode) => rec(67, (b) => b.writeUInt32LE(mode, 8))(12),
};

/**
 * A complete EMF: header (with the szlMicrometers extension), the records,
 * EOF. width × height device pixels; the frame covers exactly that area.
 */
export function emf({ width, height, handles, records }) {
  const eof = rec(14, (b) => { b.writeUInt32LE(0, 8); b.writeUInt32LE(16, 12); b.writeUInt32LE(20, 16); })(20);
  const body = [...records, eof];
  const headerSize = 108;
  const total = headerSize + body.reduce((n, r) => n + r.length, 0);
  const h = rec(1, (b) => {
    rectAt(b, 8, [0, 0, width - 1, height - 1]);
    rectAt(b, 24, [0, 0, Math.round((width * 100) / PX_PER_MM), Math.round((height * 100) / PX_PER_MM)]);
    b.writeUInt32LE(0x464d4520, 40); // " EMF"
    b.writeUInt32LE(0x10000, 44);
    b.writeUInt32LE(total, 48);
    b.writeUInt32LE(body.length + 1, 52);
    b.writeUInt16LE(handles + 1, 56);
    b.writeInt32LE(REF_MM[0] * PX_PER_MM, 72); b.writeInt32LE(REF_MM[1] * PX_PER_MM, 76);
    b.writeInt32LE(REF_MM[0], 80); b.writeInt32LE(REF_MM[1], 84);
    b.writeInt32LE(REF_MM[0] * 1000, 100); b.writeInt32LE(REF_MM[1] * 1000, 104);
  })(headerSize);
  return Buffer.concat([h, ...body]);
}
