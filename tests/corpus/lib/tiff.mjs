// Baseline TIFF (8-bit RGB, one strip), the smallest file a TIFF reader must
// accept: uncompressed, or PackBits (compression 32773, also baseline).
// Little-endian.

/** PackBits, row by row (a run never crosses a row): runs of 2-128 equal
 * bytes as (1 - n, byte), everything else as literals (n - 1, bytes...). */
function packBits(data, rowBytes) {
  const out = [];
  for (let row = 0; row < data.length; row += rowBytes) {
    const end = row + rowBytes;
    let i = row;
    while (i < end) {
      let run = 1;
      while (i + run < end && run < 128 && data[i + run] === data[i]) run += 1;
      if (run >= 2) { out.push(257 - run, data[i]); i += run; continue; }
      let lit = 1;
      while (i + lit < end && lit < 128 && !(i + lit + 1 < end && data[i + lit] === data[i + lit + 1])) lit += 1;
      out.push(lit - 1, ...data.subarray(i, i + lit));
      i += lit;
    }
  }
  return Buffer.from(out);
}

/** palette: ['RRGGBB', ...]; rects: [[x, y, w, h, paletteIndex], ...] drawn in
 * order over index 0. compression: 'none' or 'packbits'. Returns TIFF bytes
 * at 96 dpi. */
export function tiff(width, height, palette, rects, compression = 'none') {
  const colours = palette.map((hex) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)));
  const pixels = Buffer.alloc(width * height * 3);
  const fill = (x, y, w, h, c) => {
    for (let j = Math.max(0, y); j < Math.min(height, y + h); j += 1) {
      for (let i = Math.max(0, x); i < Math.min(width, x + w); i += 1) pixels.set(colours[c], (j * width + i) * 3);
    }
  };
  fill(0, 0, width, height, 0);
  for (const [x, y, w, h, c] of rects) fill(x, y, w, h, c);

  const strip = compression === 'packbits' ? packBits(pixels, width * 3) : pixels;
  const tags = [
    [256, 3, 1, width], [257, 3, 1, height], [258, 3, 3, 'bits'], [259, 3, 1, compression === 'packbits' ? 32773 : 1], [262, 3, 1, 2],
    [273, 4, 1, 'strip'], [277, 3, 1, 3], [278, 3, 1, height], [279, 4, 1, strip.length],
    [282, 5, 1, 'xres'], [283, 5, 1, 'yres'], [296, 3, 1, 2],
  ];
  const ifdOffset = 8;
  const ifdSize = 2 + tags.length * 12 + 4;
  const bitsOffset = ifdOffset + ifdSize;
  const xresOffset = bitsOffset + 6;
  const yresOffset = xresOffset + 8;
  const stripOffset = yresOffset + 8;
  const out = Buffer.alloc(stripOffset + strip.length);
  out.write('II', 0, 'latin1');
  out.writeUInt16LE(42, 2);
  out.writeUInt32LE(ifdOffset, 4);
  out.writeUInt16LE(tags.length, ifdOffset);
  tags.forEach(([tag, type, count, value], i) => {
    const at = ifdOffset + 2 + i * 12;
    out.writeUInt16LE(tag, at);
    out.writeUInt16LE(type, at + 2);
    out.writeUInt32LE(count, at + 4);
    const v = { bits: bitsOffset, strip: stripOffset, xres: xresOffset, yres: yresOffset }[value] ?? value;
    if (type === 3 && count === 1) out.writeUInt16LE(v, at + 8);
    else out.writeUInt32LE(v, at + 8);
  });
  out.writeUInt32LE(0, ifdOffset + 2 + tags.length * 12);
  [8, 8, 8].forEach((b, i) => out.writeUInt16LE(b, bitsOffset + i * 2));
  for (const at of [xresOffset, yresOffset]) { out.writeUInt32LE(96, at); out.writeUInt32LE(1, at + 4); }
  strip.copy(out, stripOffset);
  return out;
}
