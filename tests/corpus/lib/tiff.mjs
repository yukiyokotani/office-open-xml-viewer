// Baseline TIFF (uncompressed 8-bit RGB, one strip), the smallest file a
// TIFF reader must accept. Little-endian.

/** palette: ['RRGGBB', ...]; rects: [[x, y, w, h, paletteIndex], ...] drawn in
 * order over index 0. Returns TIFF bytes at 96 dpi. */
export function tiff(width, height, palette, rects) {
  const colours = palette.map((hex) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)));
  const pixels = Buffer.alloc(width * height * 3);
  const fill = (x, y, w, h, c) => {
    for (let j = Math.max(0, y); j < Math.min(height, y + h); j += 1) {
      for (let i = Math.max(0, x); i < Math.min(width, x + w); i += 1) pixels.set(colours[c], (j * width + i) * 3);
    }
  };
  fill(0, 0, width, height, 0);
  for (const [x, y, w, h, c] of rects) fill(x, y, w, h, c);

  const tags = [
    [256, 3, 1, width], [257, 3, 1, height], [258, 3, 3, 'bits'], [259, 3, 1, 1], [262, 3, 1, 2],
    [273, 4, 1, 'strip'], [277, 3, 1, 3], [278, 3, 1, height], [279, 4, 1, pixels.length],
    [282, 5, 1, 'xres'], [283, 5, 1, 'yres'], [296, 3, 1, 2],
  ];
  const ifdOffset = 8;
  const ifdSize = 2 + tags.length * 12 + 4;
  const bitsOffset = ifdOffset + ifdSize;
  const xresOffset = bitsOffset + 6;
  const yresOffset = xresOffset + 8;
  const stripOffset = yresOffset + 8;
  const out = Buffer.alloc(stripOffset + pixels.length);
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
  pixels.copy(out, stripOffset);
  return out;
}
