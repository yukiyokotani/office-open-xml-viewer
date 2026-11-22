/**
 * Test-only TrueType faces for symbol-encoding tests: one box glyph per
 * mapped code, each with its own advance, under either a symbol cmap
 * (platform 3 encoding 0, codes at U+F000+code) or a Unicode cmap (3/1, bare
 * codes). Valid enough for browser font sanitizers.
 */
export function symbolFixtureFontBytes(options: Readonly<{
  /** 'both': a symbol subtable and a Unicode one mapping the bare codes. */
  cmap: 'symbol' | 'unicode' | 'both';
  /** Inclusive code range 0x20..0xFF the face maps. */
  first: number;
  last: number;
  name?: string;
}>): Uint8Array {
  const { first, last } = options;
  const name = options.name ?? 'OoxmlSymbolFixture';
  const u16 = (value: number) => [(value >> 8) & 255, value & 255];
  const i16 = (value: number) => u16(value & 0xffff);
  const u32 = (value: number) => [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255];
  const zeros = (count: number) => new Array<number>(count).fill(0);
  const upm = 1000;
  const count = last - first + 1;
  const advanceOf = (gid: number) => (gid === 0 ? 500 : 300 + ((gid * 37) % 600));
  const outline = (advance: number) => {
    const glyph = [...i16(1), ...i16(0), ...i16(0), ...i16(advance), ...i16(700), ...u16(3), ...u16(0), 1, 1, 1, 1];
    const points = [[0, 0], [advance, 0], [advance, 700], [0, 700]];
    const xs: number[] = [];
    const ys: number[] = [];
    let px = 0;
    let py = 0;
    for (const [x, y] of points) { xs.push(...i16(x! - px)); ys.push(...i16(y! - py)); px = x!; py = y!; }
    glyph.push(...xs, ...ys);
    while (glyph.length % 4) glyph.push(0);
    return glyph;
  };
  const glyf: number[] = [];
  const loca: number[] = [...u32(0), ...u32(0)]; // .notdef: empty
  const hmtx: number[] = [...u16(advanceOf(0)), ...i16(0)];
  for (let gid = 1; gid <= count; gid++) {
    glyf.push(...outline(advanceOf(gid)));
    loca.push(...u32(glyf.length));
    hmtx.push(...u16(advanceOf(gid)), ...i16(0));
  }
  const format4 = (base: number) => {
    const start = base + first;
    const end = base + last;
    return [
      ...u16(4), ...u16(32), ...u16(0), ...u16(4), ...u16(4), ...u16(1), ...u16(0),
      ...u16(end), ...u16(0xffff), ...u16(0), ...u16(start), ...u16(0xffff),
      ...i16((1 - start) & 0xffff), ...i16(1), ...u16(0), ...u16(0),
    ];
  };
  const start = (options.cmap === 'unicode' ? 0 : 0xf000) + first;
  const end = (options.cmap === 'unicode' ? 0 : 0xf000) + last;
  const subtables: [number, number[]][] = options.cmap === 'unicode' ? [[1, format4(0)]]
    : options.cmap === 'symbol' ? [[0, format4(0xf000)]] : [[0, format4(0xf000)], [1, format4(0)]];
  const cmapTable = [...u16(0), ...u16(subtables.length)];
  let subtableOffset = 4 + subtables.length * 8;
  for (const [encoding, table] of subtables) {
    cmapTable.push(...u16(3), ...u16(encoding), ...u32(subtableOffset));
    subtableOffset += table.length;
  }
  for (const [, table] of subtables) cmapTable.push(...table);
  const nameRecords: number[] = [];
  const nameStorage: number[] = [];
  for (const [id, value] of [[1, name], [2, 'Regular'], [4, name], [6, name]] as const) {
    const bytes = [...value].flatMap((character) => u16(character.charCodeAt(0)));
    nameRecords.push(...u16(3), ...u16(1), ...u16(0x409), ...u16(id), ...u16(bytes.length), ...u16(nameStorage.length));
    nameStorage.push(...bytes);
  }
  const tables: Record<string, number[]> = {
    'OS/2': [
      ...u16(4), ...i16(500), ...u16(400), ...u16(5), ...u16(0), ...zeros(16),
      ...i16(50), ...i16(250), ...i16(0), ...zeros(10), ...zeros(16), 0x4f, 0x4f, 0x58, 0x4d,
      ...u16(0x0040), ...u16(Math.min(start, 0xffff)), ...u16(Math.min(end, 0xffff)),
      ...i16(800), ...i16(-200), ...i16(0), ...u16(800), ...u16(200),
      ...u32(options.cmap === 'unicode' ? 1 : 0x80000000), ...u32(0), ...i16(700), ...i16(700), ...u16(0), ...u16(0x20), ...u16(1),
    ],
    cmap: cmapTable,
    glyf,
    head: [
      ...u32(0x00010000), ...u32(0x00010000), ...u32(0), ...u32(0x5f0f3cf5), ...u16(0x000b),
      ...u16(upm), ...zeros(16), ...i16(0), ...i16(0), ...i16(900), ...i16(700),
      ...u16(0), ...u16(8), ...i16(2), ...i16(1), ...i16(0),
    ],
    hhea: [
      ...u32(0x00010000), ...i16(800), ...i16(-200), ...i16(0), ...u16(900),
      ...i16(0), ...i16(0), ...i16(900), ...i16(1), ...i16(0), ...i16(0), ...zeros(8), ...i16(0), ...u16(count + 1),
    ],
    hmtx,
    loca,
    maxp: [...u32(0x00010000), ...u16(count + 1), ...u16(4), ...u16(1), ...zeros(4), ...u16(2), ...zeros(16)],
    name: [...u16(0), ...u16(4), ...u16(6 + nameRecords.length), ...nameRecords, ...nameStorage],
    post: [...u32(0x00030000), ...zeros(4), ...i16(-100), ...i16(50), ...zeros(20)],
  };
  const tags = Object.keys(tables).sort();
  let selector = 0;
  while ((1 << (selector + 1)) <= tags.length) selector += 1;
  const searchRange = (1 << selector) * 16;
  const checksum = (bytes: ArrayLike<number>) => {
    let sum = 0;
    for (let index = 0; index < bytes.length; index += 4) {
      sum = (sum + ((((bytes[index] ?? 0) << 24) >>> 0) + ((bytes[index + 1] ?? 0) << 16)
        + ((bytes[index + 2] ?? 0) << 8) + (bytes[index + 3] ?? 0))) >>> 0;
    }
    return sum;
  };
  const directory: number[] = [];
  const body: number[] = [];
  let offset = 12 + 16 * tags.length;
  let headOffset = 0;
  for (const tag of tags) {
    const data = tables[tag]!;
    if (tag === 'head') headOffset = offset;
    directory.push(...[...tag].map((character) => character.charCodeAt(0)),
      ...u32(checksum(data)), ...u32(offset), ...u32(data.length));
    body.push(...data);
    while (body.length % 4) body.push(0);
    offset = 12 + 16 * tags.length + body.length;
  }
  const font = new Uint8Array([
    ...u32(0x00010000), ...u16(tags.length), ...u16(searchRange), ...u16(selector),
    ...u16(tags.length * 16 - searchRange), ...directory, ...body,
  ]);
  font.set(u32((0xb1b0afba - checksum(font)) >>> 0), headOffset + 8);
  return font;
}
