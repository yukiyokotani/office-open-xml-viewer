/** Raw line metrics read from one OpenType face. Values remain in design units;
 * format consumers decide which table and compatibility rule governs layout. */
export interface OpenTypeLineMetrics {
  readonly unitsPerEm: number;
  readonly hheaAscent: number;
  readonly hheaDescent: number;
  readonly hheaLineGap: number;
  readonly typoAscent?: number;
  readonly typoDescent?: number;
  readonly typoLineGap?: number;
  readonly winAscent?: number;
  readonly winDescent?: number;
  readonly useTypoMetrics?: boolean;
  /** True only when a Unicode cmap maps at least one East Asian code point to
   * a non-zero glyph in this face. */
  readonly hasEastAsianCmap: boolean;
}

const tagValue = (tag: string): number => (
  ((tag.charCodeAt(0) << 24) >>> 0)
  | (tag.charCodeAt(1) << 16)
  | (tag.charCodeAt(2) << 8)
  | tag.charCodeAt(3)
) >>> 0;

const TTCF = tagValue('ttcf');
const OTTO = tagValue('OTTO');
const TRUE = tagValue('true');
const TYP1 = tagValue('typ1');
const HEAD = tagValue('head');
const HHEA = tagValue('hhea');
const OS_2 = tagValue('OS/2');
const CMAP = tagValue('cmap');

// Require script-bearing ranges rather than compatibility-width punctuation or
// Latin forms, which many otherwise Latin fonts include without owning a CJK
// text face.
const EAST_ASIAN_RANGES: ReadonlyArray<readonly [number, number]> = [
  [0x1100, 0x11ff],
  [0x2e80, 0x2fff],
  [0x3040, 0x30ff],
  [0x3100, 0x312f],
  [0x3130, 0x318f],
  [0x31a0, 0x31bf],
  [0x31f0, 0x31ff],
  [0x3400, 0x4dbf],
  [0x4e00, 0x9fff],
  [0xa000, 0xa4cf],
  [0xa960, 0xa97f],
  [0xac00, 0xd7a3],
  [0xf900, 0xfaff],
  [0x20000, 0x323af],
];

function rangeFits(length: number, offset: number, size: number): boolean {
  return Number.isSafeInteger(offset) && Number.isSafeInteger(size)
    && offset >= 0 && size >= 0 && offset <= length - size;
}

function format4HasEastAsianGlyph(
  view: DataView,
  offset: number,
  availableLength: number,
): boolean {
  if (availableLength < 16) return false;
  const length = view.getUint16(offset + 2);
  if (length < 16 || length > availableLength) return false;
  const segCountX2 = view.getUint16(offset + 6);
  if (segCountX2 === 0 || segCountX2 % 2 !== 0) return false;
  const segCount = segCountX2 / 2;
  const endCodes = offset + 14;
  const startCodes = endCodes + segCount * 2 + 2;
  const deltas = startCodes + segCount * 2;
  const rangeOffsets = deltas + segCount * 2;
  if (rangeOffsets + segCount * 2 > offset + length) return false;

  // Format 4 segments are ordered and non-overlapping. Validate that invariant
  // before scanning glyph ids so a malformed font cannot make the nested loop
  // revisit the full BMP once per segment.
  let previousEnd = -1;
  for (let index = 0; index < segCount; index++) {
    const start = view.getUint16(startCodes + index * 2);
    const end = view.getUint16(endCodes + index * 2);
    if (start > end || start <= previousEnd) return false;
    previousEnd = end;
  }

  for (let index = 0; index < segCount; index++) {
    const start = view.getUint16(startCodes + index * 2);
    const end = view.getUint16(endCodes + index * 2);
    const delta = view.getInt16(deltas + index * 2);
    const rangeOffsetPosition = rangeOffsets + index * 2;
    const rangeOffset = view.getUint16(rangeOffsetPosition);
    for (const [rangeStart, rangeEnd] of EAST_ASIAN_RANGES) {
      const from = Math.max(start, rangeStart);
      const to = Math.min(end, rangeEnd);
      if (from > to) continue;
      if (rangeOffset === 0) {
        if (from < to || ((from + delta) & 0xffff) !== 0) return true;
        continue;
      }
      for (let codePoint = from; codePoint <= to; codePoint++) {
        const glyphPosition = rangeOffsetPosition + rangeOffset + (codePoint - start) * 2;
        if (glyphPosition + 2 > offset + length) break;
        const glyph = view.getUint16(glyphPosition);
        if (glyph !== 0 && ((glyph + delta) & 0xffff) !== 0) return true;
      }
    }
  }
  return false;
}

function format12Or13HasEastAsianGlyph(
  view: DataView,
  offset: number,
  availableLength: number,
  constantGlyph: boolean,
): boolean {
  if (availableLength < 16) return false;
  const length = view.getUint32(offset + 4);
  const groupCount = view.getUint32(offset + 12);
  if (length < 16 || length > availableLength || groupCount > (length - 16) / 12) return false;
  let previousEnd = -1;
  for (let index = 0; index < groupCount; index++) {
    const group = offset + 16 + index * 12;
    const start = view.getUint32(group);
    const end = view.getUint32(group + 4);
    const startGlyph = view.getUint32(group + 8);
    // Formats 12/13 require sorted, non-overlapping Unicode scalar ranges.
    // Enforcing that invariant both rejects ambiguous data and lets a face with
    // a very large cmap stop once all relevant ranges have been passed.
    if (start > end || end > 0x10ffff || start <= previousEnd) return false;
    previousEnd = end;
    if (start > EAST_ASIAN_RANGES[EAST_ASIAN_RANGES.length - 1][1]) break;
    for (const [rangeStart, rangeEnd] of EAST_ASIAN_RANGES) {
      const from = Math.max(start, rangeStart);
      const to = Math.min(end, rangeEnd);
      if (from > to) continue;
      if (constantGlyph) {
        if (startGlyph !== 0) return true;
      } else if (from < to || startGlyph + (from - start) !== 0) {
        return true;
      }
    }
  }
  return false;
}

function cmapHasEastAsianGlyph(
  view: DataView,
  table: Readonly<{ offset: number; length: number }> | undefined,
): boolean {
  if (!table || table.length < 4) return false;
  const recordCount = view.getUint16(table.offset + 2);
  if (4 + recordCount * 8 > table.length) return false;
  for (let index = 0; index < recordCount; index++) {
    const record = table.offset + 4 + index * 8;
    const platform = view.getUint16(record);
    const encoding = view.getUint16(record + 2);
    if (platform !== 0 && !(platform === 3 && (encoding === 1 || encoding === 10))) continue;
    const relativeOffset = view.getUint32(record + 4);
    if (relativeOffset > table.length - 2) continue;
    const subtable = table.offset + relativeOffset;
    const availableLength = table.length - relativeOffset;
    const format = view.getUint16(subtable);
    if (format === 4 && format4HasEastAsianGlyph(view, subtable, availableLength)) return true;
    if (format === 12 && format12Or13HasEastAsianGlyph(view, subtable, availableLength, false)) return true;
    if (format === 13 && format12Or13HasEastAsianGlyph(view, subtable, availableLength, true)) return true;
  }
  return false;
}

/**
 * Parse the table-directory metrics needed by OOXML layout from a raw sfnt or
 * TrueType Collection. This deliberately reads no family names: identity and
 * style selection belong to the font-resource loader that selected the face.
 * Malformed or unsupported input returns `null` instead of exposing partial
 * metrics to pagination.
 */
export function parseOpenTypeLineMetrics(
  bytes: Uint8Array,
  faceIndex?: number,
): OpenTypeLineMetrics | null {
  if ((faceIndex !== undefined && (!Number.isSafeInteger(faceIndex) || faceIndex < 0))
    || bytes.byteLength < 12) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let sfntOffset = 0;
  const signature = view.getUint32(0);
  if (signature === TTCF) {
    // A collection has no intrinsically correct default face. The resource
    // owner must resolve its family/style identity and pass the matching index.
    if (faceIndex === undefined) return null;
    if (bytes.byteLength < 12) return null;
    const faceCount = view.getUint32(8);
    if (faceIndex >= faceCount || !rangeFits(bytes.byteLength, 12, faceCount * 4)) return null;
    sfntOffset = view.getUint32(12 + faceIndex * 4);
  } else if (faceIndex !== undefined && faceIndex !== 0) {
    return null;
  }
  if (!rangeFits(bytes.byteLength, sfntOffset, 12)) return null;
  const scaler = view.getUint32(sfntOffset);
  if (scaler !== 0x00010000 && scaler !== OTTO && scaler !== TRUE && scaler !== TYP1) return null;
  const tableCount = view.getUint16(sfntOffset + 4);
  if (!rangeFits(bytes.byteLength, sfntOffset + 12, tableCount * 16)) return null;

  const tables = new Map<number, { offset: number; length: number }>();
  for (let index = 0; index < tableCount; index++) {
    const record = sfntOffset + 12 + index * 16;
    const tag = view.getUint32(record);
    const offset = view.getUint32(record + 8);
    const length = view.getUint32(record + 12);
    if (tables.has(tag) || !rangeFits(bytes.byteLength, offset, length)) return null;
    tables.set(tag, { offset, length });
  }
  const head = tables.get(HEAD);
  const hhea = tables.get(HHEA);
  if (!head || head.length < 20 || !hhea || hhea.length < 10) return null;
  const unitsPerEm = view.getUint16(head.offset + 18);
  // OpenType `head.unitsPerEm` is constrained to 16..16384. Rejecting values
  // outside the format contract prevents malformed resources from amplifying
  // small signed hhea fields into unbounded layout ratios.
  if (unitsPerEm < 16 || unitsPerEm > 16384) return null;

  const os2 = tables.get(OS_2);
  const hasWindowsMetrics = os2 !== undefined && os2.length >= 78;
  return Object.freeze({
    unitsPerEm,
    hheaAscent: view.getInt16(hhea.offset + 4),
    hheaDescent: view.getInt16(hhea.offset + 6),
    hheaLineGap: view.getInt16(hhea.offset + 8),
    hasEastAsianCmap: cmapHasEastAsianGlyph(view, tables.get(CMAP)),
    ...(hasWindowsMetrics ? {
      typoAscent: view.getInt16(os2.offset + 68),
      typoDescent: view.getInt16(os2.offset + 70),
      typoLineGap: view.getInt16(os2.offset + 72),
      winAscent: view.getUint16(os2.offset + 74),
      winDescent: view.getUint16(os2.offset + 76),
      useTypoMetrics: (view.getUint16(os2.offset + 62) & 0x0080) !== 0,
    } : {}),
  });
}
