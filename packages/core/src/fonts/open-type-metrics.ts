/** Raw line metrics read from one OpenType face. Values remain in design units;
 * format consumers decide which table and compatibility rule governs layout. */
export interface OpenTypeLineMetrics {
  readonly unitsPerEm: number;
  /** Positive OS/2 xAvgCharWidth / head.unitsPerEm, from this selected face.
   * This is a font-wide scalar, never a substitute for shaped text advances. */
  readonly averageCharWidthRatio?: number;
  readonly hheaAscent: number;
  readonly hheaDescent: number;
  readonly hheaLineGap: number;
  readonly typoAscent?: number;
  readonly typoDescent?: number;
  readonly typoLineGap?: number;
  readonly winAscent?: number;
  readonly winDescent?: number;
  readonly useTypoMetrics?: boolean;
  /** OS/2 ulCodePageRange1 bits 17–20 identify the Far East code-page class
   * observed by Word for Mac's automatic line allocation. `null` means the
   * OS/2 version/table does not declare code-page ranges; cmap coverage is
   * deliberately independent of this classification. */
  readonly farEastCodePage?: boolean | null;
  /** True only when a Unicode cmap maps at least one East Asian code point to
   * a non-zero glyph in this face. */
  readonly hasEastAsianCmap: boolean;
  /** Scalar ranges proven by the caller-resource parser; omitted by the legacy
   * line-metric API so its existing embedded-font policy remains unchanged. */
  readonly unicodeRanges?: readonly (readonly [number, number])[];
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
const FVAR = tagValue('fvar');

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

const MAX_CMAP_COVERAGE_RANGES = 32_768;
const MAX_CMAP_ENCODING_RECORDS = 4_096;
// Caller-resource governance, not a font-selection heuristic: limit unique
// table fan-out and cumulative scalar/group/intersection visits independently.
const MAX_CMAP_UNIQUE_SUBTABLES = 8;
const MAX_CMAP_COVERAGE_WORK = 262_144;
type CoverageWorkBudget = { remaining: number };

function format4UnicodeCoverage(
  view: DataView,
  offset: number,
  availableLength: number,
  budget: CoverageWorkBudget,
): Array<readonly [number, number]> | null {
  if (availableLength < 16) return null;
  const length = view.getUint16(offset + 2);
  if (length < 16 || length > availableLength) return null;
  const segCountX2 = view.getUint16(offset + 6);
  if (segCountX2 === 0 || segCountX2 % 2 !== 0) return null;
  const segCount = segCountX2 / 2;
  const endCodes = offset + 14;
  const startCodes = endCodes + segCount * 2 + 2;
  const deltas = startCodes + segCount * 2;
  const rangeOffsets = deltas + segCount * 2;
  if (rangeOffsets + segCount * 2 > offset + length) return null;

  const covered: Array<readonly [number, number]> = [];
  let rangeStart = -1;
  let rangeEnd = -1;
  const append = (codePoint: number) => {
    // U+0000 is a valid cmap input. Do not let the -1 sentinel make the first
    // scalar look like a continuation before a range has actually been added.
    if (covered.length > 0 && rangeEnd + 1 === codePoint) {
      rangeEnd = codePoint;
      covered[covered.length - 1] = [rangeStart, rangeEnd];
    } else {
      rangeStart = rangeEnd = codePoint;
      covered.push([codePoint, codePoint]);
    }
  };
  let previousEnd = -1;
  for (let index = 0; index < segCount; index++) {
    const start = view.getUint16(startCodes + index * 2);
    const end = view.getUint16(endCodes + index * 2);
    if (start > end || start <= previousEnd) return null;
    previousEnd = end;
    const delta = view.getInt16(deltas + index * 2);
    const rangeOffsetPosition = rangeOffsets + index * 2;
    const rangeOffset = view.getUint16(rangeOffsetPosition);
    for (let codePoint = start; codePoint <= end && codePoint < 0xffff; codePoint++) {
      if (--budget.remaining < 0) return null;
      const glyph = rangeOffset === 0
        ? (codePoint + delta) & 0xffff
        : (() => {
            const glyphPosition = rangeOffsetPosition + rangeOffset + (codePoint - start) * 2;
            if (glyphPosition + 2 > offset + length) return 0;
            const raw = view.getUint16(glyphPosition);
            return raw === 0 ? 0 : (raw + delta) & 0xffff;
          })();
      if (glyph !== 0) append(codePoint);
      if (covered.length > MAX_CMAP_COVERAGE_RANGES) return null;
    }
  }
  return covered;
}

function format12Or13UnicodeCoverage(
  view: DataView,
  offset: number,
  availableLength: number,
  constantGlyph: boolean,
  budget: CoverageWorkBudget,
): Array<readonly [number, number]> | null {
  if (availableLength < 16) return null;
  const length = view.getUint32(offset + 4);
  const groupCount = view.getUint32(offset + 12);
  if (length < 16 || length > availableLength || groupCount > (length - 16) / 12
    || groupCount > MAX_CMAP_COVERAGE_RANGES) return null;
  const covered: Array<readonly [number, number]> = [];
  let previousEnd = -1;
  for (let index = 0; index < groupCount; index++) {
    if (--budget.remaining < 0) return null;
    const group = offset + 16 + index * 12;
    const start = view.getUint32(group);
    const end = view.getUint32(group + 4);
    const startGlyph = view.getUint32(group + 8);
    if (start > end || end > 0x10ffff || start <= previousEnd) return null;
    previousEnd = end;
    if (constantGlyph) {
      if (startGlyph !== 0) covered.push([start, end]);
    } else {
      const mappedStart = startGlyph === 0 ? start + 1 : start;
      if (mappedStart <= end) covered.push([mappedStart, end]);
    }
  }
  return covered;
}

/** Coverage must hold regardless of the browser's choice of a Unicode base
 * cmap. Browser selection differs between Unicode and Windows records, so a
 * preferred-subtable ranking cannot establish resource authority. Intersect
 * all eligible base maps; any malformed/unreadable candidate or exhausted work
 * budget fails closed. Existing embedded-font parsing does not use this policy.
 */
function cmapUnicodeCoverage(
  view: DataView,
  table: Readonly<{ offset: number; length: number }> | undefined,
): readonly (readonly [number, number])[] {
  if (!table || table.length < 4) return [];
  const recordCount = view.getUint16(table.offset + 2);
  if (recordCount > MAX_CMAP_ENCODING_RECORDS || 4 + recordCount * 8 > table.length) return [];
  const seenOffsets = new Set<number>();
  const candidates: Array<Readonly<{
    subtable: number;
    availableLength: number;
    format: number;
  }>> = [];
  for (let index = 0; index < recordCount; index++) {
    const record = table.offset + 4 + index * 8;
    const platform = view.getUint16(record);
    const encoding = view.getUint16(record + 2);
    if (platform !== 0 && !(platform === 3 && (encoding === 1 || encoding === 10))) continue;
    const relativeOffset = view.getUint32(record + 4);
    if (relativeOffset > table.length - 2) return [];
    const subtable = table.offset + relativeOffset;
    const availableLength = table.length - relativeOffset;
    const format = view.getUint16(subtable);
    // OpenType cmap format 14 (Unicode platform, encoding 5) augments a base
    // map with variation sequences; it cannot independently select base glyphs.
    // It does not certify scalar coverage for an unsupported variation selector.
    if (platform === 0 && encoding === 5 && format === 14) continue;
    if (format !== 4 && format !== 12 && format !== 13) return [];
    if (seenOffsets.has(relativeOffset)) continue;
    seenOffsets.add(relativeOffset);
    if (seenOffsets.size > MAX_CMAP_UNIQUE_SUBTABLES) return [];
    candidates.push({ subtable, availableLength, format });
  }
  const budget: CoverageWorkBudget = { remaining: MAX_CMAP_COVERAGE_WORK };
  let coverage: Array<readonly [number, number]> | null = null;
  for (const candidate of candidates) {
    const parsed = candidate.format === 4
      ? format4UnicodeCoverage(view, candidate.subtable, candidate.availableLength, budget)
      : format12Or13UnicodeCoverage(
          view, candidate.subtable, candidate.availableLength, candidate.format === 13, budget,
        );
    if (!parsed?.length) return [];
    if (coverage === null) {
      coverage = parsed;
      continue;
    }
    const intersection: Array<readonly [number, number]> = [];
    let left = 0;
    let right = 0;
    while (left < coverage.length && right < parsed.length) {
      if (--budget.remaining < 0) return [];
      const a = coverage[left]!;
      const b = parsed[right]!;
      const start = Math.max(a[0], b[0]);
      const end = Math.min(a[1], b[1]);
      if (start <= end) {
        const previous = intersection.at(-1);
        if (previous && previous[1] + 1 === start) {
          intersection[intersection.length - 1] = [previous[0], end];
        } else {
          if (intersection.length >= MAX_CMAP_COVERAGE_RANGES) return [];
          intersection.push([start, end]);
        }
      }
      if (a[1] <= b[1]) left++;
      if (b[1] <= a[1]) right++;
    }
    if (!intersection.length) return [];
    coverage = intersection;
  }
  return Object.freeze((coverage ?? []).map((range) => Object.freeze(range)));
}

function rangesContainEastAsianGlyph(
  ranges: readonly (readonly [number, number])[],
): boolean {
  let rangeIndex = 0;
  let eastAsianIndex = 0;
  while (rangeIndex < ranges.length && eastAsianIndex < EAST_ASIAN_RANGES.length) {
    const [start, end] = ranges[rangeIndex]!;
    const [eastAsianStart, eastAsianEnd] = EAST_ASIAN_RANGES[eastAsianIndex]!;
    if (end < eastAsianStart) rangeIndex++;
    else if (eastAsianEnd < start) eastAsianIndex++;
    else return true;
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
  return readOpenTypeLineMetrics(bytes, faceIndex, false);
}

/** Strict coverage for any concrete resource whose design metrics may own
 * layout. The caller and embedded loaders must both check the returned cmap
 * ranges before attributing a shaped segment to this face. */
export function parseOpenTypeResourceMetrics(
  bytes: Uint8Array,
  faceIndex?: number,
): OpenTypeLineMetrics | null {
  return readOpenTypeLineMetrics(bytes, faceIndex, true);
}

function readOpenTypeLineMetrics(
  bytes: Uint8Array,
  faceIndex: number | undefined,
  resourceCoverage: boolean,
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
  // OpenType MVAR adjusts OS/2 and other global metrics for a selected variable
  // instance (https://learn.microsoft.com/en-us/typography/opentype/spec/mvar).
  // Caller-resource layout does not resolve axes or apply those deltas yet, so
  // default-instance tables cannot certify a requested weight/style. Restrict
  // this opt-in authority to static sfnt; the legacy embedded parser is unchanged.
  if (resourceCoverage && tables.has(FVAR)) return null;

  const head = tables.get(HEAD);
  const hhea = tables.get(HHEA);
  if (!head || head.length < 20 || !hhea || hhea.length < 10) return null;
  const unitsPerEm = view.getUint16(head.offset + 18);
  // OpenType `head.unitsPerEm` is constrained to 16..16384. Rejecting values
  // outside the format contract prevents malformed resources from amplifying
  // small signed hhea fields into unbounded layout ratios.
  if (unitsPerEm < 16 || unitsPerEm > 16384) return null;

  const os2 = tables.get(OS_2);
  // OpenType OS/2 xAvgCharWidth is a signed FWORD at +2. A non-positive or
  // unavailable value cannot define Word's observed inter-word-space floor.
  const averageCharWidth = os2 !== undefined && os2.length >= 4
    ? view.getInt16(os2.offset + 2)
    : 0;
  const hasWindowsMetrics = os2 !== undefined && os2.length >= 78;
  const hasCodePageRanges = os2 !== undefined && os2.length >= 86
    && view.getUint16(os2.offset) >= 1;
  const unicodeRanges = resourceCoverage
    ? cmapUnicodeCoverage(view, tables.get(CMAP)) : undefined;
  return Object.freeze({
    unitsPerEm,
    ...(averageCharWidth > 0
      ? { averageCharWidthRatio: averageCharWidth / unitsPerEm }
      : {}),
    hheaAscent: view.getInt16(hhea.offset + 4),
    hheaDescent: view.getInt16(hhea.offset + 6),
    hheaLineGap: view.getInt16(hhea.offset + 8),
    farEastCodePage: hasCodePageRanges
      ? (view.getUint32(os2.offset + 78) & 0x001e0000) !== 0
      : null,
    hasEastAsianCmap: unicodeRanges
      ? rangesContainEastAsianGlyph(unicodeRanges)
      : cmapHasEastAsianGlyph(view, tables.get(CMAP)),
    ...(unicodeRanges === undefined ? {} : { unicodeRanges }),
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
