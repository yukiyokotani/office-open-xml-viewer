/**
 * Which symbol codes one selected face draws itself (symbol-font.ts).
 *
 * Where the library holds the face's bytes (a document-embedded face, a
 * served `fonts` face whose tables were read), the cmap answers exactly
 * ({@link symbolCmapEncoding}). Otherwise (a host-installed face, an
 * application FontFace) the face is measured in the very Canvas context that
 * measures and paints the text ({@link measureSymbolFaceEncoding}).
 */
import {
  SYMBOL_FACE_UNICODE,
  SYMBOL_FACE_UNKNOWN,
  symbolCodeSet,
  symbolCodeSetHas,
  type SymbolFaceEncoding,
  type SymbolTextOptions,
  type SymbolTextSegment,
  symbolTextToUnicodeSegments,
} from './symbol-font.js';
import { parseUnicodeRange } from './css-source.js';
import { fontFaceWeightCovers } from './office-fallback.js';

const tag = (value: string): number => (
  ((value.charCodeAt(0) << 24) >>> 0) | (value.charCodeAt(1) << 16)
  | (value.charCodeAt(2) << 8) | value.charCodeAt(3)
) >>> 0;

/** Codes 0x00..0xFF mapped to a non-zero glyph by one cmap subtable, split by
 * form: at U+F000+code and at the bare code. Formats 0, 4, 6, 12 and 13 —
 * every format a BMP subtable may use. Null: malformed. */
function subtableCodes(
  view: DataView, offset: number, available: number,
): { pua: Set<number>; bare: Set<number> } | null {
  const pua = new Set<number>();
  const bare = new Set<number>();
  const add = (cp: number) => {
    if (cp >= 0xf000 && cp <= 0xf0ff) pua.add(cp - 0xf000);
    else if (cp <= 0xff) bare.add(cp);
  };
  if (available < 6) return null;
  const format = view.getUint16(offset);
  if (format === 0) {
    if (available < 262) return null;
    for (let code = 0; code < 256; code++) if (view.getUint8(offset + 6 + code) !== 0) add(code);
    return { pua, bare };
  }
  if (format === 6) {
    if (available < 10) return null;
    const first = view.getUint16(offset + 6);
    const count = view.getUint16(offset + 8);
    if (10 + count * 2 > available) return null;
    for (let index = 0; index < count; index++) {
      if (view.getUint16(offset + 10 + index * 2) !== 0) add(first + index);
    }
    return { pua, bare };
  }
  if (format === 4) {
    if (available < 16) return null;
    const length = view.getUint16(offset + 2);
    const segX2 = view.getUint16(offset + 6);
    if (length < 16 || length > available || segX2 === 0 || segX2 % 2) return null;
    const segments = segX2 / 2;
    const ends = offset + 14;
    const starts = ends + segX2 + 2;
    const deltas = starts + segX2;
    const ranges = deltas + segX2;
    if (ranges + segX2 > offset + length) return null;
    for (let index = 0; index < segments; index++) {
      const start = view.getUint16(starts + index * 2);
      const end = view.getUint16(ends + index * 2);
      if (start > end) return null;
      const delta = view.getInt16(deltas + index * 2);
      const rangeAt = ranges + index * 2;
      const rangeOffset = view.getUint16(rangeAt);
      // Only codes 0x00..0xFF and U+F000..U+F0FF matter here.
      for (const [lo, hi] of [[0, 0xff], [0xf000, 0xf0ff]] as const) {
        for (let cp = Math.max(start, lo); cp <= Math.min(end, hi); cp++) {
          let glyph: number;
          if (rangeOffset === 0) glyph = (cp + delta) & 0xffff;
          else {
            const at = rangeAt + rangeOffset + (cp - start) * 2;
            if (at + 2 > offset + length) { glyph = 0; } else {
              const raw = view.getUint16(at);
              glyph = raw === 0 ? 0 : (raw + delta) & 0xffff;
            }
          }
          if (glyph !== 0) add(cp);
        }
      }
    }
    return { pua, bare };
  }
  if (format === 12 || format === 13) {
    if (available < 16) return null;
    const length = view.getUint32(offset + 4);
    const groups = view.getUint32(offset + 12);
    if (length > available || groups > (length - 16) / 12) return null;
    for (let index = 0; index < groups; index++) {
      const group = offset + 16 + index * 12;
      const start = view.getUint32(group);
      const end = view.getUint32(group + 4);
      const startGlyph = view.getUint32(group + 8);
      for (const [lo, hi] of [[0, 0xff], [0xf000, 0xf0ff]] as const) {
        for (let cp = Math.max(start, lo); cp <= Math.min(end, hi); cp++) {
          const glyph = format === 13 ? startGlyph : startGlyph + (cp - start);
          if (glyph !== 0) add(cp);
        }
      }
    }
    return { pua, bare };
  }
  return null;
}

/**
 * Exact encoding of one face from its sfnt bytes (TrueType/OpenType, or one
 * face of a collection by `faceIndex`). Null when the bytes cannot be read.
 *
 * OpenType cmap: a font whose only subtables are Unicode ones (platform 0, or
 * 3/1, 3/10) that agree on the PUA draws a U+F0xx code only where they map it
 * and never draws a bare code as a symbol glyph. A font whose only subtable
 * of those is the symbol one (3/0) is read through it: shapers resolve a PUA
 * code through it, and a bare code 0x20..0xFF through either its own entry
 * or the F0-shifted one (HarfBuzz, DirectWrite, Core Text and Firefox all
 * apply that symbol remapping). Any other combination returns null.
 */
export function symbolCmapEncoding(bytes: Uint8Array, faceIndex?: number): SymbolFaceEncoding | null {
  try {
    if (bytes.byteLength < 12) return null;
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let sfnt = 0;
    if (view.getUint32(0) === tag('ttcf')) {
      if (faceIndex === undefined || faceIndex < 0 || faceIndex >= view.getUint32(8)) return null;
      sfnt = view.getUint32(12 + faceIndex * 4);
    } else if (faceIndex !== undefined && faceIndex !== 0) return null;
    const tables = view.getUint16(sfnt + 4);
    let cmap = -1;
    let cmapLength = 0;
    for (let index = 0; index < tables; index++) {
      const record = sfnt + 12 + index * 16;
      if (view.getUint32(record) === tag('cmap')) {
        cmap = view.getUint32(record + 8);
        cmapLength = view.getUint32(record + 12);
      }
    }
    if (cmap < 0 || cmap + cmapLength > bytes.byteLength || cmapLength < 4) return null;
    const records = view.getUint16(cmap + 2);
    if (4 + records * 8 > cmapLength) return null;
    // Per Unicode subtable: its PUA codes, and the bare codes it draws as the
    // Unicode characters (Latin) they are (`latin`, used when several served
    // faces share a family: such a face, not a later one, draws that code).
    const unicode: { pua: string; latin: string }[] = [];
    let symbol: { pua: Set<number>; bare: Set<number> } | null = null;
    for (let index = 0; index < records; index++) {
      const record = cmap + 4 + index * 8;
      const platform = view.getUint16(record);
      const encoding = view.getUint16(record + 2);
      const relative = view.getUint32(record + 4);
      if (relative >= cmapLength) return null;
      const isUnicode = platform === 0 ? encoding !== 5 : platform === 3 && (encoding === 1 || encoding === 10);
      const isSymbol = platform === 3 && encoding === 0;
      if (!isUnicode && !isSymbol) continue;
      const codes = subtableCodes(view, cmap + relative, cmapLength - relative);
      if (!codes) return null;
      if (isUnicode) {
        unicode.push({ pua: symbolCodeSet(codes.pua), latin: symbolCodeSet(codes.bare) });
        continue;
      }
      symbol ??= { pua: new Set(), bare: new Set() };
      codes.pua.forEach((code) => symbol!.pua.add(code));
      codes.bare.forEach((code) => symbol!.bare.add(code));
    }
    // Which subtable an engine reads is its choice (HarfBuzz, FreeType and
    // Core Text rank Unicode and symbol subtables differently), so a face
    // whose subtables disagree is not judged from its bytes: null, and the
    // caller measures it in its painting context instead.
    if (unicode.length > 0) {
      if (symbol || unicode.some((set) => set.pua !== unicode[0]!.pua || set.latin !== unicode[0]!.latin)) return null;
      return Object.freeze({ kind: 'known', pua: unicode[0]!.pua, bare: symbolCodeSet([]), latin: unicode[0]!.latin });
    }
    if (!symbol) return SYMBOL_FACE_UNICODE;
    return Object.freeze({
      kind: 'known',
      pua: symbolCodeSet(symbol.pua),
      bare: symbolCodeSet([...symbol.bare, ...symbol.pua].filter((code) => code >= 0x20)),
    });
  } catch {
    return null;
  }
}

/** The part of a 2D context the measurement needs. */
export interface SymbolProbeContext {
  font: string;
  direction?: CanvasDirection;
  letterSpacing?: string;
  measureText(text: string): Pick<TextMetrics, 'width'> & Partial<Pick<TextMetrics,
    'actualBoundingBoxLeft' | 'actualBoundingBoxRight' | 'actualBoundingBoxAscent' | 'actualBoundingBoxDescent'>>;
}

/** No host installs this family; listing it last ends the probe font. */
const ABSENT_FAMILY = '__ooxml_symbol_probe_absent_7c41__';
const PROBE_PX = 64;

function geometry(context: SymbolProbeContext, text: string): number[] {
  const m = context.measureText(text);
  return [m.width, m.actualBoundingBoxLeft ?? 0, m.actualBoundingBoxRight ?? 0,
    m.actualBoundingBoxAscent ?? 0, m.actualBoundingBoxDescent ?? 0];
}
const same = (a: readonly number[], b: readonly number[]) => a.every((value, index) => Object.is(value, b[index]));

/** DELETE (U+007F) and SOFT HYPHEN (U+00AD) are default-ignorable as bare
 * characters: engines draw nothing for them even in a symbol-encoded face, so
 * they cannot show the bare/PUA identity (in Chromium every other code
 * 0x21..0xFF does, C1 controls and NO-BREAK SPACE included). Their ownership
 * is therefore inconclusive, and inconclusive keeps the authored code: in a
 * face proven symbol-encoded by the other codes, their PUA forms stay in the
 * face (as in 0.89, where no table but Symbol's 0xAD converted them); their
 * bare forms are never owned. A face whose bytes the library holds is read
 * exactly instead. The space 0x20 needs no judgement: no table converts it. */
const identityProbe = (code: number) => code !== 0x7f && code !== 0xad;

/**
 * Measure which codes the face `familyText` (CSS family syntax, the selected
 * face only) draws, in the context that measures and paints the text.
 *
 * A face draws code c as a symbol glyph exactly when the shaper resolves both
 * the bare code and U+F000+c to the same glyph of that face: that identity is
 * the symbol-cmap remapping noted at {@link symbolCmapEncoding}, and a
 * Unicode-encoded or missing face draws a Latin character for the bare code
 * and a fallback (or nothing) for the PUA code. Both are measured with
 * the face followed by an absent family, and the PUA geometry must also
 * differ from the absent family's alone, so a face missing both is not
 * mistaken for one that owns both. Geometry is advance plus ink box.
 * Comparing against an absent family's fallback alone is not evidence: font
 * fallback depends on the primary face (Core Text cascades), so two fallback
 * draws can differ although neither is the face's.
 *
 * A context whose text geometry does not vary with the text (a stub context,
 * a canvas without fonts) is inconclusive: `unknown`.
 */
export function measureSymbolFaceEncoding(
  context: SymbolProbeContext | null | undefined,
  familyText: string,
  weight = 400,
  style: 'normal' | 'italic' = 'normal',
): SymbolFaceEncoding {
  if (!context) return SYMBOL_FACE_UNKNOWN;
  const previous = context.font;
  const previousDirection = context.direction;
  const previousSpacing = context.letterSpacing;
  try {
    // A bare code such as "(" is mirrored in a right-to-left context while
    // its PUA form is not; spacing is no glyph property. Measure each alone,
    // left to right, without spacing.
    if (previousDirection !== undefined) context.direction = 'ltr';
    if (previousSpacing !== undefined) context.letterSpacing = '0px';
    const prefix = `${style === 'italic' ? 'italic ' : ''}${weight} ${PROBE_PX}px `;
    const absent = `${prefix}"${ABSENT_FAMILY}"`;
    context.font = absent;
    const control = [geometry(context, 'i'), geometry(context, 'W')];
    if (!control.every((g) => g.every(Number.isFinite)) || same(control[0]!, control[1]!)) return SYMBOL_FACE_UNKNOWN;
    const absentPua: number[][] = [];
    for (let code = 0; code < 256; code++) absentPua.push(geometry(context, String.fromCharCode(0xf000 + code)));
    context.font = `${prefix}${familyText}, "${ABSENT_FAMILY}"`;
    const owned: number[] = [];
    for (let code = 0x21; code < 256; code++) {
      if (!identityProbe(code)) continue;
      const shifted = geometry(context, String.fromCharCode(0xf000 + code));
      if (!shifted.every(Number.isFinite)) return SYMBOL_FACE_UNKNOWN;
      if (!same(shifted, absentPua[code]!) && same(shifted, geometry(context, String.fromCharCode(code)))) {
        owned.push(code);
      }
    }
    if (owned.length === 0) return SYMBOL_FACE_UNICODE;
    return Object.freeze({
      kind: 'known', pua: symbolCodeSet([...owned, 0x7f, 0xad]), bare: symbolCodeSet(owned),
    });
  } catch {
    return SYMBOL_FACE_UNKNOWN;
  } finally {
    try {
      context.font = previous;
      if (previousDirection !== undefined) context.direction = previousDirection;
      if (previousSpacing !== undefined) context.letterSpacing = previousSpacing;
    } catch { /* restored best-effort */ }
  }
}

interface ProbeCanvas {
  getContext(type: '2d', options?: CanvasRenderingContext2DSettings): unknown;
}

/** A scratch canvas in the realm (and so the FontFaceSet) of `owner`. */
function scratchCanvas(owner: unknown): ProbeCanvas | null {
  try {
    const element = (owner as { canvas?: { ownerDocument?: Document } } | null)?.canvas;
    const doc = element?.ownerDocument;
    if (doc && typeof doc.createElement === 'function') {
      const canvas = doc.createElement('canvas');
      canvas.width = 48;
      canvas.height = 48;
      return canvas;
    }
    if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(48, 48);
  } catch { /* no scratch canvas */ }
  return null;
}

/**
 * Whether `familyList` paints `text` without colour in the realm of `owner`
 * (a measuring/painting 2D context): drawn in opaque black, a monochrome
 * glyph leaves only black pixels, while a colour emoji face paints its own
 * colours. Undefined (inconclusive) when no scratch canvas exists, it cannot
 * be read back, or the drawing left no ink to judge.
 */
export function createSymbolMonochromeProbe(owner: unknown, familyList: string): (text: string) => boolean | undefined {
  let ctx: CanvasRenderingContext2D | null | undefined;
  let created = false;
  const results = new Map<string, boolean | undefined>();
  return (text: string) => {
    if (results.has(text)) return results.get(text);
    if (!created) {
      created = true;
      try {
        ctx = scratchCanvas(owner)?.getContext('2d', { willReadFrequently: true }) as CanvasRenderingContext2D | null | undefined;
      } catch { ctx = null; }
    }
    let monochrome: boolean | undefined;
    if (ctx) {
      try {
        ctx.clearRect(0, 0, 48, 48);
        ctx.globalAlpha = 1;
        ctx.globalCompositeOperation = 'source-over';
        ctx.fillStyle = '#000';
        ctx.textBaseline = 'middle';
        ctx.font = `32px ${familyList}`;
        ctx.fillText(text, 4, 24);
        const data = ctx.getImageData(0, 0, 48, 48).data;
        let ink = false;
        let colour = false;
        for (let index = 0; index < data.length; index += 4) {
          if (data[index + 3]! === 0) continue;
          ink = true;
          if (data[index]! + data[index + 1]! + data[index + 2]! > 24) { colour = true; break; }
        }
        monochrome = ink ? !colour : undefined;
      } catch {
        monochrome = undefined;
      }
    }
    if (results.size >= 1024) results.delete(results.keys().next().value!);
    results.set(text, monochrome);
    return monochrome;
  };
}

/** One face selection a format asks about: the CSS family that paints (the
 * selected face, first in its family list) and, when the library holds that
 * face's bytes, their exact encoding. */
export interface SymbolFaceSelection {
  readonly familyText: string;
  readonly weight: number;
  readonly style: 'normal' | 'italic';
  readonly encoding?: SymbolFaceEncoding | null;
}

/**
 * Per-owner decisions for one set of font resources in one context. A format
 * creates a new oracle whenever its font resources change (and per render
 * context), which is its invalidation: measurements are cached only here.
 */
export interface SymbolFaceOracle {
  encoding(selection: SymbolFaceSelection): SymbolFaceEncoding;
  segments(text: string, fontFamily: string | null | undefined, selection: SymbolFaceSelection): SymbolTextSegment[];
}

export function createSymbolFaceOracle(
  context: SymbolProbeContext | null | undefined,
  fallbackFamilyList = 'sans-serif',
  /** The real 2D context when `context` wraps it: colour probes draw on a
   * scratch canvas of its realm. Defaults to `context`. */
  probeOwner: unknown = context,
): SymbolFaceOracle {
  const encodings = new Map<string, SymbolFaceEncoding>();
  const monochrome = createSymbolMonochromeProbe(probeOwner ?? context, fallbackFamilyList);
  const options = (encoding: SymbolFaceEncoding): SymbolTextOptions => ({ encoding, monochrome });
  const encoding = (selection: SymbolFaceSelection): SymbolFaceEncoding => {
    if (selection.encoding) return selection.encoding;
    const key = JSON.stringify([selection.familyText, selection.weight, selection.style]);
    let result = encodings.get(key);
    if (!result) {
      result = measureSymbolFaceEncoding(context, selection.familyText, selection.weight, selection.style);
      if (encodings.size >= 256) encodings.delete(encodings.keys().next().value!);
      encodings.set(key, result);
    }
    return result;
  };
  return Object.freeze({
    encoding,
    segments(text: string, fontFamily: string | null | undefined, selection: SymbolFaceSelection) {
      const judged = encoding(selection);
      const segments = symbolTextToUnicodeSegments(text, fontFamily, options(judged));
      const drawsNone = judged.kind === 'known' && !/[^0]/u.test(judged.pua) && !/[^0]/u.test(judged.bare);
      return drawsNone
        ? segments.map((part) => (part.mapped ? { ...part, unicodeFace: true as const } : part))
        : segments;
    },
  });
}

/** A map key for one face tuple's held encoding (CSS family, weight, style). */
export function symbolFaceKey(family: string, weight: number, style: 'normal' | 'italic'): string {
  return `${family.trim().toLowerCase()}|${weight}|${style}`;
}

/**
 * The encoding of a served family's faces for one weight/style, when every
 * face CSS matching can pick for that tuple has a read cmap (CSS Fonts 4
 * §4.5, §5.3: faces whose `unicode-range` covers a scalar are tried, the
 * last-declared first, until one has a glyph for it).
 *   - A PUA code is owned when any covering face's cmap draws it.
 *   - A bare code is owned when the first covering face that has a glyph for
 *     it draws it as a symbol code; a Unicode face that maps it draws the
 *     Unicode character instead.
 * Undefined when no face matches the tuple exactly (matching then picks a
 * nearer face or synthesizes; the caller measures instead) or a matching
 * face's bytes were not read.
 */
export function symbolEncodingOfFaces(
  faces: readonly Readonly<{
    weight: string; style: string; stretch?: string; unicodeRange: string; symbolEncoding?: SymbolFaceEncoding;
  }>[],
  weight: number,
  style: 'normal' | 'italic',
): SymbolFaceEncoding | undefined {
  // Text is requested at the normal width; a face of another stretch is
  // matched only when the family has no normal one, which is left to measurement.
  const normalWidth = (stretch: string | undefined) => !stretch || /^(?:normal|100%)$/iu.test(stretch.trim());
  const matching = faces.filter((face) => fontFaceWeightCovers(face.weight, weight)
    && face.style.trim().toLowerCase() === style && normalWidth(face.stretch));
  if (matching.length === 0) return undefined;
  const ranges: (readonly (readonly [number, number])[])[] = [];
  const encodings: Extract<SymbolFaceEncoding, { kind: 'known' }>[] = [];
  for (const face of matching) {
    const parsed = parseUnicodeRange(face.unicodeRange);
    if (face.symbolEncoding?.kind !== 'known' || !parsed) return undefined;
    ranges.push(parsed);
    encodings.push(face.symbolEncoding);
  }
  const pua: number[] = [];
  const bare: number[] = [];
  const covering = (scalar: number) => encodings.filter((_, index) =>
    ranges[index]!.some(([lo, hi]) => lo <= scalar && scalar <= hi));
  for (let code = 0; code < 256; code++) {
    if (covering(0xf000 + code).some((encoding) => symbolCodeSetHas(encoding.pua, code))) pua.push(code);
    // The last-declared covering face that draws the bare code at all decides:
    // as its symbol code (owned), or as the Unicode character (not owned).
    for (const encoding of covering(code).reverse()) {
      if (symbolCodeSetHas(encoding.bare, code)) { bare.push(code); break; }
      if (symbolCodeSetHas(encoding.latin, code)) break;
    }
  }
  return Object.freeze({ kind: 'known', pua: symbolCodeSet(pua), bare: symbolCodeSet(bare) });
}
