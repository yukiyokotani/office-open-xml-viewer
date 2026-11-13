// ── WMF text output (META_TEXTOUT / META_EXTTEXTOUT) ────────────────────────
//
// [MS-WMF] references (v18.0):
//   2.3.3.5  META_EXTTEXTOUT   Y, X, StringLength, fwOpts, optional Rectangle
//                              (present with ETO_OPAQUE or ETO_CLIPPED), String
//                              padded to 16 bits, optional Dx array.
//   2.3.3.20 META_TEXTOUT      StringLength, String (padded), YStart, XStart.
//   2.2.1.2  Font object       Height, Width, Escapement, Orientation, Weight,
//                              Italic, Underline, StrikeOut, CharSet, …,
//                              PitchAndFamily, FaceName (NUL-terminated, at
//                              most 32 bytes including the NUL).
//   2.1.1.5  CharacterSet, 2.1.2.2 ExtTextOutOptions, 2.1.2.3 TextAlignmentMode,
//   2.1.1.8  FamilyFont / 2.1.1.24 PitchFont,
//   2.3.5.11 META_SETBKCOLOR, 2.3.5.15 META_SETBKMODE, 2.3.5.24 META_SETTEXTALIGN,
//   2.3.5.25 META_SETTEXTCHAREXTRA, 2.3.5.26 META_SETTEXTCOLOR,
//   2.3.5.27 META_SETTEXTJUSTIFICATION.
// WMF playback has no SETGRAPHICSMODE, so text is always GM_COMPATIBLE: the
// LOGFONT documentation states that lfEscapement then specifies both the
// escapement and the orientation (lfOrientation is not used), and glyphs stay
// upright on the device regardless of a mirrored logical mapping.
//
// Separately positioned characters (Dx, CharExtra) are drawn as cells in
// UAX #9 visual order; cursive scripts keep their contextual forms through
// zero-width joiners chosen from the Unicode Joining_Type of each neighbour.
//
// Character spacing follows the GDI ExtTextOut model: an authored Dx array
// replaces the font's natural advances; SETTEXTCHAREXTRA is added to every
// character and SETTEXTJUSTIFICATION's BreakExtra is shared across the break
// (space) characters on display devices.
//
// Fail-closed (reported through the player's incomplete-metafile channel) only
// where neither MS-WMF nor the stored bytes define the result: no selected
// font (the playback DC default font is device-dependent, MS-WMF 3.1.5), a
// zero font height (font-mapper default), an undefined TA_UPDATECP current
// position, glyph-index strings (ETO_GLYPH_INDEX names glyphs of a font
// program we do not have), '@' vertical faces (font-mapper behaviour),
// undefined alignment codes, DEFAULT/OEM charsets with non-ASCII bytes (the
// code page is the playback system's locale), JOHAB (no decoder available),
// bytes undefined in the selected code page, and malformed records.

import { findReferenceFontMetrics } from '../fonts/reference-font-metrics.js';
import {
  hasMetafileSymbolRepertoire,
  metafileSymbolCodeToUnicode,
  SYMBOL_CHARSET_SUBSTITUTE_FACE,
} from '../fonts/metafile-symbol-encoding.js';
import { officeFaceWebAlias } from '../fonts/font-stack.js';
import { classifyFontGeneric, cjkFallbackChain, type CjkLang } from '../fonts/scripts.js';
import { quoteCssFontFamily, scopeServedFontFamilies } from '../internal/css-font-family.js';
// Direct bidi modules, not the bidi index: importing the index from this
// eagerly loaded player made Rolldown treat single-entry IIFE builds of the
// renderers as code-splitting builds.
import { getDefaultBidiEngine } from '../text/bidi/engine.js';
import { buildVisualOrder, hasStrongRtl } from '../text/bidi/line-order.js';
import { REMOVED_LEVEL } from '../text/bidi/types.js';
import { joiningType, joinsToFollowing, joinsToPreceding } from '../text/joining-type.js';
import { referenceFontDecoration } from '../text/font-decoration.js';
import { DECORATION_LINE_WIDTH_EM, STRIKE_RISE_EM } from '../text/underline.js';

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/** MS-WMF 2.2.1.2 Font object, as authored by META_CREATEFONTINDIRECT. */
export interface WmfFont {
  kind: 'font';
  /** Signed Height: < 0 character (em) height, > 0 cell height, 0 default. */
  height: number;
  width: number;
  escapement: number;
  orientation: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strikeout: boolean;
  charset: number;
  pitchAndFamily: number;
  face: string;
  /** False when the record is too short or FaceName has no NUL terminator. */
  valid: boolean;
}

// ExtTextOutOptions (MS-WMF 2.1.2.2).
export const ETO_OPAQUE = 0x0002;
export const ETO_CLIPPED = 0x0004;
const ETO_GLYPH_INDEX = 0x0010;
const ETO_RTLREADING = 0x0080;
const ETO_PDY = 0x2000;
// TextAlignmentMode (MS-WMF 2.1.2.3).
export const TA_UPDATECP = 0x0001;
const TA_RIGHT = 0x0002;
const TA_CENTER = 0x0006;
const TA_BOTTOM = 0x0008;
const TA_BASELINE = 0x0018;
const TA_RTLREADING = 0x0100;
// BackgroundMode (MS-WMF 2.1.1.4).
const OPAQUE = 0x0002;

// MS-WMF 2.3.3.5 String decoding table: CharSet → code page. Charsets
// absent from it (DEFAULT_CHARSET, MAC_CHARSET, …) decode undefined.
const ANSI_CHARSET = 0; // 1252
const SYMBOL_CHARSET = 2; // 42 (CP_SYMBOL: byte b → U+F000 + b)
const JOHAB_CHARSET = 130; // 1361
const OEM_CHARSET = 255; // 437
const SINGLE_BYTE_CODE_PAGES: Readonly<Record<number, string>> = {
  161: 'windows-1253', // GREEK_CHARSET
  162: 'windows-1254', // TURKISH_CHARSET
  163: 'windows-1258', // VIETNAMESE_CHARSET
  177: 'windows-1255', // HEBREW_CHARSET
  178: 'windows-1256', // ARABIC_CHARSET
  186: 'windows-1257', // BALTIC_CHARSET
  204: 'windows-1251', // RUSSIAN_CHARSET
  222: 'windows-874', // THAI_CHARSET
  238: 'windows-1250', // EASTEUROPE_CHARSET
};
interface DoubleByteCodePage { label: string; lead: (byte: number) => boolean; cjk: CjkLang }
const DOUBLE_BYTE_CODE_PAGES: Readonly<Record<number, DoubleByteCodePage>> = {
  // SHIFTJIS_CHARSET, code page 932: lead bytes 0x81–0x9F and 0xE0–0xFC.
  128: { label: 'shift_jis', lead: (b) => (b >= 0x81 && b <= 0x9f) || (b >= 0xe0 && b <= 0xfc), cjk: 'jp' },
  // HANGUL_CHARSET 949, GB2312_CHARSET 936, CHINESEBIG5_CHARSET 950: 0x81–0xFE.
  129: { label: 'euc-kr', lead: (b) => b >= 0x81 && b <= 0xfe, cjk: 'kr' },
  134: { label: 'gbk', lead: (b) => b >= 0x81 && b <= 0xfe, cjk: 'sc' },
  136: { label: 'big5', lead: (b) => b >= 0x81 && b <= 0xfe, cjk: 'tc' },
};

// WHATWG Encoding §9.1, index-windows-1252.txt (identifier e56d49d9,
// 2024-09-18). Decode directly so WMF ANSI text is deterministic even on
// hosts whose TextDecoder delegates this label to an ISO-8859-1 codec.
const WINDOWS_1252_C1 = [
  0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
  0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
  0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
] as const;

const decoders = new Map<string, TextDecoder | null>();
function decoderFor(label: string): TextDecoder | null {
  if (!decoders.has(label)) {
    let decoder: TextDecoder | null = null;
    try { decoder = new TextDecoder(label, { fatal: true }); } catch { decoder = null; }
    decoders.set(label, decoder);
  }
  return decoders.get(label) ?? null;
}

// Code page 437 bytes 0x80–0xFF (Unicode, Inc. MAPPINGS/VENDORS/MICSFT/PC/CP437.TXT).
const CP437_HIGH =
  'ÇüéâäàåçêëèïîìÄÅ' + 'ÉæÆôöòûùÿÖÜ¢£¥₧ƒ' + 'áíóúñÑªº¿⌐¬½¼¡«»' + '░▒▓│┤╡╢╖╕╣║╗╝╜╛┐'
  + '└┴┬├─┼╞╟╚╔╩╦╠═╬╧' + '╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀' + 'αßΓπΣσµτΦΘΩδ∞φε∩' + '≡±≥≤⌠⌡÷≈°∙·√ⁿ²■\u00a0';

// Code page 1361 (Johab, KS X 1001:1992 annex 3). Hangul is a 15-bit
// composition of initial/medial/final indices; symbols and hanja are the
// KS X 1001 repertoire rearranged two rows per lead byte.
const JOHAB_INITIAL = [-1, -1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18];
const JOHAB_MEDIAL = [-1, -1, -1, 0, 1, 2, 3, 4, -1, -1, 5, 6, 7, 8, 9, 10, -1, -1, 11, 12, 13, 14, 15, 16, -1, -1, 17, 18, 19, 20];
const JOHAB_FINAL = [-1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, -1, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27];
// Hangul Compatibility Jamo for a lone initial, medial or final.
const COMPAT_INITIAL = [0x3131, 0x3132, 0x3134, 0x3137, 0x3138, 0x3139, 0x3141, 0x3142, 0x3143, 0x3145, 0x3146, 0x3147, 0x3148, 0x3149, 0x314a, 0x314b, 0x314c, 0x314d, 0x314e];
const COMPAT_FINAL = [0, 0x3131, 0x3132, 0x3133, 0x3134, 0x3135, 0x3136, 0x3137, 0x3139, 0x313a, 0x313b, 0x313c, 0x313d, 0x313e, 0x313f, 0x3140, 0x3141, 0x3142, 0x3144, 0x3145, 0x3146, 0x3147, 0x3148, 0x314a, 0x314b, 0x314c, 0x314d, 0x314e];

function decodeJohab(bytes: Uint8Array): Decoded {
  const units: TextUnit[] = [];
  const ksx = decoderFor('euc-kr');
  for (let index = 0; index < bytes.length;) {
    const lead = bytes[index];
    if (lead < 0x80) {
      // Code page 1361 maps 0x5C to WON SIGN.
      units.push({ text: lead === 0x5c ? '\u20a9' : String.fromCharCode(lead), start: index, count: 1, face: 'authored' });
      index += 1;
      continue;
    }
    if (index + 1 >= bytes.length) return { problem: 'truncated double-byte character' };
    const trail = bytes[index + 1];
    let text: string | null = null;
    if (lead >= 0x84 && lead <= 0xd3 && ((trail >= 0x41 && trail <= 0x7e) || trail >= 0x81) && trail !== 0xff) {
      const code = (lead << 8) | trail;
      const i = JOHAB_INITIAL[(code >> 10) & 0x1f] ?? -1;
      const m = JOHAB_MEDIAL[(code >> 5) & 0x1f] ?? -1;
      const f = JOHAB_FINAL[code & 0x1f] ?? -1;
      const iFill = ((code >> 10) & 0x1f) === 1;
      const mFill = ((code >> 5) & 0x1f) === 2;
      if (i >= 0 && m >= 0 && f >= 0) text = String.fromCharCode(0xac00 + (i * 21 + m) * 28 + f);
      else if (iFill && mFill && f === 0) text = '\u3164';
      else if (i >= 0 && mFill && f === 0) text = String.fromCharCode(COMPAT_INITIAL[i]);
      else if (iFill && m >= 0 && f === 0) text = String.fromCharCode(0x314f + m);
      else if (iFill && mFill && f > 0) text = String.fromCharCode(COMPAT_FINAL[f]);
    } else if (((lead >= 0xd9 && lead <= 0xde) || (lead >= 0xe0 && lead <= 0xf9))
      && ((trail >= 0x31 && trail <= 0x7e) || (trail >= 0x91 && trail <= 0xfe))) {
      const high = trail >= 0xa1;
      const row = (lead <= 0xde ? (lead - 0xd9) * 2 + 0x21 : (lead - 0xe0) * 2 + 0x4a) + (high ? 1 : 0);
      const column = high ? trail - 0x80 : trail <= 0x7e ? trail - 0x10 : trail - 0x22;
      try { text = ksx?.decode(new Uint8Array([row | 0x80, column | 0x80])) ?? null; } catch { text = null; }
    }
    if (text === null) return { problem: 'byte undefined in code page' };
    units.push({ text, start: index, count: 2, face: 'authored' });
    index += 2;
  }
  return { units };
}

/** One drawn character: its Unicode text, the String bytes it came from (Dx
 *  entries are per byte), and the face class it must be drawn with. */
interface TextUnit {
  text: string;
  start: number;
  count: number;
  /** 'authored' = the selected face; 'symbol' = a symbol code converted to
   *  Unicode, drawn in a Unicode fallback (the symbol face would reinterpret it);
   *  'substitute' = a font code drawn in the font mapper's symbol substitute. */
  face: 'authored' | 'symbol' | 'substitute';
}

type Decoded = { units: TextUnit[] } | { problem: string };

/** Decode a byte string under a CharacterSet, or report why it is undefined.
 *  `hostFaceCovers` tells whether the authored face on this host supplies
 *  glyphs for a string (only consulted for symbol faces). */
function decodeBytes(
  bytes: Uint8Array,
  charset: number,
  face: string,
  hostFaceCovers: (text: string) => boolean,
): Decoded {
  const units: TextUnit[] = [];
  if (charset === SYMBOL_CHARSET) {
    // Symbol fonts are addressed through the F000 private range of their
    // symbol cmap. A face the host lacks is replaced by the font mapper's
    // symbol-charset substitute. A published repertoire converts a code to
    // Unicode; any other code stays font-coded in the face that draws it.
    const fontCoded = Array.from(bytes, (byte) => String.fromCharCode(0xf000 + byte)).join('');
    const substituted = !hasMetafileSymbolRepertoire(face) && bytes.length > 0 && !hostFaceCovers(fontCoded);
    const repertoire = substituted ? SYMBOL_CHARSET_SUBSTITUTE_FACE : face;
    bytes.forEach((byte, index) => {
      const mapped = metafileSymbolCodeToUnicode(byte, repertoire);
      units.push(mapped === undefined
        ? { text: fontCoded[index], start: index, count: 1, face: substituted ? 'substitute' : 'authored' }
        : { text: mapped, start: index, count: 1, face: 'symbol' });
    });
    return { units };
  }
  if (charset === ANSI_CHARSET) {
    bytes.forEach((byte, index) => units.push({
      text: String.fromCharCode(byte >= 0x80 && byte <= 0x9f ? WINDOWS_1252_C1[byte - 0x80] : byte),
      start: index, count: 1, face: 'authored',
    }));
    return { units };
  }
  if (charset === OEM_CHARSET) {
    bytes.forEach((byte, index) => units.push({
      text: String.fromCharCode(byte < 0x80 ? byte : CP437_HIGH.charCodeAt(byte - 0x80)),
      start: index, count: 1, face: 'authored',
    }));
    return { units };
  }
  if (charset === JOHAB_CHARSET) return decodeJohab(bytes);
  const single = SINGLE_BYTE_CODE_PAGES[charset];
  const double = DOUBLE_BYTE_CODE_PAGES[charset];
  if (!single && !double) return { problem: 'undefined character set' };
  const decoder = decoderFor(single ?? double!.label);
  if (!decoder) return { problem: 'character set decoder unavailable' };
  for (let index = 0; index < bytes.length;) {
    const count = double && double.lead(bytes[index]) ? 2 : 1;
    if (index + count > bytes.length) return { problem: 'truncated double-byte character' };
    let text: string;
    try { text = decoder.decode(bytes.subarray(index, index + count)); } catch { return { problem: 'byte undefined in code page' }; }
    units.push({ text, start: index, count, face: 'authored' });
    index += count;
  }
  return { units };
}

/** MS-WMF 2.2.1.2: decode a CREATEFONTINDIRECT parameter block. */
export function parseWmfFont(params: Uint8Array): WmfFont {
  const dv = new DataView(params.buffer, params.byteOffset, params.byteLength);
  const fixed = params.byteLength >= 18;
  const i16 = (offset: number) => (fixed ? dv.getInt16(offset, true) : 0);
  const faceBytes = params.subarray(18, Math.min(params.byteLength, 50));
  const nul = faceBytes.indexOf(0);
  const charset = fixed ? params[13] : 0;
  const rawFace = faceBytes.subarray(0, nul < 0 ? faceBytes.length : nul);
  // FaceName is stored in the font's own code page for DBCS faces; decode it
  // with the CharacterSet's decoder when one applies, else as Latin-1.
  const labelled = DOUBLE_BYTE_CODE_PAGES[charset]?.label ?? SINGLE_BYTE_CODE_PAGES[charset];
  let face = String.fromCharCode(...rawFace);
  if (labelled) {
    try { face = decoderFor(labelled)?.decode(rawFace) ?? face; } catch { /* keep Latin-1 */ }
  }
  return {
    kind: 'font',
    height: i16(0),
    width: i16(2),
    escapement: i16(4),
    orientation: i16(6),
    weight: i16(8),
    italic: fixed && params[10] !== 0,
    underline: fixed && params[11] !== 0,
    strikeout: fixed && params[12] !== 0,
    charset,
    pitchAndFamily: fixed ? params[17] : 0,
    face,
    valid: fixed && nul > 0,
  };
}

/** CSS generic family from PitchAndFamily (MS-WMF 2.2.2.14: family in bits
 *  4–7, FamilyFont 2.1.1.8; pitch in bits 0–1, PitchFont 2.1.1.24). An
 *  FF_DONTCARE face falls back to the shared name classifier. */
function genericFamily(font: WmfFont): string {
  switch (font.pitchAndFamily & 0xf0) {
    case 0x10: return 'serif'; // FF_ROMAN
    case 0x20: return 'sans-serif'; // FF_SWISS
    case 0x30: return 'monospace'; // FF_MODERN
    case 0x40: return 'cursive'; // FF_SCRIPT
    case 0x50: return 'fantasy'; // FF_DECORATIVE
    default: break;
  }
  if ((font.pitchAndFamily & 0x3) === 1) return 'monospace'; // FIXED_PITCH
  const generic = classifyFontGeneric(font.face);
  return generic === 'mono' ? 'monospace' : generic === 'serif' ? 'serif' : 'sans-serif';
}

function authoredStack(font: WmfFont, generic: string): string {
  const names = [font.face];
  const alias = officeFaceWebAlias(font.face);
  if (alias) names.push(alias);
  const cjk = DOUBLE_BYTE_CODE_PAGES[font.charset]?.cjk;
  if (cjk) names.push(...cjkFallbackChain(cjk, generic === 'serif' ? 'serif' : 'sans'));
  return `${names.map(quoteCssFontFamily).join(', ')}, ${generic}`;
}

/** Converted symbol text uses a Unicode face of the authored family class.
 *  An unclassified symbol face keeps the serif class of the published Symbol
 *  repertoire (its Greek and mathematical glyphs are a serif design). */
function symbolStack(font: WmfFont): string {
  return (font.pitchAndFamily & 0xf0) === 0 ? 'serif' : genericFamily(font);
}

interface FaceMetrics {
  /** GDI tmAscent / tmDescent as fractions of the em. */
  ascentEm: number;
  descentEm: number;
  /** tmAveCharWidth as a fraction of the em (evaluated only for a non-zero Width). */
  avgCharWidthEm: () => number;
}

// OpenType OS/2 xAvgCharWidth (versions 0–2): weighted lowercase + space.
const AVG_WIDTH_WEIGHTS: ReadonlyArray<readonly [string, number]> = [
  ['a', 64], ['b', 14], ['c', 27], ['d', 35], ['e', 100], ['f', 20], ['g', 14],
  ['h', 42], ['i', 63], ['j', 3], ['k', 6], ['l', 35], ['m', 20], ['n', 56],
  ['o', 56], ['p', 17], ['q', 4], ['r', 49], ['s', 56], ['t', 71], ['u', 31],
  ['v', 10], ['w', 18], ['x', 3], ['y', 18], ['z', 2], [' ', 166],
];

const PROBE_EM = 1000;

/** GDI realizes TrueType text metrics from OS/2 usWinAscent/usWinDescent and
 *  xAvgCharWidth. Use the catalogued values for the authored face; otherwise
 *  measure the face Canvas actually selects on this host. */
function faceMetrics(ctx: AnyCtx, font: WmfFont, css: (px: number) => string, weight: number): FaceMetrics | null {
  const style = font.italic ? 'italic' : 'normal';
  const profile = findReferenceFontMetrics(font.face, { weight, style })[0]
    ?? findReferenceFontMetrics(font.face, { style })[0]
    ?? findReferenceFontMetrics(font.face)[0];
  if (profile && profile.unitsPerEm > 0) {
    const upm = profile.unitsPerEm;
    const [ascent, descent] = profile.win ?? [profile.hhea[0], -profile.hhea[1]];
    const avg = profile.xAvgCharWidth;
    return {
      ascentEm: ascent / upm,
      descentEm: descent / upm,
      avgCharWidthEm: () => (avg && avg > 0 ? avg / upm : measuredAverageWidth(ctx, css)),
    };
  }
  ctx.font = css(PROBE_EM);
  const probe = ctx.measureText('Hg');
  const ascent = probe.fontBoundingBoxAscent;
  const descent = probe.fontBoundingBoxDescent;
  if (!Number.isFinite(ascent) || !Number.isFinite(descent) || ascent + descent <= 0) return null;
  return {
    ascentEm: ascent / PROBE_EM,
    descentEm: descent / PROBE_EM,
    avgCharWidthEm: () => measuredAverageWidth(ctx, css),
  };
}

function measuredAverageWidth(ctx: AnyCtx, css: (px: number) => string): number {
  ctx.font = css(PROBE_EM);
  let total = 0;
  for (const [char, weight] of AVG_WIDTH_WEIGHTS) total += ctx.measureText(char).width * weight;
  return total / 1000 / PROBE_EM;
}

const ZWJ = '\u200d';

/** UAX #9 resolved level of each unit's first code unit (paragraph level for
 *  characters removed by X9). */
function resolvedUnitLevels(texts: readonly string[], rtl: boolean): number[] {
  const logical = texts.join('');
  if (!rtl && !hasStrongRtl(logical)) return texts.map(() => 0);
  const { levels, paragraphLevel } = getDefaultBidiEngine().computeLevels(logical, rtl ? 'rtl' : 'ltr');
  let offset = 0;
  return texts.map((text) => {
    const level = levels[offset];
    offset += text.length;
    return level === REMOVED_LEVEL ? paragraphLevel : level;
  });
}
const LAM = 0x0644;
// Lam followed by one of these alefs forms the obligatory lam-alef ligature.
const LAM_ALEF_PARTNERS = new Set([0x0622, 0x0623, 0x0625, 0x0627, 0x0671]);

export interface CursiveCell { first: number; last: number; joinBefore: boolean; joinAfter: boolean }

/** Split separately positioned characters into cells and decide, from the
 *  Unicode Joining_Type of each cell's base and its neighbours, on which sides
 *  the cell joins cursively (UAX #44 / ArabicShaping.txt: a join exists where
 *  the earlier letter joins to the following one and the later letter joins
 *  to the preceding one; Transparent marks are skipped). A cell keeps its
 *  Transparent marks and a lam-alef pair stays one ligature cell. Neighbours
 *  in a different face never join. */
export function cursiveCells(
  texts: readonly string[],
  fonts: readonly string[],
  levels: readonly number[] = texts.map(() => 0),
): CursiveCell[] {
  const cells: CursiveCell[] = [];
  const base = (index: number) => texts[index].codePointAt(0) ?? 0;
  const transparent = (index: number) => [...texts[index]].every((char) => joiningType(char.codePointAt(0)!) === 'T');
  // The cell's last non-Transparent character decides its following join.
  const lastBase = (cell: CursiveCell) => {
    let index = cell.last;
    while (index > cell.first && transparent(index)) index--;
    return index;
  };
  for (let index = 0; index < texts.length; index++) {
    const last = cells.at(-1);
    // Never merge across a bidi level boundary: each cell is placed by level.
    if (last && fonts[index] === fonts[last.first] && levels[index] === levels[last.first] && texts[index].length > 0
      && (transparent(index)
        || (lastBase(last) === last.first && base(last.first) === LAM && LAM_ALEF_PARTNERS.has(base(index))))) {
      last.last = index;
      continue;
    }
    cells.push({ first: index, last: index, joinBefore: false, joinAfter: false });
  }
  // Joins skip whole cells of Transparent characters (e.g. an LRM kept in its
  // own cell by a level boundary); every cell in between gets ZWJ context.
  const allTransparent = (cell: CursiveCell) => {
    for (let index = cell.first; index <= cell.last; index++) if (!transparent(index)) return false;
    return true;
  };
  for (let k = 1; k < cells.length; k++) {
    const after = cells[k];
    if (allTransparent(after)) continue;
    let j = k - 1;
    while (j >= 0 && fonts[cells[j].first] === fonts[after.first] && allTransparent(cells[j])) j--;
    if (j < 0 || fonts[cells[j].first] !== fonts[after.first]) continue;
    if (joinsToFollowing(base(lastBase(cells[j]))) && joinsToPreceding(base(after.first))) {
      for (let m = j; m < k; m++) cells[m].joinAfter = true;
      for (let m = j + 1; m <= k; m++) cells[m].joinBefore = true;
    }
  }
  return cells;
}

const MISSING_CONTROL_FAMILY = '__ooxml_wmf_missing_face_control_5b2e__';

/** Whether Canvas draws `text` from `face` itself on this host: compare its
 *  glyph geometry against a deliberately absent control family (CSS silently
 *  substitutes a missing family, so equal geometry means fallback). */
function hostFaceCovers(ctx: AnyCtx, face: string, text: string, families?: ReadonlyMap<string, string>): boolean {
  const geometry = (family: string) => {
    ctx.font = `${PROBE_EM}px ${family}${quoteCssFontFamily(MISSING_CONTROL_FAMILY)}, serif`;
    const m = ctx.measureText(text);
    return [m.width, m.actualBoundingBoxAscent, m.actualBoundingBoxDescent, m.actualBoundingBoxLeft, m.actualBoundingBoxRight];
  };
  const selected = geometry(`${scopeServedFontFamilies(quoteCssFontFamily(face), families)}, `);
  const missing = geometry('');
  return !selected.every((value, index) => Object.is(value, missing[index]));
}

/** The DC state text output reads; PlayState in wmf.ts satisfies it. */
export interface WmfTextState {
  ctx: AnyCtx;
  orgX: number; orgY: number; extX: number; extY: number;
  vpX: number; vpY: number; vpW: number; vpH: number;
  haveExt: boolean;
  curFont: WmfFont | null;
  textColor: string;
  bkColor: string;
  bkMode: number;
  textAlign: number;
  charExtra: number;
  breakCount: number;
  breakQuotientPx: number;
  breakErrorPx: number;
  currentX: number;
  currentY: number;
  currentPositionKnown: boolean;
  /** The document's served families and their private registered families
   *  (core `LoadedFontSources.families`): text paints with the faces the
   *  document's font sources loaded, as the document's own text does. */
  fontFamilies?: ReadonlyMap<string, string>;
}

export interface WmfTextOutput {
  x: number;
  y: number;
  bytes: Uint8Array;
  options: number;
  /** Rect (MS-WMF 2.2.2.18) Left, Top, Right, Bottom in logical units. */
  rect: readonly [number, number, number, number] | null;
  /** Dx (and with ETO_PDY interleaved Dy) values, one entry per String byte. */
  dx: Int16Array | null;
}

/** Draw one text output. Returns whether anything was painted, or the reason
 *  the output is undefined for this player. */
export function drawWmfText(s: WmfTextState, out: WmfTextOutput): { drew: boolean } | { problem: string } {
  const sx = s.vpW / s.extX;
  const sy = s.vpH / s.extY;
  if (!s.haveExt) return { problem: 'logical mapping undefined' };
  if (!(Number.isFinite(sx) && Number.isFinite(sy) && sx !== 0 && sy !== 0)) return { problem: 'degenerate mapping' };
  if (out.bytes.length === 0) {
    // An empty string draws nothing but the ETO_OPAQUE rectangle.
    const opaqued = out.rect !== null && (out.options & ETO_OPAQUE) !== 0 && paintOpaqueRect(s, out.rect);
    return { drew: opaqued };
  }
  const font = s.curFont;
  if (!font) return { problem: 'no font selected' };
  if (!font.valid) return { problem: 'malformed font' };
  if (font.height === 0) return { problem: 'default font height' };
  if (font.face.startsWith('@')) return { problem: 'vertical font' };
  if (out.options & ETO_GLYPH_INDEX) return { problem: 'glyph index string' };
  const horizontal = s.textAlign & TA_CENTER;
  const vertical = s.textAlign & TA_BASELINE;
  if (horizontal === 0x4 || vertical === 0x10) return { problem: 'undefined text alignment' };
  if ((s.textAlign & TA_UPDATECP) && !s.currentPositionKnown) return { problem: 'current position undefined' };

  const { ctx } = s;
  // GDI TextOut/ExtTextOut apply no pair kerning; measure and paint alike.
  if ('fontKerning' in ctx) ctx.fontKerning = 'none';
  const weight = font.weight === 0 ? 400 : Math.min(1000, Math.max(1, font.weight));
  const style = font.italic ? 'italic ' : '';
  const generic = genericFamily(font);
  const authored = scopeServedFontFamilies(authoredStack(font, generic), s.fontFamilies);
  const symbol = symbolStack(font);
  const substitute = scopeServedFontFamilies(`${quoteCssFontFamily(SYMBOL_CHARSET_SUBSTITUTE_FACE)}, ${symbol}`, s.fontFamilies);
  const css = (stack: string) => (px: number) => `${style}${weight} ${px}px ${stack}`;

  const decoded = decodeBytes(out.bytes, font.charset, font.face, (text) => hostFaceCovers(ctx, font.face, text, s.fontFamilies));
  if ('problem' in decoded) return decoded;
  const { units } = decoded;
  const metrics = faceMetrics(ctx, font, css(authored), weight);
  if (!metrics) return { problem: 'font metrics unavailable' };
  // Font height maps through the logical y scale (MS-WMF 2.2.1.2 Height).
  const emLogical = font.height < 0 ? -font.height : font.height / (metrics.ascentEm + metrics.descentEm);
  const emPx = emLogical * Math.abs(sy);
  if (!(emPx > 0 && Number.isFinite(emPx))) return { problem: 'degenerate font size' };
  const ascentPx = metrics.ascentEm * emPx;
  const descentPx = metrics.descentEm * emPx;
  const hScale = font.width === 0 ? 1 : (Math.abs(font.width) * Math.abs(sx)) / (metrics.avgCharWidthEm() * emPx);
  if (!(hScale > 0 && Number.isFinite(hScale))) return { problem: 'degenerate font width' };
  const fontFor = (unit: TextUnit) =>
    css(unit.face === 'symbol' ? symbol : unit.face === 'substitute' ? substitute : authored)(emPx);

  // Baseline direction: escapement in tenths of a degree (sense below).
  // Logical lengths along / across it scale with the mapping.
  const theta = (font.escapement / 10) * (Math.PI / 180);
  const cos = Math.cos(theta);
  const sin = Math.sin(theta);
  // TrueType rotation follows the orientation of the logical coordinate
  // system (MS-WMF 2.1.2.1 ClipPrecision, CLIP_LH_ANGLES; LOGFONT
  // lfEscapement): counterclockwise on the device when the logical axes have
  // the device's handedness (sx·sy > 0), clockwise when one axis is mirrored.
  // Glyphs are never mirrored in GM_COMPATIBLE, so only the sense changes.
  // `phi` is the Canvas rotation (y down) of the baseline.
  const handedness = Math.sign(sx * sy) || 1;
  const phi = theta === 0 ? 0 : -theta * handedness;
  const along = Math.hypot(cos * sx, sin * sy);
  const across = Math.hypot(sin * sx, cos * sy) * Math.sign(sy);

  // Advances along the baseline (device px) and perpendicular offsets.
  // CharExtra (MS-WMF 2.3.5.25) is transformed to device units and rounded;
  // the justification quotient and its carried error were rounded the same
  // way when META_SETTEXTJUSTIFICATION was played.
  const pdy = (out.options & ETO_PDY) !== 0;
  const charExtraPx = Math.round(s.charExtra * along);
  const justified = s.breakCount > 0 && (s.breakQuotientPx !== 0 || s.breakErrorPx !== 0);
  // Widths are measured lazily and once per distinct (font, text): a long
  // string costs at most one measurement per distinct character or run.
  const widths = new Map<string, number>();
  const measure = (text: string, unitFont: string) => {
    const key = `${unitFont}\u0000${text}`;
    let width = widths.get(key);
    if (width === undefined) {
      ctx.font = unitFont;
      width = ctx.measureText(text).width * hScale;
      widths.set(key, width);
    }
    return width;
  };

  // Draw groups. Natural advances: one Canvas call per run of same-face
  // units. Justification alone: one call per word, with the break
  // characters as their own cells. Dx or CharExtra: one cell per character.
  interface Group { text: string; font: string; pen: number; rise: number; advance: number; x: number; level: number; extra: number; brk: boolean; joinBefore: boolean; joinAfter: boolean }
  const perCharacter = out.dx !== null || charExtraPx !== 0;
  const groups: Group[] = [];
  let rise = 0;
  const breakExtraFor = (unit: TextUnit) => {
    if (!justified || unit.count !== 1 || out.bytes[unit.start] !== 0x20) return 0;
    let extra = s.breakQuotientPx;
    if (s.breakErrorPx > 0) { extra += 1; s.breakErrorPx -= 1; }
    return extra;
  };
  const rtlReading = (s.textAlign & TA_RTLREADING) !== 0 || (out.options & ETO_RTLREADING) !== 0;
  if (!perCharacter) {
    // Runs placed separately (justified words, face changes) must each be a
    // single bidi level, so their visual order can be resolved between them.
    const splitRuns = justified || units.some((unit) => fontFor(unit) !== fontFor(units[0]));
    const unitLevels = splitRuns ? resolvedUnitLevels(units.map((unit) => unit.text), rtlReading) : null;
    units.forEach((unit, index) => {
      const unitFont = fontFor(unit);
      const extra = breakExtraFor(unit);
      const isBreak = justified && unit.count === 1 && out.bytes[unit.start] === 0x20;
      const level = unitLevels?.[index] ?? 0;
      const last = groups.at(-1);
      const separate = (justified && (isBreak || last?.brk === true)) || (last !== undefined && last.level !== level);
      if (last && last.font === unitFont && !separate) last.text += unit.text;
      else groups.push({ text: unit.text, font: unitFont, pen: 0, rise: 0, advance: 0, x: 0, level, extra, brk: isBreak, joinBefore: false, joinAfter: false });
    });
    // A word split at a level boundary keeps its cursive joins (ZWJ context):
    // letters join across runs of Transparent characters (e.g. LRM), as in
    // UAX #44 Joining_Type, when every run in between shares the face.
    const nonTransparent = (text: string, fromEnd: boolean) =>
      (fromEnd ? [...text].reverse() : [...text]).find((char) => joiningType(char.codePointAt(0)!) !== 'T');
    for (let k = 1; k < groups.length; k++) {
      const after = groups[k];
      const right = nonTransparent(after.text, false);
      if (after.brk || right === undefined) continue;
      let j = k - 1;
      while (j >= 0 && !groups[j].brk && groups[j].font === after.font && nonTransparent(groups[j].text, true) === undefined) j--;
      if (j < 0 || groups[j].brk || groups[j].font !== after.font) continue;
      const left = nonTransparent(groups[j].text, true)!;
      if (joinsToFollowing(left.codePointAt(0)!) && joinsToPreceding(right.codePointAt(0)!)) {
        for (let m = j; m < k; m++) groups[m].joinAfter = true;
        for (let m = j + 1; m <= k; m++) groups[m].joinBefore = true;
      }
    }
  } else {
    // Character cells: a base character with its following Transparent marks
    // (Joining_Type T), and lam with a following alef as one ligature cell.
    const cells = cursiveCells(units.map((unit) => unit.text), units.map(fontFor),
      resolvedUnitLevels(units.map((unit) => unit.text), rtlReading));
    for (const cell of cells) {
      let advance = 0;
      let cellRise = 0;
      for (let k = cell.first; k <= cell.last; k++) {
        const unit = units[k];
        advance += charExtraPx + breakExtraFor(unit);
        if (out.dx) {
          for (let b = unit.start; b < unit.start + unit.count; b++) {
            advance += out.dx[pdy ? b * 2 : b] * along;
            if (pdy) cellRise += out.dx[b * 2 + 1] * across;
          }
        }
      }
      const font = fontFor(units[cell.first]);
      const text = units.slice(cell.first, cell.last + 1).map((unit) => unit.text).join('');
      // Natural widths are taken in the cell's own joining context.
      if (!out.dx) advance += measure(`${cell.joinBefore ? ZWJ : ''}${text}${cell.joinAfter ? ZWJ : ''}`, font);
      groups.push({ text, font, pen: 0, rise, advance, x: 0, level: 0, extra: 0, brk: false, joinBefore: cell.joinBefore, joinAfter: cell.joinAfter });
      rise += cellRise;
    }
  }
  const rtl = rtlReading;
  const useCpFlag = (s.textAlign & TA_UPDATECP) !== 0;
  const needTotal = horizontal !== 0 || rtl || useCpFlag || s.bkMode === OPAQUE || font.underline || font.strikeout;
  let pen = 0;
  if (perCharacter || needTotal || groups.length > 1) {
    for (const group of groups) {
      // Same-face runs are shaped as a whole (complex scripts, no kerning).
      if (!perCharacter) {
        group.advance = measure(`${group.joinBefore ? ZWJ : ''}${group.text}${group.joinAfter ? ZWJ : ''}`, group.font) + group.extra;
      }
      group.pen = pen;
      pen += group.advance;
    }
  }
  const total = pen;
  const totalRise = rise;

  // Visual order (UAX #9, as ExtTextOut lays out bidi runs): each cell keeps
  // its own logical advance; RTL runs are placed right to left as wholes,
  // so European digits inside Hebrew keep their left-to-right order.
  let order = groups.map((_, index) => index);
  if (groups.length > 1) {
    const logicalText = groups.map((group) => group.text).join('');
    if (rtl || hasStrongRtl(logicalText)) {
      const engine = getDefaultBidiEngine();
      const { levels, paragraphLevel } = engine.computeLevels(logicalText, rtl ? 'rtl' : 'ltr');
      const starts: number[] = [];
      let offset = 0;
      for (const group of groups) { starts.push(offset); offset += group.text.length; }
      const visual = buildVisualOrder(levels, paragraphLevel, starts);
      order = visual.order;
      groups.forEach((group, index) => {
        group.level = visual.segLevels[index];
        // UAX #9 L4: a lone mirrored character at an RTL level.
        const chars = [...group.text];
        const mirrored = chars.length === 1 && (group.level & 1) ? engine.getMirror(chars[0].codePointAt(0)!) : null;
        if (mirrored !== null) group.text = String.fromCodePoint(mirrored);
      });
    }
  }
  {
    let x = 0;
    for (const index of order) { groups[index].x = x; x += groups[index].advance; }
  }

  const useCp = useCpFlag;
  const refX = useCp ? s.currentX : out.x;
  const refY = useCp ? s.currentY : out.y;
  const originX = s.vpX + (refX - s.orgX) * sx;
  const originY = s.vpY + (refY - s.orgY) * sy;
  // TA_RIGHT / TA_CENTER shift the start by the whole (or half the) advance
  // vector: along the baseline and, with ETO_PDY, across it.
  const shift = horizontal === TA_RIGHT ? -1 : horizontal === TA_CENTER ? -0.5 : 0;
  const x0 = shift * total;
  const y0 = shift * totalRise;
  const baseline = vertical === TA_BASELINE ? 0 : vertical === TA_BOTTOM ? -descentPx : ascentPx;
  const decoration = (font.underline || font.strikeout)
    ? referenceFontDecoration(font.face, weight, font.italic ? 'italic' : 'normal')
    : null;

  let opaqued = false;
  ctx.save();
  try {
    if (out.rect && (out.options & (ETO_OPAQUE | ETO_CLIPPED))) {
      const [rx, ry, rw, rh] = deviceRect(s, out.rect);
      if (out.options & ETO_OPAQUE) opaqued = paintOpaqueRect(s, out.rect);
      if (out.options & ETO_CLIPPED) {
        ctx.beginPath();
        ctx.rect(rx, ry, rw, rh);
        ctx.clip();
      }
    }
    // Local frame: x along the baseline, y down, origin at the reference point.
    let ox = originX;
    let oy = originY;
    if (phi !== 0) {
      ctx.translate(originX, originY);
      ctx.rotate(phi);
      ox = 0;
      oy = 0;
    }
    // Each character cell (or same-face run) sits at its own advance and,
    // with ETO_PDY, its own vertical offset.
    const cellX = (group: Group) => ox + x0 + group.x;
    if (s.bkMode === OPAQUE) {
      ctx.fillStyle = s.bkColor;
      for (const group of groups) {
        if (group.advance > 0) ctx.fillRect(cellX(group), oy + y0 + baseline + group.rise - ascentPx, group.advance, ascentPx + descentPx);
      }
    }
    ctx.fillStyle = s.textColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    for (const group of groups) {
      if (!group.text) continue;
      const gx = cellX(group);
      // Lone cells were mirrored above (UAX #9 L4) and are painted LTR so
      // Canvas does not mirror them again; longer runs leave bidi and
      // mirroring inside the run to Canvas.
      ctx.direction = groups.length === 1 ? (rtl ? 'rtl' : 'ltr')
        : [...group.text].length === 1 ? 'ltr' : (group.level & 1 ? 'rtl' : 'ltr');
      const gy = oy + y0 + baseline + group.rise;
      ctx.font = group.font;
      // Zero-width joiners give a separately drawn cursive cell the
      // contextual form it has in the whole string, without moving it.
      const text = `${group.joinBefore ? ZWJ : ''}${group.text}${group.joinAfter ? ZWJ : ''}`;
      if (hScale === 1) {
        ctx.fillText(text, gx, gy);
      } else {
        ctx.save();
        ctx.translate(gx, gy);
        ctx.scale(hScale, 1);
        ctx.fillText(text, 0, 0);
        ctx.restore();
      }
    }
    if (font.underline || font.strikeout) {
      ctx.fillStyle = s.textColor;
      const thickness = decoration ? decoration.underlineThicknessEm * emPx : Math.max(1, emPx * DECORATION_LINE_WIDTH_EM);
      const strike = decoration ? decoration.strikeThicknessEm * emPx : thickness;
      const underlineTop = decoration ? decoration.underlineTopEm * emPx : Math.max(2, thickness);
      const strikeTop = decoration ? decoration.strikeTopEm * emPx : -STRIKE_RISE_EM * emPx - strike / 2;
      // One rule per baseline height: adjacent cells sharing a rise merge.
      const visual = order.map((index) => groups[index]);
      let start = 0;
      while (start < visual.length) {
        let end = start + 1;
        while (end < visual.length && visual[end].rise === visual[start].rise) end++;
        const first = visual[start];
        const last = visual[end - 1];
        const from = first.x;
        const width = last.x + last.advance - first.x;
        if (width > 0) {
          const y = oy + y0 + baseline + first.rise;
          if (font.underline) ctx.fillRect(ox + x0 + from, y + underlineTop, width, thickness);
          if (font.strikeout) ctx.fillRect(ox + x0 + from, y + strikeTop, width, strike);
        }
        start = end;
      }
    }
  } finally {
    ctx.restore();
  }

  if (useCp && horizontal !== TA_CENTER) {
    // TA_LEFT moves the current position past the string; TA_RIGHT moves it to
    // the string's start; TA_CENTER leaves it unchanged.
    const signed = horizontal === TA_RIGHT ? -total : total;
    const rise = horizontal === TA_RIGHT ? -totalRise : totalRise;
    // Local (along, across) → device through the same rotation as the glyphs.
    const dxDev = signed * Math.cos(phi) - rise * Math.sin(phi);
    const dyDev = signed * Math.sin(phi) + rise * Math.cos(phi);
    s.currentX = refX + dxDev / sx;
    s.currentY = refY + dyDev / sy;
  }
  return { drew: units.length > 0 || opaqued };
}

function deviceRect(s: WmfTextState, rect: readonly [number, number, number, number]): [number, number, number, number] {
  const sx = s.vpW / s.extX;
  const sy = s.vpH / s.extY;
  const [l, t, r, b] = rect;
  const ax = s.vpX + (l - s.orgX) * sx; const bx = s.vpX + (r - s.orgX) * sx;
  const ay = s.vpY + (t - s.orgY) * sy; const by = s.vpY + (b - s.orgY) * sy;
  return [Math.min(ax, bx), Math.min(ay, by), Math.abs(bx - ax), Math.abs(by - ay)];
}

/** ETO_OPAQUE fills the rectangle with the background color; it needs no font. */
function paintOpaqueRect(s: WmfTextState, rect: readonly [number, number, number, number]): boolean {
  const [x, y, w, h] = deviceRect(s, rect);
  if (!(w > 0 && h > 0)) return false;
  s.ctx.fillStyle = s.bkColor;
  s.ctx.fillRect(x, y, w, h);
  return true;
}
