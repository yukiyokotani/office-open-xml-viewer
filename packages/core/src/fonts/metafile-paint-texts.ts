/**
 * The texts a metafile picture paints, for `unicode-range` font demand
 * (library loading policy, not OOXML rendering behaviour).
 *
 * The WMF and EMF players draw text records with the document's served faces
 * (`wmf-text.ts`, `emf.ts` `drawText`), so the scalars those records paint
 * must be demanded before the picture plays. This module walks the record
 * stream the way the players do, without drawing: it tracks only what decides
 * the painted text (the WMF object table, the selected font and the saved DC
 * levels that restore it) and decodes each text record exactly as its player
 * does. Everything a player may paint or measure for a record is returned,
 * which is a superset of what one playback paints:
 *
 * - WMF META_TEXTOUT / META_EXTTEXTOUT (MS-WMF 2.3.3.20 / 2.3.3.5): the
 *   String decoded under the selected font's CharacterSet (the player's own
 *   decoder), every unit alone and the whole string (the player draws runs of
 *   units, separately positioned cells and lone mirrored characters), ZWJ
 *   (cursive context of separately drawn cells) and the GDI metric probes
 *   (`Hg`, the weighted average-width letters). A SYMBOL_CHARSET string
 *   paints differently depending on whether the host has the face (the font
 *   mapper's substitute otherwise); both outcomes and the font-coded probe
 *   are returned. Records the player never draws (no font, malformed font,
 *   zero height, `@` vertical face, ETO_GLYPH_INDEX) paint nothing;
 * - EMF EMR_EXTTEXTOUTW (MS-EMF 2.3.5.8): the UTF-16 string, and for a string
 *   with Windows symbol cmap codes (U+F000–F0FF) every repertoire conversion
 *   the player may apply (`decodeMetafileSymbolText`). The other EMF text
 *   records (EXTTEXTOUTA, POLYTEXTOUT*, SMALLTEXTOUT) and EMF+ string records
 *   are not drawn by the players and paint nothing.
 *
 * Every text is painted literally (the players apply no further symbol-font
 * conversion), so private-use codes in it are demanded as themselves.
 * Compressed (`.wmz` / `.emz`) or unrecognised bytes play nothing.
 */
import { isEmf, isWmf } from '../image/wmf.js';
import { decodeWmfTextBytes, parseWmfFont, type WmfFont } from '../image/wmf-text.js';
import { decodeMetafileSymbolText } from './metafile-symbol-encoding.js';

/** The players' fixed probe and joiner texts: GDI text-metric probes
 * (`faceMetrics` measures `Hg`; the average width weighs `a`–`z` and space)
 * and the ZWJ that gives a separately drawn cell its cursive context. */
const WMF_PROBE_TEXTS = ['Hg', 'abcdefghijklmnopqrstuvwxyz ', '‍'] as const;

// MS-WMF 2.1.1.1 RecordType.
const META_EOF = 0x0000;
const META_SAVEDC = 0x001e;
const META_RESTOREDC = 0x0127;
const META_SELECTOBJECT = 0x012d;
const META_DELETEOBJECT = 0x01f0;
const META_TEXTOUT = 0x0521;
const META_EXTTEXTOUT = 0x0a32;
const META_CREATEFONTINDIRECT = 0x02fb;
/** Every other object record: each takes the lowest free object slot
 * (MS-WMF 3.1.4.1), as the player's `insertObject` does. */
const META_CREATE_OTHER: ReadonlySet<number> = new Set([
  0x00f7, // CREATEPALETTE
  0x01f9, // CREATEPATTERNBRUSH
  0x0142, // DIBCREATEPATTERNBRUSH
  0x02fa, // CREATEPENINDIRECT
  0x02fc, // CREATEBRUSHINDIRECT
  0x06ff, // CREATEREGION
]);
// MS-WMF 2.1.2.2 ExtTextOutOptions.
const ETO_OPAQUE = 0x0002;
const ETO_CLIPPED = 0x0004;
const ETO_GLYPH_INDEX = 0x0010;

const SYMBOL_CHARSET = 2;
const PLACEABLE_MAGIC = 0x9ac6cdd7;

// MS-EMF 2.1.1 RecordType.
const EMR_EOF = 14;
const EMR_EXTTEXTOUTW = 84;
/** The EMF player's text-unit work limit (emf.ts): a longer string throws
 * before drawing. */
const MAX_EMF_TEXT_UNITS = 0x10000;
/** The symbol repertoires `decodeMetafileSymbolText` knows. */
const EMF_SYMBOL_FACES = ['Symbol', 'MT Extra'] as const;

/** Every text the metafile `bytes` may paint or measure (see the module
 * notes), each to be demanded literally. Empty for bytes no player plays. */
export function metafilePaintTexts(bytes: Uint8Array): string[] {
  if (isWmf(bytes)) return wmfPaintTexts(bytes);
  if (isEmf(bytes)) return emfPaintTexts(bytes);
  return [];
}

type WmfObject = WmfFont | { kind: 'other' };

function wmfPaintTexts(bytes: Uint8Array): string[] {
  const texts: string[] = [];
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // The placeable header (22 bytes) precedes the 18-byte standard header.
  let pos = (bytes.length >= 4 && dv.getUint32(0, true) === PLACEABLE_MAGIC ? 22 : 0) + 18;
  const objects: (WmfObject | null)[] = [];
  let font: WmfFont | null = null;
  const saved: (WmfFont | null)[] = [];
  let probed = false;
  const u16 = (at: number) => (at + 2 <= bytes.length ? dv.getUint16(at, true) : 0);
  const i16 = (at: number) => (at + 2 <= bytes.length ? dv.getInt16(at, true) : 0);
  const insert = (object: WmfObject) => {
    const free = objects.indexOf(null);
    if (free < 0) objects.push(object);
    else objects[free] = object;
  };
  const text = (string: Uint8Array, options: number) => {
    if (!font || !font.valid || font.height === 0 || font.face.startsWith('@') || (options & ETO_GLYPH_INDEX)) return;
    if (string.length === 0) return;
    if (!probed) { probed = true; texts.push(...WMF_PROBE_TEXTS); }
    // A SYMBOL_CHARSET face either draws its font codes itself or is
    // replaced by the font mapper's substitute, as the host decides at paint
    // time (wmf-text.ts `hostFaceCovers`, which measures the font-coded
    // string): both decodings, and that probe.
    const outcomes = font.charset === SYMBOL_CHARSET ? [false, true] : [false];
    if (font.charset === SYMBOL_CHARSET) texts.push(Array.from(string, (byte) => String.fromCharCode(0xf000 + byte)).join(''));
    for (const covers of outcomes) {
      const decoded = decodeWmfTextBytes(string, font.charset, font.face, () => covers);
      if ('problem' in decoded) continue;
      for (const unit of decoded.units) texts.push(unit.text);
      texts.push(decoded.units.map((unit) => unit.text).join(''));
    }
  };
  while (pos + 6 <= bytes.length) {
    const sizeWords = dv.getUint32(pos, true);
    const fn = dv.getUint16(pos + 4, true);
    if (sizeWords < 3) break;
    const end = pos + sizeWords * 2;
    if (end > bytes.length || fn === META_EOF) break;
    const at = pos + 6;
    const remaining = end - at;
    switch (fn) {
      case META_SAVEDC:
        if (remaining === 0) saved.push(font);
        break;
      case META_RESTOREDC: {
        if (remaining !== 2) break;
        const level = i16(at);
        const index = level < 0 ? saved.length + level : level - 1;
        if (level === 0 || index < 0 || index >= saved.length) break;
        font = saved[index]!;
        saved.length = index;
        break;
      }
      case META_CREATEFONTINDIRECT:
        insert(parseWmfFont(bytes.subarray(at, end)));
        break;
      case META_SELECTOBJECT: {
        if (remaining !== 2) break;
        const object = objects[u16(at)];
        if (object?.kind === 'font') font = object;
        break;
      }
      case META_DELETEOBJECT: {
        if (remaining !== 2) break;
        const index = u16(at);
        const object = objects[index];
        if (object) {
          if (object === font) font = null;
          objects[index] = null;
        }
        break;
      }
      case META_TEXTOUT: {
        // u16 StringLength, String (WORD-padded), i16 YStart, i16 XStart.
        if (remaining < 2) break;
        const length = u16(at);
        if (remaining - 2 < length + (length & 1) + 4) break;
        text(bytes.subarray(at + 2, at + 2 + length), 0);
        break;
      }
      case META_EXTTEXTOUT: {
        // i16 Y, i16 X, i16 StringLength, u16 fwOpts, [Rect], String, [Dx].
        if (remaining < 8) break;
        const length = i16(at + 4);
        const options = u16(at + 6);
        const rect = (options & (ETO_OPAQUE | ETO_CLIPPED)) !== 0 ? 8 : 0;
        if (length < 0 || remaining - 8 < rect + length) break;
        const start = at + 8 + rect;
        text(bytes.subarray(start, start + length), options);
        break;
      }
      default:
        if (META_CREATE_OTHER.has(fn)) insert({ kind: 'other' });
        break;
    }
    pos = end;
  }
  return texts;
}

function emfPaintTexts(bytes: Uint8Array): string[] {
  const texts: string[] = [];
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let pos = 0;
  while (pos + 8 <= bytes.length) {
    const type = dv.getUint32(pos, true);
    const size = dv.getUint32(pos + 4, true);
    if (size < 8 || (size & 3) !== 0 || pos + size > bytes.length) break;
    const end = pos + size;
    if (type === EMR_EOF) break;
    // EMR_EXTTEXTOUTW: rclBounds, iGraphicsMode, exScale, eyScale, then
    // EMRTEXT at record offset 36: Reference, Chars, offString, Options.
    if (type === EMR_EXTTEXTOUTW && size >= 56) {
      const chars = dv.getUint32(pos + 44, true);
      const offString = dv.getUint32(pos + 48, true);
      const options = dv.getUint32(pos + 52, true);
      const header = options & 0x100 ? 60 : 76;
      if (pos + header <= end && chars > 0 && chars <= MAX_EMF_TEXT_UNITS
        && offString >= header && pos + offString + chars * 2 <= end) {
        let string = '';
        for (let index = 0; index < chars; index++) string += String.fromCharCode(dv.getUint16(pos + offString + index * 2, true));
        texts.push(string);
        // A SYMBOL_CHARSET font's cmap codes are converted through its
        // published repertoire when the record has no options (emf.ts).
        if (options === 0 && /[-]/.test(string)) {
          for (const face of EMF_SYMBOL_FACES) {
            const converted = decodeMetafileSymbolText(string, face, true);
            if (converted !== null) texts.push(converted);
          }
        }
      }
    }
    pos = end;
  }
  return texts;
}
