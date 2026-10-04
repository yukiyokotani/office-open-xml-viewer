import { classifyCjkFont } from '@silurus/ooxml-core';
import { referenceFontRoute } from './font-route.js';
import {
  EAST_ASIAN_SLOT_RANGES,
  MICROSOFT_JHENGHEI_RANGES,
  MS_MINCHO_GOTHIC_BITMAP_BASE64,
  PMINGLIU_RANGES,
} from './east-asian-default-coverage.js';

// PowerPoint's application defaults for a run whose East Asian or complex-
// script slot has no face (issue #1627). The parser resolves every authored
// ea/cs face, including theme tokens and the language's theme script font;
// these apply only when that leaves the slot empty: no ea/cs anywhere in the
// chain, an explicit typeface="", or a token naming an empty theme slot.
//
// Observed with PowerPoint's reference PDF engine (#1627 controls D, N, X):
// * A Latin face that is itself an East Asian face (OS/2 Far-East code page:
//   Yu Gothic, SimSun, Batang, PMingLiU, MS Mincho) draws the East Asian text.
//   Characters it lacks fell back to Microsoft JhengHei / Malgun Gothic after
//   a Japanese face (Yu Gothic, MS Mincho) and to PMingLiU / Batang after a
//   Chinese or Korean face (SimSun, Batang, PMingLiU).
// * Otherwise the Latin face's installed PANOSE serif style picks the tier:
//   serif (2-10: Perpetua, Garamond, Georgia) MS Mincho > PMingLiU > Batang,
//   sans (Corbel, Arial, Tahoma, Gill Sans MT, and any face not installed)
//   MS Gothic > Microsoft JhengHei > Malgun Gothic. Authored panose /
//   pitchFamily on <a:latin> has no effect.
// * The first tier face whose repertoire covers the whole East Asian text
//   draws it: Japanese and Traditional Chinese in MS Mincho / MS Gothic,
//   "简体中文" entirely in PMingLiU / Microsoft JhengHei, Hangul in Batang /
//   Malgun Gothic.
// * Complex scripts: Hebrew and Arabic in Arial, Thai in Angsana New,
//   Devanagari in Mangal, Tamil in Latha, whatever the Latin face. The same
//   faces follow an authored cs face that lacks a character.
//
// Issue #1689 controls (emptyea, emptyea2-4; PowerPoint 16.113 tagged PDFs,
// every glyph's face read from the embedded font name tables) refine whose
// face this is when the East Asian slot is empty:
// * The selected face S is the run's named complex-script face, else its
//   Latin face. An empty, omitted or theme-empty cs counts as unnamed. S
//   draws every East Asian-slot glyph it covers, CJK included: CS Tahoma
//   draws § ■, CS MS Mincho / DengXian / Arial Unicode MS draw Han and kana,
//   CS empty with Latin Arial draws § ■ in Arial. The language (en/ja/zh/ko)
//   and the theme Latin face play no part.
// * A non-CJK glyph S lacks takes the application symbol fallback: § in
//   Calibri even when no Calibri appears in the deck, ◆ ■ in Cambria Math.
// * CJK glyphs S lacks: a Far-East face that maps basic CJK keeps the script
//   chain above whichever slot it fills (CS MS Mincho sends 简 to Microsoft
//   JhengHei and 한 to Malgun Gothic, like Latin MS Mincho). A face that maps
//   no CJK Unified Ideograph falls back by its PANOSE serif style, now
//   including style 1 (SimSun-ExtB) with 2-10: Times New Roman, Courier New,
//   SimSun-ExtB and MingLiU-ExtB took MS Mincho; Tahoma, Lucida Sans Unicode,
//   Arial and Calibri took MS Gothic.
// * Owner decision (B), #1689: after that first CJK fallback of a Far-East
//   face without basic CJK, the next face differs per font (SimSun-ExtB sent
//   简 / 한 to JhengHei / Malgun Gothic, MingLiU-ExtB to PMingLiU / Batang)
//   in a way no general font property explains: it behaves like a per-font
//   link table. It is left to the platform's glyph fallback rather than
//   encoded by font name, and such a glyph's drawing face is unknown.

const SERIF_TIERS = ['MS Mincho', 'PMingLiU', 'Batang'] as const;
/** Application fallbacks for a non-CJK East Asian-slot glyph the selected face
 * lacks, in coverage order (#1689: § → Calibri; ◆ ■, which Calibri lacks, →
 * Cambria Math). */
export const EAST_ASIAN_SYMBOL_FALLBACK_FACES: readonly string[] = Object.freeze([
  'Calibri', 'Cambria Math',
]);
const SANS_TIERS = ['MS Gothic', 'Microsoft JhengHei', 'Malgun Gothic'] as const;

/** Complex-script application defaults, in per-glyph fallback order. */
export const COMPLEX_SCRIPT_DEFAULT_FACES: readonly string[] = Object.freeze([
  'Arial', 'Angsana New', 'Mangal', 'Latha',
]);

const THAI_RE = /\p{Script=Thai}/u;
const DEVANAGARI_RE = /\p{Script=Devanagari}/u;
const TAMIL_RE = /\p{Script=Tamil}/u;

/** The complex-script application default for one character. */
export function complexScriptDefaultFace(ch: string): string {
  if (THAI_RE.test(ch)) return 'Angsana New';
  if (DEVANAGARI_RE.test(ch)) return 'Mangal';
  if (TAMIL_RE.test(ch)) return 'Latha';
  return 'Arial';
}

let minchoBits: Uint8Array | undefined;

function decodeBitmap(): Uint8Array {
  if (minchoBits) return minchoBits;
  const binary = atob(MS_MINCHO_GOTHIC_BITMAP_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  minchoBits = bytes;
  return bytes;
}

function slotIndex(cp: number): number {
  let offset = 0;
  for (const [lo, hi] of EAST_ASIAN_SLOT_RANGES) {
    if (cp >= lo && cp <= hi) return offset + cp - lo;
    offset += hi - lo + 1;
  }
  return -1;
}

function inRanges(cp: number, ranges: readonly (readonly [number, number])[]): boolean {
  let lo = 0;
  let hi = ranges.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const [a, b] = ranges[mid];
    if (cp < a) hi = mid - 1;
    else if (cp > b) lo = mid + 1;
    else return true;
  }
  return false;
}

function coveredByMincho(cp: number): boolean {
  const i = slotIndex(cp);
  if (i < 0) return false;
  return (decodeBitmap()[i >> 3] & (1 << (i & 7))) !== 0;
}

function covers(face: string, text: string): boolean {
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    if (/\s/u.test(ch)) continue;
    const ok = face === 'MS Mincho' || face === 'MS Gothic' ? coveredByMincho(cp)
      : face === 'PMingLiU' ? inRanges(cp, PMINGLIU_RANGES)
      : face === 'Microsoft JhengHei' ? inRanges(cp, MICROSOFT_JHENGHEI_RANGES)
      : false;
    if (!ok) return false;
  }
  return true;
}

/** Shared preflight/render routing reads the compact exhaustive projection,
 * not full metric profiles, per-cut cmaps or shaping certificates. It preserves
 * the former Office-first alias/source aggregation and unknown-name defaults. */
export function isEastAsianFace(face: string): boolean { return !!(referenceFontRoute(face) & 1); }
export function isSerifLatinFace(face: string): boolean { return !!(referenceFontRoute(face) & 2); }
export function coversCjkIdeographs(face: string): boolean | undefined {
  const cjk = referenceFontRoute(face) >> 2;
  return cjk === 0 ? undefined : cjk === 2;
}

const CJK_CLASS_RE = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\u3000-\u303F\uFF00-\uFFEF]/u;

/** A CJK glyph for the empty East Asian slot's fallback: Han, kana, Hangul,
 * Bopomofo, CJK punctuation and full/half-width forms. Other East
 * Asian-slot characters (§ ° ± ◆ ■ …) take the symbol fallback. */
export function isCjkFallbackGlyph(ch: string): boolean {
  return CJK_CLASS_RE.test(ch);
}

/**
 * Faces, in fallback order, for the CJK glyphs of a run whose East Asian slot
 * is empty. `selectedFace` is S (the named cs face, else the Latin face);
 * `text` the run's East Asian-slot characters. The first face draws CJK text
 * the selected face does not.
 */
export function eastAsianDefaultFaces(selectedFace: string | null, text: string): readonly string[] {
  if (selectedFace && isEastAsianFace(selectedFace)) {
    if (coversCjkIdeographs(selectedFace) === false) {
      // A Far-East code page without basic CJK (ExtB faces): PANOSE first
      // fallback only; later faces are the platform's (decision B).
      return [isSerifLatinFace(selectedFace) ? SERIF_TIERS[0] : SANS_TIERS[0]];
    }
    return classifyCjkFont(selectedFace) === 'jp'
      ? [selectedFace, 'Microsoft JhengHei', 'Malgun Gothic']
      : [selectedFace, 'PMingLiU', 'Batang'];
  }
  const tiers = selectedFace && isSerifLatinFace(selectedFace) ? SERIF_TIERS : SANS_TIERS;
  const first = tiers.slice(0, 2).find((face) => covers(face, text));
  return first ? [first, ...tiers.filter((face) => face !== first)] : [tiers[2], tiers[0], tiers[1]];
}

/**
 * The complete font stack for East Asian-slot glyphs of a run whose East
 * Asian slot is empty: S first (coverage decides, as in Office), then the
 * symbol fallback, then the CJK fallback faces. Calibri and Cambria Math map
 * no CJK, so CJK glyphs pass them; the CJK faces follow S's own chain.
 */
export function emptyEastAsianSlotFaces(selectedFace: string | null, text: string): readonly string[] {
  const faces: string[] = [];
  const add = (face: string | null | undefined) => {
    const trimmed = face?.trim();
    if (trimmed && !faces.some((f) => f.toLowerCase() === trimmed.toLowerCase())) faces.push(trimmed);
  };
  add(selectedFace);
  for (const face of EAST_ASIAN_SYMBOL_FALLBACK_FACES) add(face);
  for (const face of eastAsianDefaultFaces(selectedFace, text)) add(face);
  return faces;
}
