import type { DocxTextRun, FieldRun } from '../types';
import type { NumberFormat } from '@silurus/ooxml-core';
import { isCjkBreakChar, isComplexScriptCodePoint, formatOrdinalNumber, parseFieldFormatSwitch, formatDateTimePicture, parseDateTimePictureSwitch } from '@silurus/ooxml-core';
import type { ParagraphLayoutRun } from '../layout/text.js';
import { RTL_PRIMARY_SUBTAGS } from '../layout/line-compatibility.js';
import { wordNeutralAttachesToActiveScript } from '../layout/script-compatibility.js';
import { type LineLayoutEnvironment } from './model.js';

/**
 * Resolve the formatting axis that actually governs a run's glyphs.
 *
 * ECMA-376 §17.3.2.30 `w:rtl` marks a run as complex-script. For such a run the
 * complex-script properties take effect — §17.3.2.4 `bCs` (bold), §17.3.2.6
 * `iCs` (italic), §17.3.2.26 `rFonts@cs` (typeface), §17.3.2.39 `szCs` (size) —
 * instead of the non-CS `b`/`i`/`rFonts@ascii`/`sz`, which apply to
 * non-complex (Latin/CJK) text. `bCs`/`iCs` are INDEPENDENT toggles: an absent
 * `bCs`/`iCs` does not inherit `b`/`i`'s value, so a complex-script run that
 * carries only `w:b`/`w:i` renders non-bold/upright (`csBold = boldCs ?? false`,
 * `csItalic = italicCs ?? false`). Thus `w:b` without `w:bCs` remains regular
 * weight, and `w:i` without `w:iCs` remains upright, while the corresponding
 * non-complex text uses the Latin-axis toggle.
 */
/**
 * Split a `w:smallCaps` (§17.3.2.33) run into maximal pieces by character class
 * for sizing. The spec reduces "all SMALL LETTER characters ... two points
 * smaller", so ONLY lowercase letters are `reduced`; uppercase letters AND every
 * non-alphabetic character (digits, punctuation) stay at the FULL run size.
 * So "Introduction" → "I" full + "NTRODUCTION" reduced (matching the heading's
 * "1."), and "co2" → "CO" reduced + "2" full. `reduced` flags the small-cap
 * pieces; the caller still uppercases every piece for display.
 *
 * Whitespace carries no glyph, so it EXTENDS the current piece rather than
 * opening a full-size one — otherwise an inter-word space between two small-cap
 * words would fragment into its own segment and corrupt trailing-space collapse
 * / line breaking. A leading run with no lowercase letter defaults to full size.
 */
export function splitSmallCapsCase(text: string): { text: string; reduced: boolean }[] {
  const out: { text: string; reduced: boolean }[] = [];
  for (const ch of text) {
    // A lowercase letter: unchanged by toLowerCase AND changed by toUpperCase.
    const isLowerLetter = ch.toLowerCase() === ch && ch.toUpperCase() !== ch;
    const reduced = /\s/.test(ch)
      ? (out[out.length - 1]?.reduced ?? false) // whitespace: keep with current piece
      : isLowerLetter;
    const last = out[out.length - 1];
    if (last && last.reduced === reduced) last.text += ch;
    else out.push({ text: ch, reduced });
  }
  return out.length ? out : [{ text, reduced: false }];
}


export type { NoteNumbering } from './model.js';


/** The label of the `ordinal`-th (1-based) automatic note of one kind. */
export function formatNoteNumber(
  ordinal: number,
  numbering: Readonly<{ format: string; start: number }> | undefined,
): string {
  if (!numbering) return String(ordinal);
  const value = numbering.start + ordinal - 1;
  return numbering.format.startsWith('custom:')
    ? formatCustomNoteNumber(value, numbering.format.slice('custom:'.length))
    : formatOrdinalNumber(value, numbering.format as NumberFormat);
}

/** Numbering sequences a custom format token can start (XSLT 1.0 §7.7.1:
 * "any other format token indicates a numbering sequence that starts with
 * that token"): the ST_NumberFormat (§17.18.59) sequence whose first value is
 * the token, e.g. `ア` (U+30A2, the §17.11.17/.18 example) for Katakana. */
const CUSTOM_TOKEN_SEQUENCES: readonly NumberFormat[] = [
  'aiueoFullWidth', 'aiueo', 'decimalEnclosedCircle', 'ideographDigital', 'ganada', 'chosung',
  'arabicAlpha', 'arabicAbjad', 'russianLower', 'russianUpper', 'thaiLetters', 'thaiNumbers',
  'hindiVowels', 'hindiConsonants', 'hindiNumbers', 'hebrew2',
];
let customTokenSequences: ReadonlyMap<string, NumberFormat> | undefined;
function customTokenSequence(token: string): NumberFormat | undefined {
  if (!customTokenSequences) {
    const sequences = new Map<string, NumberFormat>();
    for (const format of CUSTOM_TOKEN_SEQUENCES) {
      const first = formatOrdinalNumber(1, format);
      if (!sequences.has(first)) sequences.set(first, format);
    }
    customTokenSequences = sequences;
  }
  return customTokenSequences.get(token);
}

/** A custom numFmt/@format pattern (ECMA-376 §17.11.17/.18 `format`, "the
 * syntax defined by the XSLT format attribute"): its first alphanumeric
 * format token numbers the value (XSLT 1.0 §7.7.1). A run of ASCII or
 * full-width digits ending in one zero-pads to the token's length in that
 * digit family (`001`, `０１`); `a`, `A`, `i` and `I` select letters and
 * Roman numerals; another token starts the sequence whose first value it is
 * (`ア` Katakana); a token that starts no known sequence numbers in decimal.
 * A pattern such as "001, 002, 003, ..." lists sample values: only its first
 * token numbers a single note, and its separators and later tokens add no
 * prefix or suffix. */
export function formatCustomNoteNumber(value: number, pattern: string): string {
  // XSLT 1.0 §7.7.1 alphanumerics: letters and Nd/Nl/No numbers.
  const token = /[\p{L}\p{Nd}\p{Nl}\p{No}]+/u.exec(pattern)?.[0];
  if (!token || !Number.isSafeInteger(value) || value < 0) return String(value);
  for (const zero of [0x30, 0xff10]) {
    const digits = [...token].map((digit) => digit.codePointAt(0)! - zero);
    if (digits.every((digit) => digit >= 0 && digit <= 9) && digits.at(-1) === 1
      && digits.slice(0, -1).every((digit) => digit === 0)) {
      return [...String(value).padStart(digits.length, '0')]
        .map((digit) => String.fromCodePoint(zero + Number(digit))).join('');
    }
  }
  switch (token) {
    case 'a': return formatOrdinalNumber(value, 'lowerLetter');
    case 'A': return formatOrdinalNumber(value, 'upperLetter');
    case 'i': return formatOrdinalNumber(value, 'lowerRoman');
    case 'I': return formatOrdinalNumber(value, 'upperRoman');
    default: {
      const sequence = customTokenSequence(token);
      return sequence ? formatOrdinalNumber(value, sequence) : String(value);
    }
  }
}


export function findNearbyFontSize(
  runs: readonly ParagraphLayoutRun[],
  idx: number,
): number {
  // Look backwards then forwards for a text or field run to get font size
  for (let i = idx - 1; i >= 0; i--) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as DocxTextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  for (let i = idx + 1; i < runs.length; i++) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as DocxTextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  return 10; // pt fallback
}


export function resolveFieldText(f: FieldRun, environment: LineLayoutEnvironment): string {
  if (f.fieldType === 'page') {
    // ECMA-376 §17.16.5.44 PAGE — "the number of the current page". Use the
    // per-section DISPLAY number (§17.6.12 `w:start` restart), falling back to the
    // raw physical index for a single-section document without `<w:pgNumType>`.
    const n = environment.displayPageNumber ?? environment.pageIndex + 1;
    // §17.16.4.3.1 — the field's own general-formatting switch (`\* roman`, …)
    // OVERRIDES the section format (§17.6.12 `w:fmt`); it is authored ON the field.
    // No switch ⇒ the section format (or decimal for a single-section document).
    const fmt = parseFieldFormatSwitch(f.instruction) ?? environment.pageNumberFormat ?? 'decimal';
    return formatOrdinalNumber(n, fmt);
  }
  // ECMA-376 §17.16.5.42 NUMPAGES — "the number of pages in the current document".
  // This is the DOCUMENT's physical page count and is NOT affected by §17.6.12
  // page-number restart (which only shifts the DISPLAYED number). It IS still
  // subject to the field's own `\*` format switch.
  if (f.fieldType === 'numPages') {
    const fmt = parseFieldFormatSwitch(f.instruction) ?? 'decimal';
    return formatOrdinalNumber(environment.totalPages, fmt);
  }
  // ECMA-376 §17.16.5.16 DATE / §17.16.5.72 TIME — display the CURRENT date/time
  // filtered through the field's `\@` date-time picture (§17.16.4.1). The
  // "current" instant is injected via `environment.currentDateMs` (default = real time,
  // set at the render entry point) so the output is deterministic under test.
  // A field with NO `\@` picture, or one whose picture uses an unimplemented
  // token, falls back to the authored cached result (§17.16.4.1: with no picture
  // the result is formatted "in an implementation-defined manner" — we keep
  // Word's cached rendering rather than invent one).
  if (f.fieldType === 'date' || f.fieldType === 'time') {
    const picture = parseDateTimePictureSwitch(f.instruction);
    if (picture) {
      const now = new Date(environment.currentDateMs ?? Date.now());
      const formatted = formatDateTimePicture(picture, now);
      if (formatted !== null) return formatted;
    }
    return f.fallbackText;
  }
  return f.fallbackText;
}



/** Returns true when any code point of `text` permits a line break between
 *  adjacent characters (CJK / ideographic). The canonical ranges live in core's
 *  {@link isCjkBreakChar} (single source of truth across all renderers). */
export function hasCJKBreakOpportunity(text: string): boolean {
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    if (isCjkBreakChar(cp)) return true;
    i += cp > 0xffff ? 2 : 1;
  }
  return false;
}


// ECMA-376 §17.15.1.18 / §17.18.7 distinguishes punctuation-only compression
// from punctuation-plus-Japanese-kana compression. This is the reviewed
// supported subset: full-width dividing punctuation and closing forms verified
// by the registered Word fixture. U+3017, full-width !, and full-width ? remain
// full-cell in that same matrix and are deliberately excluded.
// JLReq classifies middle dot, colon, and semicolon together (cl-05); their
// whitespace belongs on both sides and must be resolved from the adjacent
// character classes, so they cannot use this trailing-side-only projection.
// Halfwidth U+FF61/U+FF64 are not full-width. Opening punctuation likewise
// needs line-start positioning rather than a pen-advance reduction. The
// implementation-note evidence for the full-width punctuation scope is
// registered in layout/line-compatibility.ts.
export const COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION = new Set([
  '、', '。', '，', '．', '」', '』', '】', '）', '］', '｝',
]);


/** Full-width Japanese kana characters for
 * `compressPunctuationAndJapaneseKana`. The ranges follow Unicode's Hiragana,
 * Katakana, Katakana Phonetic Extensions, and supplementary Kana blocks while
 * excluding halfwidth Katakana and punctuation such as U+30FB. U+30FC is the
 * shared full-width kana prolonged-sound mark. */
export function isFullWidthJapaneseKana(character: string): boolean {
  const cp = character.codePointAt(0);
  if (cp === undefined) return false;
  return (
    (cp >= 0x3041 && cp <= 0x3096)
    || (cp >= 0x309d && cp <= 0x309f)
    || (cp >= 0x30a1 && cp <= 0x30fa)
    || cp === 0x30fc
    || (cp >= 0x30fd && cp <= 0x30ff)
    || (cp >= 0x31f0 && cp <= 0x31ff)
    || (cp >= 0x1aff0 && cp <= 0x1afff)
    || (cp >= 0x1b000 && cp <= 0x1b16f)
  );
}


// `word-full-width-character-spacing-scope`
export function characterSpacingControlCompresses(
  grapheme: string,
  setting: string | undefined,
): boolean {
  switch (setting) {
    case 'compressPunctuation':
      return COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(grapheme);
    case 'compressPunctuationAndJapaneseKana':
      return COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(grapheme)
        || isFullWidthJapaneseKana(grapheme);
    case 'doNotCompress':
    default:
      return false;
  }
}


/** Shift a SEA break-offset list (issue #797) onto a suffix that drops the first
 *  `cut` UTF-16 units: keep offsets strictly greater than `cut` and rebase them.
 *  Used when a Thai/Lao/Khmer segment is split (line wrap) or resumed at a
 *  pagination boundary. A non-SEA segment (`offsets === undefined`) stays
 *  non-SEA; a SEA segment stays SEA-flagged (returns `[]` when no dictionary
 *  boundary remains, so an over-long FINAL word still takes the SEA path and is
 *  split grapheme-safely rather than by code point). */
export function rebaseSeaBreaks(offsets: readonly number[] | undefined, cut: number): readonly number[] | undefined {
  if (offsets === undefined) return undefined;
  const out: number[] = [];
  for (const o of offsets) if (o > cut) out.push(o - cut);
  return out;
}


/**
 * Binary-search the longest prefix of `text` whose rendered width fits in `maxWidth`.
 * Used for CJK overflow splitting.
 */
/** Extend an accepted split point through IMMEDIATELY FOLLOWING IDEOGRAPHIC
 *  SPACES (U+3000): the fullwidth space belongs to the line it ends, hanging
 *  past the band (JLReq line-end ideographic-space handling — the same
 *  allowance fitCJKPrefix's fit predicate applies), so a split must never
 *  strand it at the head of the next line — including the FORCE-FIT paths
 *  where the band is narrower than a single glyph (a one-glyph-wide form
 *  label column). A zero split (whole-run move / kinsoku retraction) is left
 *  untouched. */
export function extendThroughTrailingIdeographicSpaces(
  chars: string[],
  split: number,
  maximum = Number.POSITIVE_INFINITY,
): number {
  if (split <= 0 || maximum <= 0) return split;
  if (Number.isFinite(maximum) && chars[split - 1] === '\u3000') return split;
  let s = split;
  let remaining = maximum;
  while (s < chars.length && chars[s] === '\u3000' && remaining > 0) {
    s++;
    remaining--;
  }
  return s;
}


/**
 * Split a text run into layout-segment strings.
 * Each segment is an atomic unit for word-level fitting; CJK overflow is handled in layoutLines.
 */
/**
 * Decide whether a `w:lang w:bidi` tag (§17.3.2.20) designates an RTL
 * complex-script language, so the run's European digits are classified AN
 * (Word's date ordering). The tag's primary subtag (before the first '-') is
 * matched against {@link RTL_PRIMARY_SUBTAGS}. When the tag is absent OR a
 * malformed/unknown value (e.g. the "ae-AR" seen in real-world files), fall
 * back to whether the run is explicitly rtl-marked — `w:rtl` already asserts
 * the run is complex-script RTL content.
 */
export function isRtlBidiLang(langBidi: string | undefined, runIsRtl: boolean): boolean {
  if (langBidi) {
    const primary = langBidi.split('-')[0].toLowerCase();
    if (RTL_PRIMARY_SUBTAGS.has(primary)) return true;
  }
  return runIsRtl;
}


/**
 * Split `text` into maximal runs that are uniformly complex-script or not, per
 * §17.3.2.26 per-character classification. Returns `[{text, cs}]` in logical
 * order. Used only when a run has NO explicit `w:rtl`/`w:cs` (which would force
 * the whole run to cs); otherwise the caller treats the entire piece as cs.
 *
 * Under `word-neutral-script-attachment`, digits / spaces / punctuation attach
 * to the PRECEDING slice so a number embedded in Arabic ("نص 12 نص") does not
 * fragment into extra segments. A leading neutral run takes the first strong
 * slice's class.
 */
export function splitByComplexScript(text: string): { text: string; cs: boolean }[] {
  const out: { text: string; cs: boolean }[] = [];
  let curCs: boolean | null = null;
  let buf = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    // Neutral (non-letter) characters do not switch the active class; they ride
    // with whatever script is currently open (or the next one if none yet).
    if (wordNeutralAttachesToActiveScript(ch)) {
      buf += ch;
      continue;
    }
    const cs = isComplexScriptCodePoint(cp);
    if (curCs === null) {
      curCs = cs;
      buf += ch;
    } else if (cs === curCs) {
      buf += ch;
    } else {
      out.push({ text: buf, cs: curCs });
      curCs = cs;
      buf = ch;
    }
  }
  if (buf.length > 0) out.push({ text: buf, cs: curCs ?? false });
  return out;
}


/**
 * Split a (non-complex-script) string into maximal runs that are uniformly
 * East-Asian (CJK) or not, per the §17.3.2.26 ascii/eastAsia axis split. Returns
 * `[{text, ea}]` in logical order. CJK classification uses the canonical
 * {@link isCjkBreakChar} from `@silurus/ooxml-core` — the SAME predicate the body
 * wrap/justify paths use. Text-box text now feeds this splitter too (its runs are
 * adapted to body runs and run through {@link buildSegments}), so the eastAsia
 * face is picked consistently across body and shape with no name heuristics. Each
 * returned slice stays single-font when emitted, preserving the
 * measure==draw / docGrid char-grid invariant.
 *
 * Boundary rule: classification is purely per code point (every CJK code point
 * opens/continues an `ea` run; every other code point a `latin` run). This is
 * intentionally simpler than {@link splitByComplexScript}'s neutral-attachment —
 * a digit between two ideographs is Latin/ascii either way (§17.3.2.26 assigns
 * ASCII digits to the ascii face), and a single fillText anchors to the cumulative
 * whole-string advance, so the visible spacing is unchanged.
 *
 * NOTE: this split decides the FONT slot only. `linesAndChars` applies its pitch
 * to both partitions. The pre-existing non-`linesAndChars` fallback uses the
 * grid's own `EAST_ASIAN_RE` purity test (see `gridSegDeltaPx`/`eaGlyphCount`),
 * not the `ea` font-slot flag here.
 */
export function splitByEastAsia(text: string): { text: string; ea: boolean }[] {
  const out: { text: string; ea: boolean }[] = [];
  let curEa: boolean | null = null;
  let buf = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    const ea = isCjkBreakChar(cp);
    if (curEa === null || ea === curEa) {
      curEa = ea;
      buf += ch;
    } else {
      out.push({ text: buf, ea: curEa });
      curEa = ea;
      buf = ch;
    }
  }
  if (buf.length > 0) out.push({ text: buf, ea: curEa ?? false });
  return out;
}


/**
 * Split a token into maximal runs of European digits (U+0030–0039) versus the
 * separators between them, so a date in an AN-classified Arabic run can be
 * reordered group-by-group by the per-line bidi pass (which works at segment
 * granularity). "28-02-2026" → ["28","-","02","-","2026"], which the RTL reorder
 * then enters the right-to-left layout pass.
 *
 * EXCEPTION — ECMA-376 relies on UAX#9 W4: a SINGLE common separator (CS) sitting
 * between two numbers of the same type joins them into ONE number. So a decimal /
 * thousands / time separator (`.`, `,`, `:`, `/`, NBSP) flanked by European
 * digits on BOTH sides stays inside the digit group: "1234.56", "1,234.56" and
 * "12:34" are one left-to-right number, not three reorderable pieces. (A European
 * separator like `-` is ES, NOT CS, and W4's ES clause is EN-only — these run
 * digits are AN — so a hyphen still splits, preserving the date case.) Splitting
 * a decimal sent "1234.56" through the RTL segment reorder and drew it "56.1234".
 */
export function splitDigitGroups(text: string): string[] {
  const isEuDigit = (c: number) => c >= 0x30 && c <= 0x39;
  // UAX#9 Common Separator (CS) subset that can join two adjacent numbers (W4).
  // The last char is NBSP (U+00A0, e.g. a French thousands separator), itself CS;
  // a plain space is WS and never reaches here (splitTextForLayout breaks on it).
  const isJoiningCS = (ch: string) =>
    ch === '.' || ch === ',' || ch === ':' || ch === '/' || ch === ' ';
  const out: string[] = [];
  let buf = '';
  let bufDigit: boolean | null = null;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    let isDigit = isEuDigit(ch.charCodeAt(0));
    // W4: a single CS between two European digits is part of the number — keep it
    // in the current digit group so the whole number stays one (LTR) segment.
    if (!isDigit && bufDigit === true && isJoiningCS(ch) && isEuDigit(text.charCodeAt(i + 1))) {
      isDigit = true;
    }
    if (bufDigit === null || isDigit === bufDigit) {
      buf += ch;
    } else {
      out.push(buf);
      buf = ch;
    }
    bufDigit = isDigit;
  }
  if (buf.length > 0) out.push(buf);
  return out.length ? out : [text];
}


export function splitTextForLayout(text: string): string[] {
  const result: string[] = [];
  let i = 0;
  while (i < text.length) {
    let j = i;
    while (j < text.length && text[j] !== ' ') j++;
    while (j < text.length && text[j] === ' ') j++;
    if (j > i) result.push(text.slice(i, j));
    i = j;
  }
  return result.length ? result : [text];
}
