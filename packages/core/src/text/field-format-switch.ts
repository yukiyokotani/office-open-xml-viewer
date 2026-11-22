// ECMA-376 §17.16.4.3 / §17.16.4.3.1 — the field "general formatting switch"
// (`\*`). A field instruction may carry `\* <argument>` to format its numeric
// result: `\* Roman` → uppercase roman, `\* alphabetic` → lowercase letters, etc.
// This is a PER-FIELD override that takes precedence over a section-level format
// (e.g. `<w:pgNumType w:fmt>`), because the switch is authored ON the field
// itself. §17.16.4.3.1 lists every argument and its ST_NumberFormat equivalent;
// we recognise the NUMERIC, locale-independent subset that `formatOrdinalNumber`
// (number-format.ts) can render — the same subset the page-number wiring uses.
//
// The mapping table below is verbatim from §17.16.4.3.1 ("Corresponds to an
// ST_NumberFormat enumeration value of …"):
//   Arabic      → decimal        Roman  → upperRoman   roman     → lowerRoman
//   ALPHABETIC  → upperLetter    alphabetic → lowerLetter
// The switch arguments are CASE-SENSITIVE (Roman ≠ roman, ALPHABETIC ≠
// alphabetic) — §17.16.4.3.1 lists them as distinct rows with different results.
//
// The international switch arguments below are the LOCALE-INDEPENDENT ones: each
// §17.16.4.3.1 row that maps to exactly ONE ST_NumberFormat regardless of the
// document language. The LOCALE-DEPENDENT arguments — CHINESENUM1/2/3, DBNUM1–4,
// KANJINUM1–3, IROHA, AIUEO — are intentionally OMITTED: a single argument like
// `\* DBNUM1` maps to ideographDigital (ja-JP) OR koreanDigital (ko-KR)
// depending on the field's `w:lang`, which this parser does not thread. Guessing
// a locale would be a heuristic (root AGENTS.md forbids), so they return null and
// the caller keeps the section format instead of choosing the wrong script.

import type { NumberFormat } from './number-format';
import { ownValue } from '../internal/own-key';

// Case-sensitive switch-argument → ST_NumberFormat, restricted to the values
// `formatOrdinalNumber` renders natively AND unambiguous across locales. Text
// spell-out arguments (CardText, Ordinal, OrdText, BAHTTEXT, …) intentionally
// have no entry: `parseFieldFormatSwitch` returns null for them so the caller
// keeps the section format rather than silently downgrading to decimal.
const SWITCH_TO_FORMAT: Readonly<Record<string, NumberFormat>> = {
  // §17.16.4.3.1 Latin / roman / hex / dash core.
  Arabic: 'decimal',
  ArabicDash: 'numberInDash', // "- 123 -"
  Hex: 'hex', // uppercase hexadecimal
  Roman: 'upperRoman',
  roman: 'lowerRoman',
  ALPHABETIC: 'upperLetter',
  alphabetic: 'lowerLetter',
  // §17.16.4.3.1 international, locale-independent (one ST_NumberFormat each).
  ARABICABJAD: 'arabicAbjad', // ascending Abjad numerals
  ARABICALPHA: 'arabicAlpha', // Arabic alphabet
  HEBREW1: 'hebrew1', // Hebrew numerals (gematria)
  HEBREW2: 'hebrew2', // Hebrew alphabet
  HINDIARABIC: 'hindiNumbers', // Hindi numbers
  HINDILETTER1: 'hindiVowels', // Hindi vowels
  HINDILETTER2: 'hindiConsonants', // Hindi consonants
  THAIARABIC: 'thaiNumbers', // Thai numbers
  THAILETTER: 'thaiLetters', // Thai letters
  CHOSUNG: 'chosung', // Korean Chosung
  GANADA: 'ganada', // Korean Ganada
  DBCHAR: 'decimalFullWidth', // double-byte (full-width) Arabic
  SBCHAR: 'decimalHalfWidth', // single-byte (half-width) Arabic
};

/**
 * Parse a field instruction's general-formatting switch (§17.16.4.3.1) into an
 * ST_NumberFormat, or `null` when the instruction carries no numeric-format
 * switch this converter supports (no `\*`, or only `\* MERGEFORMAT`, or an
 * argument outside the native subset). The first supported `\*` argument wins;
 * unsupported ones (e.g. `MERGEFORMAT`) are skipped so `\* MERGEFORMAT \* Roman`
 * still resolves to `upperRoman`.
 */
export function parseFieldFormatSwitch(instruction: string): NumberFormat | null {
  // Match every `\* <arg>` occurrence; `<arg>` is a run of non-space chars
  // (switch arguments are single tokens — MERGEFORMAT, Roman, Arabic, …).
  const re = /\\\*\s+(\S+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(instruction)) !== null) {
    const fmt = ownValue(SWITCH_TO_FORMAT, m[1]);
    if (fmt) return fmt;
  }
  return null;
}
