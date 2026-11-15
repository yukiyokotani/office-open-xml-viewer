/**
 * XL5 furigana (phonetic-hint) geometry.
 *
 * Pure layout math for the ruby-like phonetic band Excel draws across the top
 * of a cell (ECMA-376 §18.4.6 `<rPh>` / §18.4.3 `<phoneticPr>`). Kept free of
 * any Canvas dependency (a `measure` callback is injected) so the placement is
 * unit-testable without a DOM.
 *
 * Not shared with the docx ruby renderer: docx ruby is laid out per line inside
 * the paragraph engine (grid pitch, ascent reserve, per-glyph baseline), while
 * xlsx furigana is placed relative to the CELL rectangle and the base string's
 * horizontal glyph advances. The only common idea ("small text above base") is
 * two lines of Canvas code, so a shared core primitive would be a
 * false-abstraction. The pure predicate here (offset → x-span) is xlsx-specific.
 */
import type { PhoneticAlignment, PhoneticRun } from './types.js';

/** A positioned phonetic hint ready to draw: the hint string plus the sheet-x
 *  span it occupies and how its characters should be spread inside that span. */
export interface PlacedPhonetic {
  /** The phonetic hint text (a single `<rPh>` run's reading). */
  text: string;
  /** Left edge (sheet-x, same space as the base text draw) of the hint band. */
  x: number;
  /** Width of the band the hint is laid out within. */
  width: number;
  /** How to distribute the hint's glyphs inside `[x, x+width]`:
   *  - `'start'`: draw from `x` (left-justified).
   *  - `'center'`: centre the natural hint width inside the band.
   *  - `'distribute'`: spread the glyphs so the first hugs `x` and the last
   *    hugs `x+width` (letter-spacing based). */
  spread: 'start' | 'center' | 'distribute';
}

/** Split a string into an array of user-perceived code points (so surrogate
 *  pairs — e.g. rare kanji — count as one base character, matching the spec's
 *  "character" offset semantics for `sb`/`eb`). */
export function toCodePoints(s: string): string[] {
  return Array.from(s);
}

/**
 * Place the per-word phonetic runs of a cell into sheet-x bands over the base
 * text. Handles the three PER-WORD alignments (§18.18.56 left / center /
 * distributed); `noControl` (NOT per word) is laid out by the caller with the
 * reading font, so it is not handled here.
 *
 * `measure(str)` returns the advance width of `str` in the BASE font at the
 * current scale — the same measurer the base-text draw uses — so the phonetic
 * band lines up with the glyphs it annotates; `start` is the code-point
 * offset of `str` in the base text, for a base painted in several fonts.
 * `baseLeftX` is the sheet-x where
 * the base text's first glyph starts.
 *
 * Runs whose `sb >= eb` or that fall outside the base length are skipped
 * (defensive — the spec requires `sb < eb` within bounds).
 */
export function placePhoneticRuns(
  runs: readonly PhoneticRun[],
  baseText: string,
  baseLeftX: number,
  alignment: Exclude<PhoneticAlignment, 'noControl'>,
  measure: (s: string, start: number) => number,
): PlacedPhonetic[] {
  const cps = toCodePoints(baseText);
  const n = cps.length;
  const out: PlacedPhonetic[] = [];

  for (const run of runs) {
    const sb = run.sb;
    const eb = run.eb;
    if (!(sb < eb) || sb >= n) continue;
    const clampedEb = Math.min(eb, n);
    const spanStart = baseLeftX + measure(cps.slice(0, sb).join(''), 0);
    const spanWidth = measure(cps.slice(sb, clampedEb).join(''), sb);
    const spread: PlacedPhonetic['spread'] =
      alignment === 'center' ? 'center'
      : alignment === 'distributed' ? 'distribute'
      : 'start';
    out.push({ text: run.text, x: spanStart, width: spanWidth, spread });
  }
  return out;
}
