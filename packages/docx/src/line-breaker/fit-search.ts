import type { MeasurementTextContext, VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';
import { EAST_ASIAN_RE } from '../layout/text.js';
import { textAdvanceWidth } from './advance.js';
import { verticalRunInkExtra } from './vertical-text.js';

/** Project the registered line-end allowance from the immediately preceding
 * visible East-Asian character, not from another character elsewhere in a
 * mixed-script segment. */
export function hasEastAsianVisiblePredecessor(text: string): boolean {
  const characters = [...text];
  for (let index = characters.length - 1; index >= 0; index -= 1) {
    if (characters[index] === '\u3000') continue;
    return EAST_ASIAN_RE.test(characters[index]);
  }
  return false;
}

/** Numerical tolerance of every line-fit admission, in device px. Line bands
 * are differences of authored twip geometry carried as binary64 pt values
 * (e.g. 595.3 - 56.7 - 56.7 - 361.9 = 119.99999999999994) and advances are sums of scaled
 * glyph widths, so content exactly as wide as its band can compare a few ulps
 * wider. An exact fit fits (issue #2102); this is round-off slack only, many
 * orders of magnitude below one twip, never an Office width allowance. */
export const LINE_FIT_EPSILON_PX = 1e-9;

/** The one line-fit predicate: `used` fits `available` at exact equality. */
export function fitsLineWidth(used: number, available: number): boolean {
  return used <= available + LINE_FIT_EPSILON_PX;
}

export function fitCJKPrefix(
  ctx: MeasurementTextContext,
  text: string,
  maxWidth: number,
  // ECMA-376 §17.6.5 character-grid delta (px per EA glyph, 0 when inactive).
  // The fit must compare the same advance model as the line box / draw so the
  // grid's char count and run character metrics land on the same split.
  gridDeltaPx = 0,
  // WD4 — the run's §17.3.2.43 horizontal glyph scale (1 = 100%) and §17.3.2.35
  // per-code-point character-spacing pitch in px. Threaded so a CJK run that is
  // scaled/spaced splits at the SAME cell boundary the whole-segment advance
  // model uses (measure==paint). Default (1, 0) reproduces the prior behaviour.
  charScale = 1,
  charSpacingPx = 0,
  // issue #1014 — a vertical (tbRl) run whose segment is flagged `verticalRun`:
  // fold the vo=Tr rotate-fallback ink deficit into the fit predicate too, so the
  // wrap chooses a prefix whose CORRECTED advance (the same the line box measures)
  // fits — not one that only fits by the under-reported raw width. 0 for horizontal
  // / non-under-reporting runs, so the split is byte-identical there.
  verticalRun = false,
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  /** Optional caller-owned advance authority. Production layout supplies the
   * same substring measurement used by whole-segment fit so every prefix uses
   * the canonical selected-face and OOXML pitch model. */
  measureAdvance?: (text: string) => number,
  /** Maximum trailing U+3000 characters excluded from the fit width. */
  maximumIdeographicSpaceHang = Number.POSITIVE_INFINITY,
): string {
  const chars = [...text]; // spread handles surrogate pairs
  const advanceOf = (prefix: string): number => measureAdvance?.(prefix) ?? (() => {
    const verticalExtraPx = verticalRunInkExtra(prefix, verticalRun, verticalGlyphMeasurement);
    return textAdvanceWidth(
      ctx.measureText(prefix).width + verticalExtraPx,
      prefix,
      gridDeltaPx,
      charScale,
      charSpacingPx,
    );
  })();
  // Trailing IDEOGRAPHIC SPACE (U+3000) line-end allowance: a candidate that
  // overflows ONLY because it ends in fullwidth spaces still fits — those spaces
  // hang past the line end (JLReq line-end ideographic-space handling; Word
  // does the same, which is what keeps a "char + U+3000" form label at one
  // visible glyph per line instead of alternating glyph/space lines). The
  // accepted range KEEPS the trailing spaces, so the next line starts at the
  // following visible character. Scope: trailing U+3000 in the candidate only —
  // leading/interior fullwidth spaces stay width-bearing (authored indents),
  // and ASCII-space handling is a separate, untouched mechanism. The predicate
  // stays monotone in the candidate length (appending a U+3000 never changes
  // the visible advance; appending a visible char only grows it), so the
  // binary search remains valid.
  const fitsWithHang = (endExclusive: number): boolean => {
    let visibleEnd = endExclusive;
    let remainingHang = maximumIdeographicSpaceHang;
    if (remainingHang > 0) {
      while (visibleEnd > 0 && chars[visibleEnd - 1] === '\u3000') visibleEnd--;
      const trailingCount = endExclusive - visibleEnd;
      visibleEnd += Math.max(0, trailingCount - remainingHang);
    }
    const prefix = chars.slice(0, visibleEnd).join('');
    return fitsLineWidth(advanceOf(prefix), maxWidth);
  };
  let lo = 0, hi = chars.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (fitsWithHang(mid)) lo = mid;
    else hi = mid - 1;
  }
  return chars.slice(0, lo).join('');
}
