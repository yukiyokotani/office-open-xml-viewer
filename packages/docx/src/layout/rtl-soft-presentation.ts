import type { ParagraphAcquisitionInput } from './text.js';

/** Source proof for the registered Hebrew/numeric soft presentation rule.
 * The sealed store owns this once-only scan, including negative results. */
export interface RtlSoftPresentationSource {
  readonly end: number;
  readonly lastTab: number;
}
export function rtlSoftPresentationSource(paragraph: ParagraphAcquisitionInput): RtlSoftPresentationSource | undefined {
  let offset = 0, lastTab = -1;
  let hebrew = false, digit = false, pendingDecimal = false, decimalSeen = false, valid = true;
  let token: 'none' | 'hebrew' | 'numeric' = 'none';
  for (const run of paragraph.runs) {
    if (run.type !== 'text' || run.ruby || run.fitTextVal != null || run.charScale != null
      || run.charSpacing || run.position || run.vertAlign != null || run.eastAsianVert
      || run.smallCaps || run.allCaps || run.noteRef
      || ('noteSeparatorCharacter' in run && run.noteSeparatorCharacter !== undefined)
      || ('optionalHyphen' in run && run.optionalHyphen)
      || ('noBreakRanges' in run && run.noBreakRanges?.length)
      || ('noBreakBefore' in run && run.noBreakBefore) || ('noBreakAfter' in run && run.noBreakAfter)) return undefined;
    for (const character of run.text) {
      // Whole-host controls are outside this capability, including labels
      // before the final TAB. A later cell cannot erase their bidi context.
      if (character !== '\t' && /[\p{Cc}\p{Cf}]/u.test(character)) return undefined;
      if (character === '\t') {
        lastTab = offset; hebrew = digit = pendingDecimal = decimalSeen = false;
        token = 'none'; valid = true;
      } else if (lastTab >= 0) {
        // Token state survives authored formatting seams: only U+0020 or
        // TAB separates words. Numeric tokens have at most one interior dot.
        if (character >= '0' && character <= '9') {
          if (token === 'hebrew') valid = false;
          token = 'numeric'; digit = true; pendingDecimal = false;
        } else if (character >= '\u05d0' && character <= '\u05ea') {
          if (token === 'numeric' || pendingDecimal) valid = false;
          token = 'hebrew'; hebrew = true;
        } else if (character === '.' && token === 'numeric' && !decimalSeen && !pendingDecimal) {
          decimalSeen = pendingDecimal = true;
        } else if (character === ' ') {
          if (pendingDecimal) valid = false;
          token = 'none'; decimalSeen = pendingDecimal = false;
        } else valid = false;
      }
      offset += character.length;
    }
  }
  return lastTab >= 0 && valid && !pendingDecimal && hebrew && digit
    ? Object.freeze({ end: offset, lastTab }) : undefined;
}
