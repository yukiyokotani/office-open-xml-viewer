import type { ParagraphAcquisitionInput } from './text.js';
import type { OriginalRtlArabicEofSource } from '../line-breaker/model.js';

/** Immutable lexical proof for registered word-ltr-rtl-arabic-eof-space.
 * The sealed source acquisition owns its lifetime, including a negative result.
 * Position/mode admission stays with the body acquisition, not this source fact. */
export function acquireOriginalRtlArabicEofSource(
  paragraph: ParagraphAcquisitionInput,
): OriginalRtlArabicEofSource | undefined {
  if (paragraph.runs.length !== 1) return undefined;
  const run = paragraph.runs[0]!;
  if (run.type !== 'text' || run.rtl !== true || !run.text.endsWith(' ')
    || run.ruby || ('optionalHyphen' in run && run.optionalHyphen) || run.fitTextVal != null
    || ('noBreakRanges' in run && run.noBreakRanges?.length)
    || ('noBreakBefore' in run && run.noBreakBefore) || ('noBreakAfter' in run && run.noBreakAfter)
    || run.noteRef || ('noteSeparatorCharacter' in run && run.noteSeparatorCharacter !== undefined)
    || run.eastAsianVert || run.smallCaps || run.allCaps
    || run.vertAlign != null || run.position || run.charSpacing || run.charScale != null) return undefined;
  // Closed source class: Arabic letters and authored U+0020 only. No marks,
  // digits, controls, mixed scripts or generated/transformed source windows.
  if (!/^(?:(?=\p{Script=Arabic})\p{Letter}| )+$/u.test(run.text) || !/\p{Letter}/u.test(run.text)) return undefined;
  return Object.freeze({ text: run.text, sequence: Object.freeze([
    Object.freeze({ runIndex: 0, start: 0, end: run.text.length }),
  ]) });
}
