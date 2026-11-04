import { defineCompatibilityRule } from './compatibility.js';

export const WORD_NEUTRAL_SCRIPT_ATTACHMENT = defineCompatibilityRule({
  id: 'word-neutral-script-attachment',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/compatibility.test.ts#keeps neutral characters attached to the active script slice',
  },
  description: 'Weak and neutral non-letter characters stay with the active complex-script slice instead of opening additional formatting segments. Library policy, no Office measurement: neutral characters stay in the active script slice (PR #368); [MS-OI29500] §2.1.88 does not state it.',
});

export const WORD_RTL_RUN_AMBIGUOUS_CLASS_OVERRIDE = defineCompatibilityRule({
  id: 'word-rtl-run-ambiguous-class-override',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/bidi-line.test.ts#keeps LTR word order for English text in rtl-marked runs',
  },
  description: 'Model an rtl-marked run as a higher-level UAX #9 override for punctuation and symbols only, leaving whitespace and strong letters at their ordinary classes. Evidence (observed Office behaviour): issue #366 corpus document and PR #368 (a UAX #9 higher-level protocol); no synthetic control yet.',
});

export const WORD_RTL_COMPLEX_SCRIPT_EUROPEAN_DIGITS_AN = defineCompatibilityRule({
  id: 'word-rtl-complex-script-european-digits-an',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/bidi-line.test.ts#orders an AN-classified date as 2026-02-28',
  },
  description: 'Classify European digits as Arabic Number within an Arabic or Hebrew complex-script run so UAX #9 preserves the compatible visual ordering of digit groups and separators. Library policy, no Office measurement: European digits classified as Arabic Number in RTL complex-script runs (PR #368); asserted, not measured.',
});

export const WORD_KASHIDA_FINAL_FORM_PRIORITY = defineCompatibilityRule({
  id: 'word-kashida-final-form-priority',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/kashida-priority.test.ts#uses the BaRa join (Beh->Yeh) over the final-letter join in بين',
  },
  description: 'Apply the measured kashida final-letter priority classes only at a word-final following letter instead of copying the broader Qt final-form conditions. Evidence (observed Office behaviour): issue #954, with adjudication in PR #943 and PR #970.',
});

export const WORD_VERTICAL_TU_CORNER_PLACEMENT = defineCompatibilityRule({
  id: 'word-vertical-tu-corner-placement',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/vertical-text.test.ts#does NOT ink-centre a substituted Tu comma even when ink metrics are present',
  },
  description: 'Keep a substituted vertical Tu comma or full stop at the font-designed upper-right cell position rather than ink-centering it geometrically. Evidence (observed Office behaviour): Word PDF of a corpus document in PR #793; follows the UAX #50 Tu class and JIS X 4051 corner placement.',
});

export const WORD_RUN_DIRECTION_ISOLATE = defineCompatibilityRule({
  id: 'word-run-direction-isolate',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-bidi-order',
    application: 'Microsoft Word',
    version: '16.0.20527.20088',
    platform: 'Windows',
  },
  description: 'A maximal stretch of text whose run direction (§17.3.2.30 w:rtl present: right-to-left; absent or false: left-to-right) is opposite to the paragraph direction (§17.3.1.6 w:bidi) is ordered as one directional isolate (UAX #9 LRI/RLI ... PDI): it moves as a unit and its own content is resolved in the run direction, regardless of the scripts it contains. Windows Word controls blk-0898-bidi-order (modes unset and 15, every jc including start/end, Arial 12, distinct Arabic, Hebrew, Latin and digit words): in a bidi paragraph one non-rtl run "Xray A1 A2 A3" draws Xray at the physical left with the Arabic words read right to left beside it (UAX #9 at paragraph level 1 puts Xray at the right); "A1 A2 Xray A3 A4" non-rtl draws A2 A1 Xray A4 A3, "A1 A2 Xray" draws A2 A1 Xray, "A1 A2 A3." keeps the full stop at the right, and w:cs or lang bidi on the run change nothing; the same text in rtl runs follows UAX #9 at level 1, a non-rtl Latin run beside rtl Arabic runs keeps its logical place (logical start at the right) with its own spaces inside it, and in an LTR paragraph one rtl run "Xray A1 Yoke" draws Yoke A1 Xray. Text in the paragraph direction, tabs and inline objects keep the line-level UAX #9 resolution with the WORD_RTL_RUN_AMBIGUOUS_CLASS_OVERRIDE; the order does not depend on jc or compatibility mode. Scope: a line with right-to-left text that is spread per character (WORD_LINE_END_JUSTIFICATION, pre-15 distribute) keeps the line-level UAX #9 order and line-end geometry, because the renderer\'s gap model for such lines does not place the gaps of an isolate (blk-0898-bidi-order m04 distribute, mode unset); WORD_LINE_END_SPACE_STAYS_IN_RUN is scoped out there too.',
});

export const WORD_LINE_END_SPACE_STAYS_IN_RUN = defineCompatibilityRule({
  id: 'word-line-end-space-stays-in-run',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'blk-0898-bidi-fit',
    application: 'Microsoft Word',
    version: '16.0.20527.20088',
    platform: 'Windows',
  },
  description: 'Word does not apply UAX #9 L1 to the U+0020 that ends a line: the space keeps the level of its run and is drawn on that run\'s side, at a soft wrap and before the paragraph mark alike. On a left-to-right line with right-to-left text a soft-wrapped line-end space drawn at a band edge hangs past that edge, also on unjustified lines; on a last line it stays inside the band while it fits and only the overflow hangs (WORD_RTL_JUSTIFIED_VISIBLE_FIT does the same for justified right-to-left lines). Paint: a slice is drawn in the direction its closing U+0020 resolves to when that differs from its ordering level, and every slice of that level run follows, so each gap lands between the words where UAX #9 puts it (the slices of one run share the space advance). Windows Word controls blk-0898-bidi-fit and blk-0898-bidi-order (modes 14, 15 and unset, Arial and Times New Roman 12): in an LTR paragraph of rtl Arabic runs every soft-wrapped line starts its ink at the band start (56.6pt) and a justified one ends at 538.0pt in a 56.7-538.6pt band, while the space ending the paragraph is drawn at the left of the last line (ink from 60.0pt), and in mode 14 a last line that overflows only by that space narrows it (ink from 57.1, 58.1 and 59.1pt as the band narrows by 1pt steps); the space ending a bidi paragraph inside a non-rtl Latin run is drawn at the right (ink to 535.6pt at a 538.6pt band end); in a bidi paragraph the non-rtl run "A1 A2 Xray" is drawn A2, A1, Xray with one space between each word, and an rtl run "Xray " before rtl Arabic keeps its space between Xray and the Arabic.',
});

/** Span direction of a text segment for {@link WORD_RUN_DIRECTION_ISOLATE}:
 *  its run direction; tabs and objects belong to no span. */
export function wordRunDirectionSpan(segment: unknown): 'ltr' | 'rtl' | undefined {
  const { text, rtl } = segment as { text?: unknown; rtl?: unknown };
  if (typeof text !== 'string') return undefined;
  return rtl === true ? 'rtl' : 'ltr';
}

const RTL_AMBIGUOUS_CHARACTER = /[\p{P}\p{S}]/u;

export function wordRtlAmbiguousCharacter(character: string): boolean {
  return RTL_AMBIGUOUS_CHARACTER.test(character);
}

export function wordNeutralAttachesToActiveScript(character: string): boolean {
  return !/\p{L}/u.test(character);
}

export function wordKashidaFinalFormApplies(
  beforeCodePointIndex: number,
  lastLetterCodePointIndex: number,
): boolean {
  return beforeCodePointIndex === lastLetterCodePointIndex;
}

export function wordPreservesVerticalTuCorner(
  substitutedPunctuationCodePoint: number | null,
): boolean {
  return substitutedPunctuationCodePoint !== null;
}
