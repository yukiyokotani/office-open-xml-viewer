import { defineCompatibilityRule } from './compatibility.js';

export const WORD_RUN_DIRECTION_ISOLATE = defineCompatibilityRule({
  id: 'word-run-direction-isolate',
  evidence: {
    kind: 'office-observation', syntheticFixtureId: 'bidi-run-direction-isolate',
    application: 'Microsoft Word', version: '16.0.20527', platform: 'Windows',
  },
  description: 'ECMA-376 §17.3.1.6 and §17.3.2.30 distinguish paragraph direction and run direction. Observed Word ordering, not a prescribed OOXML bidi algorithm: consecutive text runs whose authored direction opposes the paragraph resolve as one UAX #9 directional isolate. An unmarked Latin label sharing a run with Arabic in a bidi paragraph therefore stays at the physical left of that run. Windows controls cover Arabic, Hebrew, Latin and digits, modes unset/11/12/14/15, medium/high kashida and logical start/end alignments. Paragraph-direction runs, tabs and objects retain their existing resolution. Library scope and representation boundaries: only RTL paragraphs opt in; compatibility shape text without authoritative run-direction provenance, distribute/character-spread gap topology, and measured slices containing internal odd/even UAX level boundaries retain their existing ordering (#2076, #2090). These limitations are not Word-specific exceptions. This rule changes final ordering only, not line fitting or source-space budgets.',
});

export function wordRunDirectionSpan(segment: unknown): 'ltr' | 'rtl' | undefined {
  const { text, rtl } = segment as { text?: unknown; rtl?: unknown };
  return typeof text === 'string' ? rtl === true ? 'rtl' : 'ltr' : undefined;
}

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
