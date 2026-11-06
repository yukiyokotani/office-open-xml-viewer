import { isComplexScriptCodePoint, type KinsokuRules } from '@silurus/ooxml-core';
import { bidiClass } from '@silurus/ooxml-core/internal/bidi-line';
import type { DocumentLayoutSettings, SectionLayoutContext, StoryContext } from '../layout-context.js';
import type { ParagraphLayoutSource, ParagraphTextBearingRun } from './text.js';
import { COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION } from '../line-breaker/text-runs.js';



const metricSlots = ['ascii', 'highAnsi', 'eastAsia', 'complexScript'] as const;
type OrdinaryTextRun = Extract<ParagraphTextBearingRun, { type: 'text' }>;
function sameFontSlots(left: OrdinaryTextRun, right: OrdinaryTextRun): boolean {
  const a = left.fontSlots, b = right.fontSlots;
  if (!a || !b) return a === b;
  return metricSlots.every(slot => a.direct[slot] === b.direct[slot]
    && a.theme[slot] === b.theme[slot] && a.themePresent[slot] === b.themePresent[slot]);
}

const disabledViews = new WeakMap<KinsokuRules, KinsokuRules>();
const japaneseEdges = new WeakMap<KinsokuRules, KinsokuRules>();

/** ECMA-376 §17.3.1.16: explicit paragraph OFF disables its forbidden edges.
 * Share the immutable view; do not mutate the document's inherited tables.
 * Existing intrinsic atoms consult these retained tables regardless of enabled;
 * this line-edge fix deliberately does not change that separate sizing contract. */
export function paragraphKinsokuDisabled(rules: KinsokuRules): KinsokuRules {
  if (!rules.enabled) return rules;
  let result = disabledViews.get(rules);
  if (!result) {
    result = Object.freeze({ ...rules, enabled: false });
    disabledViews.set(rules, result);
  }
  return result;
}

/** word-paragraph-off-japanese-line-edges (WORD_PARAGRAPH_OFF_JAPANESE_LINE_EDGES). Registered policy for Japanese Thai-distributed paragraph-OFF
 * text in unset/11/12/14 with explicit punctuation compression still forbids
 * iteration/kana heads and opening tails. Overflow permits only the existing
 * ten full-width closing heads; the kana compression option does not enlarge
 * that permission. Chinese/Korean and mixed-language controls differ.
 * This bounded edge channel MUST NOT feed hanging, compression grouping,
 * intrinsic atoms or paragraph semantics. Unmeasured hosts stay normative OFF. */
export function paragraphOffLineEdges(
  settings: DocumentLayoutSettings, section: SectionLayoutContext,
  story: StoryContext, paragraph: ParagraphLayoutSource,
): KinsokuRules | undefined {
  const mode = settings.compat.compatibilityMode;
  if (paragraph.kinsoku !== false || !settings.kinsoku.enabled
    || (mode !== undefined && mode !== 11 && mode !== 12 && mode !== 14)
    || (settings.characterSpacingControl !== 'compressPunctuation'
      && settings.characterSpacingControl !== 'compressPunctuationAndJapaneseKana')
    || settings.hasCustomKinsokuRules || paragraph.alignment !== 'thaiDistribute'
    || story.story !== 'body' || story.containers.length !== 0 || story.quarterTurned
    || paragraph.framePr != null || paragraph.bidi === true || paragraph.wordWrap === false
    || section.textDirection !== 'lrTb' || section.sectionFlow !== undefined || section.grid.kind !== 'none'
    || paragraph.tabStops.length !== 0 || paragraph.runs.length === 0) return undefined;
  let language: string | undefined;
  let initialized = false, japanese = false;
  let first: OrdinaryTextRun | undefined;
  for (const run of paragraph.runs) {
    if (run.type !== 'text' || run.rtl === true || run.cs === true || run.ruby
      || run.vertAlign != null || run.smallCaps || run.allCaps
      || run.fitTextVal != null || run.position != null
      || ('optionalHyphen' in run && run.optionalHyphen)
      || run.noteRef !== undefined
      || ('noteSeparatorCharacter' in run && run.noteSeparatorCharacter !== undefined)
      || ('noBreakBefore' in run && run.noBreakBefore)
      || ('noBreakAfter' in run && run.noBreakAfter)
      || ('noBreakRanges' in run && Boolean(run.noBreakRanges?.length))
      || (run.charScale !== undefined && run.charScale !== 1)
      || (run.charSpacing !== undefined && run.charSpacing !== 0)) return undefined;
    // Public model runs need not carry parser-boundary acquisition facts.
    // Preserve their absence rather than asserting or inventing those fields.
    const currentLanguage = 'langEastAsia' in run ? run.langEastAsia?.toLowerCase() : undefined;
    if (currentLanguage !== undefined && currentLanguage !== 'ja-jp') return undefined;
    if (initialized && currentLanguage !== language) return undefined;
    initialized = true; language = currentLanguage;
    if (first && (run.fontSize !== first.fontSize || run.fontFamily !== first.fontFamily
      || run.fontFamilyEastAsia !== first.fontFamilyEastAsia
      || ('fontFamilyHighAnsi' in run ? run.fontFamilyHighAnsi : undefined) !== first.fontFamilyHighAnsi
      || ('fontHint' in run ? run.fontHint : undefined) !== first.fontHint
      || run.kerning !== first.kerning || !sameFontSlots(first, run)
      || run.bold !== first.bold
      || run.italic !== first.italic || run.underline !== first.underline
      || run.strikethrough !== first.strikethrough)) return undefined;
    first = run;
    japanese ||= currentLanguage === 'ja-jp' || /[\p{Script=Hiragana}\p{Script=Katakana}]/u.test(run.text);
    for (const character of run.text) {
      const code = character.codePointAt(0)!;
      const type = bidiClass(code);
      if (isComplexScriptCodePoint(code) || /[\p{Cc}\p{Cf}]/u.test(character) || type === 'R' || type === 'AL' || type === 'AN') return undefined;
    }
  }
  if (!japanese) return undefined;
  if (paragraph.overflowPunct === false) return settings.kinsoku;
  let result = japaneseEdges.get(settings.kinsoku);
  if (!result) {
    const heads = new Set(settings.kinsoku.lineStartForbidden);
    for (const character of COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION) heads.delete(character.codePointAt(0)!);
    result = Object.freeze({ ...settings.kinsoku, lineStartForbidden: heads });
    japaneseEdges.set(settings.kinsoku, result);
  }
  return result;
}
