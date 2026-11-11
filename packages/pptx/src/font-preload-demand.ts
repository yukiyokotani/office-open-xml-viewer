import { FontDemandScalarCollector } from '@silurus/ooxml-core/internal/font-scalar-demand';
import { isSymbolFontFamily, type FontPreloadDemand, type TextBody } from '@silurus/ooxml-core';
import { preparedPowerPointText } from './font-display-text.js';
import type { Slide } from './types';

/** Optional loading policy, not OOXML rendering behavior. Admit only enumerated
 * horizontal shape/table paths. All generated/unmodelled output and canonical
 * changes retain all-rule loading. The global union deliberately over-includes
 * across families and run seams; it makes no resource ownership inference.
 * Scalar rules (strong RTL, symbol PUA, normalization-inert spelling,
 * measurement sentinels, budget) are core's shared collector; this function
 * owns only the PPTX traversal and its unmodelled paint paths. */
export function collectSlideFontDemand(
  slide: Slide,
  committed: ReadonlySet<number>,
  byteAllowance: number,
): FontPreloadDemand {
  const scalars = new FontDemandScalarCollector({ budget: byteAllowance, committed });
  const body = (textBody: TextBody | null | undefined): boolean => {
    if (!textBody) return true;
    if (textBody.vert && textBody.vert !== 'horz' || 'textWarp' in textBody) return false;
    for (const para of textBody.paragraphs) {
      if (para.rtl) return false;
      // The parsed inherit sentinel produces no marker in resolveBulletLabel,
      // just like none. Actual char/autoNum/blip markers remain unmodelled.
      if (para.bullet?.type !== 'none' && para.bullet?.type !== 'inherit') return false;
      for (const run of para.runs) {
        if (run.type === 'break') continue;
        if (run.type !== 'text' || run.fontFamilySym || run.fieldType && run.fieldType !== 'slidenum') return false;
        // A symbol-font Latin face (Symbol, Wingdings 1-3, Webdings) may draw its
        // text as other characters, so the authored scalars are not the painted ones.
        if (isSymbolFontFamily((run.fontFamily ?? para.defFontFamily)?.split(',')[0]?.trim())) return false;
        if (!scalars.add(run.text) || !scalars.add(preparedPowerPointText(run, slide.slideNumber))) return false;
      }
    }
    return true;
  };
  if (scalars.isAll || slide.parseError) return 'all';
  for (const element of slide.elements) {
    if (element.type === 'shape') { if (!body(element.textBody)) return 'all'; }
    else if (element.type === 'table') {
      for (const row of element.rows) for (const cell of row.cells) if (!body(cell.textBody)) return 'all';
    } else return 'all';
  }
  return scalars.result();
}
