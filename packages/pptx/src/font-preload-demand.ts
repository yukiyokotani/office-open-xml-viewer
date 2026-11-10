import { hasStrongRtl } from '@silurus/ooxml-core';
import { FONT_TRACKING_SENTINEL, FONT_BASELINE_SENTINEL, FONT_SPACE_SENTINEL } from '@silurus/ooxml-core/internal/font-measurement-sentinels';
import type { FontPreloadDemand, TextBody } from '@silurus/ooxml-core';
import { fontPreloadInertScalar } from './font-preload-inert-text.js';
import { preparedPowerPointText } from './font-display-text.js';
import type { Slide } from './types';

/** Optional loading policy, not OOXML rendering behavior. Admit only enumerated
 * horizontal shape/table paths. All generated/unmodelled output and canonical
 * changes retain all-rule loading. The global union deliberately over-includes
 * across families and run seams; it makes no resource ownership inference.
 * Sentinels are production tracking, baseline and tab measurements in renderer
 * and core DrawingML measurement/break paths: ii, M, and space. */
export function collectSlideFontDemand(
  slide: Slide,
  committed: ReadonlySet<number>,
  byteAllowance: number,
): FontPreloadDemand {
  const points = new Set<number>();
  // Even an empty committed delta retains the JSON array representation [].
  // Charge its two bytes before admitting repeated text with no new scalars.
  let bytes = 2;
  if (bytes > byteAllowance) return 'all';
  const add = (text: string): boolean => {
    // Main painting delegates UAX9 mirroring to Canvas. Reuse its strong-RTL
    // gate; unmodelled mirrored output retains all rules, not a guessed union.
    if (hasStrongRtl(text)) return false;
    for (const scalar of text) {
      const cp = scalar.codePointAt(0) as number;
      if (!fontPreloadInertScalar(cp) || (cp >= 0xf020 && cp <= 0xf0ff)) return false;
      if (!committed.has(cp) && !points.has(cp)) {
        bytes += String(cp).length + 1;
        if (bytes > byteAllowance) return false;
        points.add(cp);
      }
    }
    return true;
  };
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
        // A Symbol Latin face draws its ASCII text as other characters (Greek,
        // operators), so the authored scalars are not the painted ones.
        if ((run.fontFamily ?? para.defFontFamily)?.trim().toLowerCase() === 'symbol') return false;
        if (!add(run.text) || !add(preparedPowerPointText(run, slide.slideNumber))) return false;
      }
    }
    return true;
  };
  if (slide.parseError || !add(FONT_BASELINE_SENTINEL + FONT_TRACKING_SENTINEL + FONT_SPACE_SENTINEL)) return 'all';
  for (const element of slide.elements) {
    if (element.type === 'shape') { if (!body(element.textBody)) return 'all'; }
    else if (element.type === 'table') {
      for (const row of element.rows) for (const cell of row.cells) if (!body(cell.textBody)) return 'all';
    } else return 'all';
  }
  return [...points];
}
