import { FontDemandScalarCollector } from '@silurus/ooxml-core/internal/font-scalar-demand';
import type { FontPreloadDemand, TextBody } from '@silurus/ooxml-core';
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
  // The family the renderer paints a slot with (renderer normalizeFontFamily,
  // before embedded aliases): a theme reference resolves to the slide theme's
  // face, an absent face to its minor face, and a ",<generic>" hint is
  // dropped. Only whether it names a symbol-encoded family matters here.
  const painted = (family: string | null | undefined): string | null | undefined => {
    if (!family) return slide.theme?.minorFont;
    if (family.startsWith('+')) return family.startsWith('+mj-') ? slide.theme?.majorFont : slide.theme?.minorFont;
    return family.split(',')[0];
  };
  const body = (textBody: TextBody | null | undefined): boolean => {
    if (!textBody) return true;
    if (textBody.vert && textBody.vert !== 'horz' || 'textWarp' in textBody) return false;
    for (const para of textBody.paragraphs) {
      if (para.rtl) return false;
      // The parsed inherit sentinel produces no marker in resolveBulletLabel,
      // just like none. A char marker paints its code in its a:buFont
      // (§21.1.2.4.3; renderer paragraphBullet), converted like run text when
      // that face is symbol-encoded. AutoNum and blip markers remain unmodelled.
      const bullet = para.bullet;
      if (bullet?.type === 'char') {
        if (!scalars.addSymbolText(bullet.char, [painted(bullet.fontFamily)])) return false;
      } else if (bullet?.type !== 'none' && bullet?.type !== 'inherit') return false;
      for (const run of para.runs) {
        if (run.type === 'break') continue;
        if (run.type !== 'text' || run.fieldType && run.fieldType !== 'slidenum') return false;
        // A symbol-encoded Latin face (Symbol, Wingdings 1-3, Webdings) or the
        // a:sym face (§21.1.2.3.10) may draw its codes as other characters
        // (renderer paragraphInputRuns): every scalar that conversion may
        // paint is demanded (core addSymbolText).
        const families = [painted(run.fontFamily ?? para.defFontFamily), run.fontFamilySym && painted(run.fontFamilySym)];
        if (!scalars.addSymbolText(run.text, families)
          || !scalars.addSymbolText(preparedPowerPointText(run, slide.slideNumber), families)) return false;
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
