import { FontDemandScalarCollector } from '@silurus/ooxml-core/internal/font-scalar-demand';
import type { FontPreloadDemand, TextBody } from '@silurus/ooxml-core';
import { chartPaintTexts, paintsTextFreeImage } from './chart-font-demand.js';
import { preparedPowerPointText } from './font-display-text.js';
import type { Slide, SlideElement } from './types';

/** Optional loading policy, not OOXML rendering behavior. Collects every
 * scalar the renderer paints for the slide's enumerated paint paths; any
 * other output (unmodelled paths, canonical changes) keeps all-rule loading.
 * The global union deliberately over-includes across families and run seams;
 * it makes no resource ownership inference. Scalar rules (strong RTL, symbol
 * PUA, normalization-inert spelling, measurement sentinels, budget) are
 * core's shared collector; this function owns only the PPTX traversal and
 * its unmodelled paint paths.
 *
 * Paint paths and what they demand:
 * - shape and table text (master, layout and slide elements alike: the model
 *   flattens inherited placeholders and group children into `elements`), in
 *   horizontal, `vert`/`vert270` (the horizontal layout under a rotation) and
 *   `mongolianVert` (painted horizontally) bodies, plus the vertical
 *   presentation forms an `eaVert` body substitutes for corner punctuation
 *   and brackets (vertical-text.ts);
 * - fields: their stored text, a slide number its actual number;
 * - char markers in their a:buFont, auto-number markers every label their
 *   scheme can format, in the marker face;
 * - images (pictures, media posters, picture fills, picture bullets): no text,
 *   unless a metafile (its text records paint with served faces) or of an
 *   unknown type (chart-font-demand.ts `paintsTextFreeImage`);
 * - charts: every scalar their labels can contain (`chartPaintTexts`).
 * Not modelled (`'all'`): equations, WordArt warps, stacked (`wordArtVert*`)
 * text, right-to-left paragraphs, region-map charts and unknown elements. */
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
    if (textBody.vert && !HORIZONTAL_GLYPH_VERTS.has(textBody.vert) || 'textWarp' in textBody) return false;
    // eaVert stands corner punctuation and brackets upright as their vertical
    // presentation forms (UAX #50 Tu/Tr; vertical-text.ts): the authored
    // scalars (measured) and every form core substitutes (painted).
    if (textBody.vert === 'eaVert' && !scalars.add(VERTICAL_PRESENTATION_FORMS)) return false;
    const add = (text: string, families: readonly (string | null | undefined)[]): boolean =>
      scalars.addSymbolText(text, families);
    for (const para of textBody.paragraphs) {
      if (para.rtl) return false;
      // The parsed inherit sentinel produces no marker in resolveBulletLabel,
      // just like none. A char marker paints its code in its a:buFont
      // (§21.1.2.4.3; renderer paragraphBullet), converted like run text when
      // that face is symbol-encoded. An auto-number label paints in its
      // a:buFont, else the first run's face; its counter is unbounded across
      // startAt values, so every scalar any scheme writes is demanded. A picture
      // marker paints no text (its image is checked with the element's).
      const bullet = para.bullet as TextBody['paragraphs'][number]['bullet'] | { type: 'blip' };
      if (bullet?.type === 'char') {
        if (!add(bullet.char, [painted(bullet.fontFamily)])) return false;
      } else if (bullet?.type === 'autoNum') {
        const firstRun = para.runs.find((run) => run.type === 'text' && !!run.fontFamily);
        const face = bullet.fontFamily ?? (firstRun?.type === 'text' ? firstRun.fontFamily : null) ?? para.defFontFamily;
        if (!add(AUTONUM_ALPHABET, [painted(face)])) return false;
      } else if (bullet?.type !== 'none' && bullet?.type !== 'inherit' && bullet?.type !== 'blip') return false;
      for (const run of para.runs) {
        if (run.type === 'break') continue;
        if (run.type !== 'text') return false;
        // A symbol-encoded Latin face (Symbol, Wingdings 1-3, Webdings) or the
        // a:sym face (§21.1.2.3.10) may draw its codes as other characters
        // (renderer paragraphInputRuns): every scalar that conversion may
        // paint is demanded (core addSymbolText). Every field paints its
        // stored text, a slide-number field its actual number
        // (preparedPowerPointText, the renderer's display text).
        const families = [painted(run.fontFamily ?? para.defFontFamily), run.fontFamilySym && painted(run.fontFamilySym)];
        const display = preparedPowerPointText(run, slide.slideNumber);
        if (!add(run.text, families) || (display !== run.text && !add(display, families))) return false;
      }
    }
    return true;
  };
  const element = (el: SlideElement): boolean => {
    if (!textFreeImages(el)) return false;
    // An equation's MCE fallback paints only when its conversion fails; the
    // equation run itself declines, the fallback is collected regardless.
    if ((el.type === 'shape' || el.type === 'table') && el.mathFallback
      && !el.mathFallback.elements.every(element)) return false;
    switch (el.type) {
      case 'shape': return body(el.textBody);
      case 'table': return el.rows.every((row) => row.cells.every((cell) => body(cell.textBody)));
      case 'picture':
      case 'media': return true;
      case 'chart': {
        const texts = chartPaintTexts(el.chart);
        return texts !== 'all' && texts.every((text) => scalars.add(text));
      }
      default: return false;
    }
  };
  if (scalars.isAll || slide.parseError) return 'all';
  if (slide.background && !textFreeImages(slide.background)) return 'all';
  if (!slide.elements.every(element)) return 'all';
  return scalars.result();
}

/** `vert` / `vert270` lay the body out horizontally under a rotation and
 * `mongolianVert` paints horizontally (renderer renderTextBody), so each
 * paints exactly its horizontal glyphs; `eaVert` adds vertical forms. */
const HORIZONTAL_GLYPH_VERTS: ReadonlySet<string> = new Set(['horz', 'vert', 'vert270', 'mongolianVert', 'eaVert']);

/** Every vertical presentation form core substitutes in eaVert text
 * (`verticalFormSubstitute` → U+FE10–FE19, `verticalBracketFormSubstitute`
 * → U+FE17/FE18 and U+FE35–FE44): the two blocks, a small superset. Kept
 * static so the collector chunk shares no module with the renderer. */
export const VERTICAL_PRESENTATION_FORMS = String.fromCharCode(
  ...Array.from({ length: 10 }, (_, i) => 0xfe10 + i),
  ...Array.from({ length: 0x20 }, (_, i) => 0xfe30 + i),
);

/** Every scalar `formatAutoNum` (autonum-label.ts) can write for any scheme
 * and counter: Arabic digits and sign, Latin letters (alphabetic and roman
 * schemes), the scheme punctuation `.()`, and full-width digits (arabicDb*).
 * Printable ASCII is a small superset; kept static so the collector chunk
 * shares no module with the renderer. */
export const AUTONUM_ALPHABET = `${String.fromCharCode(...Array.from({ length: 0x7f - 0x20 }, (_, i) => 0x20 + i))}${
  String.fromCharCode(...Array.from({ length: 10 }, (_, i) => 0xff10 + i))}`;

/** Whether every image an element (or the slide background) paints is free of
 * served-face text: any `imagePath` (picture, picture fill, picture bullet)
 * and a media poster. Chart fills are checked by chartPaintTexts, equation
 * fallbacks as elements of their own. */
function textFreeImages(value: unknown, depth = 0): boolean {
  // The model is a JSON tree (parser output, structured-cloned): a plain
  // recursive walk without a visited set; the depth bound only guards
  // against a pathological hand-built model.
  if (value === null || typeof value !== 'object') return true;
  if (depth > 64) return false;
  const record = value as Record<string, unknown>;
  if (typeof record.imagePath === 'string' && !paintsTextFreeImage(record.imagePath, record.mimeType)) return false;
  if (typeof record.posterPath === 'string' && record.posterPath
    && !paintsTextFreeImage(record.posterPath, record.posterMimeType)) return false;
  for (const key in record) {
    if (key === 'chart' || key === 'mathFallback') continue;
    const child = record[key];
    if (child !== null && typeof child === 'object' && !textFreeImages(child, depth + 1)) return false;
  }
  return true;
}
