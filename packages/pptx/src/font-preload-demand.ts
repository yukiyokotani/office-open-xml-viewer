import {
  FontDemandScalarCollector,
  VERTICAL_PRESENTATION_FORMS,
  chartPaintTexts,
  paintsTextFreeImage,
} from '@silurus/ooxml-core/internal/font-scalar-demand';
import { symbolFontKey, type FontPreloadDemand, type TextBody } from '@silurus/ooxml-core';
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
 *   presentation forms an `eaVert` or stacked (`wordArtVert*`) body
 *   substitutes for corner punctuation and brackets (vertical-text.ts, core
 *   stacked-glyphs.ts), and text warps (the flat layout's glyphs mapped
 *   through the envelope);
 * - fields: their stored text, a slide number its actual number;
 * - char markers in their a:buFont, auto-number markers every label their
 *   scheme can format, in the marker face;
 * - images (pictures, media posters, picture fills, picture bullets): no text,
 *   unless a metafile (its text records paint with served faces) or of an
 *   unknown type (core `paintsTextFreeImage`);
 * - charts: every scalar their labels can contain (core `chartPaintTexts`).
 * Right-to-left paragraphs paint their texts reordered (UAX #9), mirrored
 * characters demanded by the core collector. Not modelled (`'all'`):
 * equations, region-map charts and unknown elements. */
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
  // Whether the family a slot paints with is symbol-encoded, per authored
  // name: a deck names few families and repeats them on every run.
  const symbolFaces = new Map<string, boolean>();
  const absentFaceIsSymbol = symbolFontKey(painted(null)) !== undefined;
  const symbolFace = (family: string | null | undefined): boolean => {
    if (!family) return absentFaceIsSymbol;
    let symbol = symbolFaces.get(family);
    if (symbol === undefined) symbolFaces.set(family, symbol = symbolFontKey(painted(family)) !== undefined);
    return symbol;
  };
  // Texts already collected as plain text on this slide. Collecting a text
  // twice demands nothing new (the collector's scalars, leading joiners and
  // flags are sets), and slides repeat texts (spaces, table values, labels).
  const plain = new Set<string>();
  const addPlain = (text: string): boolean => {
    if (plain.has(text)) return true;
    plain.add(text);
    return scalars.add(text);
  };
  /** `addSymbolText(text, [painted(family), sym && painted(sym)])`: without a
   * symbol-encoded face in either slot that is `add(text)`. */
  const add = (text: string, family: string | null | undefined, sym?: string | null): boolean => {
    if (symbolFace(family) || (!!sym && symbolFace(sym))) {
      return scalars.addSymbolText(text, [painted(family), sym && painted(sym)]);
    }
    return addPlain(text);
  };
  const body = (textBody: TextBody | null | undefined): boolean => {
    if (!textBody) return true;
    if (textBody.vert && !TEXT_VERTS.has(textBody.vert)) return false;
    // eaVert and stacked (wordArtVert*) text stand corner punctuation and
    // brackets upright as their vertical presentation forms (UAX #50 Tu/Tr;
    // vertical-text.ts, core stacked-glyphs.ts): the authored scalars
    // (measured) and every form core substitutes (painted). A text warp
    // (renderWarpedText) maps the flat layout's glyphs through its envelope,
    // painting the same texts.
    if (textBody.vert && VERTICAL_FORM_VERTS.has(textBody.vert) && !scalars.add(VERTICAL_PRESENTATION_FORMS)) return false;
    for (const para of textBody.paragraphs) {
      // The parsed inherit sentinel produces no marker in resolveBulletLabel,
      // just like none. A char marker paints its code in its a:buFont
      // (§21.1.2.4.3; renderer paragraphBullet), converted like run text when
      // that face is symbol-encoded. An auto-number label paints in its
      // a:buFont, else the first run's face; its counter is unbounded across
      // startAt values, so every scalar any scheme writes is demanded. A picture
      // marker paints no text (its image is checked with the element's).
      const bullet = para.bullet as TextBody['paragraphs'][number]['bullet'] | { type: 'blip' };
      if (bullet?.type === 'char') {
        if (!add(bullet.char, bullet.fontFamily)) return false;
      } else if (bullet?.type === 'autoNum') {
        const firstRun = para.runs.find((run) => run.type === 'text' && !!run.fontFamily);
        const face = bullet.fontFamily ?? (firstRun?.type === 'text' ? firstRun.fontFamily : null) ?? para.defFontFamily;
        if (!add(AUTONUM_ALPHABET, face)) return false;
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
        const family = run.fontFamily ?? para.defFontFamily;
        const display = preparedPowerPointText(run, slide.slideNumber);
        if (!add(run.text, family, run.fontFamilySym)
          || (display !== run.text && !add(display, family, run.fontFamilySym))) return false;
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
        return texts !== 'all' && texts.every(addPlain);
      }
      default: return false;
    }
  };
  if (scalars.isAll || slide.parseError) return 'all';
  if (slide.background && !textFreeImages(slide.background)) return 'all';
  if (!slide.elements.every(element)) return 'all';
  return scalars.result();
}

/** §20.1.10.83 ST_TextVerticalType. `vert` / `vert270` lay the body out
 * horizontally under a rotation and `mongolianVert` paints horizontally
 * (renderer renderTextBody), so each paints exactly its horizontal glyphs;
 * {@link VERTICAL_FORM_VERTS} add vertical forms. */
const TEXT_VERTS: ReadonlySet<string> = new Set(['horz', 'vert', 'vert270', 'mongolianVert', 'eaVert', 'wordArtVert', 'wordArtVertRtl']);
/** Vertical-type values whose painter substitutes vertical presentation
 * forms: eaVert (vertical-text.ts) and stacked text (renderStackedTextBody). */
const VERTICAL_FORM_VERTS: ReadonlySet<string> = new Set(['eaVert', 'wordArtVert', 'wordArtVertRtl']);

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
function textFreeImages(value: unknown): boolean {
  return value === null || typeof value !== 'object' || textFreeNode(value, 0);
}

/** The model is a JSON tree (parser output, structured-cloned): a plain
 * recursive walk without a visited set; the depth bound only guards against
 * a pathological hand-built model. The walk visits every node of every
 * element, so it reads each property once: an image path is found by its
 * key among the string values, not by probing every object for it. */
function textFreeNode(node: object, depth: number): boolean {
  if (depth > 64) return false;
  if (Array.isArray(node)) {
    for (let index = 0; index < node.length; index += 1) {
      const child: unknown = node[index];
      if (child !== null && typeof child === 'object' && !textFreeNode(child, depth + 1)) return false;
    }
    return true;
  }
  const record = node as Record<string, unknown>;
  for (const key in record) {
    const child = record[key];
    if (typeof child === 'string') {
      if (key === 'imagePath') {
        if (!paintsTextFreeImage(child, record.mimeType)) return false;
      } else if (key === 'posterPath' && child && !paintsTextFreeImage(child, record.posterMimeType)) return false;
    } else if (child !== null && typeof child === 'object' && key !== 'chart' && key !== 'mathFallback'
      && !textFreeNode(child, depth + 1)) return false;
  }
  return true;
}
