import { referenceFontCoversSymbol, referenceFontCoversCjk, referenceFontSupportFacts } from './font-resource-catalogue.js';
import { analyzeFontResourceSupport, type ResourceSupport } from '@silurus/ooxml-core/internal/font-cluster-coverage';
import { findReferenceFontMetrics, type OpenTypeLineMetrics } from '@silurus/ooxml-core';
import { excelDrawingMlLineRatios } from '@silurus/ooxml-core/internal/office-auto-line';

/**
 * PowerPoint's split of a text line at its baseline (observed behaviour; the
 * DrawingML text model in ECMA-376 §21.1.2 does not say where the baseline
 * sits inside a line).
 *
 * PowerPoint for Mac 16.113.2 PDF controls (#1610). Two decks with 432
 * boxes, every face verified from the embedded PDF fonts:
 * - The line box is 1.2 × the largest run size for every face.
 * - A run's ascent share of that box is usWinAscent / (usWinAscent +
 *   usWinDescent). A face that sets fsSelection USE_TYPO_METRICS uses
 *   (sTypoAscender + sTypoLineGap) / (sTypoAscender − sTypoDescender +
 *   sTypoLineGap) instead.
 *   - Faces that separate the tables: Yu Gothic (usWin 0.765, hhea 0.798),
 *     Baskerville Old Face, Stencil, and Gabriola (typo 0.814, usWin 0.647).
 *   - Unlike Excel, there is no Far East 1.3× leading and no dependence on
 *     the face's hhea table.
 * - Font resolution follows the face PowerPoint really used, not the name:
 *   - Office-bundled faces and the macOS Supplemental faces (Times New
 *     Roman, Georgia, Courier New) are used as named. A face present in both
 *     resolves to the Supplemental copy (Times New Roman's hhea lineGap 87 is
 *     the one embedded).
 *   - PowerPoint substitutes faces that only macOS /System/Library/Fonts
 *     provides:
 *     - Palatino → Palatino Linotype;
 *     - Helvetica and Helvetica Neue → Arial (embedded under the Helvetica
 *       name with Arial's glyphs and usWin metrics);
 *     - Avenir, Menlo and Hiragino Sans → the deck's theme font or a Far
 *       East fallback.
 *   - Only the first two mappings are pinned. Every other system-only face
 *     returns undefined, and the caller keeps its ordinary line model.
 */
const OBSERVED_SUBSTITUTES: Readonly<Record<string, string>> = {
  palatino: 'Palatino Linotype',
  helvetica: 'Arial',
  'helvetica neue': 'Arial',
};

/**
 * Resolved shares keyed by face/weight/style. Font names come from document
 * content, so the cache is a bounded LRU: at most SHARE_CACHE_LIMIT entries,
 * the least recently used evicted first. A miss only re-reads the static
 * reference table, so the cap trades a little lookup work for bounded memory.
 */
export const SHARE_CACHE_LIMIT = 256;
const shareCache = new Map<string, number | null>();

/** @internal Test hook: current number of cached face keys. */
export function powerPointShareCacheSize(): number {
  return shareCache.size;
}

export function powerPointAscentShare(
  family: string,
  bold: boolean,
  italic: boolean,
): number | undefined {
  const key = `${family.trim().toLocaleLowerCase('en-US')}|${bold ? 700 : 400}|${italic ? 'i' : 'n'}`;
  const cached = shareCache.get(key);
  if (cached !== undefined) {
    shareCache.delete(key);
    shareCache.set(key, cached);
    return cached ?? undefined;
  }
  const share = resolveShare(family, bold, italic);
  shareCache.set(key, share ?? null);
  if (shareCache.size > SHARE_CACHE_LIMIT) {
    const oldest = shareCache.keys().next().value;
    if (oldest !== undefined) shareCache.delete(oldest);
  }
  return share;
}

type Profile = ReturnType<typeof findReferenceFontMetrics>[number];

/** The profiles of the copy PowerPoint lays the face out with (see above). */
function chosenProfiles(family: string, bold: boolean, italic: boolean, repertoire = false): Profile[] {
  const trimmed = family.trim();
  if (!trimmed) return [];
  const name = OBSERVED_SUBSTITUTES[trimmed.toLocaleLowerCase('en-US')] ?? trimmed;
  const style = italic ? 'italic' : 'normal';
  let profiles = findReferenceFontMetrics(name, { weight: bold ? 700 : 400, style });
  if (profiles.length === 0) {
    // A face name that is itself one cut of a family ("Calibri Light" is
    // weight 300) has no 400 or 700 profile; use that cut when the name
    // selects a single weight. #1630 controls (reference PDF export): Calibri
    // Light titles sat at the 1950 / 2500 usWin share, like Calibri. A bold
    // Calibri Light title embedded Calibri-Light itself (#1435 deck).
    const named = findReferenceFontMetrics(name, { style });
    if (named.length > 0 && named.every((p) => p.weight === named[0].weight)) profiles = named;
  }
  // A catalogued family with no italic face at any weight is drawn by
  // PowerPoint as its upright face with a synthetic slant (#1689: MS Gothic,
  // Tahoma and Microsoft Sans Serif italics embed the regular resource under a
  // [1 0 0.3333 1] text matrix). A shear leaves the vertical OS/2 metrics
  // intact, so the upright resource's metrics apply. The catalogue lists every
  // installed static face, so a missing italic profile is a missing resource,
  // not a gap in the data. Painting keeps the browser's own oblique for the
  // missing style: an accepted platform difference (owner decision (c) for
  // #1689), not an emulated 0.3333 shear.
  if (profiles.length === 0 && italic && findReferenceFontMetrics(name, { style: 'italic' }).length === 0) {
    return chosenProfiles(family, bold, false, repertoire);
  }
  // Empty-slot repertoire evidence is from Office's resources (#1689), as
  // is the existing CJK coverage catalogue. Prefer that source for cmap facts;
  // same-name system copies can omit symbols despite identical line tables.
  // Metric source precedence remains the independently measured #1610 rule.
  if (repertoire) {
    const office = profiles.filter((p) => p.source === 'office-mac');
    if (office.length > 0) return office;
  }
  const supplemental = profiles.filter((p) => p.source === 'macos-supplemental');
  return supplemental.length > 0 ? supplemental : profiles.filter((p) => p.source === 'office-mac');
}

/** Symbol presence in the Office repertoire's real/synthetic cut (see source
 * precedence above). Conflicting profiles and unrecorded faces stay unknown;
 * unioning cuts would wrongly attribute a missing italic glyph to its regular. */
export function powerPointSymbolCoverage(
  family: string, bold: boolean, italic: boolean, codePoint: number,
): boolean | undefined {
  const chosen = chosenProfiles(family, bold, italic, true);
  if (chosen.length === 0) return undefined;
  const first = referenceFontCoversSymbol(chosen[0], codePoint);
  return first !== undefined && chosen.every((p) => referenceFontCoversSymbol(p, codePoint) === first)
    ? first : undefined;
}

/** Empty-EA attribution uses the same bounded parsed certificates as
 * embedded resources. Named installed slots retain the established catalogue
 * metric policy; this is not detection of a runtime installed FontFace. */
export function powerPointCatalogueSupport(
  family: string, bold: boolean, italic: boolean, display: string,
  coverage: typeof powerPointSymbolCoverage,
): ResourceSupport {
  const profiles = chosenProfiles(family, bold, italic, true);
  if (!profiles.length) return { kind: 'unknown', reason: 'font-transform' };
  let result: ResourceSupport | undefined;
  for (const profile of profiles) {
    const support = analyzeFontResourceSupport(display, (cp) => coverage(family, bold, italic, cp), referenceFontSupportFacts(profile));
    if (support.kind === 'unknown' || (result && result.kind !== support.kind)) return { kind: 'unknown', reason: 'font-transform' };
    result = support;
  }
  return result as ResourceSupport;
}

/** CJK presence in the same real/synthetic resource cut used for metrics. */
export function powerPointCjkCoverage(
  family: string, bold: boolean, italic: boolean, codePoint: number,
): boolean | undefined {
  const chosen = chosenProfiles(family, bold, italic, true);
  if (chosen.length === 0) return undefined;
  const first = referenceFontCoversCjk(chosen[0], codePoint);
  return first !== undefined && chosen.every((p) => referenceFontCoversCjk(p, codePoint) === first)
    ? first : undefined;
}

function resolveShare(family: string, bold: boolean, italic: boolean): number | undefined {
  const chosen = chosenProfiles(family, bold, italic);
  if (chosen.length === 0) return undefined;
  let share: number | undefined;
  for (const profile of chosen) {
    let next: number | undefined;
    if (profile.typoMetrics) {
      const [ascender, descender, lineGap] = profile.typoMetrics;
      const above = ascender + Math.max(0, lineGap);
      next = above / (above - descender);
    } else if (profile.win) {
      const [ascent, descent] = profile.win;
      next = ascent / (ascent + descent);
    }
    if (next === undefined || !Number.isFinite(next) || next <= 0 || next >= 1) return undefined;
    if (share !== undefined && Math.abs(share - next) > 1e-12) return undefined;
    share = next;
  }
  return share;
}

/**
 * The glyph box behind the #1610 share, as em ratios: usWinAscent over
 * usWinDescent, or (sTypoAscender + sTypoLineGap) over −sTypoDescender for a
 * USE_TYPO_METRICS face. fontAlgn t / ctr / b position each run by this box in
 * the default line model (see `powerPointFontAlgnOffset`). Undefined when the
 * copies PowerPoint may use disagree or a table is missing.
 */
function resolveGlyphBox(family: string, bold: boolean, italic: boolean): ExcelLineBox | undefined {
  const chosen = chosenProfiles(family, bold, italic);
  if (chosen.length === 0) return undefined;
  let box: ExcelLineBox | undefined;
  for (const profile of chosen) {
    let next: ExcelLineBox | undefined;
    if (profile.typoMetrics) {
      const [ascender, descender, lineGap] = profile.typoMetrics;
      next = { ascent: (ascender + Math.max(0, lineGap)) / profile.unitsPerEm, descent: -descender / profile.unitsPerEm };
    } else if (profile.win) {
      const [ascent, descent] = profile.win;
      next = { ascent: ascent / profile.unitsPerEm, descent: descent / profile.unitsPerEm };
    }
    if (next === undefined || !(profile.unitsPerEm > 0)) return undefined;
    const share = next.ascent / (next.ascent + next.descent);
    if (!Number.isFinite(share) || share <= 0 || share >= 1) return undefined;
    if (box && (Math.abs(box.ascent - next.ascent) > 1e-12 || Math.abs(box.descent - next.descent) > 1e-12)) {
      return undefined;
    }
    box = next;
  }
  return box;
}

/**
 * The #1604 Excel natural line box (`excelDrawingMlLineRatios`) of the copy
 * PowerPoint uses, as em ratios. PowerPoint lays a body out with it when the
 * effective `a:bodyPr@compatLnSpc` is an explicit 0 (see
 * `PowerPointFaceMetrics.excel`). The copy is the same one the #1610 share
 * uses: a macOS Supplemental copy projects with the system tables (Times New
 * Roman: hhea ascender + lineGap 87, 1.150 em), an Office-bundled copy with
 * the Windows tables. Undefined when the copies disagree or a table is
 * missing; the body then keeps the ordinary line model.
 */
function resolveExcelBox(family: string, bold: boolean, italic: boolean): ExcelLineBox | undefined {
  let box: ExcelLineBox | undefined;
  for (const profile of chosenProfiles(family, bold, italic)) {
    if (profile.farEastCodePage == null || !profile.win) return undefined;
    const ratios = excelDrawingMlLineRatios({
      faceSource: profile.source === 'macos-supplemental' ? 'system' : 'office-bundle',
      unitsPerEm: profile.unitsPerEm,
      hhea: profile.hhea,
      win: profile.win,
      typoMetrics: profile.typoMetrics,
      farEastCodePage: profile.farEastCodePage,
    });
    if (!ratios) return undefined;
    if (box && (box.ascent !== ratios.ascentRatio || box.descent !== ratios.descentRatio)) return undefined;
    box = { ascent: ratios.ascentRatio, descent: ratios.descentRatio };
  }
  return box;
}

/** A face's natural line box as em ratios. */
export interface ExcelLineBox {
  ascent: number;
  descent: number;
}

/**
 * Everything PowerPoint's two line models need from one resource. Installed
 * instances are interned per family/weight/style. Embedded instances belong
 * to one registered FontFace, so same-tuple subsets cannot lend one another
 * their metrics or be merged as the same line contribution.
 */
export interface PowerPointFaceMetrics {
  /** #1610 ascent share of the 1.2 × size line box. */
  readonly share: number;
  /** The glyph box the share is taken from (see `resolveGlyphBox`);
   * undefined when the copies disagree on it. */
  readonly glyph: ExcelLineBox | undefined;
  /** #1604 natural box, used under an explicit compatLnSpc="0". */
  readonly excel: ExcelLineBox | undefined;
  /** Proven scalar coverage of a registered embedded resource. Parsed once,
   * bounded by core's cmap budgets; no font bytes or glyph-query cache retained. */
  readonly unicodeRanges?: OpenTypeLineMetrics['unicodeRanges'];
  readonly unicodePossibleRanges?: OpenTypeLineMetrics['unicodePossibleRanges'];
}

const faceCache = new Map<string, PowerPointFaceMetrics | null>();

/**
 * The line metrics of a face, or undefined when its share is unresolvable
 * (the face then adds nothing to its line's metric model; a line with no
 * known face keeps the ordinary model, #1689). Bounded LRU like the share
 * cache.
 */
export function powerPointFaceMetrics(
  family: string,
  bold: boolean,
  italic: boolean,
): PowerPointFaceMetrics | undefined {
  const key = `${family.trim().toLocaleLowerCase('en-US')}|${bold ? 700 : 400}|${italic ? 'i' : 'n'}`;
  const cached = faceCache.get(key);
  if (cached !== undefined) {
    faceCache.delete(key);
    faceCache.set(key, cached);
    return cached ?? undefined;
  }
  const share = powerPointAscentShare(family, bold, italic);
  const metrics = share === undefined ? null
    : Object.freeze({
      share, glyph: resolveGlyphBox(family, bold, italic), excel: resolveExcelBox(family, bold, italic),
    });
  faceCache.set(key, metrics);
  if (faceCache.size > SHARE_CACHE_LIMIT) {
    const oldest = faceCache.keys().next().value;
    if (oldest !== undefined) faceCache.delete(oldest);
  }
  return metrics ?? undefined;
}

/**
 * Line metrics of a concrete font resource from its own OS/2 tables (#1689:
 * PowerPoint sizes a line by the embedded resource's usWinAscent /
 * usWinDescent, or its typo metrics plus line gap under USE_TYPO_METRICS),
 * the same rule as the reference catalogue. Used for a deck-embedded face,
 * whose tables and cmap the loader retains. The #1604 compatLnSpc="0" box needs the
 * face's installation source, which an embedded part does not have, so it
 * stays undefined and that body keeps the ordinary model, as before.
 */
export function powerPointResourceFaceMetrics(metrics: OpenTypeLineMetrics): PowerPointFaceMetrics | undefined {
  const upm = metrics.unitsPerEm;
  if (!(upm > 0)) return undefined;
  let ascent: number | undefined;
  let descent: number | undefined;
  if (metrics.useTypoMetrics && metrics.typoAscent !== undefined && metrics.typoDescent !== undefined) {
    ascent = metrics.typoAscent + Math.max(0, metrics.typoLineGap ?? 0);
    descent = -metrics.typoDescent;
  } else if (metrics.winAscent !== undefined && metrics.winDescent !== undefined) {
    ascent = metrics.winAscent;
    descent = metrics.winDescent;
  }
  if (ascent === undefined || descent === undefined) return undefined;
  const share = ascent / (ascent + descent);
  if (!Number.isFinite(share) || share <= 0 || share >= 1) return undefined;
  return Object.freeze({ share, glyph: { ascent: ascent / upm, descent: descent / upm }, excel: undefined,
    unicodeRanges: metrics.unicodeRanges, unicodePossibleRanges: metrics.unicodePossibleRanges });
}

/** One run's contribution to a line: its authored size and ascent share. */
export interface PowerPointLineRun {
  sizePx: number;
  share: number;
}

/**
 * The natural line box of one line. Each run claims 1.2 × its own size,
 * split by its share. The line unions the runs' ascent and descent parts and
 * rescales them into 1.2 × the largest size.
 *
 * Measured on mixed-face lines (Arial+Meiryo in both orders, Arial+MS Gothic,
 * Calibri+Yu Gothic, Arial+Gabriola) and mixed sizes (40+100 pt): all exact.
 * Taking the largest descent instead is off by up to 5 px (1/100 in), and the
 * largest ascent by 13 px.
 */
export function powerPointNaturalLine(
  runs: readonly PowerPointLineRun[],
  largestSizePx = 0,
): { ascent: number; descent: number } {
  // ECMA-376 §21.1.2.2.5/.11: even an unresolved face contributes its authored
  // size. Known faces alone determine the split; never invent an unknown share.
  let maxSize = largestSizePx;
  let ascent = 0;
  let descent = 0;
  for (const run of runs) {
    maxSize = Math.max(maxSize, run.sizePx);
    ascent = Math.max(ascent, 1.2 * run.sizePx * run.share);
    descent = Math.max(descent, 1.2 * run.sizePx * (1 - run.share));
  }
  const height = 1.2 * maxSize;
  if (!(ascent + descent > 0)) return { ascent: height * 0.8, descent: height * 0.2 };
  const a = height * ascent / (ascent + descent);
  return { ascent: a, descent: height - a };
}

/**
 * PowerPoint rounds `spcPts` to whole points before laying out the line
 * (#1610 controls: 40.5 → 41, 45.25 → 45, 48.33 → 48, 49.25 → 49 and
 * 50.75 → 51 pt, each over eight lines). Font sizes stay exact: 10.5, 11.5,
 * 13.33, 40.5 and 55.5 pt keep a 1.2 × size line.
 */
export function powerPointExactLinePoints(points: number): number {
  return Math.floor(points + 0.5);
}

/**
 * The natural line box when the effective `a:bodyPr@compatLnSpc` is an
 * explicit 0 (ECMA-376 §21.1.2.1.1 gives no algorithm, only "decided in a
 * simplistic manner using the font scene").
 *
 * Observed with PowerPoint's reference (Windows-style) PDF export, #1619
 * controls: every left/right pair differing only in compatLnSpc 0/1, over
 * Arial, Calibri, Times New Roman, Gabriola, Meiryo, Yu Gothic, MS Gothic and
 * Aptos at 18-100 pt, lnSpc omitted / 60-150 % / 30-100 pt, spcBef/spcAft in
 * points and percent with and without spcFirstLastPara, anchors t/ctr/b and
 * mixed sizes and faces:
 * - compatLnSpc="1" renders exactly like an omitted compatLnSpc (the #1610
 *   model), whether it is authored on the slide, the layout or the master.
 * - An explicit 0 selects Excel's #1604 model instead: each run claims its
 *   own natural box (`PowerPointFaceMetrics.excel` × size), the line is the
 *   largest ascent over the largest descent (no rescale to 1.2 × size),
 *   lnSpc goes through the same `drawingMlSpacedLineBox` rule, and a
 *   percentage spcBef/spcAft is a fraction of that natural line.
 * All 546 baselines with fontAlgn omitted/auto/base, in both models, land on
 * the export's 1/100 in unit (3 sit within 0.005 unit of a rounding tie).
 * fontAlgn t/ctr/b are outside this rule (tracked separately).
 */
export function powerPointCompatOffNaturalLine(runs: readonly { sizePx: number; box: ExcelLineBox }[]): {
  ascent: number;
  descent: number;
} {
  let ascent = 0;
  let descent = 0;
  for (const run of runs) {
    ascent = Math.max(ascent, run.box.ascent * run.sizePx);
    descent = Math.max(descent, run.box.descent * run.sizePx);
  }
  return { ascent, descent };
}

/** `a:pPr@fontAlgn` values that move runs off the baseline (ST_TextFontAlignType,
 * ECMA-376 §20.1.10.62). Omitted, `auto` and `base` lay out identically. */
export type PowerPointFontAlgn = 't' | 'ctr' | 'b';

/**
 * PowerPoint's reference PDF export positions a run off its fontAlgn
 * reference by a whole number of 1/100 in (0.72 pt; see the export note
 * below). Keeping these offsets continuous misses 73 of the 203 t text boxes
 * and 80 of the 196 ctr boxes of the #1619/#1636 controls; rounding each
 * offset matches them.
 */
export const POWERPOINT_FONT_ALGN_UNIT_PT = 0.72;

/** One run (or a break / end-of-paragraph mark) contributing to a fontAlgn line. */
export interface PowerPointAlignedRun {
  sizePx: number;
  face: PowerPointFaceMetrics;
}

function runBox(run: PowerPointAlignedRun, compatOff: boolean): ExcelLineBox | undefined {
  return compatOff ? run.face.excel : run.face.glyph;
}

/**
 * The line box of a fontAlgn t / ctr / b line, split at the line's alignment
 * reference instead of at a baseline. ECMA-376 §21.1.2.2.7 names the values
 * but gives no algorithm; this is PowerPoint's behaviour in its reference
 * ("electronic distribution") PDF export, #1619 and #1636 controls:
 * t / ctr / b on Arial, Meiryo, Yu Gothic, Gabriola and MS Gothic at
 * 16-100 pt, single-size and mixed-size / mixed-face lines, lnSpc omitted,
 * 60-150 % (1 % steps over 90-110 %) and spcPts across both models' natural
 * height (1 pt steps), with compatLnSpc 0 and 1.
 *
 * - Default model: the line is the usual 1.2 × largest size (L).
 *   t references its top, ctr its middle. b references the midpoint between
 *   the line's #1610 baseline, rounded to the export unit, and its bottom.
 * - compatLnSpc="0": every run's #1604 natural box is aligned at the
 *   reference (t: tops, ctr: centres, b: descent midpoints) and the line is
 *   their union.
 *
 * lnSpc then re-divides this box with the shared DrawingML rule
 * (`drawingMlSpacedLineBox`), so a reference at the top gets
 * −0.75 (L − H) below 100 % and all extra space above it, one at the bottom
 * the reverse. The line pitch is the spaced height as for baseline lines.
 * Runs are placed off the reference by `powerPointFontAlgnOffset`.
 * Undefined when a run's face has no box for the model.
 */
export function powerPointFontAlgnReference(
  fontAlgn: PowerPointFontAlgn,
  runs: readonly PowerPointAlignedRun[],
  compatOff: boolean,
  unitPx: number,
): { ascent: number; descent: number } | undefined {
  if (runs.length === 0) return undefined;
  const boxes: ExcelLineBox[] = [];
  for (const run of runs) {
    const box = runBox(run, compatOff);
    if (!box) return undefined;
    boxes.push({ ascent: box.ascent * run.sizePx, descent: box.descent * run.sizePx });
  }
  if (!compatOff) {
    // Iterative maxima: a long run can contribute 10^5 entries, too many to
    // spread into Math.max.
    let largest = 0;
    for (const run of runs) largest = Math.max(largest, run.sizePx);
    const height = 1.2 * largest;
    if (fontAlgn === 't') return { ascent: 0, descent: height };
    if (fontAlgn === 'ctr') return { ascent: height / 2, descent: height / 2 };
    const natural = powerPointNaturalLine(runs.map((r) => ({ sizePx: r.sizePx, share: r.face.share })));
    const half = (height - Math.round(natural.ascent / unitPx) * unitPx) / 2;
    return { ascent: height - half, descent: half };
  }
  if (fontAlgn === 'b') {
    let above = 0;
    let below = 0;
    for (const b of boxes) {
      above = Math.max(above, b.ascent + b.descent / 2);
      below = Math.max(below, b.descent / 2);
    }
    return { ascent: above, descent: below };
  }
  let height = 0;
  for (const b of boxes) height = Math.max(height, b.ascent + b.descent);
  return fontAlgn === 't' ? { ascent: 0, descent: height } : { ascent: height / 2, descent: height / 2 };
}

/**
 * A run's baseline relative to its line's fontAlgn reference
 * (`powerPointFontAlgnReference`), positive downwards: t puts the run's box
 * top on the reference, ctr its centre, b the midpoint of its descent. The box
 * is the glyph box (`resolveGlyphBox`) in the default model and the #1604
 * natural box under compatLnSpc="0". Each offset is a whole export unit.
 */
export function powerPointFontAlgnOffset(
  fontAlgn: PowerPointFontAlgn,
  run: PowerPointAlignedRun,
  compatOff: boolean,
  unitPx: number,
): number | undefined {
  const box = runBox(run, compatOff);
  if (!box) return undefined;
  const ascent = box.ascent * run.sizePx;
  const descent = box.descent * run.sizePx;
  if (fontAlgn === 'b') return -Math.round(descent / 2 / unitPx) * unitPx;
  const offset = fontAlgn === 't' ? ascent : (ascent - descent) / 2;
  return Math.round(offset / unitPx) * unitPx;
}

/*
 * PDF exports quantize: each baseline lands on a whole device unit counted
 * from the anchored text top (1/100 in in the #1610 controls; 1/150 in in a
 * corpus export whose 12 pt lines alternate 13.92/14.88 pt). The unit belongs
 * to the export device, not to the slide, so layout keeps continuous
 * positions; the controls match to within half a unit.
 */
