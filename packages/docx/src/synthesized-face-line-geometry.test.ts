import { describe, expect, it } from 'vitest';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService, type GlyphRasterComparisonRequest } from './layout/text.js';
import { referenceFontLineMetrics } from './reference-font-line-metrics.js';
import { buildSegments, layoutLines, lineBoxHeight, paragraphMarkLineMetrics } from './line-layout.js';
import type { DocParagraph, DocRun } from './types.js';

const context = {
  font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
  measureText: (text: string) => ({
    width: [...text].length * 5, actualBoundingBoxAscent: 6, actualBoundingBoxDescent: 1,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

// MS Mincho is catalogued only at weight 400; its Far-East single line is
// 1.3 × usWin = 1.3 em. The host has that face loaded through an exact local()
// name under a private single-face family, so bold resolves natively.
function minchoServices(sameRaster: boolean) {
  const comparisons: GlyphRasterComparisonRequest[] = [];
  const text = createTextLayoutService({
    fonts: createFontResolver([{
      requestedFamily: 'MS Mincho', resolvedFamily: 'private-ms-mincho-400', source: 'local',
      resourceIdentity: 'office-local:local("MS Mincho")', weight: 400, style: 'normal',
      coverageProbeFamily: 'private-ms-mincho-400-coverage',
    }]),
    measurer: {
      fingerprint: `synthesized-face-fixture-${sameRaster}`,
      measure: (request) => ({
        advancePt: [...request.text].length * request.fontSizePt,
        ascentPt: request.fontSizePt * 0.8, descentPt: request.fontSizePt * 0.2,
      }),
      sameGlyphRaster: (request) => {
        comparisons.push(request);
        return sameRaster;
      },
      // The face owns everything but Hangul, which falls back.
      ownsGlyphs: (request) => !/\p{Script=Hangul}/u.test(request.text),
    },
  });
  return { layoutServices: { text }, comparisons };
}

describe('bold runs of a face catalogued only at 400 (#1848)', () => {
  it('has no catalog tuple for bold, so geometry needs the raster proof', () => {
    expect(referenceFontLineMetrics('MS Mincho', 700, 'normal')).toBeUndefined();
    expect(referenceFontLineMetrics('MS Mincho', 400, 'normal')?.lineHeightRatio).toBeCloseTo(1.3, 8);
  });

  it.each([
    { sameRaster: true, expected: 14 * 1.3 },
    // A genuine bold face (or any other painting face) keeps its Canvas box.
    { sameRaster: false, expected: 14 },
  ])('sizes a bold empty mark from the proved face: identical=$sameRaster', ({ sameRaster, expected }) => {
    const { layoutServices, comparisons } = minchoServices(sameRaster);
    const paragraph = {
      runs: [], defaultFontFamily: 'MS Mincho', defaultFontFamilyEastAsia: 'MS Mincho',
      defaultFontSize: 14, lineSpacing: null,
    } as unknown as DocParagraph;
    const mark = paragraphMarkLineMetrics(
      paragraph, 1, undefined, false, true, context, {}, null,
      layoutServices.text.fontMetrics, layoutServices.text,
      { fontSizePt: 14, weight: 700, style: 'normal',
        fonts: { ascii: 'MS Mincho', highAnsi: 'MS Mincho', eastAsia: 'MS Mincho', complexScript: 'MS Mincho' } } as never,
    );
    expect(mark.advancePx).toBeCloseTo(expected, 8);
    // The bold native route was compared with the loaded 400 face's own family.
    expect(comparisons).toContainEqual(expect.objectContaining({
      faceFamily: 'private-ms-mincho-400', weight: 700, style: 'normal',
    }));
  });

  it.each([
    { sameRaster: true, expected: 14 * 1.3 },
    { sameRaster: false, expected: 14 },
  ])('sizes a bold text line from the proved face: identical=$sameRaster', ({ sameRaster, expected }) => {
    const { layoutServices, comparisons } = minchoServices(sameRaster);
    const segments = buildSegments([{
      type: 'text', text: '注意', fontFamily: 'MS Mincho', fontFamilyEastAsia: 'MS Mincho',
      fontSize: 14, bold: true, italic: false, underline: false, strikethrough: false,
    }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices: layoutServices as never });
    const line = layoutLines(context, segments, 400, 0, 1)[0]!;
    expect(lineBoxHeight(null, line.ascent, line.descent, 1, undefined, false,
      line.intendedSingle, true, line.gridCountSingle)).toBeCloseTo(expected, 8);
    // One fixed probe of the span's script class proves the face tuple.
    expect(comparisons.map((request) => request.text)).toContain('国あ');
  });

  it('does not run the proof for a tuple the catalog already holds', () => {
    const { layoutServices, comparisons } = minchoServices(true);
    buildSegments([{
      type: 'text', text: '注意', fontFamily: 'MS Mincho', fontFamilyEastAsia: 'MS Mincho',
      fontSize: 14, bold: false, italic: false, underline: false, strikethrough: false,
    }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices: layoutServices as never });
    expect(comparisons).toEqual([]);
  });

  it('keeps a bold list-label metric run outside the synthesized geometry', () => {
    // A picture list label's metric run has no span of its own: a bold one
    // whose East Asian slot names a regular-only face keeps its Canvas box
    // (the stub's 8 + 2) even where the raster proof would succeed and would
    // give 1.3 em (rule scope). (A floating drawing's anchor host has no
    // metric at all: WORD_ANCHOR_HOST_NO_LINE_METRIC.)
    const { layoutServices, comparisons } = minchoServices(true);
    const segments = buildSegments([{
      type: 'anchorHost', fontSize: 14, fontFamily: 'MS Mincho', fontFamilyEastAsia: 'MS Mincho',
      bold: true, listLabel: 'start',
    }] as unknown as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices: layoutServices as never });
    const line = layoutLines(context, segments, 400, 0, 1)[0]!;
    expect(lineBoxHeight(null, line.ascent, line.descent, 1, undefined, false,
      line.intendedSingle, true, line.gridCountSingle)).toBeCloseTo(10, 8);
    expect(comparisons).toEqual([]);
  });

  it('declines a span whose glyphs fall back to another face', () => {
    // Korean authored in MS Mincho: the face paints the probe identically,
    // but the span's own glyphs come from a fallback face.
    const { layoutServices } = minchoServices(true);
    const segments = buildSegments([{
      type: 'text', text: '한국어', fontFamily: 'MS Mincho', fontFamilyEastAsia: 'MS Mincho',
      fontSize: 14, bold: true, italic: false, underline: false, strikethrough: false,
    }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices: layoutServices as never });
    const line = layoutLines(context, segments, 400, 0, 1)[0]!;
    expect(lineBoxHeight(null, line.ascent, line.descent, 1, undefined, false,
      line.intendedSingle, true, line.gridCountSingle)).toBeCloseTo(14, 8);
  });
});
