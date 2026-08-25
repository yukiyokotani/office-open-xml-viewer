import { describe, expect, it } from 'vitest';
import { stableFingerprint } from './fingerprint.js';
import {
  centeredLeaderGlyphOrigins,
  groupedRunBorderFragments,
  retainedEmphasisGlyphs,
  retainedTextDecorations,
  rubyPaintOperations,
} from './retained-typography.js';

const fontRoute = {
  familyList: '"Test Sans"', scope: 'native', fingerprint: 'test-font-route',
} as const;

describe('retained typography geometry', () => {
  it('retains the selected authored emphasis glyph and its authoritative face geometry', () => {
    const common = {
      origin: { xPt: 10, yPt: 30 },
      clusters: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 9 },
        { range: { start: 1, end: 2 }, offset: { xPt: 9, yPt: 0 }, advancePt: 4 },
        { range: { start: 2, end: 3 }, offset: { xPt: 13, yPt: 0 }, advancePt: 11 },
      ],
      clusterInk: [
        { text: '観', range: { start: 0, end: 1 }, ink: { xMinPt: 1, xMaxPt: 8, ascentPt: 7, descentPt: 2 } },
        { text: ' ', range: { start: 1, end: 2 }, ink: { xMinPt: 0, xMaxPt: 4, ascentPt: 0, descentPt: 0 } },
        { text: '察', range: { start: 2, end: 3 }, ink: { xMinPt: 2, xMaxPt: 10, ascentPt: 8, descentPt: 1 } },
      ],
      mark: {
        inkBounds: { xMinPt: -1, xMaxPt: 3, ascentPt: 4, descentPt: 1 },
        fontRoute, fontSizePt: 6, fontWeight: 700, fontStyle: 'italic' as const,
        color: { kind: 'explicit' as const, color: '#123456' },
      },
      scaleX: 1,
    };
    const glyphFor = (authored: string, glyph: string) => retainedEmphasisGlyphs({
      ...common, authored, glyph,
    });

    const dot = glyphFor('dot', '•');
    const comma = glyphFor('comma', '﹅');
    const circle = glyphFor('circle', '○');
    expect(dot.map((operation) => operation.text)).toEqual(['•', '•']);
    expect(comma.map((operation) => operation.text)).toEqual(['﹅', '﹅']);
    expect(circle.map((operation) => operation.text)).toEqual(['○', '○']);
    expect(circle[0]).toMatchObject({
      origin: { xPt: 13.5, yPt: 22 },
      fontRoute, fontSizePt: 6, fontWeight: 700, fontStyle: 'italic',
      inkBounds: { xMinPt: -1, xMaxPt: 3, ascentPt: 4, descentPt: 1 },
    });
    expect(circle[1]).toMatchObject({ origin: { xPt: 28, yPt: 21 } });
    expect(glyphFor('underDot', '•')[0]).toMatchObject({ origin: { xPt: 13.5, yPt: 36 } });
    expect(circle[0]).not.toHaveProperty('points');
    expect(circle[0]).not.toHaveProperty('stroke');
    expect(structuredClone(circle)).toEqual(circle);
    expect(stableFingerprint('emphasis', circle)).toBe(stableFingerprint('emphasis', structuredClone(circle)));
  });

  it('retains underline clearance and thickness from deliberately non-proportional ink', () => {
    const [underline] = retainedTextDecorations({
      origin: { xPt: 5, yPt: 10 }, advancePt: 20,
      base: {
        ascentPt: 9, descentPt: 4,
        inkBounds: { xMinPt: 0, xMaxPt: 20, ascentPt: 8, descentPt: 3 },
      },
      color: '#000000',
      underline: {
        authoredStyle: 'single', color: '#112233',
        probe: {
          ascentPt: 2, descentPt: 2,
          inkBounds: { xMinPt: 0, xMaxPt: 5, ascentPt: .5, descentPt: 1.5 },
        },
      },
    });
    // Probe ink is 2pt thick; clearing the base's 3pt descent puts its center
    // at baseline + 3 + 1 = 14pt, independent of the run font size.
    expect(underline).toMatchObject({
      kind: 'underline', widthPt: 2, color: '#112233',
      from: { xPt: 5, yPt: 14 }, to: { xPt: 25, yPt: 14 },
    });
  });

  it('omits absent optional authored underline style from the canonical retained node', () => {
    const [underline] = retainedTextDecorations({
      origin: { xPt: 5, yPt: 10 }, advancePt: 20,
      base: { ascentPt: 9, descentPt: 4 },
      color: '#000000',
      underline: {
        color: '#112233',
        probe: { ascentPt: 2, descentPt: 2 },
      },
    });

    expect(underline).toEqual({
      kind: 'underline', color: '#112233', widthPt: 2, style: 'solid',
      from: { xPt: 5, yPt: 15 }, to: { xPt: 25, yPt: 15 },
    });
    expect(underline).not.toHaveProperty('authoredStyle');
    expect(structuredClone(underline)).toEqual(underline);
  });

  it('adds an author-coloured revision underline for insertions and strike for deletions', () => {
    const base = { ascentPt: 9, descentPt: 4 };
    const inserted = retainedTextDecorations({
      origin: { xPt: 5, yPt: 10 }, advancePt: 20,
      base, color: '#000000',
      revision: { color: '#C00000', underline: { probe: { ascentPt: 2, descentPt: 2 } } },
    });
    expect(inserted).toEqual([{
      kind: 'underline', color: '#C00000', widthPt: 2, style: 'solid',
      from: { xPt: 5, yPt: 15 }, to: { xPt: 25, yPt: 15 },
    }]);
    const deleted = retainedTextDecorations({
      origin: { xPt: 5, yPt: 10 }, advancePt: 20,
      base, color: '#000000',
      revision: { color: '#0070C0', strike: { probe: { ascentPt: 2, descentPt: 2 } } },
    });
    expect(deleted).toEqual([{
      kind: 'strikethrough', color: '#0070C0', widthPt: 2, style: 'solid',
      from: { xPt: 5, yPt: 10 }, to: { xPt: 25, yPt: 10 },
    }]);
  });

  it('stacks the author-coloured revision decoration on the authored underline', () => {
    const decorations = retainedTextDecorations({
      origin: { xPt: 5, yPt: 10 }, advancePt: 20,
      base: { ascentPt: 9, descentPt: 4 }, color: '#000000',
      underline: { color: '#112233', probe: { ascentPt: 2, descentPt: 2 } },
      revision: { color: '#C00000', underline: { probe: { ascentPt: 2, descentPt: 2 } } },
    });
    expect(decorations.map((decoration) => [decoration.kind, decoration.color])).toEqual([
      ['underline', '#112233'],
      ['underline', '#C00000'],
    ]);
  });

  it('centers a shaped tab-leader sequence inside the authored tab interval', () => {
    expect(centeredLeaderGlyphOrigins({
      interval: { xPt: 10, yPt: 4, widthPt: 23, heightPt: 12 },
      baselinePt: 14,
      glyph: '.',
      advancePt: 5,
      fontRoute,
      fontSizePt: 10,
      fontWeight: 700,
      fontStyle: 'normal',
      color: { kind: 'explicit', color: '#123456' },
    })).toEqual([
      expect.objectContaining({ text: '.', origin: { xPt: 11.5, yPt: 14 } }),
      expect.objectContaining({ text: '.', origin: { xPt: 16.5, yPt: 14 } }),
      expect.objectContaining({ text: '.', origin: { xPt: 21.5, yPt: 14 } }),
      expect.objectContaining({ text: '.', origin: { xPt: 26.5, yPt: 14 } }),
    ]);
  });

  it('retains authored hpsRaise and rich shaped guide-span origins for ruby', () => {
    expect(rubyPaintOperations({
      baseOrigin: { xPt: 20, yPt: 30 },
      baseAdvancePt: 20,
      raisePt: 6,
      guideAdvancePt: 12,
      spans: [
        { text: 'ふ', offsetPt: 0, fontRoute, fontSizePt: 5, fontWeight: 400, fontStyle: 'normal', color: { kind: 'explicit', color: '#111111' } },
        { text: 'り', offsetPt: 6, fontRoute, fontSizePt: 6, fontWeight: 700, fontStyle: 'italic', color: { kind: 'explicit', color: '#222222' } },
      ],
    })).toEqual([
      expect.objectContaining({ text: 'ふ', origin: { xPt: 24, yPt: 24 }, fontSizePt: 5 }),
      expect.objectContaining({ text: 'り', origin: { xPt: 30, yPt: 24 }, fontWeight: 700, fontStyle: 'italic' }),
    ]);
  });

  it('fails explicitly when ruby has neither authored raise nor retained ink geometry', () => {
    expect(() => rubyPaintOperations({
      baseOrigin: { xPt: 20, yPt: 30 }, baseAdvancePt: 20,
      guideAdvancePt: 12, spans: [],
    })).toThrow(/ruby geometry/i);
  });

  it('groups visually adjacent equal run borders and includes owned justification slack', () => {
    const border = {
      val: 'single', color: '#123456', widthPt: 1, spacePt: 2,
      themeColor: 'accent1', themeTint: '66', shadow: false, frame: false,
    } as const;
    const fragments = groupedRunBorderFragments([
      { bounds: { xPt: 10, yPt: 5, widthPt: 12, heightPt: 10 }, trailingSlackPt: 3, border },
      { bounds: { xPt: 25, yPt: 5, widthPt: 8, heightPt: 10 }, trailingSlackPt: 0, border },
    ]);

    expect(fragments).toHaveLength(4);
    expect(fragments).toContainEqual(expect.objectContaining({
      edge: 'top', from: { xPt: 8, yPt: 3 }, to: { xPt: 35, yPt: 3 },
    }));
    expect(fragments).toContainEqual(expect.objectContaining({
      edge: 'bottom', from: { xPt: 8, yPt: 17 }, to: { xPt: 35, yPt: 17 },
    }));
  });

  it('retains authored DOCX border tokens separately from shared paint treatment', () => {
    const fragments = groupedRunBorderFragments([
      {
        bounds: { xPt: 0, yPt: 0, widthPt: 20, heightPt: 10 }, trailingSlackPt: 0,
        border: { val: 'dotDash', color: '#123456', widthPt: 2, spacePt: 0 },
      },
    ]);
    const doubled = groupedRunBorderFragments([
      {
        bounds: { xPt: 0, yPt: 0, widthPt: 20, heightPt: 10 }, trailingSlackPt: 0,
        border: { val: 'double', color: '#123456', widthPt: 3, spacePt: 0 },
      },
    ]);
    const equivalentPaint = groupedRunBorderFragments([
      {
        bounds: { xPt: 0, yPt: 0, widthPt: 20, heightPt: 10 }, trailingSlackPt: 0,
        border: { val: 'dashDotStroked', color: '#123456', widthPt: 2, spacePt: 0 },
      },
    ]);

    expect(fragments[0]).toMatchObject({
      authoredStyle: 'dotDash', style: 'dashed', dashPatternPt: [2, 4, 6, 4],
    });
    expect(equivalentPaint[0]).toMatchObject({
      authoredStyle: 'dashDotStroked', style: 'dashed', dashPatternPt: [2, 4, 6, 4],
    });
    expect(stableFingerprint('border', fragments[0])).not.toBe(
      stableFingerprint('border', equivalentPaint[0]),
    );
    expect(doubled[0]).toMatchObject({
      authoredStyle: 'double', style: 'double', dashPatternPt: [],
    });
  });
});
