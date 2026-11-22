import { describe, expect, it } from 'vitest';
import { stableFingerprint } from './fingerprint.js';
import {
  centeredLeaderGlyphOrigins,
  groupedRunBorderFragments,
  retainedEmphasisGlyphs,
  retainedTextDecorations,
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

  // Windows Word 16.0.20430 (#1896 controls), Arial at 48 pt: single
  // underline 5.04..8.52 below the baseline, thick 2.52..6.72, strike
  // 12.48..10.08 above, double strike 14.88..12.48 and 10.08..7.68.
  describe('with the face rule (Arial post / OS/2)', () => {
    const geometry = {
      underlineTopEm: 217 / 2048, underlineThicknessEm: 150 / 2048,
      strikeTopEm: -530 / 2048, strikeThicknessEm: 102 / 2048,
    };
    const base = { ascentPt: 43, descentPt: 10, inkBounds: { xMinPt: 0, xMaxPt: 40, ascentPt: 34, descentPt: 10 } };
    const probe = { ascentPt: 2, descentPt: 2 };
    const at = (input: Partial<Parameters<typeof retainedTextDecorations>[0]>) => retainedTextDecorations({
      origin: { xPt: 0, yPt: 100 }, advancePt: 40, base, color: '#000000',
      fontRule: { geometry, sizePt: 48 }, ...input,
    });
    // Word's PDF positions sit on a 0.12 pt grid.
    const band = (d: { widthPt: number; from: { yPt: number } }) =>
      [d.from.yPt - d.widthPt / 2 - 100, d.from.yPt + d.widthPt / 2 - 100];

    it('draws the single underline through descenders at the face rule', () => {
      const [u] = at({ underline: { authoredStyle: 'single', color: '#000000', probe } });
      const [top, bottom] = band(u!);
      expect(Math.abs(top - (5.04))).toBeLessThan(0.13);
      expect(Math.abs(bottom - (8.52))).toBeLessThan(0.13);
    });

    it('draws a thick underline from half the offset to the single rule centre', () => {
      const [u] = at({ underline: { authoredStyle: 'thick', color: '#000000', probe } });
      const [top, bottom] = band(u!);
      expect(Math.abs(top - (2.52))).toBeLessThan(0.13);
      expect(Math.abs(bottom - (6.72))).toBeLessThan(0.13);
    });

    it('draws single and double strikes around the face strikeout', () => {
      const [single] = at({ strike: { double: false, probe } });
      expect(Math.abs(band(single!)[0] - (-12.48))).toBeLessThan(0.13);
      expect(Math.abs(band(single!)[1] - (-10.08))).toBeLessThan(0.13);
      const [upper, lower] = at({ strike: { double: true, probe } });
      expect(Math.abs(band(upper!)[0] - (-14.88))).toBeLessThan(0.13);
      expect(Math.abs(band(lower!)[1] - (-7.68))).toBeLessThan(0.13);
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
