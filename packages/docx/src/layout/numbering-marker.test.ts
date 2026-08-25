import { describe, expect, it } from 'vitest';
import type { TextLayoutService } from './text.js';
import { resolveNumberingMarkerGeometry } from './numbering-marker.js';

const route = {
  familyList: 'serif', scope: 'generic', fingerprint: 'numbering-marker-test-route',
} as const;

function textService(advancePt: number): TextLayoutService {
  return {
    fingerprint: 'numbering-marker-test-text',
    localMetrics: {},
    resolve(request) {
      return {
        requestedFamily: request.fonts.ascii ?? 'serif',
        resolvedFamily: request.fonts.ascii ?? 'serif',
        route,
        source: 'generic',
        weight: request.weight ?? 400,
        style: request.style ?? 'normal',
        diagnostics: [],
        genericFamily: 'serif',
      };
    },
    shape(request) {
      return {
        text: request.text,
        spans: [{
          text: request.text,
          start: 0,
          end: request.text.length,
          script: 'eastAsia',
          breakBefore: true,
          font: {
            requestedFamily: 'serif', resolvedFamily: 'serif', route,
            source: 'generic', weight: request.weight ?? 400,
            style: request.style ?? 'normal', diagnostics: [], genericFamily: 'serif',
          },
          fontRoute: route,
          advancePt,
          ascentPt: 8,
          descentPt: 2,
        }],
        advancePt,
        ascentPt: 8,
        descentPt: 2,
        graphemeBoundaries: [0, request.text.length],
        clusters: [{
          range: { start: 0, end: request.text.length }, offsetPt: 0, advancePt,
        }],
        diagnostics: [],
      };
    },
  };
}

describe('resolveNumberingMarkerGeometry', () => {
  function geometryWithSuffix(suff: 'nothing' | 'space', authoredFirstIndentPt: number) {
    return resolveNumberingMarkerGeometry(
      {
        numId: 1,
        level: 0,
        format: 'decimal',
        text: '1.',
        indentLeft: 36,
        tab: 36,
        suff,
        jc: 'left',
      },
      {
        fontSizePt: 10,
        fonts: { ascii: 'serif', eastAsia: 'serif' },
        weight: 400,
        style: 'normal',
        complexScript: false,
      },
      {
        authoredFirstIndentPt,
        physicalIndentLeftPt: 36,
        tabStops: [],
        defaultTabPt: 36,
      },
      textService(10),
    );
  }

  function geometryAtCoincidentStop(alignment: 'left' | 'num') {
    return resolveNumberingMarkerGeometry(
      {
        numId: 1,
        level: 0,
        format: 'bullet',
        text: '\u30fb',
        indentLeft: 18,
        tab: 18,
        suff: 'tab',
        jc: 'left',
      },
      {
        fontSizePt: 10,
        fonts: { ascii: 'serif', eastAsia: 'serif' },
        weight: 400,
        style: 'normal',
        complexScript: false,
      },
      {
        // Marker interval relative to the 18pt paragraph indent is [-8, 2].
        // Its absolute end is therefore exactly the authored 20pt list tab.
        authoredFirstIndentPt: -8,
        physicalIndentLeftPt: 18,
        tabStops: [{ pos: 20, alignment, leader: 'none' }],
        defaultTabPt: 42.55,
      },
      textService(10),
    );
  }

  it('keeps a suffix tab on the list stop coincident with the marker end', () => {
    const geometry = geometryAtCoincidentStop('num');

    expect(geometry.markerWidthPt).toBe(10);
    expect(geometry.bodyOffsetPt).toBe(2);
  });

  it('keeps ordinary coincident tabs strictly forward', () => {
    expect(geometryAtCoincidentStop('left').bodyOffsetPt).toBeCloseTo(24.55, 10);
  });

  it.each([
    ['nothing', -36],
    ['space', -36],
  ] as const)(
    'keeps a %s suffix body at the paragraph start when its marker ends inside the hanging region',
    (suffix, authoredFirstIndentPt) => {
      expect(geometryWithSuffix(suffix, authoredFirstIndentPt).bodyOffsetPt).toBe(0);
    },
  );

  it.each([
    ['nothing', 2],
    ['space', 12],
  ] as const)(
    'keeps a positive %s suffix advance beyond the paragraph start',
    (suffix, expectedBodyOffsetPt) => {
      expect(geometryWithSuffix(suffix, -8).bodyOffsetPt).toBe(expectedBodyOffsetPt);
    },
  );
});
