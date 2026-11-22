import { describe, expect, it, vi } from 'vitest';
import { type OfficeFontFallbackRoute, type LocalFontNormalLine, canvasFontString } from '@silurus/ooxml-core';
import { loadDocxOfficeFontFallbacks, hostFontLineMetric, prepareDocxHostFontMetrics } from './host-font-line-metrics.js';
import { selectedFontLineMetric } from './line-breaker/font-metrics.js';
import { paragraphMarkLineMetrics } from './line-layout.js';
import { createTextLayoutService } from './layout/text.js';
import type { DocParagraph } from './types.js';
import { createFontResolver } from './layout/font-service.js';
import { unloadOfficeFontFallbacks } from '@silurus/ooxml-core';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';

function route(requestedFamily = 'Times New Roman', source: 'local' | 'substitute' = 'local'): OfficeFontFallbackRoute {
  return {
    requestedFamily, family: 'exact-host-face', source, resourceIdentity: 'office-local:host-resource',
    weight: 400, style: 'normal',
    metric: { family: 'exact-host-face', requestedFamily, sourceIdentity: 'office-local:host-resource' },
  };
}
const host: LocalFontNormalLine = {
  ascentRatio: 1825 / 2048, descentRatio: 443 / 2048, lineGapRatio: 87 / 2048,
};

function select(loaded: OfficeFontFallbackRoute) {
  return createFontResolver([{
    requestedFamily: loaded.requestedFamily, resolvedFamily: loaded.family,
    source: loaded.source, resourceIdentity: loaded.resourceIdentity,
  }]).resolve({ requestedFamily: loaded.requestedFamily });
}

// Browser contract the preparation relies on: font boxes scale linearly with
// the requested size (given here at a 2048px probe), and the size-adjust 0%
// coverage face measures its own glyphs as zero while fallback keeps advance.
function contractCanvas() {
  const px = (font: string) => Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? 0);
  return {
    font: '',
    measureText(text: string) {
      if (this.font.includes('_coverage')) {
        const owned = /^[\x20-\x7e]*$/.test(text);
        return { width: owned ? 0 : 100, actualBoundingBoxAscent: owned ? 0 : 80,
          actualBoundingBoxDescent: 0 } as TextMetrics;
      }
      const scale = px(this.font) / 2048;
      return { width: 1200 * scale * text.length, fontBoundingBoxAscent: 1825 * scale,
        fontBoundingBoxDescent: 443 * scale } as TextMetrics;
    },
  };
}

describe('DOCX host-face line geometry', () => {
  it('prepares measured geometry on the retained local paint route, shared across document loads', async () => {
    const registrations = new Set<FontFace>();
    const fonts = {
      add(face: FontFace) { registrations.add(face); },
      delete(face: FontFace) { return registrations.delete(face); },
      [Symbol.iterator]() { return registrations.values(); },
    } as FontFaceSet;
    class LocalFace {
      status = 'unloaded';
      readonly weight: string;
      readonly style: string;
      readonly sizeAdjust: string;
      constructor(readonly family: string, _source: string, descriptors: FontFaceDescriptors) {
        this.sizeAdjust = (descriptors as FontFaceDescriptors & { sizeAdjust?: string }).sizeAdjust ?? '100%';
        this.weight = descriptors.weight!;
        this.style = descriptors.style!;
      }
      async load() { this.status = 'loaded'; return this; }
    }
    let probes = 0;
    const context = contractCanvas();
    vi.stubGlobal('FontFace', LocalFace);
    // Each strut keeps its own style: a batch holds several at once.
    vi.stubGlobal('document', {
      fonts, body: { append() {} },
      defaultView: {
        getComputedStyle: (strut: { style: { cssText: string } }) => ({ fontFamily: /px ("[^"]*")/.exec(strut.style.cssText)?.[1] ?? '' }),
      },
      createElement: (tag: string) => tag === 'canvas' ? { getContext: () => context }
        : tag === 'div' ? { style: { cssText: '' }, attachShadow: () => ({ append() {} }), remove() {} }
        : {
          style: { cssText: '' },
          textContent: '',
          get offsetHeight() {
            probes++;
            return Math.round(2355 * Number(/(\d+)px/.exec(this.style.cssText)?.[1] ?? 0) / 2048);
          },
        },
    });
    let first: Awaited<ReturnType<typeof loadDocxOfficeFontFallbacks>> | undefined;
    let second: typeof first;
    try {
      first = await loadDocxOfficeFontFallbacks([{ family: 'Times New Roman' }], fonts);
      second = await loadDocxOfficeFontFallbacks([{ family: 'Times New Roman' }], fonts);
      const prepared = first.routes['times new roman'];
      expect(prepared.metric.lineHeightRatio).toBe(2355 / 2048);
      expect(prepared.metric.family).toBe(prepared.family);
      expect(prepared.metric.sourceIdentity).toBe(prepared.resourceIdentity);
      expect(second.routes['times new roman'].metric).toEqual(prepared.metric);
      // The requested em and half of it, once per retained resource.
      expect(probes).toBe(2);
      expect(registrations.size).toBe(2);
    } finally {
      if (first) unloadOfficeFontFallbacks(first.faces);
      if (second) unloadOfficeFontFallbacks(second.faces);
      vi.unstubAllGlobals();
    }
    expect(registrations.size).toBe(0);
  });

  it('admits DOM-realm worker measurements only for the resolved source and tuple', async () => {
    const fonts = new Set<FontFace>() as unknown as FontFaceSet;
    class ProbeFace {
      status = 'loaded';
      weight: string;
      style: string;
      sizeAdjust: string;
      constructor(readonly family: string, _source: string,
        descriptors: FontFaceDescriptors & { sizeAdjust: string }) {
        this.weight = descriptors.weight!;
        this.style = descriptors.style!;
        this.sizeAdjust = descriptors.sizeAdjust;
      }
      async load() { return this; }
    }
    vi.stubGlobal('document', undefined);
    vi.stubGlobal('FontFace', ProbeFace);
    vi.stubGlobal('OffscreenCanvas', class { getContext() { return contractCanvas(); } });
    const identity = 'office-local:local("exact-host-source")';
    const loadedRoute = { ...route(), resourceIdentity: identity,
      metric: { ...route().metric, sourceIdentity: identity } };
    const face = { family: loadedRoute.family, weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    const measurements = { resourceIdentity: identity, weight: 400, style: 'normal' as const, measured: host };
    const prepared: FontFace[] = [];
    try {
      for (const mismatch of [{ resourceIdentity: 'other-source' }, { weight: 700 }, { style: 'italic' as const }]) {
        const result = await prepareDocxHostFontMetrics({ faces: [face], routes: { host: loadedRoute }, checked: [] },
          fonts, [{ ...measurements, ...mismatch }]);
        expect(result.routes.host.metric.lineHeightRatio).toBeUndefined();
      }
      const result = await prepareDocxHostFontMetrics({ faces: [face], routes: { host: loadedRoute }, checked: [] },
        fonts, [measurements]);
      prepared.push(...result.faces);
      expect(result.routes.host.metric.lineHeightRatio).toBe(2355 / 2048);
      expect(result.routes.host.metric.requiresGlyphCoverage).toBe(true);
      expect(result.routes.host.metric.glyphCoverageProbeFamily).toBeDefined();
    } finally {
      unloadOfficeFontFallbacks(prepared);
      vi.unstubAllGlobals();
    }
    expect([...fonts]).toHaveLength(0);
  });

  it('retains one coverage probe for aliases that share a paint face', async () => {
    const fonts = new Set<FontFace>() as unknown as FontFaceSet;
    class ProbeFace {
      status = 'loaded';
      weight: string;
      style: string;
      sizeAdjust: string;
      constructor(readonly family: string, _source: string,
        descriptors: FontFaceDescriptors & { sizeAdjust: string }) {
        this.weight = descriptors.weight!;
        this.style = descriptors.style!;
        this.sizeAdjust = descriptors.sizeAdjust;
      }
      async load() { return this; }
    }
    vi.stubGlobal('document', undefined);
    vi.stubGlobal('FontFace', ProbeFace);
    vi.stubGlobal('OffscreenCanvas', class { getContext() { return contractCanvas(); } });
    const identity = 'office-local:local("exact-host-source")';
    const loadedRoute = { ...route(), resourceIdentity: identity,
      metric: { ...route().metric, sourceIdentity: identity } };
    const face = { family: loadedRoute.family, weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    let result: Awaited<ReturnType<typeof prepareDocxHostFontMetrics>> | undefined;
    try {
      result = await prepareDocxHostFontMetrics({ faces: [face],
        routes: { first: loadedRoute, alias: { ...loadedRoute, requestedFamily: 'Alias Spelling' } }, checked: [] },
        fonts, [{ resourceIdentity: identity, weight: 400, style: 'normal', measured: host }]);
      const probes = result.faces.filter((candidate) => candidate !== face);
      expect(probes).toHaveLength(1);
      expect(result.routes.first.metric.glyphCoverageProbeFamily).toBe(probes[0].family);
      expect(result.routes.alias.metric.glyphCoverageProbeFamily).toBe(probes[0].family);
    } finally {
      if (result) unloadOfficeFontFallbacks(result.faces.filter((candidate) => candidate !== face));
      vi.unstubAllGlobals();
    }
    expect([...fonts]).toHaveLength(0);
  });

  it('keeps the catalog when the engine cannot prove glyph ownership', async () => {
    // An engine that zeroes every glyph under size-adjust 0% (owned or not)
    // cannot distinguish the paint face from fallback, so the measured
    // projection is declined rather than suppressing both geometries.
    const fonts = new Set<FontFace>() as unknown as FontFaceSet;
    class ProbeFace {
      status = 'loaded';
      weight: string;
      style: string;
      sizeAdjust: string;
      constructor(readonly family: string, _source: string,
        descriptors: FontFaceDescriptors & { sizeAdjust: string }) {
        this.weight = descriptors.weight!;
        this.style = descriptors.style!;
        this.sizeAdjust = descriptors.sizeAdjust;
      }
      async load() { return this; }
    }
    vi.stubGlobal('document', undefined);
    vi.stubGlobal('FontFace', ProbeFace);
    vi.stubGlobal('OffscreenCanvas', class {
      getContext() {
        return { font: '', measureText: () => ({ width: 0, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 0 }) };
      }
    });
    const identity = 'office-local:local("exact-host-source")';
    const loadedRoute = { ...route(), resourceIdentity: identity,
      metric: { ...route().metric, sourceIdentity: identity } };
    const face = { family: loadedRoute.family, weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    try {
      const result = await prepareDocxHostFontMetrics({ faces: [face], routes: { host: loadedRoute }, checked: [] },
        fonts, [{ resourceIdentity: identity, weight: 400, style: 'normal', measured: host }]);
      expect(result.routes.host.metric).toBe(loadedRoute.metric);
      expect(result.routes.host.metric.requiresGlyphCoverage).toBeUndefined();
    } finally {
      vi.unstubAllGlobals();
    }
    expect([...fonts]).toHaveLength(0);
  });

  it('uses measured host leading ahead of the catalog and retains the exact paint resource', () => {
    const loaded = route();
    const metric = hostFontLineMetric(loaded, host);
    const selected = select(loaded);
    const line = selectedFontLineMetric(selected, metric, true);
    expect(line.resourceMetric).toBe(metric);
    expect(line.referenceMetric).toBeUndefined();
    expect(line.lineMetric?.lineHeightRatio).toBe(2355 / 2048);
    expect(line.lineMetric?.designAscentRatio).toBe(1912 / 2048);
    expect(metric.family).toBe(selected.resolvedFamily);
    expect(metric.sourceIdentity).toBe(selected.resourceIdentity);
    expect(canvasFontString(selected.route, 11, 400, 'normal')).toContain('"exact-host-face"');
    expect(selectedFontLineMetric(selected, metric, false).lineMetric).toBeUndefined();
  });

  it('uses the catalog when the host is unmeasurable, without claiming catalog resource identity', () => {
    const loaded = route();
    const metric = hostFontLineMetric(loaded, null);
    const line = selectedFontLineMetric(select(loaded), metric, true);
    expect(line.resourceMetric).toBeUndefined();
    expect(line.referenceMetric?.lineHeightRatio).toBe(2268 / 2048);
  });

  it('does not lend authored catalog geometry to a measured substitute', () => {
    const loaded = route('Meiryo', 'substitute');
    const metric = hostFontLineMetric(loaded, host);
    expect(selectedFontLineMetric(select(loaded), metric, true).lineMetric?.lineHeightRatio).toBe(2355 / 2048);
    expect(selectedFontLineMetric(select(loaded), hostFontLineMetric(loaded, null), true).lineMetric).toBeUndefined();
  });

  it('preserves the selected host usWin projection; a CSS normal line cannot replace unavailable OS/2 data', () => {
    const loaded = route('Yu Mincho');
    // This host's tables deliberately differ from the catalog and from hhea.
    const measured = wordOpenTypeAutoLineRatios({
      unitsPerEm: 1000, hheaAscent: 800, hheaDescent: -200, hheaLineGap: 400,
      winAscent: 1050, winDescent: 350, farEastCodePage: true,
    })!;
    const metric = hostFontLineMetric({ ...loaded, metric: { ...loaded.metric, ...measured } }, host);
    expect(selectedFontLineMetric(select(loaded), metric, true).lineMetric?.lineHeightRatio).toBeCloseTo(1.3 * 1.4, 12);
    // local() exposes no usWin bytes: retain #1832's catalog fallback in that case.
    const fallback = hostFontLineMetric(loaded, host);
    expect(selectedFontLineMetric(select(loaded), fallback, true).lineMetric?.lineHeightRatio)
      .toBeCloseTo(1.3 * 2636 / 2048, 12);
  });

  it('sizes an East Asian-probed empty mark from its Latin host face; the unpainted probe needs no glyph', () => {
    // §17.3.1.29: an empty paragraph's mark line is its own font's line box.
    // Japanese-language paragraphs probe the mark with an East Asian glyph,
    // which a Latin face authored in the East Asian slot does not own. The mark
    // paints nothing, so that missing glyph must not discard the measured host
    // geometry (nor fall through to the fallback face that would paint it).
    vi.stubGlobal('OffscreenCanvas', class { getContext() { return contractCanvas(); } });
    try {
      const loaded = route();
      const metric = Object.freeze({
        ...hostFontLineMetric(loaded, host), requiresGlyphCoverage: true,
        glyphCoverageProbeFamily: 'exact-host-face_coverage',
      });
      const text = createTextLayoutService({
        fonts: createFontResolver([{
          requestedFamily: loaded.requestedFamily, resolvedFamily: loaded.family,
          source: loaded.source, resourceIdentity: loaded.resourceIdentity,
        }]),
        measurer: {
          fingerprint: 'east-asian-mark-probe',
          // A fallback face would paint the probe with a different box.
          measure: () => ({ advancePt: 10, ascentPt: 12, descentPt: 3 }),
        },
        fontMetrics: { host: metric },
      });
      const paragraph = {
        runs: [], defaultFontFamily: 'Times New Roman',
        defaultFontFamilyEastAsia: 'Times New Roman', defaultFontSize: 14, lineSpacing: null,
      } as unknown as DocParagraph;
      const mark = paragraphMarkLineMetrics(
        paragraph, 1, undefined, false, true, undefined, {}, null, {}, text,
      );
      expect(mark.advancePx).toBeCloseTo(14 * 2355 / 2048, 8);
      expect(mark.ascentPx).toBeCloseTo(14 * 1912 / 2048, 8);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
