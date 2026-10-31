import { afterEach, describe, expect, it, vi } from 'vitest';
import { loadOfficeFontFallbacks, unloadOfficeFontFallbacks } from '@silurus/ooxml-core';
import type { OfficeFontFallbackRequest } from '@silurus/ooxml-core';
import { PptxPresentation } from './presentation.js';
import { pptxSlideOfficeFontRequests } from './google-fonts.js';
import { officeFaceWeight } from './powerpoint-line-metrics.js';
import { buildFont, layoutParagraph } from './renderer.js';
import { acquireOfficeRoutes } from './office-route-acquisition.js';
import type { Paragraph, Slide } from './types.js';

// Line metrics admit a face's catalog geometry only when that face is present
// (its exact local() route was loaded). Acquisition must therefore probe every
// authored tuple, under the same weight the metric lookup uses.

function fakeHost() {
  const sources: string[] = [];
  class Face {
    status = 'unloaded';
    constructor(public family: string, public source: string, public descriptors: unknown) { sources.push(source); }
    async load() { this.status = 'loaded'; return this; }
  }
  vi.stubGlobal('FontFace', Face);
  const set = { add() {}, delete() { return true; }, *[Symbol.iterator]() {} } as unknown as FontFaceSet;
  return { sources, set };
}

function owner() {
  const value = Object.create(PptxPresentation.prototype) as {
    _destroyed: boolean; _embeddedFontTuples: Set<string>; _officeFontLoads: Map<unknown, unknown>;
    _officeRoutesForRequests(requests: OfficeFontFallbackRequest[], set: FontFaceSet): Promise<{ routes: Record<string, unknown>; installed: string[] }>;
  };
  value._destroyed = false;
  value._embeddedFontTuples = new Set();
  value._officeFontLoads = new Map();
  return value;
}

function paragraph(): Paragraph {
  return {
    runs: [{ type: 'text', text: 'A', fontFamily: 'Arial', fontSize: 60, bold: false, italic: false },
      { type: 'break', fontFamily: 'Meiryo', fontSize: 16 }],
    endRunProperties: { fontFamily: 'Yu Gothic', fontSize: 16 },
    alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null, spaceLine: null,
    lvl: 0, bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [],
  } as unknown as Paragraph;
}

afterEach(() => vi.unstubAllGlobals());

describe('PPTX exact local() routes for absent-face line metrics', () => {
  it('acquires every authored family, not only one per weight/style slot', async () => {
    const { set } = fakeHost();
    const routes = await owner()._officeRoutesForRequests([
      { family: 'Arial', weight: 400, style: 'normal' },
      { family: 'Calibri', weight: 400, style: 'normal' }], set);
    expect(Object.keys(routes.routes).sort()).toEqual(['arial', 'calibri']);
  });

  it('does not let an earlier failed tuple poison a later family', async () => {
    const { set } = fakeHost();
    const value = owner();
    // No catalog 400 profile exists for Calibri Light: no route, but no poisoning.
    await value._officeRoutesForRequests([{ family: 'Calibri Light', weight: 400, style: 'normal' }], set);
    const routes = await value._officeRoutesForRequests([{ family: 'Arial', weight: 400, style: 'normal' }], set);
    expect(routes.routes.arial).toBeDefined();
  });

  it('loads each slide\'s new tuples in one bounded loader call', async () => {
    const calls: OfficeFontFallbackRequest[][] = [];
    const cache = new Map();
    const load = async (requests: OfficeFontFallbackRequest[]) => { calls.push(requests); return { faces: [], routes: {} }; };
    await acquireOfficeRoutes(cache, [{ family: 'Arial' }, { family: 'Calibri' }, { family: 'Arial' }], load);
    await acquireOfficeRoutes(cache, [{ family: 'Arial' }, { family: 'Georgia', weight: 700 }], load);
    expect(calls.map((batch) => batch.map((request) => request.family))).toEqual([['Arial', 'Calibri'], ['Georgia']]);
    // Family presence is a separate cached probe, independent of styled tuples.
    const presence = await acquireOfficeRoutes(cache, [{ family: 'Arial', presenceOnly: true }],
      async (requests) => { calls.push(requests); return { faces: [], routes: {}, installed: ['arial'] }; });
    expect(presence.installed).toEqual(['arial']);
  });

  it('returns only the routes this call requested, independent of earlier slides', async () => {
    const cache = new Map();
    const route = (family: string) => ({ family } as never);
    const load = async () => ({ faces: [], routes: { arial: route('arial'), 'yu gothic': route('yu gothic') } });
    await acquireOfficeRoutes(cache, [{ family: 'Arial' }, { family: 'Yu Gothic' }], load);
    const later = await acquireOfficeRoutes(cache, [{ family: 'Arial' }], load);
    expect(Object.keys(later.routes)).toEqual(['arial']);
  });

  it('treats families named after prototype members like any other family', async () => {
    const names = ['__proto__', 'constructor', 'toString', 'hasOwnProperty'];
    const requests = names.map((family) => ({ family }));
    // No loaded route: no name may be marked installed or routed via Object.prototype.
    const none = await acquireOfficeRoutes(new Map(), requests, async () => ({ faces: [], routes: {} }));
    expect(none).toEqual({ routes: {}, installed: [] });
    // A worker delivers loaded routes through structured clone: `__proto__` is an own key there.
    const route = { family: '__ooxml_local___proto__' } as never;
    const cloned = structuredClone(JSON.parse('{"__proto__": {"family": "__ooxml_local___proto__"}}'));
    const loaded = await acquireOfficeRoutes(new Map(), requests, async () => ({ faces: [], routes: cloned }));
    expect(loaded.installed).toEqual(['__proto__']);
    expect(Object.entries(loaded.routes)).toEqual([['__proto__', route]]);
  });

  it('probes an installed named cut under its catalogued weight', async () => {
    expect(officeFaceWeight('Calibri Light', false, false)).toBe(300);
    // A bold run synthesizes bold on the Light face, which is still present.
    expect(officeFaceWeight('Calibri Light', true, false)).toBe(300);
    expect(officeFaceWeight('Arial', true, false)).toBe(700);
    const slide = { elements: [{ type: 'shape', textBody: { paragraphs: [{
      ...paragraph(), runs: [{ type: 'text', text: 'T', fontFamily: 'Calibri Light', fontSize: 40, bold: true, italic: false }],
      endRunProperties: undefined,
    }] } }] } as unknown as Slide;
    const requests = pptxSlideOfficeFontRequests(slide, null, null);
    expect(requests).toEqual([{ family: 'Calibri Light', weight: 300, style: 'normal' },
      { family: 'Calibri Light', presenceOnly: true }]);
    const { set } = fakeHost();
    const loaded = await loadOfficeFontFallbacks(requests, set);
    try {
      expect(loaded.routes['calibri light:300:normal']).toBeDefined();
    } finally {
      unloadOfficeFontFallbacks(loaded.faces);
    }
  });

  it('requests the faces of line breaks and end marks, which contribute line metrics', () => {
    const requests = pptxSlideOfficeFontRequests(
      { elements: [{ type: 'shape', textBody: { paragraphs: [paragraph()] } }] } as unknown as Slide, null, null);
    expect(requests.map((request) => request.family)).toEqual(expect.arrayContaining(['Arial', 'Meiryo', 'Yu Gothic']));
  });

  it('uses a non-Calibri exact private alias in both layout measurement and retained paint style', () => {
    const fonts: string[] = [];
    const ctx = { font: '', measureText: () => { fonts.push(ctx.font); return { width: 10 }; } } as unknown as CanvasRenderingContext2D;
    const alias = '__exact_arial';
    const rc = { themeMajorFont: null, themeMinorFont: null, dpr: 1,
      officeFontRoutes: { arial: { family: alias } } } as unknown as Parameters<typeof buildFont>[4];
    const p = paragraph();
    p.runs = p.runs.slice(0, 1);
    (p as { endRunProperties?: unknown }).endRunProperties = undefined;
    const lines = layoutParagraph(ctx, p, 1000, 60, '000000', 1 / 12700, 0, false, false, 1, 1, rc);
    expect(buildFont(false, false, 60, 'Arial', rc)).toContain(alias);
    expect(fonts.every((font) => font.includes(alias))).toBe(true);
    expect(lines[0]!.segments[0]!.font).toContain(alias);
  });

  it('derives family presence from a loaded route of any requested tuple', async () => {
    const load = async () => ({ faces: [], routes: { 'yu gothic': { family: 'x' } as never } });
    const acquired = await acquireOfficeRoutes(new Map(), [{ family: 'Yu Gothic' }, { family: 'Arial' }], load);
    expect(acquired.installed).toEqual(['yu gothic']);
  });

  it('does not report a family shadowed by an application FontFace as installed', async () => {
    // A declared FontFace owns the authored family name for every tuple, so the
    // system face's presence or routes cannot describe what paints.
    const declared = [{ family: '"Arial"', status: 'loaded' }] as unknown as FontFace[];
    const set = { [Symbol.iterator]: () => declared[Symbol.iterator]() } as unknown as FontFaceSet;
    const load = async () => ({ faces: [], routes: { 'arial:700:normal': { family: 'x' } as never }, installed: ['arial', 'georgia'] });
    const acquired = await acquireOfficeRoutes(new Map(), [
      { family: 'Arial', weight: 700, style: 'normal' }, { family: 'Arial', presenceOnly: true },
      { family: 'Georgia', presenceOnly: true },
    ], load, set);
    expect(Object.keys(acquired.routes)).toEqual(['arial:700:normal']);
    expect(acquired.installed).toEqual(['georgia']);
  });
});
