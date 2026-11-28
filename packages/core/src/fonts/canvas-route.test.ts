import { describe, expect, it } from 'vitest';
import { canvasFontRouteFaces, canvasFontRouteWithLanguage, canvasFontString, createCanvasFontRoute } from './canvas-route.js';

describe('Canvas font route', () => {
  it('serializes the immutable complete family list without adding a fallback', () => {
    const route = createCanvasFontRoute('"Times New Roman", serif', 'native');
    expect(canvasFontString(route, 12, 700, 'italic'))
      .toBe('italic 700 12px "Times New Roman", serif');
    expect(createCanvasFontRoute('"Times New Roman", serif', 'native').fingerprint)
      .toBe(route.fingerprint);
    expect(createCanvasFontRoute('"Times New Roman", serif', 'registered').fingerprint)
      .not.toBe(route.fingerprint);
    expect(createCanvasFontRoute('a:b', 'native').fingerprint)
      .not.toBe(createCanvasFontRoute('a%3Ab', 'native').fingerprint);
  });

  it('makes the language part of the route identity and leaves untagged routes unchanged', () => {
    const route = createCanvasFontRoute('"Times New Roman", serif', 'native');
    expect(route.fingerprint).toBe('canvas-font-route-v1:native:%22Times%20New%20Roman%22%2C%20serif');
    expect('lang' in route).toBe(false);
    expect(canvasFontRouteWithLanguage(route, undefined)).toBe(route);
    expect(canvasFontRouteFaces(route)).toBe(route);

    const serbian = canvasFontRouteWithLanguage(route, 'sr-rs');
    expect(serbian.lang).toBe('sr-rs');
    expect(serbian.fingerprint).toBe(`${route.fingerprint}:lang:sr-rs`);
    expect([serbian.familyList, serbian.scope]).toEqual([route.familyList, route.scope]);
    expect(canvasFontString(serbian, 12, 400, 'normal')).toBe(canvasFontString(route, 12, 400, 'normal'));
    expect(Object.isFrozen(serbian)).toBe(true);
    // One object per route and language; the faces are the route it came from.
    expect(canvasFontRouteWithLanguage(route, 'sr-rs')).toBe(serbian);
    expect(canvasFontRouteWithLanguage(serbian, 'sr-rs')).toBe(serbian);
    expect(canvasFontRouteFaces(serbian)).toBe(route);
    expect(canvasFontRouteWithLanguage(route, 'ru-ru').fingerprint).not.toBe(serbian.fingerprint);
    expect(() => canvasFontRouteWithLanguage(serbian, 'ru-ru')).toThrow(TypeError);
    // A family list cannot spell another route's language suffix.
    expect(createCanvasFontRoute(`${route.familyList}:lang:sr-rs`, 'native').fingerprint)
      .not.toBe(serbian.fingerprint);
  });
});
