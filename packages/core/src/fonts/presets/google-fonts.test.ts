import { describe, expect, it } from 'vitest';
import { assertFontSourceDescriptors, fontCssApiRequests, type FontCssApiSource } from '../source.js';
import { GOOGLE_FONTS_CATALOG, GOOGLE_FONTS_SUBSTITUTIONS, GOOGLE_NOTO_CJK_NAMING, googleFonts } from './google-fonts.js';
import { OFFICE_FONT_SUBSTITUTIONS } from '../substitutions.js';
import { cjkFallbackChain, NON_CJK_SANS_FALLBACKS, NON_CJK_SERIF_FALLBACKS } from '../scripts.js';
import { arabicFallbackFamilies } from '../font-stack.js';

/**
 * Oracle: every request the loaders issued with `useGoogleFonts: true`, frozen
 * from the former per-format tables (`DOCX_GOOGLE_FONTS`, `PPTX_GOOGLE_FONTS`,
 * `XLSX_GOOGLE_FONTS`, each exactly the former `GOOGLE_FONT_SUBSTITUTES` +
 * `SCRIPT_GOOGLE_FONTS`) as of 0.89.4, which no longer exist. Row: authored family key, the former
 * `loadFamily` (the family the request serves, when recorded), script scope,
 * and the CSS2 `family=` value. The preset must issue byte-identical requests.
 */
const LEGACY_REQUESTS: readonly (readonly [string, string | null, 'arabic' | null, string])[] = [
  ["calibri", "Carlito", null, "Carlito:ital,wght@0,400;0,700;1,400;1,700"],
  ["cambria", "Caladea", null, "Caladea:ital,wght@0,400;0,700;1,400;1,700"],
  ["franklin gothic book", "Libre Franklin", null, "Libre+Franklin:ital,wght@0,400;0,600;0,700;1,400;1,600;1,700"],
  ["franklin gothic medium", "Libre Franklin", null, "Libre+Franklin:ital,wght@0,400;0,600;0,700;1,400;1,600;1,700"],
  ["nunito sans", null, null, "Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700"],
  ["nunito", null, null, "Nunito:ital,wght@0,400;0,700;1,400;1,700"],
  ["open sans", null, null, "Open+Sans:ital,wght@0,400;0,700;1,400;1,700"],
  ["roboto", null, null, "Roboto:ital,wght@0,400;0,700;1,400;1,700"],
  ["lato", null, null, "Lato:ital,wght@0,400;0,700;1,400;1,700"],
  ["montserrat", null, null, "Montserrat:ital,wght@0,400;0,700;1,400;1,700"],
  ["poppins", null, null, "Poppins:ital,wght@0,400;0,700;1,400;1,700"],
  ["raleway", null, null, "Raleway:ital,wght@0,400;0,700;1,400;1,700"],
  ["playfair display", null, null, "Playfair+Display:ital,wght@0,400;0,700;1,400;1,700"],
  ["ubuntu", null, null, "Ubuntu:ital,wght@0,400;0,700;1,400;1,700"],
  ["sakkal majalla", "Noto Naskh Arabic", "arabic", "Noto+Naskh+Arabic:wght@400;700"],
  ["traditional arabic", "Noto Naskh Arabic", "arabic", "Noto+Naskh+Arabic:wght@400;700"],
  ["simplified arabic", "Noto Naskh Arabic", "arabic", "Noto+Naskh+Arabic:wght@400;700"],
  ["arabic typesetting", "Noto Naskh Arabic", "arabic", "Noto+Naskh+Arabic:wght@400;700"],
  ["univers next arabic", "Noto Sans Arabic", "arabic", "Noto+Sans+Arabic:wght@400;700"],
  ["noto naskh arabic", "Noto Naskh Arabic", null, "Noto+Naskh+Arabic:wght@400;700"],
  ["noto sans arabic", "Noto Sans Arabic", null, "Noto+Sans+Arabic:wght@400;700"],
  ["noto sans kr", null, null, "Noto+Sans+KR:wght@400;700"],
  ["noto sans sc", null, null, "Noto+Sans+SC:wght@400;700"],
  ["noto sans tc", null, null, "Noto+Sans+TC:wght@400;700"],
  ["noto sans hk", null, null, "Noto+Sans+HK:wght@400;700"],
  ["noto sans jp", null, null, "Noto+Sans+JP:wght@400;700"],
  ["noto serif kr", null, null, "Noto+Serif+KR:wght@400;700"],
  ["noto serif sc", null, null, "Noto+Serif+SC:wght@400;700"],
  ["noto serif tc", null, null, "Noto+Serif+TC:wght@400;700"],
  ["noto serif jp", null, null, "Noto+Serif+JP:wght@400;700"],
  ["noto sans cjk sc", "Noto Sans SC", null, "Noto+Sans+SC:wght@400;700"],
  ["noto serif cjk sc", "Noto Serif SC", null, "Noto+Serif+SC:wght@400;700"],
  ["noto sans cjk tc", "Noto Sans TC", null, "Noto+Sans+TC:wght@400;700"],
  ["noto serif cjk tc", "Noto Serif TC", null, "Noto+Serif+TC:wght@400;700"],
  ["noto sans cjk jp", "Noto Sans JP", null, "Noto+Sans+JP:wght@400;700"],
  ["noto serif cjk jp", "Noto Serif JP", null, "Noto+Serif+JP:wght@400;700"],
  ["noto sans cjk kr", "Noto Sans KR", null, "Noto+Sans+KR:wght@400;700"],
  ["noto serif cjk kr", "Noto Serif KR", null, "Noto+Serif+KR:wght@400;700"],
  ["noto sans cjk hk", "Noto Sans HK", null, "Noto+Sans+HK:wght@400;700"],
  ["noto sans", null, null, "Noto+Sans:wght@400;700"],
  ["noto serif", null, null, "Noto+Serif:wght@400;700"],
  ["noto sans devanagari", null, null, "Noto+Sans+Devanagari:wght@400;700"],
  ["noto sans thai", null, null, "Noto+Sans+Thai:wght@400;700"],
  ["noto sans hebrew", null, null, "Noto+Sans+Hebrew:wght@400;700"],
  ["noto serif hebrew", null, null, "Noto+Serif+Hebrew:wght@400;700"],
];

const legacyUrl = (family: string, origin = 'https://fonts.googleapis.com') =>
  `${origin}/css2?family=${family}&display=swap`;

/** The single request the preset issues for one authored family. */
function onlyRequest(source: FontCssApiSource, family: string) {
  const requests = fontCssApiRequests(source, family);
  expect(requests, family).toHaveLength(1);
  return requests[0];
}

describe('googleFonts() preset', () => {
  it('issues byte-identical requests and the same served family and scope per authored name', () => {
    const source = googleFonts();
    for (const [key, served, script, family] of LEGACY_REQUESTS) {
      const request = onlyRequest(source, key);
      expect(request.url, key).toBe(legacyUrl(family));
      expect(request.family.toLowerCase(), key).toBe((served ?? key).toLowerCase());
      const substitutes = served !== null && served.toLowerCase() !== key;
      expect(request.substitution?.script ?? null, key).toBe(substitutes ? script : null);
      expect(request.substitution !== undefined, key).toBe(substitutes);
    }
  });

  it('answers only the former tables\' names plus the substitute families themselves', () => {
    const source = googleFonts();
    const answered = new Set([
      ...Object.keys(GOOGLE_FONTS_CATALOG),
      ...(source.substitutes ?? []).filter((entry) => fontCssApiRequests(source, entry.for).length > 0)
        .map((entry) => entry.for),
    ].map((family) => family.toLowerCase()));
    // The one deliberate widening: a document that authors a served
    // substitute family by name (as LibreOffice output does) requests it
    // directly; the former tables keyed those families only as substitutes.
    expect(answered).toEqual(new Set([...LEGACY_REQUESTS.map(([key]) => key),
      'carlito', 'caladea', 'libre franklin']));
    // Distinct Office faces without a compatible substitute stay unanswered.
    for (const family of ['Calibri Light', 'Cambria Math', 'Noto Serif CJK HK', 'Noto Serif HK']) {
      expect(fontCssApiRequests(source, family), family).toEqual([]);
    }
  });

  it('serves every library fallback family a renderer stack may append', () => {
    const source = googleFonts();
    const fallbacks = new Set([
      ...(['kr', 'sc', 'tc', 'hk', 'jp'] as const).flatMap((lang) =>
        [...cjkFallbackChain(lang, 'sans'), ...cjkFallbackChain(lang, 'serif')]),
      ...NON_CJK_SANS_FALLBACKS, ...NON_CJK_SERIF_FALLBACKS,
      ...arabicFallbackFamilies(true), ...arabicFallbackFamilies(false),
    ]);
    for (const family of fallbacks) {
      expect(onlyRequest(source, family).family, family).toBe(family);
    }
  });

  it('moves every request to a validated custom origin and nothing else (#1497)', () => {
    for (const [cssOrigin, origin] of [
      ['https://fonts.googleapis.cn', 'https://fonts.googleapis.cn'],
      ['https://fonts.internal.example:8443/', 'https://fonts.internal.example:8443'],
      ['http://localhost:8080', 'http://localhost:8080'],
    ] as const) {
      const source = googleFonts({ cssOrigin });
      for (const [key, , , family] of LEGACY_REQUESTS) {
        expect(onlyRequest(source, key).url).toBe(legacyUrl(family, origin));
      }
      expect(assertFontSourceDescriptors(structuredClone([source]))).toEqual([source]);
    }
    for (const cssOrigin of ['fonts.googleapis.cn', 'ftp://fonts.example', 'https://fonts.example/css2',
      'https://fonts.example?x=1', 'https://user@fonts.example', 'https://fonts.example//', 7]) {
      expect(() => googleFonts({ cssOrigin } as { cssOrigin: string }), String(cssOrigin)).toThrow(TypeError);
    }
    expect(() => googleFonts({ origin: 'https://fonts.example' } as never)).toThrow(TypeError);
  });

  it('uses the exported default substitutes, which substitutes replaces entirely', () => {
    expect(GOOGLE_FONTS_SUBSTITUTIONS).toEqual([...OFFICE_FONT_SUBSTITUTIONS, ...GOOGLE_NOTO_CJK_NAMING]);
    expect(Object.isFrozen(GOOGLE_FONTS_SUBSTITUTIONS)).toBe(true);
    expect(googleFonts().substitutes).toEqual(GOOGLE_FONTS_SUBSTITUTIONS);
    expect(googleFonts({ substitutes: GOOGLE_FONTS_SUBSTITUTIONS })).toEqual(googleFonts());

    const own = googleFonts({ substitutes: [{ for: 'Corporate Sans', use: 'Open Sans' }] });
    expect(own.substitutes).toEqual([{ for: 'Corporate Sans', use: 'Open Sans' }]);
    expect(fontCssApiRequests(own, 'Corporate Sans').map((request) => request.family)).toEqual(['Open Sans']);
    // Replaced, not merged: Calibri now asks Google for Calibri only, which
    // the catalog does not list, so nothing is requested.
    expect(fontCssApiRequests(own, 'Calibri')).toEqual([]);

    const extended = googleFonts({ substitutes: [...GOOGLE_FONTS_SUBSTITUTIONS, { for: 'Corporate Sans', use: 'Lato' }] });
    expect(fontCssApiRequests(extended, 'Calibri').map((request) => request.family)).toEqual(['Carlito']);
    expect(fontCssApiRequests(extended, 'Corporate Sans').map((request) => request.family)).toEqual(['Lato']);
    expect(googleFonts({ substitutes: [] }).substitutes).toEqual([]);
    expect(() => googleFonts({ substitutes: [{ for: 'A' }] } as never)).toThrow(TypeError);
  });
});
