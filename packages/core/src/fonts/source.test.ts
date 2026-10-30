import { afterEach, describe, expect, it, vi } from 'vitest';
import { OoxmlError } from '../errors/ooxml-error.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_FACE_BYTES,
  HARD_MAX_FONT_SOURCE_URL_LENGTH,
} from '../worker/resource-policy.generated.js';
import {
  FONT_SOURCE_PROTOCOL,
  assertFontSourceDescriptors,
  fontCssApi,
  fontCssApiRequests,
  fontCss,
  fontFiles,
  fontResolver,
  dataUrlDecodedByteLength,
  fontResolverFunction,
  validateFontFaceDeclaration,
  type FontResolveFn,
} from './source.js';
import { OFFICE_FONT_SUBSTITUTIONS } from './substitutions.js';

function limitOf(run: () => unknown): unknown {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe('font-source-limit');
    return (error as OoxmlError).violation;
  }
  throw new Error('expected a font-source limit');
}

afterEach(() => vi.unstubAllGlobals());

describe('font source descriptors', () => {
  it('produce frozen, structured-clone-safe descriptors a worker re-validates identically', () => {
    const bytes = new Uint8Array([0, 1, 0, 0]).buffer;
    const resolver = fontResolver(() => []);
    const sources = [
      fontCssApi({ urlTemplate: 'https://fonts.example/css?family={family}:{query}', query: '400,700',
        substitutes: OFFICE_FONT_SUBSTITUTIONS, credentials: 'omit' }),
      fontCss({ url: 'https://cdn.example/office.css' }),
      fontFiles({ faces: [
        { family: 'Carlito', url: 'https://cdn.example/Carlito.woff2', weight: '400', unicodeRange: 'U+0-FF' },
        { family: 'Meiryo', data: bytes },
      ], credentials: 'omit' }),
      resolver,
    ];
    expect(sources.map((source) => source.kind)).toEqual(['font-css-api', 'font-css', 'font-files', 'font-resolver']);
    for (const source of sources) {
      expect(source.protocol).toBe(FONT_SOURCE_PROTOCOL);
      expect(Object.isFrozen(source)).toBe(true);
    }
    // The posted copy is a different object, so the worker validates it; the
    // result must describe the same source (no application function inside).
    const crossed = assertFontSourceDescriptors(structuredClone(sources));
    expect(crossed).toEqual(sources);
    expect(crossed[3]).toEqual({ protocol: FONT_SOURCE_PROTOCOL, kind: 'font-resolver', id: resolver.id });
    // Factory descriptors pass through without copying their face bytes.
    expect(assertFontSourceDescriptors(sources)[2]).toBe(sources[2]);
  });

  it('resolves relative application URLs against the calling realm, never in the receiving realm', () => {
    vi.stubGlobal('document', { baseURI: 'https://app.example/viewer/index.html' });
    const css = fontCss({ url: '../fonts/office.css' });
    const files = fontFiles({ faces: [{ family: 'Carlito', url: '/fonts/Carlito.woff2' }] });
    expect(css.url).toBe('https://app.example/fonts/office.css');
    expect(files.faces[0].url).toBe('https://app.example/fonts/Carlito.woff2');
    expect(() => assertFontSourceDescriptors([{ ...structuredClone(css), url: '/fonts/office.css' }]))
      .toThrow(TypeError);
  });

  it('rejects malformed descriptors with TypeError before any load', () => {
    const invalid: Array<() => unknown> = [
      () => fontCssApi({ urlTemplate: 'https://fonts.example/css' }),
      () => fontCssApi({ urlTemplate: 'https://fonts.example/css?family={family}&{subset}' }),
      () => fontCssApi({ urlTemplate: 'https://fonts.example/css?family={family}:{query}' }),
      () => fontCssApi({ urlTemplate: '/css?family={family}' }),
      () => fontCssApi({ urlTemplate: 'ftp://fonts.example/css?family={family}' }),
      () => fontCssApi({ urlTemplate: 'https://fonts.example/css?family={family}', families: { Lato: true, lato: true } }),
      () => fontCssApi({ urlTemplate: 'https://fonts.example/css?family={family}', credentials: 'include' as 'omit' }),
      () => fontCss({ url: 'blob:https://app.example/1' }),
      () => fontCss({ url: 'https://cdn.example/a.css', headers: {} } as never),
      () => fontFiles({ faces: [] }),
      () => fontFiles({ faces: [{ family: 'A', url: 'https://a.example/a.woff2' }], credentials: 'include' as 'omit' }),
      () => fontFiles({ faces: [{ family: 'A' }] }),
      () => fontFiles({ faces: [{ family: 'A', url: 'https://a.example/a.woff2', data: new ArrayBuffer(4) }] }),
      () => fontFiles({ faces: [{ family: 'A', data: new Uint8Array(4) as unknown as ArrayBuffer }] }),
      () => fontFiles({ faces: [{ family: 'A', data: new ArrayBuffer(0) }] }),
      () => fontFiles({ faces: [{ family: ' A', url: 'https://a.example/a.woff2' }] }),
      () => fontFiles({ faces: [{ family: 'A', url: 'javascript:alert(1)' }] }),
      () => fontFiles({ faces: [{ family: 'A', url: 'https://a.example/a.woff2' }],
        substitutes: [{ for: 'Calibri', use: 'Carlito', script: 'latin' as 'arabic' }] }),
      () => fontResolver('resolve' as unknown as FontResolveFn),
      () => assertFontSourceDescriptors({}),
      () => assertFontSourceDescriptors([{ protocol: 'other/v1', kind: 'font-files', faces: [] }]),
      () => assertFontSourceDescriptors([{ protocol: FONT_SOURCE_PROTOCOL, kind: 'font-resolver', id: 'x' }]),
    ];
    for (const run of invalid) expect(run).toThrow(TypeError);
  });

  it('fails closed with font-source-limit beyond the generated ceilings', () => {
    expect(limitOf(() => fontFiles({ faces: [{ family: 'A',
      data: new ArrayBuffer(HARD_MAX_FONT_SOURCE_FACE_BYTES + 1) }] })))
      .toEqual({ resource: 'font-source', quota: 'face', metric: 'bytes',
        limit: HARD_MAX_FONT_SOURCE_FACE_BYTES, observed: HARD_MAX_FONT_SOURCE_FACE_BYTES + 1 });
    const long = `https://a.example/${'a'.repeat(HARD_MAX_FONT_SOURCE_URL_LENGTH)}`;
    expect(limitOf(() => fontCss({ url: long }))).toMatchObject({ quota: 'url', metric: 'length' });
    expect(limitOf(() => fontFiles({ faces: Array.from({ length: HARD_MAX_FONT_SOURCE_DECLARATIONS + 1 },
      () => ({ family: 'A', url: 'https://a.example/a.woff2' })) })))
      .toMatchObject({ quota: 'declarations', metric: 'entry-count' });
    // A data: URL is charged against the face byte ceiling, not the URL
    // length ceiling, so small inline faces are accepted.
    const inline = `data:font/woff2;base64,${'A'.repeat(HARD_MAX_FONT_SOURCE_URL_LENGTH * 2)}`;
    expect(fontFiles({ faces: [{ family: 'A', url: inline }] }).faces[0].url).toBe(inline);
  });

  it('charges a data: URL its decoded body size in either encoding, never its header', () => {
    const face = HARD_MAX_FONT_SOURCE_FACE_BYTES;
    const accepts = (url: string) => expect(validateFontFaceDeclaration({ family: 'A', url }).url).toBe(url);
    const rejects = (url: string, observed: number) =>
      expect(limitOf(() => validateFontFaceDeclaration({ family: 'A', url })))
        .toEqual({ resource: 'font-source', quota: 'face', metric: 'bytes', limit: face, observed });
    // base64: exactly the ceiling (face is a multiple of 3) vs one byte more.
    accepts(`data:font/woff2;base64,${'A'.repeat(face / 3 * 4)}`);
    rejects(`data:font/woff2;base64,${'A'.repeat(face / 3 * 4 + 2)}==`, face + 1);
    // Percent-encoded (non-base64) bodies decode one byte per character.
    accepts(`data:font/woff2,${'A'.repeat(face)}`);
    rejects(`data:font/woff2,${'A'.repeat(face + 1)}`, face + 1);
    expect(() => validateFontFaceDeclaration({ family: 'A', url: 'data:font/woff2;base64' })).toThrow(TypeError);
  });

  it('counts data: bodies as the URL parser and Fetch decode them', () => {
    // Fragment, header, tab/newline and base64 whitespace/padding excluded;
    // %XX is one byte, non-ASCII its UTF-8 length (lone surrogate → U+FFFD).
    expect(dataUrlDecodedByteLength('data:,ab%41%zz#ignored')).toBe(6);
    expect(dataUrlDecodedByteLength('data:,\u00e9\u4e00\u{1f600}\ud800')).toBe(2 + 3 + 4 + 3);
    expect(dataUrlDecodedByteLength('DATA:text/plain ; BASE64 ,QU\tJD %20 RA==')).toBe(4);
    expect(dataUrlDecodedByteLength('data:;base64,QUJD%3D')).toBe(3);
    expect(dataUrlDecodedByteLength('data:no-comma')).toBeUndefined();
  });

  it('brands face buffers by their internal slots, not by tags or own properties', () => {
    const spoofed = { [Symbol.toStringTag]: 'ArrayBuffer', byteLength: 4 };
    expect(() => fontFiles({ faces: [{ family: 'A', data: spoofed as unknown as ArrayBuffer }] })).toThrow(TypeError);
    expect(() => fontFiles({ faces: [{ family: 'A',
      data: new Proxy(new ArrayBuffer(4), {}) }] })).toThrow(TypeError);
    const oversized = new ArrayBuffer(HARD_MAX_FONT_SOURCE_FACE_BYTES + 1);
    Object.defineProperty(oversized, 'byteLength', { value: 1 });
    expect(limitOf(() => fontFiles({ faces: [{ family: 'A', data: oversized }] })))
      .toMatchObject({ quota: 'face', observed: HARD_MAX_FONT_SOURCE_FACE_BYTES + 1 });
    const detached = new ArrayBuffer(4);
    structuredClone(detached, { transfer: [detached] });
    const resizable = new (ArrayBuffer as unknown as new (n: number, o: object) => ArrayBuffer)(4, { maxByteLength: 8 });
    for (const data of [new Uint8Array(4), new DataView(new ArrayBuffer(4)), new SharedArrayBuffer(4),
      detached, resizable]) {
      expect(() => fontFiles({ faces: [{ family: 'A', data: data as ArrayBuffer }] })).toThrow(TypeError);
    }
  });

  it('keeps the resolver function in this realm and shares one id per function', () => {
    const resolve: FontResolveFn = () => [];
    const first = fontResolver(resolve);
    const second = fontResolver(resolve, { substitutes: OFFICE_FONT_SUBSTITUTIONS });
    const other = fontResolver(() => []);
    expect(second.id).toBe(first.id);
    expect(other.id).not.toBe(first.id);
    expect(fontResolverFunction(second)).toBe(resolve);
    expect(fontResolverFunction(structuredClone(first))).toBeUndefined();
  });
});

describe('font-css-api request composition', () => {
  it('requests the authored family first, then declared substitutes, encoding the family only', () => {
    const source = fontCssApi({
      urlTemplate: 'https://fonts.example/css?family={family}:{query}',
      query: '400,700',
      substitutes: OFFICE_FONT_SUBSTITUTIONS,
    });
    expect(fontCssApiRequests(source, 'Calibri')).toEqual([
      { family: 'Calibri', url: 'https://fonts.example/css?family=Calibri:400,700' },
      { family: 'Carlito', url: 'https://fonts.example/css?family=Carlito:400,700',
        substitution: { for: 'Calibri', use: 'Carlito' } },
    ]);
    expect(fontCssApiRequests(source, 'Noto Sans & Co')[0].url)
      .toBe('https://fonts.example/css?family=Noto+Sans+%26+Co:400,700');
  });

  it('restricts requests to the allow-list, using its spelling and per-family query', () => {
    const source = fontCssApi({
      urlTemplate: 'https://fonts.example/css?family={family}:{query}',
      query: 'wght@400',
      families: { 'Open Sans': true, 'Libre Franklin': 'wght@400;600' },
      substitutes: [{ for: 'Franklin Gothic Book', use: 'Libre Franklin' }],
    });
    expect(fontCssApiRequests(source, 'OPEN SANS')).toEqual([
      { family: 'Open Sans', url: 'https://fonts.example/css?family=Open+Sans:wght@400' },
    ]);
    expect(fontCssApiRequests(source, 'franklin gothic book')).toEqual([
      { family: 'Libre Franklin', url: 'https://fonts.example/css?family=Libre+Franklin:wght@400;600',
        substitution: { for: 'Franklin Gothic Book', use: 'Libre Franklin' } },
    ]);
    expect(fontCssApiRequests(source, 'Comic Sans MS')).toEqual([]);
  });

  it('keeps allow-list families named after prototype members as own entries', () => {
    // A caller allow-list built from data (JSON, structured clone) carries
    // `__proto__` as an own key; plain-object assignment used to drop it.
    const source = fontCssApi({
      urlTemplate: 'https://fonts.example/css?family={family}',
      families: JSON.parse('{"__proto__": true, "constructor": true}') as Record<string, true>,
    });
    for (const descriptor of [source, structuredClone(source)]) {
      expect(fontCssApiRequests(descriptor, '__proto__')).toEqual([
        { family: '__proto__', url: 'https://fonts.example/css?family=__proto__' },
      ]);
      expect(fontCssApiRequests(descriptor, 'constructor')).toHaveLength(1);
      expect(fontCssApiRequests(descriptor, 'toString')).toEqual([]);
    }
  });

  it('bounds every completed request URL, including encoding expansion', () => {
    const template = 'https://fonts.example/css?family={family}:{query}';
    expect(limitOf(() => fontCssApi({ urlTemplate: template, query: 'q'.repeat(9000) })))
      .toMatchObject({ quota: 'url', metric: 'length' });
    // The URL parser percent-encodes a non-ASCII query (é → %C3%A9).
    expect(limitOf(() => fontCssApi({ urlTemplate: template, query: '\u00e9'.repeat(2100) })))
      .toMatchObject({ quota: 'url', limit: HARD_MAX_FONT_SOURCE_URL_LENGTH });
    expect(limitOf(() => fontCssApi({ urlTemplate: template, families: { [`F${'\u4e00'.repeat(1000)}`]: '400' } })))
      .toMatchObject({ quota: 'url' });
    // A document-supplied family is checked when its request is composed.
    const source = fontCssApi({ urlTemplate: template, query: '400' });
    expect(limitOf(() => fontCssApiRequests(source, '\u4e00'.repeat(1000)))).toMatchObject({ quota: 'url' });
    expect(fontCssApiRequests(source, 'A\ud800')[0].url).toBe('https://fonts.example/css?family=A%EF%BF%BD:400');
  });
});
