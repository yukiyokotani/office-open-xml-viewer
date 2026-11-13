import { afterEach, describe, expect, it, vi } from 'vitest';
import { OoxmlError } from '../errors/ooxml-error.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES,
} from '../worker/resource-policy.generated.js';
import { fetchFontStylesheet, parseFontFaceCss, tokenizeCss } from './css-source.js';
import { parseUnicodeRange } from './unicode-range.js';

const SHEET = 'https://cdn.internal.example/styles/office.css';

const face = (body: string) => `@font-face { ${body} }`;

function limitQuota(run: () => unknown): string | undefined {
  try {
    run();
  } catch (error) {
    if (error instanceof OoxmlError && error.code === 'font-source-limit') return error.violation?.quota;
    throw error;
  }
  return undefined;
}

describe('parseFontFaceCss', () => {
  it('reads machine-generated per-family CSS with every descriptor', () => {
    const css = `/* latin */
@font-face {
  font-family: 'Carlito';
  font-style: italic;
  font-weight: 700;
  font-stretch: 75% 125%;
  font-display: swap;
  src: url(https://fonts.gstatic.example/s/carlito/x.woff2) format('woff2');
  unicode-range: U+0000-00FF, U+0131, U+02??;
}`;
    expect(parseFontFaceCss(css, SHEET)).toEqual([{
      family: 'Carlito',
      url: 'https://fonts.gstatic.example/s/carlito/x.woff2',
      style: 'italic',
      weight: '700',
      stretch: '75% 125%',
      unicodeRange: 'U+0000-00FF, U+0131, U+02??',
    }]);
  });

  it.each([
    ['url(../files/a.woff2)', 'https://cdn.internal.example/files/a.woff2'],
    ["url('/files/a.woff2')", 'https://cdn.internal.example/files/a.woff2'],
    ['url( "a.woff2" )', 'https://cdn.internal.example/styles/a.woff2'],
    ['url(//assets.internal.example/a.woff2)', 'https://assets.internal.example/a.woff2'],
    ['url(a\\ b.woff2)', 'https://cdn.internal.example/styles/a%20b.woff2'],
  ])('resolves %s against the stylesheet URL', (src, expected) => {
    expect(parseFontFaceCss(face(`font-family: X; src: ${src};`), SHEET)[0]?.url).toBe(expected);
  });

  it('ignores invalid declarations for every descriptor, so an earlier valid one stays', () => {
    const parse = (body: string) => parseFontFaceCss(face(`font-family: X; src: url(a.woff2); ${body}`), SHEET)[0];
    expect(parse('font-weight: 400; font-weight: 1001;').weight).toBe('400');
    expect(parse('font-weight: 400; font-weight: 0;').weight).toBe('400');
    expect(parse('font-weight: 400; font-weight: 100 200 300;').weight).toBe('400');
    expect(parse('font-weight: 400; font-weight: heavy;').weight).toBe('400');
    expect(parse('font-weight: 400; font-weight: 100 900;').weight).toBe('100 900');
    expect(parse('font-weight: 400; font-weight: \\62old;').weight).toBe('bold');
    expect(parse('font-weight: 400; font-weight: auto;').weight).toBeUndefined();
    expect(parse('font-style: italic; font-style: oblique 100deg;').style).toBe('italic');
    expect(parse('font-style: italic; font-style: oblique 0.3turn;').style).toBe('italic');
    expect(parse('font-style: italic; font-style: slanted;').style).toBe('italic');
    expect(parse('font-style: italic; font-style: oblique -10deg 0.1rad;').style).toBe('oblique -10deg 0.1rad');
    expect(parse('font-stretch: 75%; font-stretch: -5%;').stretch).toBe('75%');
    expect(parse('font-stretch: 75%; font-stretch: narrow;').stretch).toBe('75%');
    expect(parse('font-stretch: 75%; font-stretch: condensed 125%;').stretch).toBe('condensed 125%');
    expect(parse('unicode-range: U+0-FF; unicode-range: U+110000;').unicodeRange).toBe('U+0-FF');
    expect(parse('unicode-range: U+0-FF; unicode-range: U+30-20;').unicodeRange).toBe('U+0-FF');
    expect(parse('unicode-range: U+0-FF; unicode-range: u+3000-30ff ,U+4E00;').unicodeRange).toBe('u+3000-30ff, U+4E00');
    // src: an invalid declaration keeps the earlier source...
    expect(parse('src: url(b.woff2) format(not-a-format);').url).toBe(`${SHEET.replace('office.css', '')}a.woff2`);
    expect(parse('src: url(b c.woff2);').url).toBe(`${SHEET.replace('office.css', '')}a.woff2`);
    expect(parse('src: url(b.woff2) tech(unknown-tech);').url).toBe(`${SHEET.replace('office.css', '')}a.woff2`);
    // ...while a valid one replaces it, even when it offers nothing usable here.
    expect(parse('src: local(X);')).toBeUndefined();
    expect(parse('src: url(b.woff2) tech(variations);')).toBeUndefined();
    expect(parse('src: url(b.woff2) format("woff2");').url).toBe(`${SHEET.replace('office.css', '')}b.woff2`);
    // font-family: an invalid value keeps the earlier family.
    expect(parseFontFaceCss(face('font-family: X; font-family: serif; src: url(a.woff2);'), SHEET)[0].family).toBe('X');
  });

  it('parses src by the fallback grammar as Chrome does (one format argument, unsupported formats dropped)', () => {
    const base = 'https://cdn.internal.example/styles/';
    const url = (body: string) => parseFontFaceCss(face(`font-family: X; src: url(a.woff2); ${body}`), SHEET)[0]?.url;
    // An invalid later declaration leaves the earlier source.
    expect(url('src: url(b.woff2) format("bogus");')).toBe(`${base}a.woff2`);
    expect(url('src: url(b.woff2) format("woff2", "woff");')).toBe(`${base}a.woff2`);
    expect(url('src: url(b.woff2) format(svg);')).toBe(`${base}a.woff2`);
    expect(url('src: url(b.woff2) format("embedded-opentype");')).toBe(`${base}a.woff2`);
    expect(url('src: url(b.woff2) format("truetype-aat");')).toBe(`${base}a.woff2`);
    expect(url('src: local(x) format(woff2);')).toBe(`${base}a.woff2`);
    expect(url('src: url(b.woff2) tech(bogus);')).toBe(`${base}a.woff2`);
    // Invalid entries are dropped inside a valid declaration.
    expect(url('src: url(b.woff2) format(bogus), url(c.woff2);')).toBe(`${base}c.woff2`);
    expect(url('src: url(b c), url(d.woff2);')).toBe(`${base}d.woff2`);
    expect(url('src: url(b.woff2) format("woff2-variations");')).toBe(`${base}b.woff2`);
    expect(url('src: url(b.woff2) format(WOFF2);')).toBe(`${base}b.woff2`);
    // Valid declarations with nothing the library registers drop the rule.
    expect(url('src: url();')).toBeUndefined();
    expect(url('src: url(b.ttc) format(collection);')).toBeUndefined();
  });

  it('tokenizes per CSS Syntax §4.3.1', () => {
    expect(tokenizeCss('#url( #-x #1 # @font-face @1 @ <!-- --> -->x')).toEqual([
      { type: 'hash', value: 'url', id: true }, { type: '(' }, { type: 'whitespace' },
      { type: 'hash', value: '-x', id: true }, { type: 'whitespace' },
      { type: 'hash', value: '1', id: false }, { type: 'whitespace' },
      { type: 'delim', value: '#' }, { type: 'whitespace' },
      { type: 'at-keyword', value: 'font-face' }, { type: 'whitespace' },
      { type: 'delim', value: '@' }, { type: 'number', value: 1 }, { type: 'whitespace' },
      { type: 'delim', value: '@' }, { type: 'whitespace' },
      { type: 'CDO' }, { type: 'whitespace' }, { type: 'CDC' }, { type: 'whitespace' },
      { type: 'CDC' }, { type: 'ident', value: 'x' },
    ]);
    expect(tokenizeCss('+.5 -1e3 1e 10% 5px -x - + . 1.')).toEqual([
      { type: 'number', value: 0.5 }, { type: 'whitespace' },
      { type: 'number', value: -1000 }, { type: 'whitespace' },
      { type: 'dimension', value: 1, unit: 'e' }, { type: 'whitespace' },
      { type: 'percentage', value: 10 }, { type: 'whitespace' },
      { type: 'dimension', value: 5, unit: 'px' }, { type: 'whitespace' },
      { type: 'ident', value: '-x' }, { type: 'whitespace' },
      { type: 'delim', value: '-' }, { type: 'whitespace' },
      { type: 'delim', value: '+' }, { type: 'whitespace' },
      { type: 'delim', value: '.' }, { type: 'whitespace' },
      { type: 'number', value: 1 }, { type: 'delim', value: '.' },
    ]);
    expect(tokenizeCss('u\\72l(a{b) url( "q" ) f(x) url(a b) "s\\\nt" "bad\n\\\nx \\ y /*c*/;:,[]{}')).toEqual([
      { type: 'url', value: 'a{b' }, { type: 'whitespace' },
      { type: 'function', value: 'url' }, { type: 'whitespace' }, { type: 'string', value: 'q' }, { type: 'whitespace' }, { type: ')' },
      { type: 'whitespace' }, { type: 'function', value: 'f' }, { type: 'ident', value: 'x' }, { type: ')' },
      { type: 'whitespace' }, { type: 'bad-url' }, { type: 'whitespace' },
      { type: 'string', value: 'st' }, { type: 'whitespace' }, { type: 'bad-string' }, { type: 'whitespace' },
      // A backslash before a newline is a delim; before a space it escapes it.
      { type: 'delim', value: '\\' }, { type: 'whitespace' }, { type: 'ident', value: 'x' }, { type: 'whitespace' },
      { type: 'ident', value: ' y' }, { type: 'whitespace' },
      { type: 'comment' }, { type: ';' }, { type: ':' }, { type: ',' }, { type: '[' }, { type: ']' }, { type: '{' }, { type: '}' },
    ]);
  });

  it('keeps a hash token before "(" structural, as Chrome does', () => {
    expect(parseFontFaceCss('@media print { --x:#url({) } @font-face{font-family:Y;src:url(y)} }', SHEET)).toEqual([]);
    expect(parseFontFaceCss('@media print { --x:#u } @font-face{font-family:Y;src:url(y)}', SHEET).map((entry) => entry.family))
      .toEqual(['Y']);
  });

  it('accepts math functions in the fallback grammar', () => {
    const parse = (body: string) => parseFontFaceCss(face(`font-family: X; src: url(a.woff2); ${body}`), SHEET)[0];
    expect(parse('font-weight: 400; font-weight: calc(700);').weight).toBe('calc(700)');
    expect(parse('font-weight: 400; font-weight: calc(100) max(800, 900);').weight).toBe('calc(100) max(800, 900)');
    expect(parse('font-style: normal; font-style: oblique calc(5deg * 2);').style).toBe('oblique calc(5deg * 2)');
    expect(parse('font-stretch: 75%; font-stretch: clamp(50%, 80%, 100%);').stretch).toBe('clamp(50%, 80%, 100%)');
    expect(parse('font-weight: 400; font-weight: var(--w);').weight).toBe('400');
  });

  it('lets the host FontFace parser validate descriptors, walking back from the last declaration', () => {
    const G = globalThis as { FontFace?: unknown };
    const original = G.FontFace;
    const probes: { descriptors: FontFaceDescriptors; loads: number; observed: boolean }[] = [];
    // Stands in for the host CSS parser: rejects by status (spec) or by
    // throwing (older engines); accepts anything else, calc() included.
    class HostFontFace {
      status: FontFaceLoadStatus = 'unloaded';
      readonly loaded: Promise<unknown>;
      private readonly probe: (typeof probes)[number];
      constructor(_family: string, _source: string, descriptors: FontFaceDescriptors = {}) {
        if (descriptors.weight === 'heavy') throw new DOMException('bad', 'SyntaxError');
        this.probe = { descriptors, loads: 0, observed: false };
        probes.push(this.probe);
        const invalid = descriptors.weight === '1001' || descriptors.style === 'slanted';
        if (invalid) this.status = 'error';
        const loaded = invalid ? Promise.reject(new DOMException('bad', 'SyntaxError')) : new Promise(() => undefined);
        const probe = this.probe;
        this.loaded = { catch: (handler: (error: unknown) => unknown) => { probe.observed = true; return loaded.catch(handler); } } as Promise<unknown>;
      }
      load() { this.probe.loads++; return this.loaded; }
    }
    G.FontFace = HostFontFace;
    try {
      const [declaration] = parseFontFaceCss(face(`font-family: X; src: url(a.woff2);
        font-weight: 300; font-weight: calc(700); font-weight: heavy; font-weight: 1001;
        font-style: italic; font-style: slanted; font-stretch: 75%;`), SHEET);
      expect(declaration).toMatchObject({ weight: 'calc(700)', style: 'italic', stretch: '75%' });
      // Walked back 1001 → heavy (thrown) → calc(700); memoized; never loaded.
      expect(probes.map((probe) => probe.descriptors.weight).filter(Boolean)).toEqual(['1001', 'calc(700)']);
      expect(probes.every((probe) => probe.loads === 0 && probe.observed)).toBe(true);
    } finally {
      G.FontFace = original;
    }
  });

  it('finds structure with the same token rules as values (escaped url, numbers)', () => {
    const urls = (css: string) => parseFontFaceCss(css, SHEET).map((entry) => entry.url);
    // An escaped url( is still a url token: its braces are not structure.
    expect(urls('@font-face { font-family: X; src: u\\72l(x{y.woff2) }')).toEqual(['https://cdn.internal.example/styles/x%7By.woff2']);
    expect(urls('@font-face { font-family: X; src: URL(x}y.woff2) }')).toEqual(['https://cdn.internal.example/styles/x%7Dy.woff2']);
    // A quoted url("…") is a function; its argument string hides the braces.
    expect(urls('@font-face { font-family: X; src: url("x}y.woff2") }')).toEqual(['https://cdn.internal.example/styles/x%7Dy.woff2']);
    // "1url(" is a dimension followed by "(": not a url token.
    expect(urls('@font-face { font-family: X; src: url(a.woff2); --x: 1url(}) } @font-face { font-family: Y; src: url(b.woff2) }'))
      .toEqual(['https://cdn.internal.example/styles/a.woff2', 'https://cdn.internal.example/styles/b.woff2']);
  });

  it('consumes string tokens per CSS Syntax §4.3.5', () => {
    const family = (value: string) => parseFontFaceCss(face(`font-family: ${value}; src: url(a.woff2);`), SHEET)[0]?.family;
    // A hex escape takes one following white space, including a newline (CR LF is one).
    expect(family('"A\\49\nnter"')).toBe('AInter');
    expect(family('"A\\49\r\nnter"')).toBe('AInter');
    expect(family('"A\\49\fnter"')).toBe('AInter');
    // An escaped newline continues the string.
    expect(family('"A\\\nB"')).toBe('AB');
    expect(family('"A\\\r\nB"')).toBe('AB');
    expect(family("'A\\'B'")).toBe("A'B");
    expect(family('"A\\"B"')).toBe('A"B');
    // An unescaped newline is a bad string: the declaration is invalid.
    // It ends before the newline; the quote after B opens another (bad) string
    // ending at the next newline, so the block structure resumes after it.
    expect(parseFontFaceCss('@font-face { font-family: X; font-family: "A\nB"\n; src: url(a.woff2); }', SHEET)[0]?.family)
      .toBe('X');
    // Without that second newline the string swallows the rest of the block.
    expect(parseFontFaceCss('@font-face { font-family: X; font-family: "A\nB"; src: url(a.woff2); }', SHEET)).toEqual([]);
    // Strings feed url() too.
    expect(parseFontFaceCss(face('font-family: X; src: url("a\\2e woff2");'), SHEET)[0].url)
      .toBe('https://cdn.internal.example/styles/a.woff2');
  });

  it('decodes escapes in at-keywords and descriptor names', () => {
    expect(parseFontFaceCss('@font\\-face { font\\-family: X; s\\72 c: url(a.woff2); }', SHEET)
      .map((entry) => entry.family)).toEqual(['X']);
    expect(parseFontFaceCss('@\\66ont-face { font-family: Y; src: url(a.woff2); }', SHEET)
      .map((entry) => entry.family)).toEqual(['Y']);
    expect(parseFontFaceCss(face('font-family: X; src: u\\72l(a.woff2);'), SHEET)[0].url)
      .toBe('https://cdn.internal.example/styles/a.woff2');
  });

  it('consumes unquoted url tokens per CSS Syntax §4.3.6', () => {
    const url = (src: string) => parseFontFaceCss(face(`font-family: X; src: ${src};`), SHEET)[0]?.url;
    const base = 'https://cdn.internal.example/styles/';
    // The white space ending a hex escape belongs to the escape.
    expect(url('url(font\\2e woff2)')).toBe(`${base}font.woff2`);
    expect(url('url(font\\2e\r\nwoff2)')).toBe(`${base}font.woff2`);
    expect(url('url(  font.woff2\t)')).toBe(`${base}font.woff2`);
    expect(url('url(a\\ b.woff2)')).toBe(`${base}a%20b.woff2`);
    expect(url('url(a\\)b.woff2)')).toBe(`${base}a)b.woff2`);
    expect(url('url(a\\(b.woff2)')).toBe(`${base}a(b.woff2`);
    expect(url('url(a\\"b\\\'c.woff2)')).toBe(`${base}a%22b'c.woff2`);
    // Bad-url tokens drop the entry; the next entry still parses.
    for (const bad of ['url(a b.woff2)', 'url(a"b.woff2)', "url(a'b.woff2)", 'url(a(b.woff2)',
      'url(a\\\nb.woff2)', 'url(a\u0001b.woff2)']) {
      expect(url(bad), bad).toBeUndefined();
      expect(url(`${bad}, url(ok.woff2)`), bad).toBe(`${base}ok.woff2`);
    }
  });

  it('accepts only url() sources: local() skipped, first supported format chosen', () => {
    const [declaration] = parseFontFaceCss(face(`font-family: X; src:
      local("X Regular"), url(x.eot) format("embedded-opentype"), url(x.var.woff2) format(woff2) tech(variations),
      url(x.woff) format("woff"), url(x.ttf);`), SHEET);
    expect(declaration.url).toBe('https://cdn.internal.example/styles/x.woff');
    // local()-only rules register nothing.
    expect(parseFontFaceCss(face('font-family: X; src: local(X);'), SHEET)).toEqual([]);
  });

  it('keeps data: URLs verbatim (including unquoted ";base64,") and accepts blob:', () => {
    const data = 'data:font/woff2;base64,d09GMgABAAAAAA==';
    expect(parseFontFaceCss(face(`font-family: X; src: url(${data}) format("woff2");`), SHEET)[0].url).toBe(data);
    const blob = 'blob:https://app.example/1f0e';
    expect(parseFontFaceCss(face(`font-family: X; src: url("${blob}");`), SHEET)[0].url).toBe(blob);
  });

  it.each([
    'javascript:alert(1)',
    'file:///etc/fonts/x.ttf',
    'ftp://files.example/x.ttf',
    'chrome://x/y.woff',
  ])('drops a rule whose source resolves to %s', (url) => {
    expect(parseFontFaceCss(face(`font-family: X; src: url("${url}");`), SHEET)).toEqual([]);
  });

  it('interprets only top-level @font-face rules', () => {
    const css = `
@charset "utf-8";
@import url(other.css);
body { font-family: "A"; src: url(no.woff2); }
@media print { @font-face { font-family: Nested; src: url(nested.woff2); } }
@supports (display: grid) { .a { color: red } }
<!-- @font-face { font-family: Cdo; src: url(cdo.woff2); } -->
@FONT-FACE { font-family: Upper; src: url(upper.woff2); }`;
    expect(parseFontFaceCss(css, SHEET).map((entry) => entry.family)).toEqual(['Cdo', 'Upper']);
  });

  it('is not misled by braces, semicolons and comments inside strings and url tokens', () => {
    const css = `
/* @font-face { font-family: Commented; src: url(c.woff2); } */
@font-face {
  font-family: "Brace } ; Name";
  src: /* } */ url("a;b}.woff2");
}
@font-face { font-family: Second; src: url(x/*not-a-comment*/.woff2); }`;
    expect(parseFontFaceCss(css, SHEET)).toEqual([
      { family: 'Brace } ; Name', url: 'https://cdn.internal.example/styles/a;b%7D.woff2' },
      { family: 'Second', url: 'https://cdn.internal.example/styles/x/*not-a-comment*/.woff2' },
    ]);
  });

  it('decodes family escapes and rejects reserved or malformed names', () => {
    const family = (value: string) => parseFontFaceCss(face(`font-family: ${value}; src: url(a.woff2);`), SHEET)[0]?.family;
    expect(family('"Noto \\53 ans"')).toBe('Noto Sans');
    expect(family('Noto   Sans  JP')).toBe('Noto Sans JP');
    expect(family('serif')).toBeUndefined();
    expect(family('inherit')).toBeUndefined();
    expect(family('1Font')).toBeUndefined();
    expect(family('"unterminated')).toBeUndefined();
  });

  it('consumes escapes per CSS Syntax §4.3.7', () => {
    const family = (value: string) => parseFontFaceCss(face(`font-family: ${value}; src: url(a.woff2);`), SHEET)[0]?.family;
    // The white space ending a hex escape belongs to the escape.
    expect(family('\\49 nter')).toBe('Inter');
    expect(family('\\49\r\nnter')).toBe('Inter');
    expect(family('\\49  nter')).toBe('I nter');
    expect(family('\\000049nter')).toBe('Inter');
    expect(family('A\\0 B')).toBe('A\uFFFDB');
    expect(family('A\\D800 B')).toBe('A\uFFFDB');
    expect(family('A\\110000 B')).toBe('A\uFFFDB');
    expect(family('A\\1F600 B')).toBe('A\u{1F600}B');
    expect(family('A\\\u{1F600}B')).toBe('A\u{1F600}B');
    expect(family('A\\ B')).toBe('A B');
    expect(family('"A\\49 nter"')).toBe('AInter');
    // A string line continuation is removed.
    expect(family('"Line\\\nCont"')).toBe('LineCont');
    // A backslash at end of input is U+FFFD in an identifier.
    expect(parseFontFaceCss('@font-face { src: url(a.woff2); font-family: X\\}', SHEET)).toEqual([]);
    expect(parseFontFaceCss('@font-face { src: url(a.woff2); font-family: X\\', SHEET)).toEqual([]);
  });

  it('drops invalid, !important and block-valued descriptors as CSS does', () => {
    const [declaration] = parseFontFaceCss(face(`
      font-family: X; src: url(a.woff2);
      font-weight: 400; font-weight: 700 !important;
      font-style: {italic};
      unicode-range: U+0000-00FF; unicode-range: U+0041-0030;`), SHEET);
    // An invalid later declaration leaves the earlier valid one in place.
    expect(declaration).toEqual({
      family: 'X', url: 'https://cdn.internal.example/styles/a.woff2', weight: '400', unicodeRange: 'U+0000-00FF',
    });
    const [noRange] = parseFontFaceCss(face('font-family: X; src: url(a.woff2); unicode-range: U+110000;'), SHEET);
    expect(noRange.unicodeRange).toBeUndefined();
  });

  it('drops an @font-face block left unterminated at end of input', () => {
    expect(parseFontFaceCss('@font-face { font-family: X; src: url(a.woff2);', SHEET)).toEqual([]);
    expect(parseFontFaceCss('@font-face { font-family: X; src: url(a.woff2); } @font-face { font-family: Y', SHEET)
      .map((entry) => entry.family)).toEqual(['X']);
  });

  it('scans deeply nested and unbalanced input linearly without recursion', () => {
    const depth = 200_000;
    const css = `.a ${'{'.repeat(depth)}${'}'.repeat(depth)} ${'('.repeat(depth)} ;
      ${face('font-family: After; src: url(a.woff2);')}`;
    const started = performance.now();
    // The open parentheses of the hostile prelude swallow the rest of the sheet.
    expect(parseFontFaceCss(css, SHEET)).toEqual([]);
    expect(parseFontFaceCss(`${'['.repeat(depth)}`, SHEET)).toEqual([]);
    expect(performance.now() - started).toBeLessThan(2000);
  });

  it('rejects a stylesheet beyond the byte ceiling and rules beyond the declaration ceiling', () => {
    expect(limitQuota(() => parseFontFaceCss(' '.repeat(HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES + 1), SHEET)))
      .toBe('stylesheet');
    // Counted as UTF-8 bytes, not UTF-16 units.
    expect(limitQuota(() => parseFontFaceCss('あ'.repeat(HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES / 3 + 1), SHEET)))
      .toBe('stylesheet');
    const rules = face('font-family: X; src: url(a.woff2);').repeat(HARD_MAX_FONT_SOURCE_DECLARATIONS + 1);
    expect(limitQuota(() => parseFontFaceCss(rules, SHEET))).toBe('declarations');
    expect(parseFontFaceCss(face('font-family: X; src: url(a.woff2);').repeat(HARD_MAX_FONT_SOURCE_DECLARATIONS), SHEET))
      .toHaveLength(HARD_MAX_FONT_SOURCE_DECLARATIONS);
  });

  it('rejects an overlong face URL instead of dropping it', () => {
    expect(limitQuota(() => parseFontFaceCss(face(`font-family: X; src: url(${'a'.repeat(9000)}.woff2);`), SHEET)))
      .toBe('url');
  });
});

describe('parseUnicodeRange', () => {
  it('accepts spans and wildcards and rejects anything else', () => {
    expect(parseUnicodeRange('U+0-7F, u+30??')).toEqual([[0, 0x7f], [0x3000, 0x30ff]]);
    expect(parseUnicodeRange('U+7F-0')).toBeNull();
    expect(parseUnicodeRange('U+0-110000')).toBeNull();
    expect(parseUnicodeRange('U+0000 - 00FF')).toBeNull();
    expect(parseUnicodeRange('U+???????')).toBeNull();
    expect(parseUnicodeRange('')).toBeNull();
  });
});

describe('fetchFontStylesheet', () => {
  const G = globalThis as { fetch?: unknown };
  const originalFetch = G.fetch;
  afterEach(() => { G.fetch = originalFetch; });

  function streamed(chunks: Uint8Array[], init: { url?: string; contentLength?: number } = {}) {
    let cancelled = false;
    let bodyCancelled = false;
    let index = 0;
    return {
      response: {
        ok: true, status: 200, url: init.url ?? '',
        headers: { get: (name: string) => name === 'content-length' && init.contentLength !== undefined ? String(init.contentLength) : null },
        body: {
          getReader: () => ({
            read: async () => index < chunks.length ? { done: false, value: chunks[index++] } : { done: true },
            cancel: async () => { cancelled = true; },
          }),
          cancel: async () => { bodyCancelled = true; },
        },
        text: async () => { throw new Error('body must be streamed'); },
      },
      cancelled: () => cancelled,
      bodyCancelled: () => bodyCancelled,
      reads: () => index,
    };
  }

  it('decodes a streamed body and resolves relative URLs against the final URL', async () => {
    const css = face("font-family: 'Interé'; src: url(../fonts/i.woff2);");
    const bytes = new TextEncoder().encode(css);
    const { response } = streamed([bytes.slice(0, 20), bytes.slice(20)], { url: 'https://cdn2.example/v2/a.css' });
    const fetchMock = vi.fn(async () => response);
    G.fetch = fetchMock;
    const result = await fetchFontStylesheet('https://fonts.example/a.css', undefined, new AbortController().signal);
    expect(result).toEqual({
      ok: true,
      finalUrl: 'https://cdn2.example/v2/a.css',
      declarations: [{ family: 'Interé', url: 'https://cdn2.example/fonts/i.woff2' }],
    });
    expect(fetchMock).toHaveBeenCalledWith('https://fonts.example/a.css',
      expect.objectContaining({ credentials: 'same-origin' }));
  });

  it('stops reading once the byte ceiling is exceeded, failing the stylesheet', async () => {
    const chunk = new Uint8Array(256 * 1024).fill(0x20);
    const stream = streamed([chunk, chunk, chunk, chunk, chunk, chunk, chunk, chunk]);
    let requestSignal: AbortSignal | undefined;
    G.fetch = async (_url: string, init: RequestInit) => { requestSignal = init.signal ?? undefined; return stream.response; };
    const caller = new AbortController();
    expect(await fetchFontStylesheet('https://fonts.example/a.css', 'omit', caller.signal)).toEqual({
      ok: false, reason: `Font source limit exceeded (stylesheet): bytes ${5 * 256 * 1024} > ${HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES}`,
    });
    expect(stream.reads()).toBe(5);
    expect(stream.cancelled()).toBe(true);
    // The request is aborted although the caller's signal stays live.
    expect(requestSignal?.aborted).toBe(true);
    expect(caller.signal.aborted).toBe(false);
  });

  it('fails a declared Content-Length beyond the ceiling before reading', async () => {
    const stream = streamed([], { contentLength: HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES + 1 });
    let requestSignal: AbortSignal | undefined;
    G.fetch = async (_url: string, init: RequestInit) => { requestSignal = init.signal ?? undefined; return stream.response; };
    expect(await fetchFontStylesheet('https://fonts.example/a.css', undefined, new AbortController().signal))
      .toMatchObject({ ok: false, reason: expect.stringContaining('Font source limit exceeded (stylesheet)') });
    expect(stream.reads()).toBe(0);
    expect(stream.bodyCancelled()).toBe(true);
    expect(requestSignal?.aborted).toBe(true);
  });

  it('reports transport failures and empty stylesheets as failures, not errors', async () => {
    const signal = new AbortController().signal;
    G.fetch = async () => ({ ok: false, status: 404, text: async () => '' });
    expect(await fetchFontStylesheet('https://fonts.example/a.css', undefined, signal)).toEqual({ ok: false, reason: 'HTTP 404' });
    G.fetch = async () => { throw new TypeError('Failed to fetch'); };
    expect(await fetchFontStylesheet('https://fonts.example/a.css', undefined, signal)).toEqual({ ok: false, reason: 'Failed to fetch' });
    G.fetch = async () => ({ ok: true, status: 200, text: async () => 'body { color: red }' });
    expect(await fetchFontStylesheet('https://fonts.example/a.css', undefined, signal))
      .toEqual({ ok: false, reason: 'no usable @font-face rule' });
  });
});
