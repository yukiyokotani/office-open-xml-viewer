/**
 * CSS font sources: `font-css-api` (one stylesheet per served family) and
 * `font-css` (one stylesheet describing many families), plus the
 * hardened `@font-face` parser and bounded stylesheet fetch they share.
 *
 * Stylesheets are untrusted input (a third-party CDN, a mirror, an internal
 * service). The parser is a single linear scan over CSS Syntax Level 3
 * structure (comments, strings, url tokens and `()`/`[]`/`{}` blocks; no
 * recursion, so nesting depth costs nothing) and interprets only the
 * `font-family`, `src`, `font-weight`, `font-style`, `font-stretch` and
 * `unicode-range` descriptors of top-level `@font-face` rules. Everything
 * else (other at-rules, qualified rules, `@font-face` nested in conditional
 * rules, `!important` or block-valued descriptors) is skipped. Input is
 * preprocessed (§3.3) and tokens follow §4.3: string and url tokens, escapes
 * (decoded in identifiers, at-keywords, descriptor names, units, strings and
 * urls), bad strings and bad urls; structure (blocks, declaration and
 * entry separators) is found with the same token rules, so an escaped
 * `u\72l(` hides its braces exactly as a literal `url(` does. The last valid
 * declaration of each descriptor wins (an invalid one never overrides an
 * earlier valid one). Validity of `font-weight`, `font-style`,
 * `font-stretch` and `unicode-range` is decided by the host's own CSS parser
 * through an unregistered, never-loaded probe `FontFace` wherever the realm
 * has one (window or worker), so the full grammar (math functions included)
 * matches the browser that will paint. Without `FontFace` (Node) a fallback
 * CSS Fonts 4 §4 grammar applies; it accepts math functions unevaluated and
 * treats `auto` as the initial value, where a browser that does not support
 * `auto` ignores the declaration. `src` validity is the host's as well,
 * probed per entry (`new FontFace(family, entry)` parses without fetching);
 * the url to register is then taken from the last valid declaration with the
 * library's own tokenizer, and the fallback grammar follows Chrome (one
 * `format()` argument; svg, embedded-opentype and unknown formats or techs
 * invalidate their entry). `font-family` uses the library's own parse. All
 * structure comes from one CSS Syntax §4.3.1 tokenizer ({@link consumeToken}):
 * hash, at-keyword, numeric, ident-like (url, function), string, CDO/CDC and
 * delim tokens. Library policy on top of CSS Fonts Level 4:
 *
 * - only `url()` sources are accepted: `local()` would let a stylesheet probe
 *   or paint host fonts, which is the host-fallback path's job. The first
 *   `url()` whose `format()` is absent or one of woff2 / woff / truetype /
 *   opentype is used; entries with `tech()` are skipped (support cannot be
 *   evaluated here);
 * - relative URLs resolve against the stylesheet's final URL after redirects
 *   (#1497); results must be http(s) / blob: / data:, and `data:` URLs are
 *   charged their decoded size against the face ceiling (PR 1 policy);
 * - an invalid `unicode-range` is dropped as CSS drops an invalid declaration,
 *   so the face covers every scalar and is loaded whenever its family is;
 * - an `@font-face` block left unterminated at end of input is dropped rather
 *   than closed (a truncated stylesheet must not register a partial rule);
 * - the stylesheet is bounded by `HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES` while
 *   it is read and its rules by `HARD_MAX_FONT_SOURCE_DECLARATIONS`; both
 *   reject with `OoxmlError('font-source-limit')`. The body is decoded as
 *   UTF-8 (`@charset` and `Content-Type` charset are not interpreted).
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES,
} from '../worker/resource-policy.generated.js';
import { claimFromIndex, indexFontDeclarations } from './files-source.js';
import type { FamilyResolution, FontSourceFailure, FontStylesheet, SourceCache } from './source-cache.js';
import {
  fontCssApiRequests,
  validateFontFaceDeclaration,
  type FontCssApiSource,
  type FontCssSource,
  type FontFaceDeclaration,
  type FontSourceCredentials,
} from './source.js';
import { fontFamilyKey, scopedFamilyKey, type FontSubstituteScripts } from './substitutions.js';

// ---- unicode-range ------------------------------------------------------------------

/**
 * Parse a CSS `unicode-range` value (CSS Fonts 4 §4.5: comma-separated
 * `U+hex`, `U+hex-hex` or `U+hex?` wildcard). Returns null unless every
 * component is recognised and within U+0–U+10FFFF with start ≤ end: a range
 * is an exclusion hint, never a support certificate, so an unparseable value
 * means "covers everything".
 */
export function parseUnicodeRange(value: string | undefined): readonly (readonly [number, number])[] | null {
  if (!value?.trim()) return null;
  const ranges: [number, number][] = [];
  for (const component of value.split(',')) {
    const part = component.trim();
    let lo: number, hi: number;
    const span = /^U\+([0-9A-F]{1,6})(?:-([0-9A-F]{1,6}))?$/i.exec(part);
    const wildcard = /^U\+([0-9A-F]{0,5})(\?{1,6})$/i.exec(part);
    if (span) {
      lo = parseInt(span[1], 16);
      hi = parseInt(span[2] ?? span[1], 16);
    } else if (wildcard && wildcard[1].length + wildcard[2].length <= 6) {
      lo = parseInt(wildcard[1] + '0'.repeat(wildcard[2].length), 16);
      hi = parseInt(wildcard[1] + 'F'.repeat(wildcard[2].length), 16);
    } else {
      return null;
    }
    if (lo > hi || hi > 0x10ffff) return null;
    ranges.push([lo, hi]);
  }
  return ranges;
}

// ---- CSS scanning -----------------------------------------------------------------

const isWhitespace = (unit: number) => unit === 0x20 || unit === 0x09 || unit === 0x0a || unit === 0x0c || unit === 0x0d;
const isNameUnit = (unit: number) =>
  (unit >= 0x30 && unit <= 0x39) || (unit >= 0x41 && unit <= 0x5a) || (unit >= 0x61 && unit <= 0x7a)
  || unit === 0x2d || unit === 0x5f || unit >= 0x80;

/** Index after a comment starting at `at` (`/*`), or the end of input. */
function skipComment(css: string, at: number): number {
  const end = css.indexOf('*/', at + 2);
  return end < 0 ? css.length : end + 2;
}

/** One CSS token (CSS Syntax §4.3.1). Comments are reported as tokens so
 * callers can keep source positions; nothing else is folded away. */
export type CssToken =
  | { readonly type: 'whitespace' | 'comment' | 'CDO' | 'CDC' | 'bad-string' | 'bad-url' }
  | { readonly type: 'string' | 'ident' | 'function' | 'at-keyword' | 'url' | 'delim'; readonly value: string }
  | { readonly type: 'hash'; readonly value: string; readonly id: boolean }
  | { readonly type: 'number' | 'percentage'; readonly value: number }
  | { readonly type: 'dimension'; readonly value: number; readonly unit: string }
  | { readonly type: '(' | ')' | '[' | ']' | '{' | '}' | ',' | ':' | ';' };

const PUNCTUATION = '()[]{},:;';
const isDigit = (char: string | undefined) => char !== undefined && char >= '0' && char <= '9';

/** "Check if three code points would start a number" (§4.3.10). */
function startsNumber(text: string, at: number): boolean {
  const char = text[at];
  if (isDigit(char)) return true;
  if (char === '.') return isDigit(text[at + 1]);
  if (char === '+' || char === '-') return isDigit(text[at + 1]) || (text[at + 1] === '.' && isDigit(text[at + 2]));
  return false;
}

/** "Consume an ident-like token" (§4.3.4): an ident, a function token, or a
 * url token when the decoded name is `url` and the argument is unquoted. */
function consumeIdentLike(text: string, at: number): { token: CssToken; end: number } {
  const ident = consumeIdent(text, at);
  if (text[ident.end] !== '(') return { token: { type: 'ident', value: ident.value }, end: ident.end };
  if (ident.value.toLowerCase() === 'url') {
    let first = ident.end + 1;
    while (first < text.length && isWhitespace(text.charCodeAt(first))) first++;
    if (text[first] !== '"' && text[first] !== "'") {
      const url = consumeUrlToken(text, ident.end + 1);
      return { token: url.value === null ? { type: 'bad-url' } : { type: 'url', value: url.value }, end: url.end };
    }
  }
  return { token: { type: 'function', value: ident.value }, end: ident.end + 1 };
}

/** "Consume a token" (§4.3.1) at `at` (input already preprocessed, §3.3). */
export function consumeToken(text: string, at: number): { token: CssToken; end: number } {
  const char = text[at];
  if (char === '/' && text[at + 1] === '*') return { token: { type: 'comment' }, end: skipComment(text, at) };
  if (isWhitespace(text.charCodeAt(at))) {
    let end = at + 1;
    while (end < text.length && isWhitespace(text.charCodeAt(end))) end++;
    return { token: { type: 'whitespace' }, end };
  }
  if (char === '"' || char === "'") {
    const string = consumeStringToken(text, at);
    return { token: string.bad ? { type: 'bad-string' } : { type: 'string', value: string.value }, end: string.end };
  }
  if (startsNumber(text, at)) {
    NUMBER.lastIndex = at;
    const match = NUMBER.exec(text) as RegExpExecArray;
    const value = Number(match[0]);
    const end = at + match[0].length;
    if (startsIdentifier(text, end)) {
      const unit = consumeIdent(text, end);
      return { token: { type: 'dimension', value, unit: unit.value.toLowerCase() }, end: unit.end };
    }
    if (text[end] === '%') return { token: { type: 'percentage', value }, end: end + 1 };
    return { token: { type: 'number', value }, end };
  }
  if (char === '#' && (isNameUnit(text.charCodeAt(at + 1)) || startsValidEscape(text, at + 1))) {
    const id = startsIdentifier(text, at + 1);
    const name = consumeIdent(text, at + 1);
    return { token: { type: 'hash', value: name.value, id }, end: name.end };
  }
  if (char === '-' && text.startsWith('-->', at)) return { token: { type: 'CDC' }, end: at + 3 };
  if (char === '<' && text.startsWith('<!--', at)) return { token: { type: 'CDO' }, end: at + 4 };
  if (char === '@' && startsIdentifier(text, at + 1)) {
    const name = consumeIdent(text, at + 1);
    return { token: { type: 'at-keyword', value: name.value }, end: name.end };
  }
  if (startsIdentifier(text, at)) return consumeIdentLike(text, at);
  if (PUNCTUATION.includes(char)) return { token: { type: char as '(' }, end: at + 1 };
  const code = text.codePointAt(at) as number;
  return { token: { type: 'delim', value: String.fromCodePoint(code) }, end: at + (code > 0xffff ? 2 : 1) };
}

/** The whole token stream of `text` (preprocessed per §3.3 first). */
export function tokenizeCss(text: string): CssToken[] {
  const input = preprocessCss(text);
  const tokens: CssToken[] = [];
  for (let index = 0; index < input.length;) {
    const { token, end } = consumeToken(input, index);
    tokens.push(token);
    index = end;
  }
  return tokens;
}

/** CSS Syntax §3.3 input preprocessing: CR LF, CR and FF become LF; NUL and
 * lone surrogates become U+FFFD. */
function preprocessCss(css: string): string {
  return css.replace(/\r\n?|\f/g, '\n')
    .replace(/\0|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, REPLACEMENT);
}

/**
 * Scan tokens from `at` until a punctuation token in `stops` appears outside
 * any block. Returns its index, or the end of input. `(`, `[`, `{` and
 * function tokens open blocks; only the matching closer ends one (a
 * mismatched closer is an ordinary token, CSS Syntax §5.4.8).
 */
function scanUntil(css: string, at: number, stops: string): number {
  const expected: string[] = [];
  let index = at;
  while (index < css.length) {
    const { token, end } = consumeToken(css, index);
    const type = token.type;
    if (expected.length === 0 && type.length === 1 && stops.includes(type)) return index;
    if (type === '(' || type === 'function') expected.push(')');
    else if (type === '[') expected.push(']');
    else if (type === '{') expected.push('}');
    else if (expected.length && type === expected[expected.length - 1]) expected.pop();
    index = end;
  }
  return css.length;
}

/** Index of the `}` closing the block opened at `at`, or -1 at end of input. */
function blockEnd(css: string, at: number): number {
  const end = scanUntil(css, at + 1, '}');
  return end < css.length ? end : -1;
}

/** `text` with every comment token replaced by a space. */
function stripComments(text: string): string {
  let out = '';
  for (let index = 0; index < text.length;) {
    const { token, end } = consumeToken(text, index);
    out += token.type === 'comment' ? ' ' : text.slice(index, end);
    index = end;
  }
  return out;
}

/** Split `text` at top-level occurrences of `separator`. */
function splitTopLevel(text: string, separator: string): string[] {
  const parts: string[] = [];
  let start = 0;
  while (start <= text.length) {
    const end = scanUntil(text, start, separator);
    parts.push(text.slice(start, end));
    start = end + 1;
  }
  return parts;
}

const isNewline = (char: string | undefined) => char === '\n' || char === '\r' || char === '\f';
const isHexDigit = (unit: number) =>
  (unit >= 0x30 && unit <= 0x39) || (unit >= 0x41 && unit <= 0x46) || (unit >= 0x61 && unit <= 0x66);
const REPLACEMENT = String.fromCharCode(0xfffd);

/**
 * Consume an escaped code point (CSS Syntax §4.3.7); `at` is the backslash,
 * already known to start a valid escape (not followed by a newline). One to
 * six hex digits are followed by one optional white space, which belongs to
 * the escape (CR LF counts as one, as input preprocessing makes it one LF);
 * zero, a surrogate or a value beyond U+10FFFF is U+FFFD. A backslash at end
 * of input is U+FFFD; otherwise the next code point stands for itself (NUL is
 * U+FFFD after preprocessing).
 */
function consumeEscape(text: string, at: number): { value: string; end: number } {
  let index = at + 1;
  if (index >= text.length) return { value: REPLACEMENT, end: index };
  let digits = 0;
  while (digits < 6 && index + digits < text.length && isHexDigit(text.charCodeAt(index + digits))) digits++;
  if (digits === 0) {
    const code = text.codePointAt(index) as number;
    return { value: code === 0 ? REPLACEMENT : String.fromCodePoint(code), end: index + (code > 0xffff ? 2 : 1) };
  }
  const code = parseInt(text.slice(index, index + digits), 16);
  index += digits;
  if (text.startsWith('\r\n', index)) index += 2;
  else if (index < text.length && isWhitespace(text.charCodeAt(index))) index++;
  const value = code === 0 || (code >= 0xd800 && code <= 0xdfff) || code > 0x10ffff
    ? REPLACEMENT : String.fromCodePoint(code);
  return { value, end: index };
}

/**
 * Consume a string token (CSS Syntax §4.3.5) starting at the quote at `at`.
 * The string ends at its matching quote or at end of input (a parse error,
 * still a string). An unescaped newline makes a bad string, ending before the
 * newline. A backslash before a newline continues the string (both removed),
 * a backslash at end of input contributes nothing, and any other backslash
 * starts an escape ({@link consumeEscape}: a hex escape takes one following
 * white space).
 */
function consumeStringToken(text: string, at: number): { value: string; end: number; bad: boolean } {
  const quote = text[at];
  let value = '';
  let index = at + 1;
  while (index < text.length) {
    const char = text[index];
    if (char === quote) return { value, end: index + 1, bad: false };
    if (isNewline(char)) return { value, end: index, bad: true };
    if (char !== '\\') { value += char; index++; continue; }
    const next = text[index + 1];
    if (next === undefined) { index++; continue; }
    if (isNewline(next)) { index += text.startsWith('\r\n', index + 1) ? 3 : 2; continue; }
    const escape = consumeEscape(text, index);
    value += escape.value;
    index = escape.end;
  }
  return { value, end: index, bad: false };
}

/** The value of `text` when it is exactly one (non-bad) string token, else null. */
function cssString(text: string): string | null {
  if (text[0] !== '"' && text[0] !== "'") return null;
  const token = consumeStringToken(text, 0);
  return token.bad || token.end !== text.length ? null : token.value;
}

/** Reserved words a `<family-name>` cannot be spelled as unquoted (CSS-wide
 * keywords, `default`, and the generic families). */
const RESERVED_FAMILY_IDENTS = new Set([
  'inherit', 'initial', 'unset', 'revert', 'revert-layer', 'default',
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'math', 'emoji', 'fangsong',
]);

const isNameStart = (char: string | undefined) =>
  char !== undefined && (/[A-Za-z_]/.test(char) || char.charCodeAt(0) >= 0x80);
const startsValidEscape = (text: string, at: number) => text[at] === '\\' && !isNewline(text[at + 1]);

/** Whether an identifier starts at `at` (CSS Syntax §4.3.9). */
function startsIdentifier(text: string, at: number): boolean {
  if (text[at] === '-') {
    return text[at + 1] === '-' || isNameStart(text[at + 1]) || startsValidEscape(text, at + 1);
  }
  return isNameStart(text[at]) || startsValidEscape(text, at);
}

/** Consume an identifier sequence (CSS Syntax §4.3.11) at `at`, decoding escapes. */
function consumeIdent(text: string, at: number): { value: string; end: number } {
  let value = '';
  let index = at;
  while (index < text.length) {
    if (startsValidEscape(text, index)) {
      const escape = consumeEscape(text, index);
      value += escape.value;
      index = escape.end;
    } else if (isNameUnit(text.charCodeAt(index))) {
      value += text[index++];
    } else {
      break;
    }
  }
  return { value, end: index };
}

/** The decoded identifiers of a white-space separated identifier sequence,
 * tokenized as CSS does (the white space ending a hex escape belongs to the
 * escape), or null when anything else appears. */
function identifierSequence(text: string): string[] | null {
  const idents: string[] = [];
  let index = 0;
  while (index < text.length) {
    if (isWhitespace(text.charCodeAt(index))) { index++; continue; }
    if (!startsIdentifier(text, index)) return null;
    const ident = consumeIdent(text, index);
    index = ident.end;
    if (index < text.length && !isWhitespace(text.charCodeAt(index))) return null;
    idents.push(ident.value);
  }
  return idents;
}

/** `font-family` descriptor (CSS Fonts 4 §4.1, `<family-name>`): one string,
 * or identifiers joined by single spaces. */
function parseFamily(value: string): string | null {
  const text = cssTrim(value);
  if (text.startsWith('"') || text.startsWith("'")) {
    const family = cssString(text)?.trim();
    return family ? family : null;
  }
  const idents = identifierSequence(text);
  if (!idents?.length) return null;
  if (idents.length === 1 && RESERVED_FAMILY_IDENTS.has(idents[0].toLowerCase())) return null;
  const family = idents.join(' ').trim();
  return family || null;
}

const cssTrim = (text: string) => text.replace(/^[ \t\n\f\r]+|[ \t\n\f\r]+$/g, '');

const isNonPrintable = (unit: number) =>
  unit <= 0x08 || unit === 0x0b || (unit >= 0x0e && unit <= 0x1f) || unit === 0x7f;

/**
 * Consume a url token whose contents start at `at` (just after `url(`), per
 * CSS Syntax §4.3.6. Leading and trailing white space is dropped; white space
 * elsewhere, a quote, `(`, a non-printable code point or a backslash that does
 * not start a valid escape makes a bad url, whose remnants are consumed up to
 * the closing `)` (§4.3.14, valid escapes included). Valid escapes are
 * consumed as in identifiers (§4.3.7). End of input closes the token.
 */
function consumeUrlToken(text: string, at: number): { value: string | null; end: number } {
  let index = at;
  while (index < text.length && isWhitespace(text.charCodeAt(index))) index++;
  let value = '';
  const remnants = () => {
    while (index < text.length) {
      if (text[index] === ')') return { value: null, end: index + 1 };
      if (startsValidEscape(text, index)) index = consumeEscape(text, index).end;
      else index++;
    }
    return { value: null, end: index };
  };
  while (index < text.length) {
    const char = text[index];
    const unit = text.charCodeAt(index);
    if (char === ')') return { value, end: index + 1 };
    if (isWhitespace(unit)) {
      while (index < text.length && isWhitespace(text.charCodeAt(index))) index++;
      if (index >= text.length) return { value, end: index };
      if (text[index] === ')') return { value, end: index + 1 };
      return remnants();
    }
    if (char === '"' || char === "'" || char === '(' || isNonPrintable(unit)) return remnants();
    if (char === '\\') {
      if (!startsValidEscape(text, index)) return remnants();
      const escape = consumeEscape(text, index);
      value += escape.value;
      index = escape.end;
      continue;
    }
    value += char;
    index++;
  }
  return { value, end: index };
}

/** The component values the descriptor grammars read (CSS Syntax §4). White
 * space and comments are dropped; escapes are decoded in every token. */
type ValueToken =
  | { readonly type: 'ident'; readonly value: string }
  | { readonly type: 'function'; readonly name: string; readonly args: string }
  | { readonly type: 'url'; readonly value: string }
  | { readonly type: 'string'; readonly value: string }
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'percentage'; readonly value: number }
  | { readonly type: 'dimension'; readonly value: number; readonly unit: string }
  | { readonly type: 'comma' }
  | { readonly type: 'delim'; readonly value: string }
  /** A bad string, bad url or unterminated function: invalid wherever it appears. */
  | { readonly type: 'bad' };

const NUMBER = /[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?/y;

/** Descriptor-value view of the token stream ({@link consumeToken}): white
 * space and comments dropped, a function token taking its arguments up to the
 * matching `)`. A bad string, a bad url or an unterminated function becomes a
 * `bad` token, which no grammar accepts; every other token kind is a delim. */
function tokenizeValue(text: string): ValueToken[] {
  const tokens: ValueToken[] = [];
  let index = 0;
  while (index < text.length) {
    const { token, end } = consumeToken(text, index);
    switch (token.type) {
      case 'whitespace': case 'comment': break;
      case 'string': case 'ident': case 'url': tokens.push({ type: token.type, value: token.value }); break;
      case 'number': case 'percentage': tokens.push({ type: token.type, value: token.value }); break;
      case 'dimension': tokens.push({ type: 'dimension', value: token.value, unit: token.unit }); break;
      case 'bad-string': case 'bad-url': tokens.push({ type: 'bad' }); break;
      case ',': tokens.push({ type: 'comma' }); break;
      case 'function': {
        const close = scanUntil(text, end, ')');
        if (close >= text.length) { tokens.push({ type: 'bad' }); return tokens; }
        tokens.push({ type: 'function', name: token.value.toLowerCase(), args: text.slice(end, close) });
        index = close + 1;
        continue;
      }
      default: tokens.push({ type: 'delim', value: text.slice(index, end) }); break;
    }
    index = end;
  }
  return tokens;
}

/** Split tokens at top-level commas. */
function commaSeparated(tokens: readonly ValueToken[]): ValueToken[][] {
  const groups: ValueToken[][] = [[]];
  for (const token of tokens) {
    if (token.type === 'comma') groups.push([]);
    else groups[groups.length - 1].push(token);
  }
  return groups;
}

/** CSS Values 4 math functions. The fallback grammar accepts them in place
 * of a number, percentage or angle without evaluating them. */
const MATH_FUNCTIONS = new Set([
  'calc', 'min', 'max', 'clamp', 'round', 'mod', 'rem', 'sin', 'cos', 'tan', 'asin', 'acos', 'atan',
  'atan2', 'pow', 'sqrt', 'hypot', 'log', 'exp', 'abs', 'sign',
]);
const mathFunction = (token: ValueToken) =>
  (token.type === 'function' && MATH_FUNCTIONS.has(token.name) ? `${token.name}(${token.args})` : null);

const keyword = (token: ValueToken | undefined) => (token?.type === 'ident' ? token.value.toLowerCase() : undefined);

/**
 * A descriptor parse: null when the declaration is invalid (CSS ignores it,
 * so an earlier valid declaration of the descriptor stays), `{}` for `auto`
 * (the initial value: no descriptor is passed to `FontFace`), otherwise the
 * normalized value passed to `FontFace`.
 */
type DescriptorValue = { readonly value?: string } | null;

/** One or two values, or the single keyword `auto`. */
function rangeDescriptor(value: string, single: (token: ValueToken) => string | null): DescriptorValue {
  const tokens = tokenizeValue(value);
  if (tokens.length === 0 || tokens.length > 2) return null;
  if (tokens.length === 1 && keyword(tokens[0]) === 'auto') return {};
  const values = tokens.map(single);
  return values.every((entry) => entry !== null) ? { value: values.join(' ') } : null;
}

/** `font-weight` (CSS Fonts 4 §4.3): `auto | <font-weight-absolute>{1,2}`,
 * absolute = `normal | bold | <number [1,1000]>`. */
function parseWeight(value: string): DescriptorValue {
  return rangeDescriptor(value, (token) => {
    const word = keyword(token);
    if (word === 'normal' || word === 'bold') return word;
    if (token.type === 'number') return token.value >= 1 && token.value <= 1000 ? String(token.value) : null;
    return mathFunction(token);
  });
}

const STRETCH_KEYWORDS = new Set([
  'normal', 'ultra-condensed', 'extra-condensed', 'condensed', 'semi-condensed',
  'semi-expanded', 'expanded', 'extra-expanded', 'ultra-expanded',
]);

/** `font-stretch` (CSS Fonts 4 §4.4, font-width): `auto | <'font-width'>{1,2}`,
 * a keyword or a non-negative percentage. */
function parseStretch(value: string): DescriptorValue {
  return rangeDescriptor(value, (token) => {
    const word = keyword(token);
    if (word !== undefined) return STRETCH_KEYWORDS.has(word) ? word : null;
    if (token.type === 'percentage') return token.value >= 0 ? `${token.value}%` : null;
    return mathFunction(token);
  });
}

const DEGREES_PER_UNIT: Readonly<Record<string, number>> = { deg: 1, grad: 0.9, rad: 180 / Math.PI, turn: 360 };

/** `font-style` (CSS Fonts 4 §4.2): `auto | normal | italic |
 * oblique [ <angle [-90deg,90deg]>{1,2} ]?`. */
function parseStyle(value: string): DescriptorValue {
  const tokens = tokenizeValue(value);
  if (tokens.length === 0) return null;
  const word = keyword(tokens[0]);
  if (tokens.length === 1 && word === 'auto') return {};
  if (tokens.length === 1 && (word === 'normal' || word === 'italic')) return { value: word };
  if (word !== 'oblique' || tokens.length > 3) return null;
  const angles: string[] = [];
  for (const token of tokens.slice(1)) {
    const math = mathFunction(token);
    if (math) { angles.push(math); continue; }
    if (token.type !== 'dimension' || !Object.hasOwn(DEGREES_PER_UNIT, token.unit)) return null;
    const degrees = token.value * DEGREES_PER_UNIT[token.unit];
    if (!(degrees >= -90 && degrees <= 90)) return null;
    angles.push(`${token.value}${token.unit}`);
  }
  return { value: ['oblique', ...angles].join(' ') };
}

const FORMAT_KEYWORDS = new Set(['collection', 'embedded-opentype', 'opentype', 'svg', 'truetype', 'woff', 'woff2']);
/** Formats the fallback grammar treats as supported (CSS Fonts 4 §4.3.1: an
 * entry with an unsupported format is dropped at parse time). Chosen to match
 * Chrome, the reference engine: svg and embedded-opentype are unsupported,
 * and the legacy `-variations` strings are accepted. */
const FALLBACK_SUPPORTED_FORMATS = new Set([
  'woff2', 'woff', 'truetype', 'opentype', 'collection',
  'woff2-variations', 'woff-variations', 'truetype-variations', 'opentype-variations',
]);
/** Formats whose file the library registers (a collection is not). */
const ACCEPTED_FORMATS = new Set([
  'woff2', 'woff', 'truetype', 'opentype',
  'woff2-variations', 'woff-variations', 'truetype-variations', 'opentype-variations',
]);
const TECH_KEYWORDS = new Set([
  'features-opentype', 'features-aat', 'features-graphite', 'color-colrv0', 'color-colrv1', 'color-svg',
  'color-sbix', 'color-cbdt', 'variations', 'palettes', 'incremental-patch', 'incremental-range', 'incremental-auto',
]);

/** One `src` entry by the fallback grammar: null when invalid, else whether
 * it is a url entry the library can register, and that url. */
function srcEntry(entry: readonly ValueToken[]): { readonly url: string | null } | null {
  const [head, ...rest] = entry;
  if (!head) return null;
  if (head.type === 'function' && head.name === 'local') {
    return rest.length === 0 && parseFamily(head.args) !== null ? { url: null } : null;
  }
  let url: string | null = null;
  if (head.type === 'url') url = head.value;
  else if (head.type === 'function' && head.name === 'url') {
    const args = tokenizeValue(head.args);
    url = args.length === 1 && args[0].type === 'string' ? args[0].value : null;
    if (url === null) return null;
  } else {
    return null;
  }
  let index = 0;
  let accepted = url !== '';
  const next = rest[index];
  if (next?.type === 'function' && next.name === 'format') {
    // Exactly one argument: a string or a format keyword.
    const args = tokenizeValue(next.args);
    if (args.length !== 1) return null;
    const [format] = args;
    if (format.type !== 'string' && !(format.type === 'ident' && FORMAT_KEYWORDS.has(format.value.toLowerCase()))) {
      return null;
    }
    const name = format.value.toLowerCase();
    if (!FALLBACK_SUPPORTED_FORMATS.has(name)) return null;
    accepted &&= ACCEPTED_FORMATS.has(name);
    index++;
  }
  const tech = rest[index];
  if (tech?.type === 'function' && tech.name === 'tech') {
    const groups = commaSeparated(tokenizeValue(tech.args));
    if (!groups.every((group) => group.length === 1 && TECH_KEYWORDS.has(keyword(group[0]) ?? ''))) return null;
    // Support for a font technology cannot be evaluated here.
    accepted = false;
    index++;
  }
  if (index !== rest.length) return null;
  return { url: accepted ? url : null };
}

/**
 * `src` (CSS Fonts 4 §4.3): `[ <url> [format(<font-format>)]? [tech(<font-tech>#)]? | local(<family-name>) ]#`,
 * parsed forgivingly: an invalid entry is dropped, and the declaration is
 * invalid (null) only when no entry is valid. Entry validity is the host's
 * (`entryValid`, a probe FontFace per entry) where the realm has FontFace,
 * the fallback grammar otherwise. `url` is the first valid entry the library
 * registers (a non-empty url with an accepted or absent format and no
 * `tech()`), or null when the valid declaration offers none (only `local()`,
 * say), which drops the rule.
 */
function parseSrc(value: string, entryValid: ((text: string) => boolean) | null): { readonly url: string | null } | null {
  let valid = false;
  let url: string | null = null;
  for (const text of splitTopLevel(value, ',')) {
    const parsed = srcEntry(tokenizeValue(text));
    const isValid = entryValid ? entryValid(cssTrim(text)) : parsed !== null;
    if (!isValid) continue;
    valid = true;
    if (url === null && parsed?.url) url = parsed.url;
  }
  return valid ? { url } : null;
}

/** `unicode-range` (CSS Fonts 4 §4.5): valid only when every range parses.
 * An escape in the value is treated as invalid (ignored): the face then keeps
 * an earlier range or covers every scalar, never fewer than CSS would. */
function parseUnicodeRangeDescriptor(value: string): DescriptorValue {
  const text = cssTrim(stripComments(value));
  if (text.includes('\\') || !parseUnicodeRange(text)) return null;
  return { value: text.split(',').map(cssTrim).join(', ') };
}

type FaceDescriptorKey = 'weight' | 'style' | 'stretch' | 'unicodeRange';

/** The fallback grammars (CSS Fonts 4 §4), used where no `FontFace` exists. */
const GRAMMAR: Readonly<Record<FaceDescriptorKey, (value: string) => DescriptorValue>> = {
  weight: parseWeight,
  style: parseStyle,
  stretch: parseStretch,
  unicodeRange: parseUnicodeRangeDescriptor,
};

const PROBE_FAMILY = 'ooxml-descriptor-probe';

/**
 * Descriptor validation by the host's own CSS parser, where the realm has
 * `FontFace` (window and worker alike). CSS Font Loading §2.1: the
 * constructor parses each descriptor with the `@font-face` grammar and, on a
 * parse failure, sets `status` to "error" (older engines throw a SyntaxError
 * instead; both count as invalid). The probe uses a `local()` source and is
 * never added to a FontFaceSet or loaded, so it fetches nothing; its
 * rejected `loaded` promise is observed. Results are memoized per parse.
 * Chrome's behaviour (status "error", a rejected `loaded`, no request, no
 * FontFaceSet entry; main thread and worker) is checked by
 * `tests/patterns/font-descriptor-host-validation.spec.ts`.
 */
interface HostValidator {
  descriptor(key: FaceDescriptorKey, text: string): DescriptorValue;
  /** One `src` entry, probed as the whole source (`new FontFace(family, src)`
   * parses src without fetching until `load()`). */
  srcEntry(text: string): boolean;
}

function hostDescriptorValidator(): HostValidator | null {
  const Face = (globalThis as { FontFace?: typeof FontFace }).FontFace;
  if (typeof Face !== 'function') return null;
  const memo = new Map<string, boolean>();
  const probe = (memoKey: string, create: () => FontFace) => {
    let valid = memo.get(memoKey);
    if (valid === undefined) {
      try {
        const face = create();
        face.loaded?.catch(() => undefined);
        valid = face.status !== 'error';
      } catch {
        valid = false;
      }
      memo.set(memoKey, valid);
    }
    return valid;
  };
  return {
    descriptor: (key, text) => (probe(`${key}\n${text}`,
      () => new Face(PROBE_FAMILY, `local(${PROBE_FAMILY})`, { [key]: text })) ? { value: text } : null),
    srcEntry: (text) => probe(`src\n${text}`, () => new Face(PROBE_FAMILY, text)),
  };
}

/**
 * Declarations of one `@font-face` body; null when the rule is unusable.
 * Every descriptor follows CSS cascade within the block: an invalid
 * declaration is ignored, so a later invalid value never replaces an earlier
 * valid one, while a later valid value always does. Descriptors the library
 * does not pass to `FontFace` (font-display, font-feature-settings, the
 * metric overrides, …) are not interpreted.
 */
function parseFontFaceBody(
  body: string,
  stylesheetUrl: string,
  location: string,
  validateHost: HostValidator | null,
): FontFaceDeclaration | null {
  let family: string | null = null;
  const srcCandidates: string[] = [];
  const candidates: Record<FaceDescriptorKey, string[]> = { weight: [], style: [], stretch: [], unicodeRange: [] };
  for (const raw of splitTopLevel(body, ';')) {
    const declaration = stripComments(raw);
    const colon = scanUntil(declaration, 0, ':{');
    if (colon >= declaration.length || declaration[colon] !== ':') continue;
    // The name is one identifier token (escapes decoded, ASCII case-insensitive).
    const names = identifierSequence(declaration.slice(0, colon));
    if (names?.length !== 1) continue;
    const name = names[0].toLowerCase();
    const value = declaration.slice(colon + 1);
    // A block-valued or !important descriptor is invalid in @font-face.
    if (scanUntil(value, 0, '{') < value.length || /!\s*important\s*$/i.test(value)) continue;
    switch (name) {
      case 'font-family': family = parseFamily(value) ?? family; break;
      case 'src': srcCandidates.push(value); break;
      case 'font-weight': candidates.weight.push(cssTrim(value)); break;
      case 'font-style': candidates.style.push(cssTrim(value)); break;
      case 'font-stretch': candidates.stretch.push(cssTrim(value)); break;
      case 'unicode-range': candidates.unicodeRange.push(cssTrim(value)); break;
      default: break;
    }
  }
  // The last valid declaration of each descriptor wins: walk back from the
  // last until one validates. The host parser decides where it exists, the
  // fallback grammar otherwise (`auto` there resets to the initial value).
  const descriptors: { weight?: string; style?: string; stretch?: string; unicodeRange?: string } = {};
  for (const key of Object.keys(candidates) as FaceDescriptorKey[]) {
    const texts = candidates[key];
    for (let index = texts.length - 1; index >= 0; index--) {
      const parsed = validateHost ? validateHost.descriptor(key, texts[index]) : GRAMMAR[key](texts[index]);
      if (parsed === null) continue;
      if (parsed.value !== undefined) descriptors[key] = parsed.value;
      break;
    }
  }
  // src: the last valid declaration wins likewise (walking back), then the
  // url to register is extracted from it with the library's own tokenizer.
  let src: string | null = null;
  for (let index = srcCandidates.length - 1; index >= 0; index--) {
    const parsed = parseSrc(srcCandidates[index], validateHost ? validateHost.srcEntry : null);
    if (parsed === null) continue;
    src = parsed.url;
    break;
  }
  if (!family || !src) return null;
  let resolved = src;
  if (!/^\s*data:/i.test(src)) {
    try {
      resolved = new URL(src, stylesheetUrl).href;
    } catch {
      return null;
    }
  }
  try {
    return validateFontFaceDeclaration({ family, url: resolved, ...descriptors }, location);
  } catch (error) {
    // A limit breach (oversized URL or data: face) fails closed; any other
    // invalid rule is dropped like an invalid CSS declaration.
    if (error instanceof OoxmlError) throw error;
    return null;
  }
}

function utf8ByteLength(text: string): number {
  let bytes = 0;
  for (let index = 0; index < text.length; index++) {
    const unit = text.charCodeAt(index);
    if (unit < 0x80) bytes += 1;
    else if (unit < 0x800) bytes += 2;
    else if (unit >= 0xd800 && unit <= 0xdbff && index + 1 < text.length
      && text.charCodeAt(index + 1) >= 0xdc00 && text.charCodeAt(index + 1) <= 0xdfff) { bytes += 4; index++; }
    else bytes += 3;
  }
  return bytes;
}

/**
 * Parse the `@font-face` rules of an untrusted stylesheet into validated
 * declarations, in stylesheet order. `stylesheetUrl` is the absolute final
 * URL the stylesheet was served from (relative font URLs resolve against it).
 */
export function parseFontFaceCss(css: string, stylesheetUrl: string): FontFaceDeclaration[] {
  const size = utf8ByteLength(css);
  if (size > HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES) {
    throw fontSourceLimitError('stylesheet', HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES, size);
  }
  // CSS Syntax §3.3 input preprocessing: CR LF, CR and FF become LF; NUL and
  // lone surrogates become U+FFFD.
  css = preprocessCss(css);
  const declarations: FontFaceDeclaration[] = [];
  const validateHost = hostDescriptorValidator();
  let index = 0;
  while (index < css.length) {
    // CSS Syntax §5.4.1 consume a list of rules (top level): white space,
    // comments, CDO and CDC are skipped; an at-keyword starts an at-rule,
    // anything else a qualified rule.
    const { token, end: tokenEnd } = consumeToken(css, index);
    const type = token.type;
    if (type === 'whitespace' || type === 'comment' || type === 'CDO' || type === 'CDC') { index = tokenEnd; continue; }
    const atName = token.type === 'at-keyword' ? token.value : undefined;
    const isAtRule = atName !== undefined;
    const nameEnd = isAtRule ? tokenEnd : index;
    // An at-rule prelude ends at `;` or `{`; a qualified rule's only at `{`.
    const stop = scanUntil(css, nameEnd, isAtRule ? ';{' : '{');
    if (stop >= css.length) break;
    if (css[stop] === ';') { index = stop + 1; continue; }
    const end = blockEnd(css, stop);
    const isFontFace = atName?.toLowerCase() === 'font-face'
      && stripComments(css.slice(nameEnd, stop)).trim() === '';
    if (isFontFace && end >= 0) {
      const declaration = parseFontFaceBody(css.slice(stop + 1, end), stylesheetUrl,
        `stylesheet @font-face #${declarations.length + 1}`, validateHost);
      if (declaration) {
        if (declarations.length >= HARD_MAX_FONT_SOURCE_DECLARATIONS) {
          throw fontSourceLimitError('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, declarations.length + 1);
        }
        declarations.push(declaration);
      }
    }
    if (end < 0) break;
    index = end + 1;
  }
  return declarations;
}

// ---- bounded stylesheet fetch ---------------------------------------------------

interface ByteReader {
  read(): Promise<{ done: boolean; value?: Uint8Array }>;
  cancel(reason?: unknown): Promise<void>;
}

interface StylesheetResponse {
  readonly ok: boolean;
  readonly status?: number;
  readonly url?: string;
  readonly headers?: { get(name: string): string | null };
  readonly body?: { getReader(): ByteReader; cancel?(reason?: unknown): Promise<void> } | null;
  text(): Promise<string>;
}

/** Cancel an unread response body (the caller also aborts the request), so a
 * rejected response stops downloading. */
function discardBody(response: StylesheetResponse): void {
  try {
    response.body?.cancel?.().catch(() => undefined);
  } catch {
    /* a locked or consumed body: the request abort still stops it */
  }
}

async function readBoundedText(response: StylesheetResponse): Promise<string> {
  const limit = HARD_MAX_FONT_SOURCE_STYLESHEET_BYTES;
  const declared = Number(response.headers?.get('content-length') ?? NaN);
  if (Number.isSafeInteger(declared) && declared > limit) {
    discardBody(response);
    throw fontSourceLimitError('stylesheet', limit, declared);
  }
  const reader = typeof response.body?.getReader === 'function' ? response.body.getReader() : null;
  if (!reader) {
    const text = await response.text();
    const size = utf8ByteLength(text);
    if (size > limit) throw fontSourceLimitError('stylesheet', limit, size);
    return text;
  }
  const decoder = new TextDecoder('utf-8');
  let size = 0;
  let text = '';
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (!value) continue;
    size += value.byteLength;
    if (size > limit) {
      reader.cancel().catch(() => undefined);
      throw fontSourceLimitError('stylesheet', limit, size);
    }
    text += decoder.decode(value, { stream: true });
  }
  return text + decoder.decode();
}

/**
 * Fetch and parse one stylesheet in the painting realm. Credentials are
 * `'same-origin'` unless the source opted into `'omit'`; never `'include'`,
 * never custom headers. A network error, a non-OK status, a final URL that is
 * not http(s) or a stylesheet without any usable rule is a failure (host
 * fallback); a limit breach rejects.
 */
export async function fetchFontStylesheet(
  url: string,
  credentials: FontSourceCredentials | undefined,
  signal: AbortSignal,
): Promise<FontStylesheet | FontSourceFailure> {
  if (typeof fetch !== 'function') return { ok: false, reason: 'fetch is unavailable' };
  // A request of our own, so every early rejection below aborts the download
  // even while the caller's signal stays live.
  const request = new AbortController();
  const onAbort = () => request.abort();
  signal.addEventListener('abort', onAbort, { once: true });
  if (signal.aborted) request.abort();
  let complete = false;
  let response: StylesheetResponse;
  let text: string;
  try {
    response = await fetch(url, { credentials: credentials ?? 'same-origin', signal: request.signal }) as StylesheetResponse;
    if (!response.ok) {
      discardBody(response);
      return { ok: false, reason: `HTTP ${response.status ?? 'error'}` };
    }
    text = await readBoundedText(response);
    complete = true;
  } catch (error) {
    if (error instanceof OoxmlError) throw error;
    if (signal.aborted) throw error;
    return { ok: false, reason: error instanceof Error ? error.message : String(error) };
  } finally {
    signal.removeEventListener('abort', onAbort);
    if (!complete) request.abort();
  }
  const finalUrl = response.url || url;
  if (!/^https?:/i.test(finalUrl)) return { ok: false, reason: 'stylesheet redirected off http(s)' };
  const declarations = parseFontFaceCss(text, finalUrl);
  if (declarations.length === 0) return { ok: false, reason: 'no usable @font-face rule' };
  return { ok: true, declarations, finalUrl };
}

/** Join the stylesheet on behalf of `holder` (the family resolution using it). */
function joinStylesheet(
  cache: SourceCache,
  url: string,
  credentials: FontSourceCredentials | undefined,
  signal: AbortSignal,
  holder: object,
): Promise<FontStylesheet | FontSourceFailure> {
  return cache.stylesheets.join(url, (taskSignal) => fetchFontStylesheet(url, credentials, taskSignal), signal, holder);
}

/** Family index of a parsed stylesheet, living exactly as long as it. */
const stylesheetIndexes = new WeakMap<FontStylesheet, ReadonlyMap<string, readonly FontFaceDeclaration[]>>();

// ---- adapters -----------------------------------------------------------------------

/**
 * `font-css-api`: one request per served family (the demanded family, then each
 * declared substitute's `use`, filtered by the allow-list), issued in
 * parallel. The best-precedence request that yields rules of its served family
 * claims the demanded family; when every request fails the family is final
 * (failed); when the allow-list admits no candidate the family is absent.
 */
async function resolveCssApiFamily(
  source: FontCssApiSource,
  cache: SourceCache,
  family: string,
  signal: AbortSignal,
  holder: object,
  scripts: FontSubstituteScripts,
): Promise<FamilyResolution> {
  const requests = fontCssApiRequests(source, family, scripts);
  if (requests.length === 0) return { kind: 'absent' };
  const sheets = await Promise.all(requests.map((request) =>
    joinStylesheet(cache, request.url, source.credentials, signal, holder)));
  const reasons: string[] = [];
  for (let index = 0; index < requests.length; index++) {
    const request = requests[index];
    const sheet = sheets[index];
    if (!sheet.ok) { reasons.push(sheet.reason); continue; }
    // A per-family stylesheet answers only for the family it was requested for.
    const key = fontFamilyKey(request.family);
    const declarations = sheet.declarations.filter((declaration) => fontFamilyKey(declaration.family) === key);
    if (declarations.length === 0) { reasons.push(`no @font-face rule for ${request.family}`); continue; }
    return {
      kind: 'claimed',
      served: declarations[0].family,
      ...(request.substitution ? { substitution: request.substitution } : {}),
      declarations,
    };
  }
  return { kind: 'failed', reason: reasons.join('; ') };
}

/** `font-css`: fetch the stylesheet once per source identity, then answer
 * every family from its index. A failed stylesheet fails every family asked of
 * it (no other source is consulted). */
async function resolveFontFaceCssFamily(
  source: FontCssSource,
  cache: SourceCache,
  family: string,
  signal: AbortSignal,
  holder: object,
  scripts: FontSubstituteScripts,
): Promise<FamilyResolution> {
  const sheet = await joinStylesheet(cache, source.url, source.credentials, signal, holder);
  if (!sheet.ok) return { kind: 'failed', reason: sheet.reason };
  let index = stylesheetIndexes.get(sheet);
  if (!index) { index = indexFontDeclarations(sheet.declarations); stylesheetIndexes.set(sheet, index); }
  return claimFromIndex(index, family, source.substitutes, scripts);
}

/** Resolve demanded families against a CSS source through the shared
 * resolution cache (one in-flight resolution per source identity × family),
 * held by `cache` (the requesting session's handle) until it releases. */
export async function resolveCssSource(
  source: FontCssApiSource | FontCssSource,
  cache: SourceCache,
  families: readonly string[],
  signal: AbortSignal,
  scriptsOf: (family: string) => FontSubstituteScripts = () => 'all',
): Promise<ReadonlyMap<string, FamilyResolution>> {
  const entries = await Promise.all(families.map(async (family) => {
    const key = fontFamilyKey(family);
    const scripts = scriptsOf(family);
    const resolution = await cache.families.join(scopedFamilyKey(family, scripts), (taskSignal, task) => source.kind === 'font-css-api'
      ? resolveCssApiFamily(source, cache, family, taskSignal, task, scripts)
      : resolveFontFaceCssFamily(source, cache, family, taskSignal, task, scripts), signal, cache);
    return [key, resolution] as const;
  }));
  return new Map(entries);
}
