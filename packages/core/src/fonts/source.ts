/**
 * Font sources: application-owned font provision, as structured-clone-safe
 * descriptors.
 *
 * A font source answers library-owned demand ({@link FontDemand}: the authored
 * families the host lacks and the library's script fallback families) with
 * ordinary `@font-face` declarations
 * ({@link FontFaceDeclaration}: a URL or owned bytes plus standard
 * descriptors). Sources may declare family substitutions
 * (`substitutions.ts`); the renderers consume only the substitutions whose
 * served face actually loaded. No source kind is provider-specific: a hosted
 * CSS API such as Google Fonts is a {@link fontCssApi} descriptor filled with that
 * provider's data (`presets/google-fonts.ts`).
 *
 * Descriptors are plain frozen data so a load request can carry them to a
 * render worker unchanged. The one exception is {@link fontResolver}: its
 * application function never crosses a realm. The descriptor carries only an
 * id, and the function stays in a main-realm `WeakMap`.
 *
 * Every factory validates its input before any load, throwing `TypeError` for
 * a malformed descriptor and `OoxmlError('font-source-limit')` for a value
 * beyond a generated hard ceiling. {@link assertFontSourceDescriptors} applies
 * the same validation to descriptors that crossed a realm. Relative URLs are
 * resolved against the calling realm's base URL at factory time, so a
 * descriptor only ever holds absolute URLs and a worker never resolves an
 * application URL against its own script URL.
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { nullRecord } from '../internal/own-key.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_FACE_BYTES,
  HARD_MAX_FONT_SOURCE_URL_LENGTH,
} from '../worker/resource-policy.generated.js';
import type { FontSubstituteScript } from './substitute-script.js';
import { fontFamilyKey, substitutionsFor, type FontSubstituteScripts, type FontSubstitution } from './substitutions.js';

export type { FontSubstitution } from './substitutions.js';

export const FONT_SOURCE_PROTOCOL = 'ooxml-font-source/v1';

/**
 * Credentials of the requests the library sends for a source. Never
 * `'include'` and never custom headers: an authenticated cross-origin service
 * answers through {@link fontResolver}.
 *
 * - `'same-origin'` (the default) matches CSS font fetching: the stylesheet is
 *   fetched with it and face files are left to `FontFace.load()`, which sends
 *   credentials to the same origin only.
 * - `'omit'` sends no cookies or HTTP authentication at all: the stylesheet
 *   and every `url` face of a claimed family are fetched by the library with
 *   `credentials: 'omit'` (face bytes then count like `data` faces). The
 *   faces are downloaded when the family is claimed, not per `unicode-range`
 *   subset as the browser would.
 */
export type FontSourceCredentials = 'omit' | 'same-origin';

interface FontSourceBase {
  readonly protocol: typeof FONT_SOURCE_PROTOCOL;
  readonly kind: string;
  readonly substitutes?: readonly FontSubstitution[];
}

/** A CSS API answering one request per served family. */
export interface FontCssApiSource extends FontSourceBase {
  readonly kind: 'font-css-api';
  /** Absolute http(s) URL template with `{family}` (required) and `{query}`. */
  readonly urlTemplate: string;
  /** `{query}` text for families without an allow-list override. */
  readonly query?: string;
  /** Optional allow-list of served families, matched case-insensitively. Each
   * key's spelling is the `{family}` sent; its value is a `{query}` override
   * or `true` for the default query. When present, only these are requested. */
  readonly families?: Readonly<Record<string, string | true>>;
  readonly credentials?: FontSourceCredentials;
}

/** One static stylesheet describing any number of families. */
export interface FontCssSource extends FontSourceBase {
  readonly kind: 'font-css';
  readonly url: string;
  readonly credentials?: FontSourceCredentials;
}

/** Faces declared statically by the application. */
export interface FontFilesSource extends FontSourceBase {
  readonly kind: 'font-files';
  readonly faces: readonly FontFaceDeclaration[];
  readonly credentials?: FontSourceCredentials;
}

/** Application resolution on the main thread; see {@link fontResolver}. */
export interface FontResolverSource extends FontSourceBase {
  readonly kind: 'font-resolver';
  /** Main-realm identity of the resolver function. */
  readonly id: string;
}

export type FontSource = FontCssApiSource | FontCssSource | FontFilesSource | FontResolverSource;

/** One `@font-face` a source answers with. Exactly one of `url` / `data`. */
export interface FontFaceDeclaration {
  /** CSS family it registers under. */
  readonly family: string;
  /** Absolute http(s), blob: or data: URL. Fetched by `FontFace.load()` in
   * the painting realm, except that the library downloads it itself for a
   * source with `credentials: 'omit'` and, in worker mode, in the page for a
   * {@link fontResolver} answer (see {@link FontSourceCredentials}). */
  readonly url?: string;
  /** Owned font bytes; transferred once to a worker. */
  readonly data?: ArrayBuffer;
  readonly weight?: string;
  readonly style?: string;
  readonly stretch?: string;
  /** Standard CSS `unicode-range`; enables lazy subset loading. */
  readonly unicodeRange?: string;
}

/**
 * What a {@link fontResolver} is asked for: the families one document (or one
 * progressive delta) needs that no earlier call answered.
 *
 * Contract: answer each family with every face you have for it (all weights,
 * styles and `unicode-range` subsets), or with nothing. A family is resolved
 * once per document, so faces left out are never asked for again; the
 * renderer picks the weight and style per run and loads `unicode-range`
 * subsets only for the characters painted.
 */
export interface FontDemand {
  /** Authored family names not proven installed on the host (trimmed, original case). */
  readonly families: readonly string[];
  /** Library fallback families the renderer puts in CSS tails for the scripts
   * present in the text (for example Noto families for CJK). */
  readonly fallbackFamilies: readonly string[];
}

export interface FontResolveContext {
  /** Aborted when the requesting document is destroyed or the font ceiling expires. */
  readonly signal: AbortSignal;
}

export type FontResolveFn = (
  demand: FontDemand,
  context: FontResolveContext,
) => readonly FontFaceDeclaration[] | Promise<readonly FontFaceDeclaration[]>;

export interface FontCssApiOptions {
  readonly urlTemplate: string;
  readonly query?: string;
  readonly families?: Readonly<Record<string, string | true>>;
  readonly substitutes?: readonly FontSubstitution[];
  readonly credentials?: FontSourceCredentials;
}

export interface FontCssOptions {
  readonly url: string;
  readonly substitutes?: readonly FontSubstitution[];
  readonly credentials?: FontSourceCredentials;
}

export interface FontFilesOptions {
  readonly faces: readonly FontFaceDeclaration[];
  readonly substitutes?: readonly FontSubstitution[];
  readonly credentials?: FontSourceCredentials;
}

export interface FontResolverOptions {
  readonly substitutes?: readonly FontSubstitution[];
}

// ---- validation primitives --------------------------------------------------

type Fields = Readonly<Record<string, unknown>>;

function requireRecord(value: unknown, location: string): Fields {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${location} must be an object`);
  }
  return value as Fields;
}

function requireKnownKeys(value: Fields, allowed: readonly string[], location: string): void {
  for (const key of Object.keys(value)) {
    if (!allowed.includes(key)) throw new TypeError(`${location} has unknown property "${key}"`);
  }
}

function requireName(value: unknown, location: string): string {
  if (typeof value !== 'string' || value.trim() === '' || value.trim() !== value) {
    throw new TypeError(`${location} must be a non-empty string without surrounding white space`);
  }
  return value;
}

function optionalDescriptor(value: unknown, location: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || value.trim() === '') {
    throw new TypeError(`${location} must be a non-empty string`);
  }
  return value;
}

function optionalCredentials(value: unknown, location: string): FontSourceCredentials | undefined {
  if (value === undefined || value === 'omit' || value === 'same-origin') return value;
  throw new TypeError(`${location} must be 'omit' or 'same-origin'`);
}

// Intrinsic ArrayBuffer accessors, read once: brand checks must not trust
// Symbol.toStringTag, own properties or a Proxy (none of which structured clone
// or FontFace would accept), and the size charged must be the internal one.
type Getter = (this: unknown) => unknown;
function intrinsicGetter(prototype: object | undefined, name: string): Getter | undefined {
  return prototype ? Object.getOwnPropertyDescriptor(prototype, name)?.get as Getter | undefined : undefined;
}
const ARRAY_BUFFER_BYTE_LENGTH = intrinsicGetter(ArrayBuffer.prototype, 'byteLength') as Getter;
const ARRAY_BUFFER_DETACHED = intrinsicGetter(ArrayBuffer.prototype, 'detached');
const ARRAY_BUFFER_RESIZABLE = intrinsicGetter(ArrayBuffer.prototype, 'resizable');
const SHARED_ARRAY_BUFFER_BYTE_LENGTH = intrinsicGetter(
  (globalThis as { SharedArrayBuffer?: { prototype: object } }).SharedArrayBuffer?.prototype, 'byteLength');

function hasBrand(getter: Getter | undefined, value: object): boolean {
  if (!getter) return false;
  try {
    Reflect.apply(getter, value, []);
    return true;
  } catch {
    return false;
  }
}

/**
 * The internal byte length of a genuine, owned, fixed-length `ArrayBuffer`.
 * Views, `SharedArrayBuffer`s (shared memory cannot be owned or transferred),
 * resizable buffers (they could grow past the ceiling after admission),
 * detached and empty buffers, and look-alike objects are rejected.
 */
function requireOwnedArrayBuffer(value: unknown, location: string): number {
  if (typeof value !== 'object' || value === null) throw new TypeError(`${location} must be an ArrayBuffer`);
  if (ArrayBuffer.isView(value)) throw new TypeError(`${location} must be an ArrayBuffer, not a view`);
  if (hasBrand(SHARED_ARRAY_BUFFER_BYTE_LENGTH, value)) {
    throw new TypeError(`${location} must be an ArrayBuffer, not a SharedArrayBuffer`);
  }
  if (!hasBrand(ARRAY_BUFFER_BYTE_LENGTH, value)) throw new TypeError(`${location} must be an ArrayBuffer`);
  if (ARRAY_BUFFER_DETACHED && Reflect.apply(ARRAY_BUFFER_DETACHED, value, []) === true) {
    throw new TypeError(`${location} must not be detached`);
  }
  if (ARRAY_BUFFER_RESIZABLE && Reflect.apply(ARRAY_BUFFER_RESIZABLE, value, []) === true) {
    throw new TypeError(`${location} must not be resizable`);
  }
  const byteLength = Reflect.apply(ARRAY_BUFFER_BYTE_LENGTH, value, []) as number;
  if (byteLength === 0) throw new TypeError(`${location} must not be empty or detached`);
  return byteLength;
}

const isAsciiWhitespace = (byte: number) =>
  byte === 0x09 || byte === 0x0a || byte === 0x0c || byte === 0x0d || byte === 0x20;
const hexValue = (unit: number) =>
  unit >= 0x30 && unit <= 0x39 ? unit - 0x30
    : unit >= 0x41 && unit <= 0x46 ? unit - 0x37
      : unit >= 0x61 && unit <= 0x66 ? unit - 0x57 : -1;

/**
 * Decoded body size of a `data:` URL, computed in one pass without
 * materialising the body, following the URL parser (outer C0/space trimmed,
 * tab/LF/CR removed, fragment excluded, non-ASCII UTF-8 encoded) and the Fetch
 * data: URL processor (header up to the first comma; body percent-decoded;
 * `;base64` bodies forgiving-base64 decoded: ASCII whitespace removed, up to
 * two trailing `=` dropped when the length is a multiple of four, 3 bytes per
 * 4 characters). The header never counts. Returns undefined without a comma
 * (not a data: URL any fetch would accept).
 */
export function dataUrlDecodedByteLength(url: string): number | undefined {
  let end = url.length;
  while (end > 0 && url.charCodeAt(end - 1) <= 0x20) end--;
  let start = 0;
  while (start < end && url.charCodeAt(start) <= 0x20) start++;
  const units: number[] = [];
  let position = start;
  // Next code unit after URL-parser tab/newline removal, or -1 at the end.
  const next = (): number => {
    while (position < end) {
      const unit = url.charCodeAt(position++);
      if (unit !== 0x09 && unit !== 0x0a && unit !== 0x0d) return unit;
    }
    return -1;
  };
  for (let index = 0; index < 5; index++) units.push(next());
  if (String.fromCharCode(...units).toLowerCase() !== 'data:') return undefined;
  let header = '';
  for (let unit = next(); ; unit = next()) {
    if (unit === -1 || unit === 0x23 /* # */) return undefined;
    if (unit === 0x2c /* , */) break;
    if (header.length > HARD_MAX_FONT_SOURCE_URL_LENGTH) {
      throw fontSourceLimitError('url', HARD_MAX_FONT_SOURCE_URL_LENGTH, header.length);
    }
    header += String.fromCharCode(unit);
  }
  const base64 = /;\u0020*base64$/i.test(header.replace(/^[\t\n\f\r ]+|[\t\n\f\r ]+$/g, ''));
  let bytes = 0;
  let trailingEquals = 0;
  const add = (byte: number) => {
    if (!base64) { bytes++; return; }
    if (isAsciiWhitespace(byte)) return;
    bytes++;
    trailingEquals = byte === 0x3d ? trailingEquals + 1 : 0;
  };
  for (let unit = next(); unit !== -1 && unit !== 0x23; unit = next()) {
    if (unit === 0x25 /* % */) {
      const save = position;
      const high = hexValue(next());
      const low = high < 0 ? -1 : hexValue(next());
      if (low >= 0) { add(high * 16 + low); continue; }
      position = save;
      add(0x25);
    } else if (unit < 0x80) {
      add(unit);
    } else {
      // Non-ASCII is UTF-8 percent-encoded by the URL parser; a lone
      // surrogate becomes U+FFFD. None of these bytes is whitespace or '='.
      const pair = unit >= 0xd800 && unit <= 0xdbff && position < end
        && url.charCodeAt(position) >= 0xdc00 && url.charCodeAt(position) <= 0xdfff;
      if (pair) position++;
      const length = pair ? 4 : unit < 0x800 ? 2 : 3;
      for (let index = 0; index < length; index++) add(0x80);
    }
  }
  if (!base64) return bytes;
  if (bytes % 4 === 0) bytes -= Math.min(trailingEquals, 2);
  return Math.floor(bytes * 3 / 4);
}

/** The calling realm's base for relative application URLs; undefined in a
 * realm without one (Node), where relative URLs are rejected. */
function realmBaseUrl(): string | undefined {
  const scope = globalThis as { document?: { baseURI?: unknown }; location?: { href?: unknown } };
  const base = scope.document?.baseURI ?? scope.location?.href;
  return typeof base === 'string' ? base : undefined;
}

/**
 * Validate one URL. Non-`data:` URLs are bounded by
 * `HARD_MAX_FONT_SOURCE_URL_LENGTH` (as written and after URL-parser
 * encoding). A `data:` URL is bounded by the face byte ceiling instead,
 * charged its decoded body size ({@link dataUrlDecodedByteLength}), so inline
 * faces obey the same byte budget as `data` buffers.
 */
function requireUrl(
  value: unknown,
  location: string,
  protocols: readonly string[],
  base: string | undefined,
): string {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new TypeError(`${location} must be a non-empty URL string`);
  }
  const isData = /^data:/i.test(value);
  if (isData && protocols.includes('data:')) {
    const decoded = dataUrlDecodedByteLength(value);
    if (decoded === undefined) throw new TypeError(`${location} must be a data: URL with a body`);
    if (decoded > HARD_MAX_FONT_SOURCE_FACE_BYTES) {
      throw fontSourceLimitError('face', HARD_MAX_FONT_SOURCE_FACE_BYTES, decoded);
    }
  } else if (value.length > HARD_MAX_FONT_SOURCE_URL_LENGTH) {
    throw fontSourceLimitError('url', HARD_MAX_FONT_SOURCE_URL_LENGTH, value.length);
  }
  let parsed: URL;
  try {
    parsed = base === undefined ? new URL(value) : new URL(value, base);
  } catch {
    throw new TypeError(`${location} must be an absolute URL`);
  }
  if (!protocols.includes(parsed.protocol)) {
    throw new TypeError(`${location} must use ${protocols.join(' / ')}`);
  }
  const href = parsed.href;
  if (!isData && href.length > HARD_MAX_FONT_SOURCE_URL_LENGTH) {
    throw fontSourceLimitError('url', HARD_MAX_FONT_SOURCE_URL_LENGTH, href.length);
  }
  return isData ? value : href;
}

const STYLESHEET_PROTOCOLS = ['https:', 'http:'] as const;
const FACE_PROTOCOLS = ['https:', 'http:', 'blob:', 'data:'] as const;

const SUBSTITUTE_SCRIPTS = { arabic: true } satisfies Record<FontSubstituteScript, true>;

function normalizeSubstitutes(value: unknown, location: string): readonly FontSubstitution[] | undefined {
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) throw new TypeError(`${location} must be an array`);
  return Object.freeze(value.map((entry: unknown, index) => {
    const at = `${location}[${index}]`;
    const fields = requireRecord(entry, at);
    requireKnownKeys(fields, ['for', 'use', 'script'], at);
    const script = fields.script;
    if (script !== undefined
      && !(typeof script === 'string' && Object.prototype.hasOwnProperty.call(SUBSTITUTE_SCRIPTS, script))) {
      throw new TypeError(`${at}.script must be one of: ${Object.keys(SUBSTITUTE_SCRIPTS).join(', ')}`);
    }
    return Object.freeze({
      for: requireName(fields.for, `${at}.for`),
      use: requireName(fields.use, `${at}.use`),
      ...(script === undefined ? {} : { script: script as FontSubstituteScript }),
    });
  }));
}

/**
 * Validate one face declaration: exactly one of `url` / `data`, `data` an
 * owned non-empty `ArrayBuffer` within the face ceiling, `url` an absolute
 * http(s) / blob: / data: URL (relative URLs resolve against `base`).
 * The buffer is referenced, not copied.
 */
export function validateFontFaceDeclaration(
  value: unknown,
  location = 'font face declaration',
  base: string | undefined = undefined,
): FontFaceDeclaration {
  const fields = requireRecord(value, location);
  requireKnownKeys(fields, ['family', 'url', 'data', 'weight', 'style', 'stretch', 'unicodeRange'], location);
  const family = requireName(fields.family, `${location}.family`);
  const hasUrl = fields.url !== undefined;
  const hasData = fields.data !== undefined;
  if (hasUrl === hasData) throw new TypeError(`${location} must have exactly one of url or data`);
  let url: string | undefined;
  let data: ArrayBuffer | undefined;
  if (hasUrl) {
    url = requireUrl(fields.url, `${location}.url`, FACE_PROTOCOLS, base);
  } else {
    const byteLength = requireOwnedArrayBuffer(fields.data, `${location}.data`);
    if (byteLength > HARD_MAX_FONT_SOURCE_FACE_BYTES) {
      throw fontSourceLimitError('face', HARD_MAX_FONT_SOURCE_FACE_BYTES, byteLength);
    }
    data = fields.data as ArrayBuffer;
  }
  const weight = optionalDescriptor(fields.weight, `${location}.weight`);
  const style = optionalDescriptor(fields.style, `${location}.style`);
  const stretch = optionalDescriptor(fields.stretch, `${location}.stretch`);
  const unicodeRange = optionalDescriptor(fields.unicodeRange, `${location}.unicodeRange`);
  return Object.freeze({
    family,
    ...(url === undefined ? {} : { url }),
    ...(data === undefined ? {} : { data }),
    ...(weight === undefined ? {} : { weight }),
    ...(style === undefined ? {} : { style }),
    ...(stretch === undefined ? {} : { stretch }),
    ...(unicodeRange === undefined ? {} : { unicodeRange }),
  });
}

const TEMPLATE_PLACEHOLDER = /\{[^{}]*\}/g;

function normalizeCssApi(fields: Fields, location: string): FontCssApiSource {
  requireKnownKeys(fields, ['protocol', 'kind', 'urlTemplate', 'query', 'families', 'substitutes', 'credentials'], location);
  const urlTemplate = fields.urlTemplate;
  if (typeof urlTemplate !== 'string' || urlTemplate.trim() === '') {
    throw new TypeError(`${location}.urlTemplate must be a non-empty string`);
  }
  if (urlTemplate.length > HARD_MAX_FONT_SOURCE_URL_LENGTH) {
    throw fontSourceLimitError('url', HARD_MAX_FONT_SOURCE_URL_LENGTH, urlTemplate.length);
  }
  const placeholders: readonly string[] = urlTemplate.match(TEMPLATE_PLACEHOLDER) ?? [];
  for (const placeholder of placeholders) {
    if (placeholder !== '{family}' && placeholder !== '{query}') {
      throw new TypeError(`${location}.urlTemplate has unknown placeholder ${placeholder}`);
    }
  }
  if (!placeholders.includes('{family}')) {
    throw new TypeError(`${location}.urlTemplate must contain {family}`);
  }
  if (/[{}]/.test(urlTemplate.replace(TEMPLATE_PLACEHOLDER, ''))) {
    throw new TypeError(`${location}.urlTemplate has an unbalanced brace`);
  }
  // The template itself must already be absolute: placeholders are filled by
  // string substitution, never resolved against a base.
  requireUrl(fillTemplate(urlTemplate, 'family', 'query'), `${location}.urlTemplate`, STYLESHEET_PROTOCOLS, undefined);
  const query = optionalDescriptor(fields.query, `${location}.query`);
  let families: Readonly<Record<string, string | true>> | undefined;
  if (fields.families !== undefined) {
    const entries = requireRecord(fields.families, `${location}.families`);
    const seen = new Set<string>();
    // Keys are family names: a null-prototype copy keeps `__proto__` as an own
    // allow-list entry (plain assignment would invoke the inherited setter).
    const copy = nullRecord<string | true>();
    for (const [family, value] of Object.entries(entries)) {
      const at = `${location}.families[${JSON.stringify(family)}]`;
      requireName(family, `${at} key`);
      const key = fontFamilyKey(family);
      if (seen.has(key)) throw new TypeError(`${at} duplicates another family`);
      seen.add(key);
      if (value !== true && (typeof value !== 'string' || value.trim() === '')) {
        throw new TypeError(`${at} must be true or a non-empty query string`);
      }
      copy[family] = value;
    }
    families = Object.freeze(copy);
  }
  if (placeholders.includes('{query}') && query === undefined
    && (!families || Object.values(families).some((value) => value === true))) {
    throw new TypeError(`${location}.query is required when urlTemplate contains {query}`);
  }
  // Fail at the factory for every request whose family is known now: each
  // allow-listed family, or the default query with an empty family. Requests
  // for document-supplied families are checked again when composed.
  if (families) {
    for (const [family, value] of Object.entries(families)) {
      composeCssApiUrl(urlTemplate, family, value === true ? query : value);
    }
  } else {
    composeCssApiUrl(urlTemplate, '', query);
  }
  const substitutes = normalizeSubstitutes(fields.substitutes, `${location}.substitutes`);
  const credentials = optionalCredentials(fields.credentials, `${location}.credentials`);
  return Object.freeze({
    protocol: FONT_SOURCE_PROTOCOL,
    kind: 'font-css-api',
    urlTemplate,
    ...(query === undefined ? {} : { query }),
    ...(families === undefined ? {} : { families }),
    ...(substitutes === undefined ? {} : { substitutes }),
    ...(credentials === undefined ? {} : { credentials }),
  });
}

function normalizeFontFaceCss(fields: Fields, location: string, base: string | undefined): FontCssSource {
  requireKnownKeys(fields, ['protocol', 'kind', 'url', 'substitutes', 'credentials'], location);
  const url = requireUrl(fields.url, `${location}.url`, STYLESHEET_PROTOCOLS, base);
  const substitutes = normalizeSubstitutes(fields.substitutes, `${location}.substitutes`);
  const credentials = optionalCredentials(fields.credentials, `${location}.credentials`);
  return Object.freeze({
    protocol: FONT_SOURCE_PROTOCOL,
    kind: 'font-css',
    url,
    ...(substitutes === undefined ? {} : { substitutes }),
    ...(credentials === undefined ? {} : { credentials }),
  });
}

function normalizeFontFiles(fields: Fields, location: string, base: string | undefined): FontFilesSource {
  requireKnownKeys(fields, ['protocol', 'kind', 'faces', 'substitutes', 'credentials'], location);
  const faces = fields.faces;
  if (!Array.isArray(faces) || faces.length === 0) {
    throw new TypeError(`${location}.faces must be a non-empty array`);
  }
  if (faces.length > HARD_MAX_FONT_SOURCE_DECLARATIONS) {
    throw fontSourceLimitError('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, faces.length);
  }
  const declarations = Object.freeze(faces.map((face: unknown, index) =>
    validateFontFaceDeclaration(face, `${location}.faces[${index}]`, base)));
  const substitutes = normalizeSubstitutes(fields.substitutes, `${location}.substitutes`);
  const credentials = optionalCredentials(fields.credentials, `${location}.credentials`);
  return Object.freeze({
    protocol: FONT_SOURCE_PROTOCOL,
    kind: 'font-files',
    faces: declarations,
    ...(substitutes === undefined ? {} : { substitutes }),
    ...(credentials === undefined ? {} : { credentials }),
  });
}

const RESOLVER_ID = /^font-resolver:[1-9]\d*$/;

function normalizeFontResolver(fields: Fields, location: string): FontResolverSource {
  requireKnownKeys(fields, ['protocol', 'kind', 'id', 'substitutes'], location);
  if (typeof fields.id !== 'string' || !RESOLVER_ID.test(fields.id)) {
    throw new TypeError(`${location}.id must be a font resolver id`);
  }
  const substitutes = normalizeSubstitutes(fields.substitutes, `${location}.substitutes`);
  return Object.freeze({
    protocol: FONT_SOURCE_PROTOCOL,
    kind: 'font-resolver',
    id: fields.id,
    ...(substitutes === undefined ? {} : { substitutes }),
  });
}

function normalizeFontSource(value: unknown, location: string, base: string | undefined): FontSource {
  const fields = requireRecord(value, location);
  if (fields.protocol !== FONT_SOURCE_PROTOCOL) {
    throw new TypeError(`${location} has unsupported protocol ${String(fields.protocol)}`);
  }
  switch (fields.kind) {
    case 'font-css-api': return normalizeCssApi(fields, location);
    case 'font-css': return normalizeFontFaceCss(fields, location, base);
    case 'font-files': return normalizeFontFiles(fields, location, base);
    case 'font-resolver': return normalizeFontResolver(fields, location);
    default: throw new TypeError(`${location} has unsupported kind ${String(fields.kind)}`);
  }
}

// ---- factories ----------------------------------------------------------------

/** Descriptors this realm's factories produced (already validated, frozen). */
const factoryDescriptors = new WeakSet<FontSource>();

function issued<T extends FontSource>(source: T): T {
  factoryDescriptors.add(source);
  return source;
}

function factoryOptions(value: unknown, name: string, allowed: readonly string[]): Fields {
  const fields = requireRecord(value, `${name}() options`);
  requireKnownKeys(fields, allowed, `${name}() options`);
  return fields;
}

/** Any CSS API that answers one stylesheet per family through a URL template
 * (`{family}` is percent-encoded with `+` for spaces; `{query}` is inserted
 * verbatim from descriptor data, never from document content). */
export function fontCssApi(options: FontCssApiOptions): FontCssApiSource {
  const fields = factoryOptions(options, 'fontCssApi', ['urlTemplate', 'query', 'families', 'substitutes', 'credentials']);
  return issued(normalizeCssApi({ ...fields, protocol: FONT_SOURCE_PROTOCOL, kind: 'font-css-api' }, 'fontCssApi()'));
}

/** One static stylesheet (self-hosted or kit CSS) describing many families. */
export function fontCss(options: FontCssOptions): FontCssSource {
  const fields = factoryOptions(options, 'fontCss', ['url', 'substitutes', 'credentials']);
  return issued(normalizeFontFaceCss({ ...fields, protocol: FONT_SOURCE_PROTOCOL, kind: 'font-css' },
    'fontCss()', realmBaseUrl()));
}

/** Faces declared up front by the application (URLs or owned bytes). */
export function fontFiles(options: FontFilesOptions): FontFilesSource {
  const fields = factoryOptions(options, 'fontFiles', ['faces', 'substitutes', 'credentials']);
  return issued(normalizeFontFiles({ ...fields, protocol: FONT_SOURCE_PROTOCOL, kind: 'font-files' },
    'fontFiles()', realmBaseUrl()));
}

const resolverIds = new WeakMap<FontResolveFn, string>();
const resolverFunctions = new WeakMap<FontResolverSource, FontResolveFn>();
let nextResolverId = 1;

/**
 * Application resolution for authenticated, dynamic or in-memory fonts. The
 * function runs on the main thread only; in worker mode the worker asks main
 * to run it (`resolver-bridge.ts`), so credentials and network clients stay in
 * the application, and main also downloads every http(s) / blob: `url` the
 * function returns and sends the bytes, so cookies, HTTP authentication and
 * Referer checks see the page in both modes. The same function always yields
 * the same id, so documents sharing it share resolutions. See
 * {@link FontDemand} for what to answer.
 *
 * Returned `data` buffers are owned by the library from then on (cached and
 * shared across documents of this realm); the application must not modify
 * them. In worker mode the worker receives a copy; nothing is detached.
 */
export function fontResolver(resolve: FontResolveFn, options: FontResolverOptions = {}): FontResolverSource {
  if (typeof resolve !== 'function') throw new TypeError('fontResolver() requires a function');
  const fields = factoryOptions(options, 'fontResolver', ['substitutes']);
  let id = resolverIds.get(resolve);
  if (id === undefined) {
    id = `font-resolver:${nextResolverId++}`;
    resolverIds.set(resolve, id);
  }
  const source = normalizeFontResolver({ ...fields, protocol: FONT_SOURCE_PROTOCOL, kind: 'font-resolver', id },
    'fontResolver()');
  resolverFunctions.set(source, resolve);
  return issued(source);
}

/** The application function behind a {@link fontResolver} descriptor of this
 * realm, or undefined (a descriptor from another realm or not from the factory). */
export function fontResolverFunction(source: FontResolverSource): FontResolveFn | undefined {
  return resolverFunctions.get(source);
}

/**
 * Validate a `fonts` option value, in the realm that received it. Descriptors
 * issued by this realm's factories are returned as they are; anything else
 * (descriptors that crossed a realm, or hand-written objects) is validated as
 * strictly as the factories validate and returned as frozen copies. URLs must
 * already be absolute here.
 */
export function assertFontSourceDescriptors(value: unknown): readonly FontSource[] {
  if (!Array.isArray(value)) throw new TypeError('fonts must be an array of font sources');
  return Object.freeze(value.map((entry: unknown, index) =>
    factoryDescriptors.has(entry as FontSource)
      ? entry as FontSource
      : normalizeFontSource(entry, `fonts[${index}]`, undefined)));
}

// ---- font-css-api request composition -------------------------------------------

function fillTemplate(template: string, family: string, query: string): string {
  return template.split('{family}').join(family).split('{query}').join(query);
}

/** CSS API family encoding: percent-encoding with `+` for spaces. A lone
 * surrogate (possible in document-supplied names) is sent as U+FFFD, as the
 * URL parser would, instead of throwing. */
function encodeCssApiFamily(family: string): string {
  const wellFormed = family.replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, '\ufffd');
  return encodeURIComponent(wellFormed).replace(/%20/g, '+');
}

/** The request URL exactly as fetched (URL-parser serialization), bounded by
 * `HARD_MAX_FONT_SOURCE_URL_LENGTH` after family encoding and any further
 * percent-encoding of the query. */
function composeCssApiUrl(template: string, family: string, query: string | undefined): string {
  const filled = fillTemplate(template, encodeCssApiFamily(family), query ?? '');
  if (filled.length > HARD_MAX_FONT_SOURCE_URL_LENGTH) {
    throw fontSourceLimitError('url', HARD_MAX_FONT_SOURCE_URL_LENGTH, filled.length);
  }
  const href = new URL(filled).href;
  if (href.length > HARD_MAX_FONT_SOURCE_URL_LENGTH) {
    throw fontSourceLimitError('url', HARD_MAX_FONT_SOURCE_URL_LENGTH, href.length);
  }
  return href;
}

/** One stylesheet request a {@link FontCssApiSource} issues for a demanded family. */
export interface FontCssApiRequest {
  /** Served family the stylesheet is requested for (allow-list spelling). */
  readonly family: string;
  readonly url: string;
  /** Present when the request serves a declared substitute. */
  readonly substitution?: FontSubstitution;
}

/**
 * Requests for one demanded family, in precedence order: the family itself,
 * then each declared substitute's `use`, keeping only families that pass the
 * allow-list. One request per served family; duplicates are dropped. Throws
 * `OoxmlError('font-source-limit')` when a completed URL exceeds
 * `HARD_MAX_FONT_SOURCE_URL_LENGTH`.
 */
export function fontCssApiRequests(
  source: FontCssApiSource,
  demandedFamily: string,
  scripts: FontSubstituteScripts = 'all',
): readonly FontCssApiRequest[] {
  const candidates: { family: string; substitution?: FontSubstitution }[] = [
    { family: demandedFamily.trim() },
    ...substitutionsFor(source.substitutes, demandedFamily, scripts).map((substitution) => ({
      family: substitution.use, substitution,
    })),
  ];
  const requests: FontCssApiRequest[] = [];
  const urls = new Set<string>();
  for (const candidate of candidates) {
    if (!candidate.family) continue;
    let served = candidate.family;
    let query = source.query;
    if (source.families) {
      const key = fontFamilyKey(candidate.family);
      const entry = Object.entries(source.families).find(([family]) => fontFamilyKey(family) === key);
      if (!entry) continue;
      served = entry[0];
      if (entry[1] !== true) query = entry[1];
    }
    const url = composeCssApiUrl(source.urlTemplate, served, query);
    if (urls.has(url)) continue;
    urls.add(url);
    requests.push({
      family: served, url,
      ...(candidate.substitution ? { substitution: candidate.substitution } : {}),
    });
  }
  return requests;
}
