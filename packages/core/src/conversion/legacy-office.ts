import { resolveOoxmlContainer, toArrayBuffer } from '../errors/cfb-guard.js';
import { sniffCfb, sniffLegacyOfficeFormat } from '../errors/cfb-sniff.js';
import type { OoxmlFormat } from '../errors/ooxml-error.js';
import {
  LegacyOfficeConversionError,
  type LegacyOfficeFormat,
} from './legacy-office-error.js';

export {
  LegacyOfficeConversionError,
  type LegacyOfficeConversionFailureReason,
  type LegacyOfficeFormat,
} from './legacy-office-error.js';

/** Bytes and fixed same-family formats handed to an opt-in converter. */
export interface LegacyOfficeConversionInput {
  /**
   * Exactly-sized view owned by this conversion request. A converter may
   * transfer (and therefore detach) its backing ArrayBuffer. It must not retain
   * or mutate the view after the returned promise settles.
   */
  readonly bytes: Uint8Array;
  readonly from: LegacyOfficeFormat;
  readonly to: OoxmlFormat;
  /** Maximum number of generated package bytes the host will admit. */
  readonly maxOutputBytes: number;
  /** Aborted for caller cancellation or the configured conversion timeout. */
  readonly signal: AbortSignal;
}

/** Converter-owned diagnostics must never contain document content or source identifiers. */
export interface LegacyOfficeConversionResult {
  /**
   * Standalone OOXML package bytes. Ownership transfers to the caller when the
   * conversion promise resolves; the converter must not retain or mutate them.
   */
  readonly bytes: Uint8Array | ArrayBuffer;
  /** Stable engine identifier suitable for provenance records. */
  readonly engine?: string;
  /** Engine/build version suitable for provenance records. */
  readonly engineVersion?: string;
  /**
   * Converter-computed SHA-256 of `bytes`, encoded as 64 lowercase hexadecimal
   * characters. The host validates the representation but does not repeat the
   * potentially expensive full-output digest.
   */
  readonly outputSha256?: string;
  /** Content-free loss/compatibility diagnostics. */
  readonly warnings?: readonly string[];
}

/** Implementation-neutral, asynchronous legacy Office conversion contract. */
export interface LegacyOfficeConverter {
  convert(input: Readonly<LegacyOfficeConversionInput>): Promise<LegacyOfficeConversionResult>;
}

/** Content-free record emitted after validated conversion. */
export interface LegacyOfficeConversionRecord {
  readonly from: LegacyOfficeFormat;
  readonly to: OoxmlFormat;
  readonly inputBytes: number;
  readonly outputBytes: number;
  readonly engine?: string;
  readonly engineVersion?: string;
  /** Converter-computed digest; verify independently when the converter is outside the trust boundary. */
  readonly outputSha256?: string;
  readonly warnings?: readonly string[];
}

/** Converter and resource policy for one explicitly enabled legacy format. */
export interface LegacyOfficeFormatConversionOptions {
  /** Converter implementation for this legacy format. */
  readonly converter: LegacyOfficeConverter;
  /** Cooperative caller cancellation for conversion. */
  readonly signal?: AbortSignal;
  /** Conversion deadline in milliseconds. Default: 120000 (two minutes). */
  readonly timeoutMs?: number;
  /** Legacy source admission limit. Default: 256 MiB; hard ceiling: 1 GiB. */
  readonly maxInputBytes?: number;
  /** Generated OOXML admission limit. Default: 512 MiB; hard ceiling: 1 GiB. */
  readonly maxOutputBytes?: number;
  /** Receives one content-free record after output validation succeeds. */
  readonly onResult?: (result: Readonly<LegacyOfficeConversionRecord>) => void | Promise<void>;
}

/**
 * Per-format legacy conversion opt-ins. Each omitted format remains rejected,
 * even when another format uses the same converter implementation.
 */
export interface LegacyOfficeConversionOptions {
  readonly doc?: LegacyOfficeFormatConversionOptions;
  readonly xls?: LegacyOfficeFormatConversionOptions;
  readonly ppt?: LegacyOfficeFormatConversionOptions;
}

export interface NormalizedOfficeInput {
  /** OOXML ZIP bytes ready for the existing parser path. */
  readonly bytes: Uint8Array;
  /** Present only when a legacy conversion actually ran. */
  readonly conversion?: Readonly<LegacyOfficeConversionRecord>;
}

export const DEFAULT_LEGACY_CONVERSION_TIMEOUT_MS = 120_000;
export const DEFAULT_MAX_LEGACY_INPUT_BYTES = 256 * 1024 * 1024;
export const DEFAULT_MAX_CONVERTED_OOXML_BYTES = 512 * 1024 * 1024;
export const HARD_MAX_LEGACY_CONVERSION_BYTES = 1024 * 1024 * 1024;

const FAMILY: Readonly<Record<OoxmlFormat, LegacyOfficeFormat>> = {
  docx: 'doc',
  xlsx: 'xls',
  pptx: 'ppt',
};

const EXPECTED_MAIN_PART: Readonly<Record<OoxmlFormat, Readonly<{
  partName: string;
  contentType: string;
}>>> = {
  docx: {
    partName: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  },
  xlsx: {
    partName: '/xl/workbook.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  },
  pptx: {
    partName: '/ppt/presentation.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  },
};

const MAX_CONTENT_TYPES_BYTES = 1024 * 1024;
const MAX_VALIDATION_ENTRIES = 100_000;
const MAX_CENTRAL_DIRECTORY_BYTES = 64 * 1024 * 1024;
const MAX_TIMER_DELAY_MS = 0x7fffffff;
const EOCD_MIN_BYTES = 22;
const MAX_ZIP_COMMENT_BYTES = 0xffff;

/**
 * Normalize one browser/Node input into OOXML bytes. Existing ZIP and encrypted
 * OOXML behavior is unchanged. Only a classified legacy CFB and an explicitly
 * supplied converter enter the conversion path.
 */
export async function normalizeOfficeInput(
  bytes: Uint8Array | ArrayBuffer,
  target: OoxmlFormat,
  conversion?: LegacyOfficeConversionOptions,
  password?: string,
): Promise<NormalizedOfficeInput> {
  const inspected = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const from = FAMILY[target];
  const selected = conversion?.[from];
  if (sniffCfb(inspected) !== 'legacy-binary-format' || selected === undefined) {
    return { bytes: await resolveOoxmlContainer(inspected, password) };
  }

  const inputByteLength = inspected.byteLength;
  const classifiedFormat = sniffLegacyOfficeFormat(inspected);
  if (classifiedFormat !== null && classifiedFormat !== from) {
    throw new LegacyOfficeConversionError('unsupported-input', from, target);
  }
  const maxInputBytes = normalizeByteLimit(
    selected.maxInputBytes,
    DEFAULT_MAX_LEGACY_INPUT_BYTES,
    `legacyConversion.${from}.maxInputBytes`,
  );
  const maxOutputBytes = normalizeByteLimit(
    selected.maxOutputBytes,
    DEFAULT_MAX_CONVERTED_OOXML_BYTES,
    `legacyConversion.${from}.maxOutputBytes`,
  );
  const timeoutMs = normalizeTimeout(selected.timeoutMs, from);

  if (inputByteLength > maxInputBytes) {
    throw new LegacyOfficeConversionError('source-too-large', from, target);
  }
  if (selected.signal?.aborted) {
    throw new LegacyOfficeConversionError('aborted', from, target);
  }
  const input = exactBytes(inspected);

  const controller = new AbortController();
  let cancellationReason: 'aborted' | 'timeout' | undefined;
  let rejectCancellation: ((error: LegacyOfficeConversionError) => void) | undefined;
  const cancellation = new Promise<never>((_resolve, reject) => {
    rejectCancellation = reject;
  });
  const cancel = (reason: 'aborted' | 'timeout'): void => {
    if (cancellationReason !== undefined) return;
    cancellationReason = reason;
    controller.abort();
    rejectCancellation?.(new LegacyOfficeConversionError(reason, from, target));
  };
  const onAbort = (): void => cancel('aborted');
  selected.signal?.addEventListener('abort', onAbort, { once: true });
  const timer = setTimeout(() => cancel('timeout'), timeoutMs);

  try {
    let pending: Promise<LegacyOfficeConversionResult>;
    try {
      pending = Promise.resolve(selected.converter.convert({
        bytes: input,
        from,
        to: target,
        maxOutputBytes,
        signal: controller.signal,
      }));
    } catch (error) {
      pending = Promise.reject(error);
    }

    let converted: LegacyOfficeConversionResult;
    try {
      converted = await Promise.race([pending, cancellation]);
    } catch (error) {
      if (cancellationReason !== undefined) {
        throw new LegacyOfficeConversionError(cancellationReason, from, target);
      }
      if (error instanceof LegacyOfficeConversionError) {
        throw new LegacyOfficeConversionError(error.reason, from, target);
      }
      throw new LegacyOfficeConversionError('failed', from, target);
    }

    let output: Uint8Array;
    let diagnostics: Readonly<{
      engine?: string;
      engineVersion?: string;
      outputSha256?: string;
      warnings?: readonly string[];
    }>;
    try {
      output = conversionResultBytes(converted, from, target, maxOutputBytes);
      diagnostics = normalizeDiagnostics(converted);
    } catch (error) {
      if (error instanceof LegacyOfficeConversionError) throw error;
      throw new LegacyOfficeConversionError('invalid-output', from, target);
    }
    await validateConvertedOoxml(output, target);
    if (cancellationReason !== undefined) {
      throw new LegacyOfficeConversionError(cancellationReason, from, target);
    }

    const record = freezeConversionRecord({
      from,
      to: target,
      inputBytes: inputByteLength,
      outputBytes: output.byteLength,
      ...diagnostics,
    });
    notifyConversionObserver(selected.onResult, record);
    return { bytes: output, conversion: record };
  } finally {
    clearTimeout(timer);
    selected.signal?.removeEventListener('abort', onAbort);
  }
}

/**
 * Validate a converter-produced ZIP and its package content types before the
 * bytes can reach any format parser. This deliberately validates converted
 * output only, leaving the compatibility surface of ordinary OOXML loads
 * unchanged.
 */
export async function validateConvertedOoxml(
  bytes: Uint8Array | ArrayBuffer,
  target: OoxmlFormat,
): Promise<void> {
  const from = FAMILY[target];
  try {
    const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const manifest = inspectConvertedZip(input);
    const expectedMainEntry = asciiCaseFold(EXPECTED_MAIN_PART[target].partName.slice(1));
    if (!manifest.entryNames.has(expectedMainEntry)) throw new Error('missing main part');
    const contentTypesBytes = await inflateContentTypes(input, manifest.contentTypes);
    const xml = new TextDecoder('utf-8', { fatal: true }).decode(contentTypesBytes);
    validateContentTypesXml(xml, target);
  } catch (error) {
    if (error instanceof LegacyOfficeConversionError) throw error;
    throw new LegacyOfficeConversionError('invalid-output', from, target);
  }
}

interface ZipEntry {
  readonly flags: number;
  readonly method: number;
  readonly crc32: number;
  readonly compressedBytes: number;
  readonly uncompressedBytes: number;
  readonly localHeaderOffset: number;
  readonly name: string;
}

interface ConvertedZipManifest {
  readonly contentTypes: ZipEntry;
  readonly entryNames: ReadonlySet<string>;
}

/**
 * Apply the converter boundary's bounded OPC/ZIP preflight. ECMA-376 Part 2
 * §7.2.3.2 and §7.3.7 define the content-types stream, §7.3.6 forbids
 * ZIP encryption and compression methods other than DEFLATE, and Annex B
 * requires local and central header consistency.
 */
function inspectConvertedZip(bytes: Uint8Array): ConvertedZipManifest {
  if (bytes.byteLength < EOCD_MIN_BYTES) throw new Error('missing end record');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOffset = findEocd(view);
  const disk = view.getUint16(eocdOffset + 4, true);
  const centralDisk = view.getUint16(eocdOffset + 6, true);
  const entriesOnDisk = view.getUint16(eocdOffset + 8, true);
  const entryCount = view.getUint16(eocdOffset + 10, true);
  const centralBytes = view.getUint32(eocdOffset + 12, true);
  const centralOffset = view.getUint32(eocdOffset + 16, true);
  if (disk !== 0 || centralDisk !== 0 || entriesOnDisk !== entryCount) {
    throw new Error('multi-disk zip');
  }
  if (entryCount === 0xffff || centralBytes === 0xffffffff || centralOffset === 0xffffffff) {
    throw new Error('zip64 is not admitted by the pre-parser validator');
  }
  if (entryCount > MAX_VALIDATION_ENTRIES || centralBytes > MAX_CENTRAL_DIRECTORY_BYTES) {
    throw new Error('central directory budget exceeded');
  }
  if (!boundedRange(centralOffset, centralBytes, eocdOffset)) {
    throw new Error('central directory out of range');
  }

  let offset = centralOffset;
  let contentTypes: ZipEntry | undefined;
  const entryNames = new Set<string>();
  for (let index = 0; index < entryCount; index++) {
    if (!boundedRange(offset, 46, centralOffset + centralBytes)) throw new Error('short entry');
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('bad central signature');
    const flags = view.getUint16(offset + 8, true);
    const method = view.getUint16(offset + 10, true);
    const crc = view.getUint32(offset + 16, true);
    const compressedBytes = view.getUint32(offset + 20, true);
    const uncompressedBytes = view.getUint32(offset + 24, true);
    const nameBytes = view.getUint16(offset + 28, true);
    const extraBytes = view.getUint16(offset + 30, true);
    const commentBytes = view.getUint16(offset + 32, true);
    const startDisk = view.getUint16(offset + 34, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const entryBytes = 46 + nameBytes + extraBytes + commentBytes;
    if (!boundedRange(offset, entryBytes, centralOffset + centralBytes)) {
      throw new Error('entry out of range');
    }
    if (
      compressedBytes === 0xffffffff
      || uncompressedBytes === 0xffffffff
      || localHeaderOffset === 0xffffffff
      || startDisk !== 0
    ) throw new Error('zip64 or split entry');
    const name = decodeZipName(bytes.subarray(offset + 46, offset + 46 + nameBytes), flags);
    if ((flags & 0x0001) !== 0) throw new Error('encrypted zip entry');
    if (method !== 0 && method !== 8) throw new Error('unsupported compression method');
    validateOpcZipName(name);
    const equivalentName = asciiCaseFold(name);
    if (entryNames.has(equivalentName)) throw new Error('duplicate zip entry');
    entryNames.add(equivalentName);
    if (isActiveEntryName(name)) throw new Error('active content entry is not admitted');
    if (name === '[Content_Types].xml') {
      if (contentTypes !== undefined) throw new Error('duplicate content types');
      contentTypes = {
        flags,
        method,
        crc32: crc,
        compressedBytes,
        uncompressedBytes,
        localHeaderOffset,
        name,
      };
    }
    offset += entryBytes;
  }
  if (offset !== centralOffset + centralBytes || contentTypes === undefined) {
    throw new Error('inconsistent central directory');
  }
  return { contentTypes, entryNames };
}

function findEocd(view: DataView): number {
  const minimum = Math.max(0, view.byteLength - EOCD_MIN_BYTES - MAX_ZIP_COMMENT_BYTES);
  for (let offset = view.byteLength - EOCD_MIN_BYTES; offset >= minimum; offset--) {
    if (view.getUint32(offset, true) !== 0x06054b50) continue;
    const commentBytes = view.getUint16(offset + 20, true);
    if (offset + EOCD_MIN_BYTES + commentBytes === view.byteLength) return offset;
  }
  throw new Error('missing end record');
}

function decodeZipName(bytes: Uint8Array, flags: number): string {
  // The package-defined name is ASCII. UTF-8 names are accepted; legacy ZIP
  // encodings cannot spell this target differently and are not decoded here.
  if ((flags & 0x0800) === 0 && bytes.some((byte) => byte > 0x7f)) {
    throw new Error('unsupported entry-name encoding');
  }
  return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
}

function validateOpcZipName(name: string): void {
  if (name.length === 0 || name.startsWith('/') || name.includes('\\') || name.includes('\0')) {
    throw new Error('invalid OPC zip item name');
  }
  const segments = name.split('/');
  const finalIndex = segments.length - 1;
  if (segments.some((segment, index) => (
    segment === '.'
    || segment === '..'
    || (segment.length === 0 && index !== finalIndex)
  ))) throw new Error('invalid OPC zip item path');
}

function isActiveEntryName(name: string): boolean {
  return name.toLowerCase().split('/').some((segment) => (
    segment === 'activex'
    || segment.includes('vba')
  ));
}

async function inflateContentTypes(bytes: Uint8Array, entry: ZipEntry): Promise<Uint8Array> {
  if ((entry.flags & 0x0001) !== 0) throw new Error('encrypted zip entry');
  if (entry.uncompressedBytes > MAX_CONTENT_TYPES_BYTES) throw new Error('content types too large');
  if (!boundedRange(entry.localHeaderOffset, 30, bytes.byteLength)) {
    throw new Error('local header out of range');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const offset = entry.localHeaderOffset;
  if (view.getUint32(offset, true) !== 0x04034b50) throw new Error('bad local signature');
  const localFlags = view.getUint16(offset + 6, true);
  const localMethod = view.getUint16(offset + 8, true);
  const nameBytes = view.getUint16(offset + 26, true);
  const extraBytes = view.getUint16(offset + 28, true);
  if (localFlags !== entry.flags || localMethod !== entry.method) throw new Error('header mismatch');
  const dataOffset = offset + 30 + nameBytes + extraBytes;
  if (!boundedRange(dataOffset, entry.compressedBytes, bytes.byteLength)) {
    throw new Error('compressed data out of range');
  }
  const localName = decodeZipName(bytes.subarray(offset + 30, offset + 30 + nameBytes), localFlags);
  if (localName !== entry.name) throw new Error('local name mismatch');
  const compressed = bytes.subarray(dataOffset, dataOffset + entry.compressedBytes);

  let output: Uint8Array;
  if (entry.method === 0) {
    if (entry.compressedBytes !== entry.uncompressedBytes) throw new Error('stored size mismatch');
    output = compressed.slice();
  } else if (entry.method === 8) {
    output = await inflateRawBounded(compressed, entry.uncompressedBytes);
  } else {
    throw new Error('unsupported compression method');
  }
  if (crc32(output) !== entry.crc32) throw new Error('crc mismatch');
  return output;
}

async function inflateRawBounded(compressed: Uint8Array, expectedBytes: number): Promise<Uint8Array> {
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('deflate validator unavailable');
  }
  const source = new Blob([compressed as BlobPart]).stream();
  const reader = source.pipeThrough(new DecompressionStream('deflate-raw')).getReader();
  const output = new Uint8Array(expectedBytes);
  let offset = 0;
  try {
    while (true) {
      const chunk = await reader.read();
      if (chunk.done) break;
      if (offset + chunk.value.byteLength > expectedBytes) throw new Error('inflated size mismatch');
      output.set(chunk.value, offset);
      offset += chunk.value.byteLength;
    }
  } finally {
    reader.releaseLock();
  }
  if (offset !== expectedBytes) throw new Error('inflated size mismatch');
  return output;
}

function validateContentTypesXml(xml: string, target: OoxmlFormat): void {
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/u.test(xml)) {
    throw new Error('invalid XML character');
  }
  const tags = scanXmlStartTags(xml);
  const root = tags.shift();
  if (root === undefined || root.depth !== 0 || localXmlName(root.name) !== 'Types') {
    throw new Error('not content-types XML');
  }
  const rootPrefix = xmlPrefix(root.name);
  const namespaceAttribute = rootPrefix === '' ? 'xmlns' : `xmlns:${rootPrefix}`;
  if (root.attributes.get(namespaceAttribute) !== 'http://schemas.openxmlformats.org/package/2006/content-types') {
    throw new Error('wrong content-types namespace');
  }
  if ([...root.attributes.keys()].some((attribute) => (
    attribute !== 'xmlns' && !attribute.startsWith('xmlns:')
  ))) throw new Error('unexpected root attribute');
  const expected = EXPECTED_MAIN_PART[target];
  let foundExpectedMain = false;
  const allTypes: string[] = [];
  const overrides = new Set<string>();
  const defaults = new Set<string>();
  for (const tag of tags) {
    if (tag.depth !== 1) throw new Error('invalid content-types nesting');
    if (xmlPrefix(tag.name) !== rootPrefix) throw new Error('foreign namespace');
    const name = localXmlName(tag.name);
    if (name !== 'Override' && name !== 'Default') throw new Error('unexpected content-types element');
    const contentType = tag.attributes.get('ContentType');
    if (contentType === undefined || tag.attributes.size !== 2) throw new Error('invalid content type declaration');
    allTypes.push(asciiCaseFold(contentType));
    if (name === 'Override') {
      const partName = tag.attributes.get('PartName');
      const equivalentPartName = partName === undefined ? undefined : asciiCaseFold(partName);
      if (equivalentPartName === undefined || overrides.has(equivalentPartName)) {
        throw new Error('invalid override');
      }
      overrides.add(equivalentPartName);
      if (
        equivalentPartName === asciiCaseFold(expected.partName)
        && asciiCaseFold(contentType) === expected.contentType
      ) {
        foundExpectedMain = true;
      }
    } else {
      const extension = tag.attributes.get('Extension');
      const equivalentExtension = extension === undefined ? undefined : asciiCaseFold(extension);
      if (equivalentExtension === undefined || defaults.has(equivalentExtension)) {
        throw new Error('invalid default');
      }
      defaults.add(equivalentExtension);
    }
  }
  if (!foundExpectedMain) throw new Error('wrong package family');
  if (allTypes.some(isActiveContentType)) throw new Error('active content is not admitted');
}

interface ParsedXmlTag {
  readonly name: string;
  readonly attributes: ReadonlyMap<string, string>;
  readonly depth: number;
  readonly selfClosing: boolean;
}

function scanXmlStartTags(xml: string): ParsedXmlTag[] {
  const tags: ParsedXmlTag[] = [];
  const stack: string[] = [];
  let sawRoot = false;
  let closedRoot = false;
  let cursor = 0;
  while (cursor < xml.length) {
    const start = xml.indexOf('<', cursor);
    const textEnd = start < 0 ? xml.length : start;
    if (!/^\s*$/u.test(xml.slice(cursor, textEnd))) throw new Error('unexpected XML text');
    if (start < 0) {
      cursor = xml.length;
      break;
    }
    if (xml.startsWith('<!--', start)) {
      const end = xml.indexOf('-->', start + 4);
      if (end < 0) throw new Error('unterminated XML comment');
      cursor = end + 3;
      continue;
    }
    if (xml.startsWith('<?', start)) {
      const end = xml.indexOf('?>', start + 2);
      if (end < 0) throw new Error('unterminated processing instruction');
      cursor = end + 2;
      continue;
    }
    if (xml.startsWith('<!', start)) throw new Error('DTD/CDATA is not admitted');
    const end = findXmlTagEnd(xml, start + 1);
    const body = xml.slice(start + 1, end);
    cursor = end + 1;
    if (/^\s*\//u.test(body)) {
      const closingName = parseXmlClosingTag(body);
      if (stack.pop() !== closingName) throw new Error('mismatched XML closing tag');
      if (stack.length === 0) closedRoot = true;
      continue;
    }
    if (closedRoot) throw new Error('multiple XML roots');
    const parsed = parseXmlStartTag(body);
    const depth = stack.length;
    if (depth === 0) {
      if (sawRoot) throw new Error('multiple XML roots');
      sawRoot = true;
    }
    tags.push({ ...parsed, depth });
    if (parsed.selfClosing) {
      if (depth === 0) closedRoot = true;
    } else {
      stack.push(parsed.name);
    }
  }
  if (!sawRoot || stack.length !== 0 || !closedRoot) throw new Error('incomplete XML document');
  return tags;
}

function findXmlTagEnd(xml: string, start: number): number {
  let quote = '';
  for (let index = start; index < xml.length; index++) {
    const char = xml[index] as string;
    if (quote !== '') {
      if (char === quote) quote = '';
      continue;
    }
    if (char === '"' || char === "'") quote = char;
    else if (char === '>') return index;
  }
  throw new Error('unterminated XML tag');
}

function parseXmlStartTag(body: string): Omit<ParsedXmlTag, 'depth'> {
  let cursor = 0;
  cursor = skipXmlSpace(body, cursor);
  const nameStart = cursor;
  while (cursor < body.length && /[A-Za-z0-9_.:-]/.test(body[cursor] as string)) cursor++;
  const name = body.slice(nameStart, cursor);
  if (!/^[A-Za-z_][A-Za-z0-9_.:-]*$/.test(name)) throw new Error('invalid XML name');
  const attributes = new Map<string, string>();
  let selfClosing = false;
  while (true) {
    cursor = skipXmlSpace(body, cursor);
    if (cursor === body.length) break;
    if (body[cursor] === '/' && cursor + 1 === body.length) {
      selfClosing = true;
      break;
    }
    const attrStart = cursor;
    while (cursor < body.length && /[A-Za-z0-9_.:-]/.test(body[cursor] as string)) cursor++;
    const attribute = body.slice(attrStart, cursor);
    if (!/^[A-Za-z_][A-Za-z0-9_.:-]*$/.test(attribute) || attributes.has(attribute)) {
      throw new Error('invalid XML attribute');
    }
    cursor = skipXmlSpace(body, cursor);
    if (body[cursor] !== '=') throw new Error('missing attribute value');
    cursor = skipXmlSpace(body, cursor + 1);
    const quote = body[cursor];
    if (quote !== '"' && quote !== "'") throw new Error('unquoted attribute');
    const valueStart = ++cursor;
    const valueEnd = body.indexOf(quote, valueStart);
    if (valueEnd < 0 || body.slice(valueStart, valueEnd).includes('<')) {
      throw new Error('invalid attribute value');
    }
    attributes.set(attribute, decodeXmlAttributeValue(body.slice(valueStart, valueEnd)));
    cursor = valueEnd + 1;
  }
  return { name, attributes, selfClosing };
}

function parseXmlClosingTag(body: string): string {
  const match = /^\s*\/\s*([A-Za-z_][A-Za-z0-9_.:-]*)\s*$/u.exec(body);
  if (match === null) throw new Error('invalid XML closing tag');
  return match[1] as string;
}

function skipXmlSpace(value: string, start: number): number {
  let cursor = start;
  while (cursor < value.length && /[\t\n\r ]/.test(value[cursor] as string)) cursor++;
  return cursor;
}

function localXmlName(name: string): string {
  const colon = name.indexOf(':');
  return colon < 0 ? name : name.slice(colon + 1);
}

function xmlPrefix(name: string): string {
  const colon = name.indexOf(':');
  return colon < 0 ? '' : name.slice(0, colon);
}

function decodeXmlAttributeValue(value: string): string {
  let output = '';
  let cursor = 0;
  while (cursor < value.length) {
    const amp = value.indexOf('&', cursor);
    if (amp < 0) return output + value.slice(cursor);
    output += value.slice(cursor, amp);
    const semicolon = value.indexOf(';', amp + 1);
    if (semicolon < 0) throw new Error('unterminated XML entity');
    const entity = value.slice(amp + 1, semicolon);
    const named: Readonly<Record<string, string>> = {
      amp: '&',
      apos: "'",
      gt: '>',
      lt: '<',
      quot: '"',
    };
    const namedValue = named[entity];
    if (namedValue !== undefined) {
      output += namedValue;
    } else {
      const hex = entity.startsWith('#x') ? entity.slice(2) : undefined;
      const decimal = entity.startsWith('#') && hex === undefined ? entity.slice(1) : undefined;
      const digits = hex ?? decimal;
      const radix = hex === undefined ? 10 : 16;
      if (digits === undefined || !new RegExp(radix === 16 ? '^[0-9A-Fa-f]+$' : '^[0-9]+$').test(digits)) {
        throw new Error('unsupported XML entity');
      }
      const codePoint = Number.parseInt(digits, radix);
      if (
        !Number.isInteger(codePoint)
        || !isXml10CodePoint(codePoint)
      ) throw new Error('invalid XML character reference');
      output += String.fromCodePoint(codePoint);
    }
    cursor = semicolon + 1;
  }
  return output;
}

function isXml10CodePoint(codePoint: number): boolean {
  return codePoint === 0x09
    || codePoint === 0x0a
    || codePoint === 0x0d
    || (codePoint >= 0x20 && codePoint <= 0xd7ff)
    || (codePoint >= 0xe000 && codePoint <= 0xfffd)
    || (codePoint >= 0x10000 && codePoint <= 0x10ffff);
}

function asciiCaseFold(value: string): string {
  return value.replace(/[A-Z]/g, (character) => character.toLowerCase());
}

function isActiveContentType(value: string): boolean {
  return value.includes('macro')
    || value.includes('vba')
    || value.includes('activex');
}

function exactBytes(bytes: Uint8Array | ArrayBuffer): Uint8Array {
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
  if (
    bytes.byteOffset === 0
    && bytes.byteLength === bytes.buffer.byteLength
    && bytes.buffer instanceof ArrayBuffer
  ) return bytes;
  return new Uint8Array(toArrayBuffer(bytes));
}

function conversionResultBytes(
  result: LegacyOfficeConversionResult,
  from: LegacyOfficeFormat,
  to: OoxmlFormat,
  maxOutputBytes: number,
): Uint8Array {
  if (result === null || typeof result !== 'object') {
    throw new LegacyOfficeConversionError('invalid-output', from, to);
  }
  const bytes = result.bytes;
  if (!(bytes instanceof Uint8Array) && !(bytes instanceof ArrayBuffer)) {
    throw new LegacyOfficeConversionError('invalid-output', from, to);
  }
  if (bytes.byteLength > maxOutputBytes) {
    throw new LegacyOfficeConversionError('output-too-large', from, to);
  }
  return exactBytes(bytes);
}

function normalizeDiagnostics(result: LegacyOfficeConversionResult): Readonly<{
  engine?: string;
  engineVersion?: string;
  outputSha256?: string;
  warnings?: readonly string[];
}> {
  const engine = normalizeDiagnosticText(result.engine, 'engine', 128);
  const engineVersion = normalizeDiagnosticText(result.engineVersion, 'engineVersion', 128);
  const outputSha256 = result.outputSha256;
  if (outputSha256 !== undefined && !/^[0-9a-f]{64}$/u.test(outputSha256)) {
    throw new Error('invalid output SHA-256');
  }
  const sourceWarnings = result.warnings;
  let warnings: readonly string[] | undefined;
  if (sourceWarnings !== undefined) {
    if (!Array.isArray(sourceWarnings) || sourceWarnings.length > 64) {
      throw new Error('invalid warnings');
    }
    warnings = sourceWarnings.map((warning) => {
      const normalized = normalizeDiagnosticText(warning, 'warning', 512);
      if (normalized === undefined) throw new Error('invalid warning');
      return normalized;
    });
  }
  return {
    ...(engine === undefined ? {} : { engine }),
    ...(engineVersion === undefined ? {} : { engineVersion }),
    ...(outputSha256 === undefined ? {} : { outputSha256 }),
    ...(warnings === undefined ? {} : { warnings }),
  };
}

function normalizeDiagnosticText(
  value: unknown,
  _name: string,
  maxLength: number,
): string | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value !== 'string'
    || value.length === 0
    || value.length > maxLength
    || /[\u0000-\u001f\u007f]/.test(value)
  ) throw new Error('invalid diagnostic text');
  return value;
}

function normalizeByteLimit(value: number | undefined, fallback: number, name: string): number {
  const resolved = value ?? fallback;
  if (!Number.isSafeInteger(resolved) || resolved <= 0 || resolved > HARD_MAX_LEGACY_CONVERSION_BYTES) {
    throw new TypeError(`${name} must be a positive safe integer no greater than 1 GiB`);
  }
  return resolved;
}

function normalizeTimeout(value: number | undefined, format: LegacyOfficeFormat): number {
  const resolved = value ?? DEFAULT_LEGACY_CONVERSION_TIMEOUT_MS;
  if (!Number.isSafeInteger(resolved) || resolved <= 0 || resolved > MAX_TIMER_DELAY_MS) {
    throw new TypeError(`legacyConversion.${format}.timeoutMs must be a positive integer no greater than 2147483647`);
  }
  return resolved;
}

function freezeConversionRecord(
  source: LegacyOfficeConversionRecord,
): Readonly<LegacyOfficeConversionRecord> {
  const warnings = source.warnings === undefined ? undefined : Object.freeze([...source.warnings]);
  return Object.freeze({
    from: source.from,
    to: source.to,
    inputBytes: source.inputBytes,
    outputBytes: source.outputBytes,
    ...(source.engine === undefined ? {} : { engine: source.engine }),
    ...(source.engineVersion === undefined ? {} : { engineVersion: source.engineVersion }),
    ...(source.outputSha256 === undefined ? {} : { outputSha256: source.outputSha256 }),
    ...(warnings === undefined ? {} : { warnings }),
  });
}

function notifyConversionObserver(
  observer: LegacyOfficeFormatConversionOptions['onResult'],
  record: Readonly<LegacyOfficeConversionRecord>,
): void {
  if (observer === undefined) return;
  try {
    void Promise.resolve(observer(record)).catch(() => {});
  } catch {
    // Diagnostics are observational and cannot change a successful load.
  }
}

function boundedRange(offset: number, length: number, upperBound: number): boolean {
  return Number.isSafeInteger(offset)
    && Number.isSafeInteger(length)
    && offset >= 0
    && length >= 0
    && offset <= upperBound
    && length <= upperBound - offset;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}
