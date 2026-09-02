import { describe, expect, it } from 'vitest';
import { OoxmlError, OoxmlResourceLimitError } from '../errors/ooxml-error.js';
import {
  OoxmlDecodedImageLimitError,
  isOoxmlDecodedImageLimitError,
  type OoxmlDecodedImageLimitMetric,
} from '../image/pixel-budget.js';
import { isTiffDecodeError, TiffDecodeError } from '../image/tiff-contract.js';
import {
  deserializeWorkerError,
  parseResourceLimitError,
  serializeWorkerError,
  type WorkerErrorPayload,
} from './error-wire.js';

const USAGE = {
  archiveEntryCount: 1,
  declaredInflatedBytes: 6,
  largestInflatedEntryBytes: 6,
  distinctInflatedBytes: 6,
  operationInflatedBytes: 6,
};

describe('worker error wire', () => {
  const rustError =
    'OOXML_RESOURCE_LIMIT:{"code":"ooxml-resource-limit","details":{"stage":"decompression","violation":{"format":"xlsx","operation":"parse","resource":"archive-entry","part":"xl/worksheets/sheet1.xml","metric":"actual-inflated-bytes","limit":5,"observed":6,"configurable":true,"usage":{"archiveEntryCount":1,"declaredInflatedBytes":6,"largestInflatedEntryBytes":6,"distinctInflatedBytes":6,"operationInflatedBytes":6}}}}';

  it('parses an exact Rust envelope into the public typed error', () => {
    const error = parseResourceLimitError(new Error(rustError));
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    expect(error).toMatchObject({ code: 'ooxml-resource-limit' });
    expect(error?.details.violation).toMatchObject({
      format: 'xlsx',
      operation: 'parse',
      resource: 'archive-entry',
      part: 'xl/worksheets/sheet1.xml',
      metric: 'actual-inflated-bytes',
      limit: 5,
      observed: 6,
      configurable: true,
      usage: expect.objectContaining({ largestInflatedEntryBytes: 6 }),
    });
  });

  it('survives worker serialization and structured clone as a real typed error', () => {
    const wire = structuredClone(serializeWorkerError(new Error(rustError)));
    const error = deserializeWorkerError(wire);
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    expect((error as OoxmlResourceLimitError).details.violation).toMatchObject({
      resource: 'archive-entry',
      part: 'xl/worksheets/sheet1.xml',
    });
  });

  it.each([
    'image-dimension',
    'image-pixels',
    'active-decoded-bytes',
  ] satisfies OoxmlDecodedImageLimitMetric[])(
    'preserves the %s decoded-image quota error across the worker boundary',
    (metric) => {
      const original = new OoxmlDecodedImageLimitError(
        metric,
        128 * 1024 * 1024,
        192 * 1024 * 1024,
      );
      const restored = deserializeWorkerError(structuredClone(serializeWorkerError(original)));

      expect(restored).toBeInstanceOf(OoxmlDecodedImageLimitError);
      expect(restored).toMatchObject({
        code: 'ooxml-decoded-image-limit',
        metric,
        limit: 128 * 1024 * 1024,
        observed: 192 * 1024 * 1024,
      });
    },
  );

  it('preserves a decoded-image quota error from another bundled core realm', () => {
    const foreignQuota = {
      name: 'OoxmlDecodedImageLimitError',
      message: 'OOXML decoded image limit exceeded: image-pixels 11 > 10',
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 10,
      observed: 11,
    };
    expect(isOoxmlDecodedImageLimitError(foreignQuota)).toBe(true);
    const restored = deserializeWorkerError(structuredClone(serializeWorkerError(foreignQuota)));

    expect(restored).toBeInstanceOf(OoxmlDecodedImageLimitError);
    expect(restored).toMatchObject(foreignQuota);
  });

  it('uses canonical decoded-image text instead of caller-controlled fields', () => {
    const foreignQuota = {
      name: () => 'uncloneable name',
      message: () => 'uncloneable message',
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 10,
      observed: 11,
    };

    const wire = serializeWorkerError(foreignQuota);
    expect(wire).toEqual({
      message: 'OOXML decoded image limit exceeded: image-pixels 11 > 10',
      errorName: 'OoxmlDecodedImageLimitError',
      code: 'ooxml-decoded-image-limit',
      decodedImage: { metric: 'image-pixels', limit: 10, observed: 11 },
    });
    expect(isOoxmlDecodedImageLimitError(foreignQuota)).toBe(false);
    expect(() => structuredClone(wire)).not.toThrow();
  });

  it.each([
    ['function-valued metric', { metric: () => 'image-pixels' }],
    ['unknown metric', { metric: 'future-metric' }],
    ['fractional limit', { limit: 10.5 }],
    ['negative limit', { limit: -1 }],
    ['unsafe observed value', { observed: Number.MAX_SAFE_INTEGER + 1 }],
    ['non-crossing values', { limit: 11, observed: 11 }],
  ])('rejects decoded-image quota shapes with %s', (_name, overrides) => {
    const candidate = {
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 10,
      observed: 11,
      ...overrides,
    };

    expect(isOoxmlDecodedImageLimitError(candidate)).toBe(false);
    const wire = serializeWorkerError(candidate);
    expect(wire).not.toHaveProperty('decodedImage');
    expect(() => structuredClone(wire)).not.toThrow();
    expect(deserializeWorkerError(wire)).not.toBeInstanceOf(OoxmlDecodedImageLimitError);
  });

  it('contains throwing decoded-image accessors and proxies', () => {
    const throwingAccessor = {
      code: 'ooxml-decoded-image-limit',
      get metric(): never {
        throw new Error('must not escape validation');
      },
      limit: 10,
      observed: 11,
    };
    const throwingProxy = new Proxy({}, {
      get(): never {
        throw new Error('must not escape validation');
      },
    });

    for (const candidate of [throwingAccessor, throwingProxy]) {
      expect(() => isOoxmlDecodedImageLimitError(candidate)).not.toThrow();
      expect(isOoxmlDecodedImageLimitError(candidate)).toBe(false);
      const wire = serializeWorkerError(candidate);
      expect(() => structuredClone(wire)).not.toThrow();
      expect(deserializeWorkerError(wire)).not.toBeInstanceOf(OoxmlDecodedImageLimitError);
    }
  });

  it('snapshots mutable decoded-image worker fields exactly once', () => {
    const reads = { decodedImage: 0, metric: 0, limit: 0, observed: 0 };
    const decodedImage = {
      get metric(): string {
        reads.metric++;
        return reads.metric === 1 ? 'image-pixels' : 'future-metric';
      },
      get limit(): number {
        reads.limit++;
        return reads.limit === 1 ? 10 : -1;
      },
      get observed(): number {
        reads.observed++;
        return reads.observed === 1 ? 11 : Number.NaN;
      },
    };
    const payload = {
      message: 'caller-controlled text',
      code: 'ooxml-decoded-image-limit',
      get decodedImage(): typeof decodedImage {
        reads.decodedImage++;
        return decodedImage;
      },
    } as WorkerErrorPayload;

    const restored = deserializeWorkerError(payload);
    expect(restored).toBeInstanceOf(OoxmlDecodedImageLimitError);
    expect(restored).toMatchObject({ metric: 'image-pixels', limit: 10, observed: 11 });
    expect(reads).toEqual({ decodedImage: 1, metric: 1, limit: 1, observed: 1 });
  });

  it('contains throwing decoded-image worker payload getters', () => {
    const payload = new Proxy({} as WorkerErrorPayload, {
      get(): never {
        throw new Error('must not escape deserialization');
      },
    });
    const nested = {
      message: 'quota',
      code: 'ooxml-decoded-image-limit',
      decodedImage: new Proxy({}, {
        get(): never {
          throw new Error('must not escape nested deserialization');
        },
      }),
    } as WorkerErrorPayload;

    for (const candidate of [payload, nested]) {
      expect(() => deserializeWorkerError(candidate)).not.toThrow();
      expect(deserializeWorkerError(candidate)).not.toBeInstanceOf(OoxmlDecodedImageLimitError);
    }
  });

  it('rejects an unknown decoded-image metric as an ordinary error', () => {
    const wire = structuredClone(serializeWorkerError(
      new OoxmlDecodedImageLimitError('image-pixels', 10, 11),
    ));
    if (wire.decodedImage) Object.assign(wire.decodedImage, { metric: 'future-metric' });

    expect(deserializeWorkerError(wire)).not.toBeInstanceOf(OoxmlDecodedImageLimitError);
  });

  it('preserves a diagnostic TIFF failure across the worker boundary', () => {
    const original = new TiffDecodeError('Unsupported TIFF compression: 5');
    const restored = deserializeWorkerError(structuredClone(serializeWorkerError(original)));

    expect(restored).toBeInstanceOf(TiffDecodeError);
    expect(isTiffDecodeError(restored)).toBe(true);
    expect(restored).toMatchObject({
      name: 'TiffDecodeError',
      code: 'ooxml-tiff-decode',
      message: 'Unsupported TIFF compression: 5',
    });
  });

  it('preserves a foreign-realm TIFF failure across the worker boundary', () => {
    const foreignTiff = {
      name: 'TiffDecodeError',
      message: 'Unsupported TIFF compression: 5',
      code: 'ooxml-tiff-decode',
    };
    expect(isTiffDecodeError(foreignTiff)).toBe(true);

    const wire = structuredClone(serializeWorkerError(foreignTiff));
    expect(wire).toEqual({
      message: foreignTiff.message,
      errorName: 'TiffDecodeError',
      code: 'ooxml-tiff-decode',
    });
    expect(deserializeWorkerError(wire)).toBeInstanceOf(TiffDecodeError);
  });

  it('contains throwing TIFF error accessors and proxies', () => {
    const throwingAccessor = Object.defineProperty({}, 'code', {
      get(): never {
        throw new Error('must not escape TIFF validation');
      },
    });
    const throwingProxy = new Proxy({}, {
      getPrototypeOf(): never {
        throw new Error('must not escape TIFF validation');
      },
    });
    const throwingMessage = {
      name: 'TiffDecodeError',
      code: 'ooxml-tiff-decode',
      get message(): never {
        throw new Error('must not escape TIFF message validation');
      },
    };

    for (const candidate of [throwingAccessor, throwingProxy, throwingMessage]) {
      expect(() => isTiffDecodeError(candidate)).not.toThrow();
      expect(isTiffDecodeError(candidate)).toBe(false);
      const wire = serializeWorkerError(candidate);
      expect(() => structuredClone(wire)).not.toThrow();
      expect(deserializeWorkerError(wire)).not.toBeInstanceOf(TiffDecodeError);
    }
  });

  it('encodes public errors field-by-field before structured clone', () => {
    const parsed = parseResourceLimitError(new Error(rustError))!;
    const error = new OoxmlResourceLimitError('unsafe extras', {
      ...parsed.details,
      violation: {
        ...parsed.details.violation,
        unsafeFunction: () => undefined,
        unsafeObject: { callback: () => undefined },
      } as never,
    });
    const wire = serializeWorkerError(error);

    expect(() => structuredClone(error.details)).not.toThrow();
    expect(() => structuredClone(wire)).not.toThrow();
    expect(wire.resourceLimit?.violation).not.toHaveProperty('unsafeFunction');
    expect(wire.resourceLimit?.violation).not.toHaveProperty('unsafeObject');
  });

  it('drops non-string fields from ordinary errors before structured clone', () => {
    const error = new Error('unsafe code') as Error & { code: unknown };
    error.code = { callback: () => undefined };
    const wire = serializeWorkerError(error);

    expect(wire).not.toHaveProperty('code');
    expect(() => structuredClone(wire)).not.toThrow();
  });

  it('contains caller-defined throwing error accessors', () => {
    const error = new Error('hidden');
    Object.defineProperty(error, 'code', {
      get() {
        throw new Error('must not escape serialization');
      },
    });

    expect(serializeWorkerError(error)).toEqual({
      message: 'Worker operation failed with an unreadable error',
      errorName: 'Error',
    });
  });

  it('sanitizes mutated typed OOXML errors before structured clone', () => {
    const error = new OoxmlError('encrypted', 'unsafe typed error');
    Object.assign(error, {
      name: { callback: () => undefined },
      code: { callback: () => undefined },
    });
    const wire = serializeWorkerError(error);

    expect(wire).toEqual({
      message: 'unsafe typed error',
      errorName: 'OoxmlError',
    });
    expect(() => structuredClone(wire)).not.toThrow();
  });

  it('downgrades invalid public resource-limit details without cloning unsafe fields', () => {
    const error = new OoxmlResourceLimitError('unsafe details', {
      stage: 'parsing',
      violation: {
        format: 'xlsx',
        operation: 'parse',
        resource: (() => 'worksheet-row') as never,
        metric: 'projected-bytes',
        limit: 1,
        observed: 2,
        configurable: false,
        usage: USAGE,
      },
    });
    const wire = serializeWorkerError(error);

    expect(wire).toEqual({
      message: 'Invalid OOXML resource-limit error payload',
      errorName: 'Error',
    });
    expect(() => structuredClone(wire)).not.toThrow();
  });

  it('downgrades a resource-limit error whose public fields were mutated at runtime', () => {
    const parsed = parseResourceLimitError(new Error(rustError))!;
    Object.assign(parsed, {
      code: { callback: () => undefined },
      details: null,
    });
    const wire = serializeWorkerError(parsed);

    expect(wire).toEqual({
      message: 'Invalid OOXML resource-limit error payload',
      errorName: 'Error',
    });
    expect(() => structuredClone(wire)).not.toThrow();
  });

  it.each([false, true])(
    'represents a configurable=%s archive entry-count metric without a dummy part',
    (configurable) => {
    const error = new OoxmlResourceLimitError('too many entries', {
      stage: 'container',
      violation: {
        format: 'docx',
        operation: 'open',
        resource: 'archive',
        metric: 'entry-count',
        limit: 20_000,
        observed: 20_001,
        configurable,
        usage: { ...USAGE, archiveEntryCount: 20_001 },
      },
    });
    const restored = deserializeWorkerError(serializeWorkerError(error));
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details.violation).not.toHaveProperty('part');
    },
  );

  it('preserves the hard central-directory metadata limit', () => {
    const error = new OoxmlResourceLimitError('central directory is too large', {
      stage: 'container',
      violation: {
        format: 'xlsx',
        operation: 'open',
        resource: 'archive',
        metric: 'central-directory-bytes',
        limit: 16 * 1024 * 1024,
        observed: 16 * 1024 * 1024 + 1,
        configurable: false,
        usage: USAGE,
      },
    });
    const restored = deserializeWorkerError(serializeWorkerError(error));
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details.violation).toMatchObject({
      resource: 'archive',
      metric: 'central-directory-bytes',
      configurable: false,
    });
    expect((restored as OoxmlResourceLimitError).details.violation).not.toHaveProperty('part');
  });

  it('preserves non-configurable parser-buffer limits and their part context', () => {
    const error = new OoxmlResourceLimitError('worksheet row is too large', {
      stage: 'parsing',
      violation: {
        format: 'xlsx',
        operation: 'parse-sheet',
        resource: 'worksheet-row',
        metric: 'projected-bytes',
        part: 'xl/worksheets/sheet1.xml',
        limit: 8 * 1024 * 1024,
        observed: 8 * 1024 * 1024 + 1,
        configurable: false,
        usage: USAGE,
      },
    });
    const restored = deserializeWorkerError(serializeWorkerError(error));
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details).toMatchObject({
      stage: 'parsing',
      violation: {
        resource: 'worksheet-row',
        metric: 'projected-bytes',
        part: 'xl/worksheets/sheet1.xml',
        configurable: false,
      },
    });
  });

  it('preserves a non-configurable hard per-entry violation', () => {
    const hardError = rustError.replace('"configurable":true', '"configurable":false');
    const error = parseResourceLimitError(hardError);
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    expect(error?.details.violation.configurable).toBe(false);
  });

  it.each([new TypeError('bad input'), new RangeError('bad range')])(
    'preserves %s across the worker wire',
    (original) => {
      const error = deserializeWorkerError(serializeWorkerError(original));
      expect(error).toBeInstanceOf(original.constructor);
      expect(error.message).toBe(original.message);
    },
  );

  it('does not classify malformed or wrapped envelopes as resource limits', () => {
    expect(parseResourceLimitError('OOXML_RESOURCE_LIMIT:{"code":"wrong"}')).toBeUndefined();
    expect(parseResourceLimitError(`xlsx parser: ${rustError}`)).toBeUndefined();
  });

  it('does not echo an invalid Rust resource payload across the worker boundary', () => {
    const invalid = rustError.replace(
      'xl/worksheets/sheet1.xml',
      'file:///Users/private/document.xml',
    );
    expect(serializeWorkerError(new Error(invalid))).toEqual({
      message: 'Invalid OOXML resource-limit payload',
      errorName: 'Error',
    });
  });

  it('rejects invalid discriminants at the worker boundary', () => {
    expect(
      deserializeWorkerError({
        message: 'invalid details',
        code: 'ooxml-resource-limit',
        resourceLimit: {
          stage: 'container',
          violation: {
            format: 'pptx',
            operation: 'open',
            resource: 'archive-entry',
            metric: 'actual-inflated-bytes',
            part: 'ppt/slides/slide1.xml',
            limit: 5,
            observed: 6,
            configurable: true,
            usage: USAGE,
          },
        },
      }),
    ).not.toBeInstanceOf(OoxmlResourceLimitError);

    const wrongParserStage = new OoxmlResourceLimitError('wrong stage', {
      stage: 'parsing',
      violation: {
        format: 'xlsx',
        operation: 'parse-sheet',
        resource: 'xml-event',
        metric: 'bytes',
        limit: 1,
        observed: 2,
        configurable: false,
        usage: USAGE,
      },
    });
    const payload = structuredClone(serializeWorkerError(wrongParserStage));
    if (payload.resourceLimit) Object.assign(payload.resourceLimit, { stage: 'container' });
    expect(deserializeWorkerError(payload)).not.toBeInstanceOf(OoxmlResourceLimitError);

    const missingRequiredPart = structuredClone(serializeWorkerError(new Error(rustError)));
    if (missingRequiredPart.resourceLimit) {
      delete (missingRequiredPart.resourceLimit.violation as { part?: string }).part;
    }
    expect(deserializeWorkerError(missingRequiredPart))
      .not.toBeInstanceOf(OoxmlResourceLimitError);
  });

  it('preserves a valid future resource/metric pair without a host release', () => {
    const future = structuredClone(serializeWorkerError(new Error(rustError)));
    if (future.resourceLimit) {
      Object.assign(future.resourceLimit, { stage: 'serialization' });
      Object.assign(future.resourceLimit.violation, {
        resource: 'slide-wire',
        metric: 'retained-bytes',
        configurable: false,
      });
      delete (future.resourceLimit.violation as { part?: string }).part;
    }
    const restored = deserializeWorkerError(future);
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details).toMatchObject({
      stage: 'serialization',
      violation: { resource: 'slide-wire', metric: 'retained-bytes' },
    });
  });

  it('rejects invalid permutations of resource and metric names already known to the host', () => {
    const invalid = structuredClone(serializeWorkerError(new Error(rustError)));
    if (invalid.resourceLimit) {
      Object.assign(invalid.resourceLimit, { stage: 'container' });
      Object.assign(invalid.resourceLimit.violation, {
        resource: 'archive-entry',
        metric: 'entry-count',
      });
      delete (invalid.resourceLimit.violation as { part?: string }).part;
    }

    expect(deserializeWorkerError(invalid)).not.toBeInstanceOf(OoxmlResourceLimitError);
  });

  it('rejects fractional counters and unsafe strings', () => {
    const fractional = structuredClone(serializeWorkerError(new Error(rustError)));
    if (fractional.resourceLimit) {
      Object.assign(fractional.resourceLimit.violation, { limit: 5.5 });
    }
    expect(deserializeWorkerError(fractional)).not.toBeInstanceOf(OoxmlResourceLimitError);

    const unsafe = structuredClone(serializeWorkerError(new Error(rustError)));
    if (unsafe.resourceLimit) {
      Object.assign(unsafe.resourceLimit.violation, { resource: 'bad\nresource' });
    }
    expect(deserializeWorkerError(unsafe)).not.toBeInstanceOf(OoxmlResourceLimitError);

    const unsafePart = structuredClone(serializeWorkerError(new Error(rustError)));
    if (unsafePart.resourceLimit) {
      Object.assign(unsafePart.resourceLimit.violation, {
        part: '../../private/document.xml',
      });
    }
    expect(deserializeWorkerError(unsafePart)).not.toBeInstanceOf(OoxmlResourceLimitError);
  });

  it('accepts the Rust wire-safe redaction token as a typed required part', () => {
    const redacted = structuredClone(serializeWorkerError(new Error(rustError)));
    if (redacted.resourceLimit) {
      Object.assign(redacted.resourceLimit.violation, {
        part: 'untrusted-archive-entry',
      });
    }
    const restored = deserializeWorkerError(redacted);
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details.violation.part)
      .toBe('untrusted-archive-entry');
  });
});
