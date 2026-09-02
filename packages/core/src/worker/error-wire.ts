import {
  OoxmlError,
  OoxmlResourceLimitError,
  type OoxmlErrorCode,
  type OoxmlErrorStage,
  type OoxmlFormat,
  type OoxmlResourceLimitErrorDetails,
  type OoxmlResourceUsageSnapshot,
  type OoxmlResourceViolation,
} from '../errors/ooxml-error.js';
import {
  OoxmlDecodedImageLimitError,
  getOoxmlDecodedImageLimitDetails,
  type OoxmlDecodedImageLimitMetric,
} from '../image/pixel-budget.js';
import { TiffDecodeError, getTiffDecodeErrorDetails } from '../image/tiff-contract.js';
import {
  PullSessionInsufficientCreditError,
  isPullSessionInsufficientCreditDetails,
  parsePullSessionInsufficientCreditError,
  type PullSessionInsufficientCreditDetails,
} from './pull-credit-error.js';

const RESOURCE_LIMIT_PREFIX = 'OOXML_RESOURCE_LIMIT:';
const MAX_IDENTIFIER_LENGTH = 128;
const MAX_OPERATION_LENGTH = 256;
const MAX_PART_LENGTH = 4_096;

/** Structured-clone-safe error payload shared by all OOXML workers. */
export interface WorkerErrorPayload {
  message: string;
  errorName?: string;
  code?: string;
  resourceLimit?: OoxmlResourceLimitErrorDetails;
  insufficientCredit?: PullSessionInsufficientCreditDetails;
  decodedImage?: {
    metric: OoxmlDecodedImageLimitMetric;
    limit: number;
    observed: number;
  };
}

interface RustResourceLimitPayload {
  code: 'ooxml-resource-limit';
  details: OoxmlResourceLimitErrorDetails;
}

function isNonNegativeSafeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function isUsage(value: unknown): value is OoxmlResourceUsageSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const usage = value as Partial<OoxmlResourceUsageSnapshot>;
  return (
    isNonNegativeSafeInteger(usage.archiveEntryCount) &&
    isNonNegativeSafeInteger(usage.declaredInflatedBytes) &&
    (usage.largestInflatedEntryBytes === undefined ||
      isNonNegativeSafeInteger(usage.largestInflatedEntryBytes)) &&
    isNonNegativeSafeInteger(usage.distinctInflatedBytes) &&
    isNonNegativeSafeInteger(usage.operationInflatedBytes)
  );
}

/** Decode one canonical Rust usage checkpoint shared by all format workers. */
export function decodeOoxmlResourceUsage(bytes: Uint8Array): OoxmlResourceUsageSnapshot {
  let value: unknown;
  try {
    value = JSON.parse(new TextDecoder().decode(bytes)) as unknown;
  } catch {
    throw new TypeError('OOXML resource usage checkpoint is not valid JSON');
  }
  if (!isUsage(value)) {
    throw new TypeError('OOXML resource usage checkpoint is invalid');
  }
  return value;
}

function isFormat(value: unknown): value is OoxmlFormat {
  return value === 'docx' || value === 'xlsx' || value === 'pptx';
}

function isStage(value: unknown): value is OoxmlErrorStage {
  return value === 'container' || value === 'decompression' || value === 'parsing'
    || value === 'serialization' || value === 'layout' || value === 'rendering'
    || value === 'worker';
}

function isBoundedText(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= maximum
    && !/[\u0000-\u001f\u007f]/u.test(value);
}

function isIdentifier(value: unknown): value is string {
  return isBoundedText(value, MAX_IDENTIFIER_LENGTH) && /^[a-z0-9][a-z0-9-]*$/u.test(value);
}

/** Accept only a relative OPC/ZIP part address, never an external or local
 * absolute address. Package extensions may introduce new top-level segments,
 * so validation is structural rather than a closed directory allow-list. */
function isSafePartAddress(value: unknown): value is string {
  if (!isBoundedText(value, MAX_PART_LENGTH)) return false;
  if (
    value.startsWith('/') ||
    value.startsWith('\\') ||
    value.includes('\\') ||
    value.includes('?') ||
    value.includes('#') ||
    value.includes('://') ||
    /^[a-z]:/iu.test(value)
  ) {
    return false;
  }
  return value.split('/').every((segment) => segment !== '' && segment !== '.' && segment !== '..');
}

type KnownPairRule = Readonly<{
  stage: OoxmlErrorStage;
  part: 'required' | 'forbidden' | 'optional';
  configurable?: false;
}>;

const KNOWN_PAIR_RULES = new Map<string, KnownPairRule>([
  ['archive-entry:declared-inflated-bytes', { stage: 'container', part: 'required' }],
  ['archive-entry:actual-inflated-bytes', { stage: 'decompression', part: 'required' }],
  ['archive:entry-count', { stage: 'container', part: 'forbidden' }],
  ['archive:central-directory-bytes', { stage: 'container', part: 'forbidden', configurable: false }],
  ['archive:distinct-inflated-bytes', { stage: 'decompression', part: 'required' }],
  ['xml-event:bytes', { stage: 'parsing', part: 'optional', configurable: false }],
  ['xml-context:bytes', { stage: 'parsing', part: 'optional', configurable: false }],
  ['xml-tree:depth', { stage: 'parsing', part: 'optional', configurable: false }],
  ['worksheet-row:projected-bytes', { stage: 'parsing', part: 'optional', configurable: false }],
  ['worksheet-shell:projected-bytes', { stage: 'parsing', part: 'optional', configurable: false }],
]);
const KNOWN_RESOURCES = new Set(
  [...KNOWN_PAIR_RULES.keys()].map((pair) => pair.slice(0, pair.indexOf(':'))),
);
const KNOWN_METRICS = new Set(
  [...KNOWN_PAIR_RULES.keys()].map((pair) => pair.slice(pair.indexOf(':') + 1)),
);

function isViolation(value: unknown): value is OoxmlResourceViolation {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const data = value as Partial<OoxmlResourceViolation>;
  if (
    !isFormat(data.format) ||
    !isBoundedText(data.operation, MAX_OPERATION_LENGTH) ||
    !isIdentifier(data.resource) ||
    !isIdentifier(data.metric) ||
    !isNonNegativeSafeInteger(data.limit) ||
    !isNonNegativeSafeInteger(data.observed) ||
    typeof data.configurable !== 'boolean' ||
    !isUsage(data.usage)
  ) {
    return false;
  }
  return !('part' in data) || isSafePartAddress(data.part);
}

function isResourceLimitDetails(value: unknown): value is OoxmlResourceLimitErrorDetails {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const details = value as Partial<OoxmlResourceLimitErrorDetails>;
  if (!isStage(details.stage) || !isViolation(details.violation)) return false;
  const violation = details.violation;
  const rule = KNOWN_PAIR_RULES.get(`${violation.resource}:${violation.metric}`);
  if (!rule) {
    // A newer worker may introduce a new resource or metric before its host is
    // upgraded. Preserve that record, but do not accept a nonsensical
    // permutation made solely from the vocabulary this host already knows.
    return !(
      KNOWN_RESOURCES.has(violation.resource) &&
      KNOWN_METRICS.has(violation.metric)
    );
  }
  if (details.stage !== rule.stage) return false;
  if (rule.configurable === false && violation.configurable !== false) return false;
  if (rule.part === 'required') return violation.part !== undefined;
  if (rule.part === 'forbidden') return violation.part === undefined;
  return true;
}

function usageForWire(usage: OoxmlResourceUsageSnapshot): OoxmlResourceUsageSnapshot {
  return {
    archiveEntryCount: usage.archiveEntryCount,
    declaredInflatedBytes: usage.declaredInflatedBytes,
    ...(usage.largestInflatedEntryBytes === undefined
      ? {}
      : { largestInflatedEntryBytes: usage.largestInflatedEntryBytes }),
    distinctInflatedBytes: usage.distinctInflatedBytes,
    operationInflatedBytes: usage.operationInflatedBytes,
  };
}

function detailsForWire(
  details: unknown,
): OoxmlResourceLimitErrorDetails | undefined {
  if (!isResourceLimitDetails(details)) return undefined;
  const violation = details.violation;
  const candidate = {
    stage: details.stage,
    violation: {
      format: violation.format,
      operation: violation.operation,
      resource: violation.resource,
      metric: violation.metric,
      ...(violation.part === undefined ? {} : { part: violation.part }),
      limit: violation.limit,
      observed: violation.observed,
      configurable: violation.configurable,
      usage: usageForWire(violation.usage),
    },
  };
  return isResourceLimitDetails(candidate) ? candidate : undefined;
}

function resourceLimitMessage(details: OoxmlResourceLimitErrorDetails): string {
  const violation = details.violation;
  const location = violation.part ? ` for ${violation.part}` : '';
  return `OOXML resource limit exceeded${location}: ${violation.metric} ${violation.observed} > ${violation.limit}`;
}

/** Parse only an exact Rust resource-limit envelope, never a wrapped substring. */
export function parseResourceLimitError(error: unknown): OoxmlResourceLimitError | undefined {
  const text = error instanceof Error ? error.message : String(error);
  if (!text.startsWith(RESOURCE_LIMIT_PREFIX)) return undefined;
  let value: unknown;
  try {
    value = JSON.parse(text.slice(RESOURCE_LIMIT_PREFIX.length));
  } catch {
    return undefined;
  }
  if (!value || typeof value !== 'object') return undefined;
  const data = value as Partial<RustResourceLimitPayload>;
  if (data.code !== 'ooxml-resource-limit' || !isResourceLimitDetails(data.details)) {
    return undefined;
  }
  return new OoxmlResourceLimitError(resourceLimitMessage(data.details), data.details);
}

function serializeWorkerErrorUnchecked(error: unknown): WorkerErrorPayload {
  const decodedImage = getOoxmlDecodedImageLimitDetails(error);
  if (decodedImage) {
    const canonical = new OoxmlDecodedImageLimitError(
      decodedImage.metric,
      decodedImage.limit,
      decodedImage.observed,
    );
    return {
      message: canonical.message,
      errorName: canonical.name,
      code: canonical.code,
      decodedImage,
    };
  }
  const tiff = getTiffDecodeErrorDetails(error);
  if (tiff) {
    return {
      message: tiff.message,
      errorName: 'TiffDecodeError',
      code: 'ooxml-tiff-decode',
    };
  }
  const insufficientCredit = parsePullSessionInsufficientCreditError(error);
  if (insufficientCredit) {
    return {
      message: insufficientCredit.message,
      errorName: insufficientCredit.name,
      code: insufficientCredit.code,
      insufficientCredit: {
        requiredBytes: insufficientCredit.requiredBytes,
        offeredBytes: insufficientCredit.offeredBytes,
      },
    };
  }
  const typed =
    error instanceof OoxmlError || error instanceof OoxmlResourceLimitError
      ? error
      : parseResourceLimitError(error);
  if (typed instanceof OoxmlResourceLimitError) {
    const resourceLimit = detailsForWire(typed.details);
    if (!resourceLimit) {
      return {
        message: 'Invalid OOXML resource-limit error payload',
        errorName: 'Error',
      };
    }
    return {
      message: typeof typed.message === 'string' ? typed.message : resourceLimitMessage(resourceLimit),
      errorName: 'OoxmlResourceLimitError',
      code: 'ooxml-resource-limit',
      resourceLimit,
    };
  }
  if (typed instanceof OoxmlError) {
    return {
      message: typeof typed.message === 'string' ? typed.message : String(typed.message),
      errorName: isBoundedText(typed.name, MAX_IDENTIFIER_LENGTH) ? typed.name : 'OoxmlError',
      ...(isIdentifier(typed.code) ? { code: typed.code } : {}),
    };
  }
  const sourceMessage = error instanceof Error ? error.message : String(error);
  if (typeof sourceMessage === 'string' && sourceMessage.startsWith(RESOURCE_LIMIT_PREFIX)) {
    return {
      message: 'Invalid OOXML resource-limit payload',
      errorName: 'Error',
    };
  }
  const ordinary = error instanceof Error ? error : new Error(sourceMessage);
  const details = ordinary as Error & { code?: unknown };
  return {
    message: typeof ordinary.message === 'string' ? ordinary.message : String(ordinary.message),
    errorName: isBoundedText(ordinary.name, MAX_IDENTIFIER_LENGTH) ? ordinary.name : 'Error',
    ...(typeof details.code === 'string' ? { code: details.code } : {}),
  };
}

/** Convert an arbitrary worker-side error to a structured-clone-safe payload. */
export function serializeWorkerError(error: unknown): WorkerErrorPayload {
  try {
    return serializeWorkerErrorUnchecked(error);
  } catch {
    // Host objects and caller-mutated Error instances may expose throwing
    // accessors. No property from such an object is safe to forward.
    return {
      message: 'Worker operation failed with an unreadable error',
      errorName: 'Error',
    };
  }
}

const OOXML_ERROR_CODES = new Set<OoxmlErrorCode>([
  'encrypted',
  'invalid-password',
  'unsupported-encryption',
  'legacy-binary-format',
  'not-ooxml',
]);

const DECODED_IMAGE_LIMIT_METRICS = {
  'image-dimension': true,
  'image-pixels': true,
  'active-decoded-bytes': true,
} satisfies Record<OoxmlDecodedImageLimitMetric, true>;

function isDecodedImageLimitMetric(value: unknown): value is OoxmlDecodedImageLimitMetric {
  return typeof value === 'string'
    && Object.prototype.hasOwnProperty.call(DECODED_IMAGE_LIMIT_METRICS, value);
}

function decodedImageLimitDetailsFromWire(
  code: string | undefined,
  value: unknown,
): Readonly<{
  metric: OoxmlDecodedImageLimitMetric;
  limit: number;
  observed: number;
}> | undefined {
  if (code !== 'ooxml-decoded-image-limit' || !value || typeof value !== 'object') {
    return undefined;
  }
  const candidate = value as {
    readonly metric?: unknown;
    readonly limit?: unknown;
    readonly observed?: unknown;
  };
  const metric = candidate.metric;
  const limit = candidate.limit;
  const observed = candidate.observed;
  if (!isDecodedImageLimitMetric(metric)
    || !isNonNegativeSafeInteger(limit)
    || !isNonNegativeSafeInteger(observed)
    || observed <= limit) return undefined;
  return { metric, limit, observed };
}

function deserializeWorkerErrorUnchecked(payload: WorkerErrorPayload): Error {
  // Snapshot every top-level discriminant once. Apart from containing hostile
  // accessors, this prevents a mutable Proxy from passing one branch's checks
  // and then supplying different values to its constructor.
  const rawMessage = payload.message as unknown;
  const rawErrorName = payload.errorName as unknown;
  const rawCode = payload.code as unknown;
  const decodedImagePayload = payload.decodedImage as unknown;
  const insufficientCreditPayload = payload.insufficientCredit as unknown;
  const resourceLimitPayload = payload.resourceLimit as unknown;
  const message = typeof rawMessage === 'string'
    ? rawMessage
    : 'Worker operation failed with an invalid error payload';
  const errorName = isBoundedText(rawErrorName, MAX_IDENTIFIER_LENGTH)
    ? rawErrorName
    : undefined;
  const code = typeof rawCode === 'string' ? rawCode : undefined;
  const decodedImage = decodedImageLimitDetailsFromWire(code, decodedImagePayload);
  if (decodedImage) {
    return new OoxmlDecodedImageLimitError(
      decodedImage.metric,
      decodedImage.limit,
      decodedImage.observed,
    );
  }
  if (code === 'ooxml-tiff-decode') {
    return new TiffDecodeError(message);
  }
  if (
    code === 'ooxml-insufficient-credit'
    && isPullSessionInsufficientCreditDetails(insufficientCreditPayload)
  ) {
    return new PullSessionInsufficientCreditError(insufficientCreditPayload);
  }
  if (
    code === 'ooxml-resource-limit' &&
    isResourceLimitDetails(resourceLimitPayload)
  ) {
    return new OoxmlResourceLimitError(message, resourceLimitPayload);
  }
  if (code && OOXML_ERROR_CODES.has(code as OoxmlErrorCode)) {
    return new OoxmlError(code as OoxmlErrorCode, message);
  }
  const error =
    errorName === 'TypeError'
      ? new TypeError(message)
      : errorName === 'RangeError'
        ? new RangeError(message)
        : new Error(message);
  if (errorName) error.name = errorName;
  if (code !== undefined) Object.assign(error, { code });
  return error;
}

/** Reconstruct a real Error subclass after a payload crosses a worker boundary. */
export function deserializeWorkerError(payload: WorkerErrorPayload): Error {
  try {
    return deserializeWorkerErrorUnchecked(payload);
  } catch {
    return new Error('Worker operation failed with an unreadable error payload');
  }
}
