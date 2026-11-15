/**
 * `OoxmlError('font-source-limit')` construction and validation.
 *
 * Font sources are application configuration, so their ceilings are library
 * resource policy (generated from `packages/ooxml-common/resource-policy.json`),
 * not OOXML rules. Each quota has exactly one measurement axis; the worker
 * error decoder accepts only these pairings with `observed > limit`.
 */
import {
  OoxmlError,
  type OoxmlFontSourceViolation,
} from './ooxml-error.js';

const FONT_SOURCE_QUOTA_METRICS = {
  stylesheet: 'bytes',
  declarations: 'entry-count',
  families: 'entry-count',
  face: 'bytes',
  retained: 'bytes',
  url: 'length',
} as const;

/** The quotas this version of the library emits (the public
 * `OoxmlFontSourceQuota` is open-ended for later versions). */
export type KnownFontSourceQuota = keyof typeof FONT_SOURCE_QUOTA_METRICS;

function isNonNegativeSafeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

export function fontSourceLimitMessage(violation: OoxmlFontSourceViolation): string {
  return `Font source limit exceeded (${violation.quota}): ${violation.metric} ${violation.observed} > ${violation.limit}`;
}

/** Typed rejection for one breached font-source ceiling. */
export function fontSourceLimitError(
  quota: KnownFontSourceQuota,
  limit: number,
  observed: number,
): OoxmlError {
  const violation: OoxmlFontSourceViolation = {
    resource: 'font-source',
    quota,
    metric: FONT_SOURCE_QUOTA_METRICS[quota],
    limit,
    observed,
  };
  return new OoxmlError('font-source-limit', fontSourceLimitMessage(violation), violation);
}

/** Copy a structurally valid violation record, or undefined. */
export function fontSourceViolationFromWire(value: unknown): OoxmlFontSourceViolation | undefined {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
  const candidate = value as Partial<Record<keyof OoxmlFontSourceViolation, unknown>>;
  const { resource, quota, metric, limit, observed } = candidate;
  if (resource !== 'font-source' || typeof quota !== 'string'
    || !Object.prototype.hasOwnProperty.call(FONT_SOURCE_QUOTA_METRICS, quota)) return undefined;
  const typedQuota = quota as KnownFontSourceQuota;
  if (metric !== FONT_SOURCE_QUOTA_METRICS[typedQuota]
    || !isNonNegativeSafeInteger(limit) || !isNonNegativeSafeInteger(observed)
    || observed <= limit) return undefined;
  return { resource, quota: typedQuota, metric: FONT_SOURCE_QUOTA_METRICS[typedQuota], limit, observed };
}
