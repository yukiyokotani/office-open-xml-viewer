/**
 * Machine-readable code for a typed input failure, at load or during lazy reads.
 *
 * Container failures can be detected before parsing (sniffCfb/decryptOoxml);
 * parsers and configured model sources can also report typed failures.
 * This is the seed of the broader typed-error surface tracked as PD4 (OoxmlError
 * typed errors). Add codes here rather than throwing bare `Error(string)`, so
 * callers can `switch` on `err.code` instead of matching message text.
 *
 *   - `'encrypted'`             — password-protected, but no `password` was
 *     supplied (pass `LoadOptions.password` to decrypt).
 *   - `'invalid-password'`      — a `password` was supplied but did not match.
 *   - `'unsupported-encryption'`— encrypted with a scheme other than Agile
 *     (Standard / Extensible / a legacy binary encryptor), which this library
 *     cannot decrypt (PD8 implements Agile only).
 *   - `'legacy-binary-format'`  — a raw .doc / .xls / .ppt (not OOXML).
 *   - `'not-ooxml'`             — a CFB of an unrecognised kind, or otherwise
 *     not an OOXML package: unreadable as ZIP (including empty input), a ZIP
 *     without the OPC `[Content_Types].xml` stream, or an OPC package without
 *     the format's main part. Raised by the parser at package admission and
 *     reconstructed across the worker boundary.
 *   - `'unsupported-content'`  — a configured model source cannot represent
 *     the input; details identify its format and stable rejection category.
 *   - `'unsupported-image-effect'` — a malformed CT_Blip
 *     chain, an unavailable image-effect/fill surface, or unrepresentable
 *     image-effect/fill geometry at render time.
 *   - `'invalid-numbering'`     — a DOCX numbering level that Word cannot open
 *     normally (invalid paragraph ilvl or an out-of-range level definition).
 *   - `'font-source-limit'`     — a font source factory received a value
 *     beyond a generated hard ceiling (declaration count, face bytes or URL
 *     length); {@link OoxmlError.violation} names it. Thrown by the factory
 *     (a misconfigured source fails fast). A ceiling reached while a document
 *     loads never throws: the families concerned keep their host fallback
 *     with a warning, like a network failure.
 */
export type OoxmlErrorCode =
  | 'encrypted'
  | 'invalid-password'
  | 'unsupported-encryption'
  | 'legacy-binary-format'
  | 'not-ooxml'
  | 'unsupported-image-effect'
  | 'invalid-numbering'
  | 'unsupported-content'
  | 'font-source-limit';

export type OoxmlErrorStage =
  | 'container'
  | 'decompression'
  | 'parsing'
  | 'serialization'
  | 'layout'
  | 'rendering'
  | 'worker';

/** A source's declared representational boundary, not a parser crash or quota.
 * Format is a bounded source-owned identifier; core has no input-format registry.
 * No document text, paths, native diagnostics or URLs belong in this record. */
export interface OoxmlUnsupportedContentDetails {
  readonly format: string;
  readonly reason: 'source-subset';
}

/** Font-source quota, named after its generated `HARD_MAX_FONT_SOURCE_*`
 * ceiling. Library resource policy, not an OOXML rule. */
export type OoxmlFontSourceQuota =
  | 'stylesheet'
  | 'declarations'
  | 'families'
  | 'face'
  | 'retained'
  | 'url';

/** Violation record carried by `OoxmlError('font-source-limit')`. It names
 * the quota and its measurement only: no family, URL or font bytes. */
export interface OoxmlFontSourceViolation {
  readonly resource: 'font-source';
  readonly quota: OoxmlFontSourceQuota;
  readonly metric: 'bytes' | 'entry-count' | 'length';
  readonly limit: number;
  readonly observed: number;
}

/**
 * Typed error thrown by load factories and the image paint pipeline for failures
 * that carry a stable, programmatic {@link OoxmlErrorCode} (e.g. a
 * password-protected or legacy-binary file detected from its container magic).
 *
 * Workers send validated code/details payloads; the receiving viewer restores
 * a real OoxmlError. Source-module open and lazy archive operations use the
 * same constructor in their owning realm before crossing the worker boundary.
 */
export class OoxmlError extends Error {
  readonly code: OoxmlErrorCode;

  readonly details?: OoxmlUnsupportedContentDetails;

  /** Present only for `'font-source-limit'`. */
  readonly violation?: OoxmlFontSourceViolation;

  constructor(
    code: OoxmlErrorCode,
    message: string,
    details?: OoxmlUnsupportedContentDetails | OoxmlFontSourceViolation,
  ) {
    super(message);
    this.name = 'OoxmlError';
    this.code = code;
    if (code === 'unsupported-content' && details && 'format' in details) {
      this.details = Object.freeze({ format: details.format, reason: details.reason });
    }
    if (code === 'font-source-limit' && details && 'quota' in details) {
      this.violation = Object.freeze({
        resource: 'font-source',
        quota: details.quota,
        metric: details.metric,
        limit: details.limit,
        observed: details.observed,
      });
    }
    // Restore the prototype chain for environments that down-level `extends
    // Error` (e.g. older TS `target`), so `instanceof OoxmlError` holds.
    Object.setPrototypeOf(this, OoxmlError.prototype);
  }
}

export type OoxmlFormat = 'docx' | 'xlsx' | 'pptx';

export interface OoxmlResourceUsageSnapshot {
  readonly archiveEntryCount: number;
  readonly declaredInflatedBytes: number;
  /** Largest actual decompressed size observed for one ZIP entry. */
  readonly largestInflatedEntryBytes?: number;
  readonly distinctInflatedBytes: number;
  readonly operationInflatedBytes: number;
}

type ExtensibleLiteral<Known extends string> = Known | (string & Record<never, never>);

/**
 * Resource family reported by a policy or hard-quota violation.
 *
 * The known literals provide editor completion. The string tail is deliberate:
 * adding a future format-owned unit must not break exhaustive switches compiled
 * against an older host while a newer worker is already able to report it.
 */
export type OoxmlResourceName = ExtensibleLiteral<
  | 'archive'
  | 'archive-entry'
  | 'xml-event'
  | 'xml-context'
  | 'xml-tree'
  | 'worksheet-row'
  | 'worksheet-shell'
>;

/** Measurement axis used by an OOXML resource violation. Extensible by design. */
export type OoxmlResourceMetric = ExtensibleLiteral<
  | 'declared-inflated-bytes'
  | 'actual-inflated-bytes'
  | 'entry-count'
  | 'central-directory-bytes'
  | 'distinct-inflated-bytes'
  | 'bytes'
  | 'depth'
  | 'projected-bytes'
>;

/** Stable public violation record. Valid resource/metric pairings are enforced
 * by the emitting parser and the worker decoder rather than by a closed public
 * union that would require a breaking expansion for every new hard quota. */
export interface OoxmlResourceViolation {
  readonly format: OoxmlFormat;
  readonly operation: string;
  readonly resource: OoxmlResourceName;
  readonly metric: OoxmlResourceMetric;
  readonly part?: string;
  readonly limit: number;
  readonly observed: number;
  readonly configurable: boolean;
  readonly usage: OoxmlResourceUsageSnapshot;
}

export interface OoxmlResourceLimitErrorDetails {
  readonly stage: OoxmlErrorStage;
  readonly violation: OoxmlResourceViolation;
}

/** Deterministic rejection caused by a measured OOXML resource-policy breach. */
export class OoxmlResourceLimitError extends Error {
  readonly code = 'ooxml-resource-limit' as const;
  readonly details: OoxmlResourceLimitErrorDetails;

  constructor(message: string, details: OoxmlResourceLimitErrorDetails) {
    super(message);
    this.name = 'OoxmlResourceLimitError';
    const source = details.violation;
    const violation: OoxmlResourceViolation = Object.freeze({
      format: source.format,
      operation: source.operation,
      resource: source.resource,
      metric: source.metric,
      ...(source.part === undefined ? {} : { part: source.part }),
      limit: source.limit,
      observed: source.observed,
      configurable: source.configurable,
      usage: Object.freeze({
        archiveEntryCount: source.usage.archiveEntryCount,
        declaredInflatedBytes: source.usage.declaredInflatedBytes,
        ...(source.usage.largestInflatedEntryBytes === undefined
          ? {}
          : { largestInflatedEntryBytes: source.usage.largestInflatedEntryBytes }),
        distinctInflatedBytes: source.usage.distinctInflatedBytes,
        operationInflatedBytes: source.usage.operationInflatedBytes,
      }),
    });
    this.details = Object.freeze({
      stage: details.stage,
      violation,
    });
    Object.setPrototypeOf(this, OoxmlResourceLimitError.prototype);
  }
}
