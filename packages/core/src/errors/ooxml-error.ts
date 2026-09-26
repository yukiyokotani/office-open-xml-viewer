/**
 * Machine-readable code for a typed load-time failure.
 *
 * The container-level failures the `load()` factories detect on the main thread
 * before handing bytes to the parser worker (see `sniffCfb` / `decryptOoxml`).
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
 *   - `'invalid-numbering'`     — a DOCX numbering level that Word cannot open
 *     normally (invalid paragraph ilvl or an out-of-range level definition).
 */
export type OoxmlErrorCode =
  | 'encrypted'
  | 'invalid-password'
  | 'unsupported-encryption'
  | 'legacy-binary-format'
  | 'not-ooxml'
  | 'invalid-numbering';

export type OoxmlErrorStage =
  | 'container'
  | 'decompression'
  | 'parsing'
  | 'serialization'
  | 'layout'
  | 'rendering'
  | 'worker';

/**
 * Typed error thrown by the docx / pptx / xlsx `load()` factories for failures
 * that carry a stable, programmatic {@link OoxmlErrorCode} (e.g. a
 * password-protected or legacy-binary file detected from its container magic).
 *
 * Note on workers: `instanceof OoxmlError` does not survive a structured-clone
 * across the worker boundary. Detection that needs a typed error is therefore
 * done on the main thread (before the worker is involved) so a genuine
 * `OoxmlError` instance is thrown to the caller. Errors that must cross the
 * worker boundary should carry the `code` string and be reconstructed on the
 * main side.
 */
export class OoxmlError extends Error {
  readonly code: OoxmlErrorCode;

  constructor(code: OoxmlErrorCode, message: string) {
    super(message);
    this.name = 'OoxmlError';
    this.code = code;
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
