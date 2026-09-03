import type {
  LegacyOfficeConversionOptions,
  OoxmlResourceLimits,
  OoxmlResourceMetrics,
} from '@silurus/ooxml-core';

/** Resource policy, diagnostics, and cancellation shared by every Node session. */
export interface OoxmlNodeSessionOptions {
  /** Password for an Agile-encrypted OOXML container. */
  password?: string;
  /** Opt-in legacy DOC/XLS/PPT normalization before parser-WASM initialization. */
  legacyConversion?: LegacyOfficeConversionOptions;
  /** Package-level inflated ZIP admission limits. */
  resourceLimits?: OoxmlResourceLimits;
  /** @deprecated Use `resourceLimits.maxArchiveEntryBytes`. Scheduled for
   * removal in a future breaking release. */
  maxZipEntryBytes?: number;
  /** Emit one content-free resource report for the terminal session outcome. */
  debug?: boolean;
  /** Receive the same terminal report without enabling console output. */
  onResourceMetrics?: (metrics: OoxmlResourceMetrics) => void;
  /** Cooperatively abort conversion, initialization, or active work; synchronous WASM cannot be preempted. */
  signal?: AbortSignal;
}
