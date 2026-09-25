import type {
  ModelSource,
  OoxmlResourceLimits,
  OoxmlResourceMetrics,
} from '@silurus/ooxml-core';

/** Resource policy, diagnostics, and cancellation shared by every Node session. */
export interface OoxmlNodeSessionOptions {
  /** Password for an Agile-encrypted OOXML container. */
  password?: string;
  /**
   * Application-supplied sources for input that is not an OOXML package (see
   * `LoadOptions.modelSources`). Each must target the session's format; the
   * first whose `claim()` accepts the bytes opens them in this realm.
   */
  modelSources?: readonly ModelSource[];
  /** Package-level inflated ZIP admission limits. */
  resourceLimits?: OoxmlResourceLimits;
  /** @deprecated Use `resourceLimits.maxArchiveEntryBytes`. Scheduled for
   * removal in a future breaking release. */
  maxZipEntryBytes?: number;
  /** Emit one content-free resource report for the terminal session outcome. */
  debug?: boolean;
  /** Receive the same terminal report without enabling console output. */
  onResourceMetrics?: (metrics: OoxmlResourceMetrics) => void;
  /** Cooperatively abort initialization or active work; synchronous WASM cannot be preempted. */
  signal?: AbortSignal;
}
