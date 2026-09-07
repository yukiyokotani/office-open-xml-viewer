import {
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  parseResourceLimitError,
} from '@silurus/ooxml-core/worker';
import { normalizePresentationBootstrap } from '../presentation-preflight.js';
import type { PptxSlideCursorArchive } from '../slide-cursor-operation.js';
import type { PresentationBootstrap } from '../worker-protocol.js';

export interface PptxNodeAcquisitionOptions {
  readonly resourceLimits?: import('@silurus/ooxml-core').OoxmlResourceLimits;
  readonly maxZipEntryBytes?: number;
  readonly debug?: boolean;
  readonly onResourceMetrics?: (metrics: import('@silurus/ooxml-core').OoxmlResourceMetrics) => void;
  readonly signal?: AbortSignal;
}

export interface PptxNodeArchive extends PptxSlideCursorArchive {
  free(): void;
  assert_healthy(): void;
  presentation_bootstrap(): Uint8Array;
  resource_usage(): Uint8Array;
  extract_image(path: string): Uint8Array;
  extract_media(path: string): Uint8Array;
}

export interface PptxNodeAcquisition {
  readonly archive: PptxNodeArchive;
  readonly bootstrap: PresentationBootstrap;
  readonly metrics: OoxmlResourceMetricsSession;
  closeArchive(): void;
}

export interface PptxOwnedArchiveSource {
  readonly archive: PptxNodeArchive;
  readonly sourceByteLength: number;
  closeArchive(): void;
}

/** Admit an already-owned PPTX archive into the format session. */
export function acquirePptxSessionFromArchive(
  source: PptxOwnedArchiveSource,
  options: PptxNodeAcquisitionOptions = {},
  existingMetrics?: OoxmlResourceMetricsSession,
): PptxNodeAcquisition {
  let closed = false;
  const closeArchive = (): void => {
    if (closed) return;
    closed = true;
    source.closeArchive();
  };
  let metrics = existingMetrics;
  try {
    if (!Number.isSafeInteger(source.sourceByteLength) || source.sourceByteLength < 0) {
      throw new RangeError('PPTX sourceByteLength must be a non-negative safe integer');
    }
    const resourceOptions = normalizeLoadResourceOptions(options);
    metrics ??= new OoxmlResourceMetricsSession({
      enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
      format: 'pptx',
      mode: 'node',
      scope: 'session',
      policy: resourceOptions.policy,
      onMetrics: resourceOptions.onResourceMetrics,
      emitToConsole: resourceOptions.debug,
    });
    if (!existingMetrics) metrics.setSourceBytes(source.sourceByteLength);
    throwIfAborted(options.signal);
    const bootstrap = normalizePresentationBootstrap(JSON.parse(
      new TextDecoder().decode(source.archive.presentation_bootstrap()),
    ) as PresentationBootstrap);
    throwIfAborted(options.signal);
    metrics.checkpoint('presentation bootstrap ready');
    return { archive: source.archive, bootstrap, metrics, closeArchive };
  } catch (error) {
    try { closeArchive(); } catch {}
    const normalized = parseResourceLimitError(error) ?? error;
    metrics?.fail(normalized);
    throw normalized;
  }
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('PPTX presentation session was aborted');
  error.name = 'AbortError';
  throw error;
}
