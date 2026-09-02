import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import { decodeOoxmlResourceUsage } from '@silurus/ooxml-core/worker';

const ARCHIVE_USAGE_UNAVAILABLE = 'xlsx resource usage is unavailable';

export interface XlsxArchiveBootstrap<TWorkbook> {
  readonly workbook: TWorkbook;
  readonly usage: OoxmlResourceUsageSnapshot | undefined;
}

/**
 * Read the workbook projection first, then attach archive accounting when the
 * retained ZIP opened successfully. A corrupt container intentionally produces
 * a degraded workbook without an archive usage ledger; that missing diagnostic
 * must not replace the already-materialized container failure.
 */
export function readXlsxArchiveBootstrap<TWorkbook>(
  readWorkbook: () => TWorkbook,
  readUsage: () => Uint8Array,
): XlsxArchiveBootstrap<TWorkbook> {
  const workbook = readWorkbook();
  try {
    return {
      workbook,
      usage: decodeOoxmlResourceUsage(readUsage()),
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message === ARCHIVE_USAGE_UNAVAILABLE) {
      return { workbook, usage: undefined };
    }
    throw error;
  }
}
