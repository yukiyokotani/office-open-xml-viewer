import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import { decodeOoxmlResourceUsage } from '@silurus/ooxml-core/worker';

export interface XlsxArchiveBootstrap<TWorkbook> {
  readonly workbook: TWorkbook;
  readonly usage: OoxmlResourceUsageSnapshot;
}

/**
 * Read the workbook projection first, then attach archive accounting. The
 * archive constructor admits only an OPC workbook package, so a retained
 * archive always has a usage ledger; any usage failure is a real error.
 */
export function readXlsxArchiveBootstrap<TWorkbook>(
  readWorkbook: () => TWorkbook,
  readUsage: () => Uint8Array,
): XlsxArchiveBootstrap<TWorkbook> {
  const workbook = readWorkbook();
  return { workbook, usage: decodeOoxmlResourceUsage(readUsage()) };
}
