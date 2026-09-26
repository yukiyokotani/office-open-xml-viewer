import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import { decodeOoxmlResourceUsage } from '@silurus/ooxml-core/worker';

export interface XlsxArchiveBootstrap<TWorkbook> {
  readonly workbook: TWorkbook;
  readonly usage: OoxmlResourceUsageSnapshot | undefined;
}

/**
 * Read the workbook projection first, then attach archive accounting. The
 * archive constructor admits only an OPC workbook package, so a retained
 * archive always has a usage ledger; any usage failure is a real error. A
 * model source that keeps no ZIP ledger reports `undefined` instead.
 */
export function readXlsxArchiveBootstrap<TWorkbook>(
  readWorkbook: () => TWorkbook,
  readUsage: () => Uint8Array | undefined,
): XlsxArchiveBootstrap<TWorkbook> {
  const workbook = readWorkbook();
  // `undefined`: the loaded model source has no ZIP accounting.
  const usage = readUsage();
  return {
    workbook,
    usage: usage === undefined ? undefined : decodeOoxmlResourceUsage(usage),
  };
}
