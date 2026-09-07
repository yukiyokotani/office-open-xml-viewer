import type { LegacyOfficeConversionOptions } from '@silurus/ooxml-core';
import type { XlsxWorkbook } from './workbook.js';

interface BoundLegacyConversion {
  readonly options?: LegacyOfficeConversionOptions;
  readonly cleanup: () => void;
}

/** Retain direct-source cancellation wiring for the owned workbook lifetime. */
export function settleLegacyXlsLoad(
  pending: Promise<XlsxWorkbook>,
  conversion: BoundLegacyConversion,
): Promise<XlsxWorkbook> {
  const selected = conversion.options?.xls;
  if (selected === undefined) return pending;
  if (!('source' in selected)) return pending.finally(conversion.cleanup);
  return pending.then(
    (workbook) => {
      workbook._retainLegacyXlsSignalCleanup(conversion.cleanup);
      return workbook;
    },
    (error: unknown) => {
      conversion.cleanup();
      throw error;
    },
  );
}
