import { PartFailureLog, partFailureLogOf } from '@silurus/ooxml-core/internal/part-failures';

/**
 * A worksheet, chartsheet or dialogsheet part that failed to parse (#1974).
 * The workbook still loads; that sheet shows a `Sheet "…" could not be
 * displayed` placeholder and every other sheet is unaffected. A valid legacy
 * dialogsheet is not a failure.
 *
 * `message` is the parser's part-tagged diagnostic (for example
 * `xl/worksheets/sheet3.xml: …`), copied and clamped to 1 KiB. `message` and
 * `sheetName` come from the file: treat them as untrusted text (use
 * `textContent`, never `innerHTML`).
 */
export interface XlsxSheetPartFailure {
  readonly kind: 'sheet';
  readonly sheetIndex: number;
  readonly sheetName: string;
  readonly message: string;
}

/**
 * A shared workbook part that failed to parse (for example
 * `xl/sharedStrings.xml`). The workbook still opens; no placeholder is painted
 * for it. Clamped and untrusted like {@link XlsxSheetPartFailure.message}.
 */
export interface XlsxWorkbookPartFailure {
  readonly kind: 'workbook';
  readonly message: string;
}

export type XlsxPartFailure = XlsxSheetPartFailure | XlsxWorkbookPartFailure;

/** @internal */
export const XLSX_WORKBOOK_FAILURE_KEY = 'workbook';

/** @internal */
export function createXlsxPartFailureLog(): PartFailureLog<XlsxPartFailure> {
  return new PartFailureLog<XlsxPartFailure>(
    (failure) => failure.kind === 'workbook' ? XLSX_WORKBOOK_FAILURE_KEY : failure.sheetIndex,
    (left, right) => (left.kind === 'workbook' ? -1 : left.sheetIndex)
      - (right.kind === 'workbook' ? -1 : right.sheetIndex),
  );
}

/** Raised in the render worker when a function option must be resolved on
 * the main thread for a sheet the main thread has not yet admitted. @internal */
export class XlsxErrorPlaceholderRequired extends Error {
  constructor(readonly sheetIndex: number) {
    super('XLSX errorPlaceholder must be resolved on the main thread');
  }
}

/** The engine's failure log (created on first use). @internal */
export function xlsxPartFailureLog(engine: object): PartFailureLog<XlsxPartFailure> {
  return partFailureLogOf(engine, createXlsxPartFailureLog);
}
