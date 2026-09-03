import {
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS,
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES,
  HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES,
  HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS,
  HARD_MAX_XLSX_WORKSHEET_CELLS,
  HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
  HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
  HARD_MAX_XLSX_WORKSHEET_ROWS,
} from '@silurus/ooxml-core/worker';
import {
  OoxmlResourceLimitError,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import {
  cappedAdd,
  measureStructuralJson,
  utf8Bytes,
} from '@silurus/ooxml-core/internal/resource-measurement';
import type { Row, Worksheet } from './types.js';

export const XLSX_MAX_MATERIALIZED_ROWS = HARD_MAX_XLSX_WORKSHEET_ROWS;
export const XLSX_MAX_MATERIALIZED_CELLS = HARD_MAX_XLSX_WORKSHEET_CELLS;
export const XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES =
  HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES;
export const XLSX_MAX_MATERIALIZED_JSON_BYTES = HARD_MAX_XLSX_WORKSHEET_JSON_BYTES;
export const XLSX_MAX_CACHED_ROWS = HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS;
export const XLSX_MAX_CACHED_CELLS = HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS;
export const XLSX_MAX_CACHED_OWNED_UTF8_BYTES =
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES;
export const XLSX_MAX_CACHED_JSON_BYTES = HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES;

export interface WorksheetModelUsage {
  rows: number;
  cells: number;
  ownedUtf8Bytes: number;
}

export interface WorksheetCacheUsage {
  rows: number;
  cells: number;
  ownedUtf8Bytes: number;
  jsonBytes: number;
}

const ZERO_RESOURCE_USAGE: OoxmlResourceUsageSnapshot = Object.freeze({
  archiveEntryCount: 0,
  declaredInflatedBytes: 0,
  distinctInflatedBytes: 0,
  operationInflatedBytes: 0,
});

export function measureRows(rows: readonly Row[]): WorksheetModelUsage {
  const cells = rows.reduce(
    (total, row) => cappedAdd(total, row.cells.length, XLSX_MAX_MATERIALIZED_CELLS),
    0,
  );
  return {
    rows: rows.length,
    cells,
    // This is deliberately cell-scoped. It includes every retained string in
    // Cell.value (rich/phonetic formatting and the discriminator included) and
    // formula text. Shared values must be resolved before calling this helper,
    // so repeated shared-string references are charged once per materialized
    // cell. Ancillary worksheet strings are covered by the exact JSON ceiling.
    ownedUtf8Bytes: rows.reduce((rowTotal, row) => row.cells.reduce((cellTotal, cell) => {
      const valueBytes = measureStructuralJson(
        cell.value,
        XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
      ).stringValueUtf8Bytes;
      const formulaBytes = cell.formula === undefined
        ? 0
        : utf8Bytes(cell.formula, XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES);
      return cappedAdd(
        cellTotal,
        cappedAdd(valueBytes, formulaBytes, XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES),
        XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
      );
    }, rowTotal), 0),
  };
}

export function measureWorksheet(worksheet: Worksheet): WorksheetModelUsage & { jsonBytes: number } {
  return completeWorksheetUsage(worksheet, measureRows(worksheet.rows));
}

/** Complete an incrementally accumulated row/cell measurement with the exact
 * retained JSON size. Streaming callers already measured each row chunk, so
 * repeating that full traversal at terminal admission adds cost but no safety. */
export function completeWorksheetUsage(
  worksheet: Worksheet,
  model: WorksheetModelUsage,
): WorksheetCacheUsage {
  const measured = measureStructuralJson(worksheet, Math.max(
    XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
    XLSX_MAX_MATERIALIZED_JSON_BYTES,
  ));
  return { ...model, jsonBytes: measured.jsonBytes };
}

export function addWorksheetUsage(
  current: WorksheetModelUsage,
  addition: WorksheetModelUsage,
): WorksheetModelUsage {
  return {
    rows: cappedAdd(current.rows, addition.rows, XLSX_MAX_MATERIALIZED_ROWS),
    cells: cappedAdd(current.cells, addition.cells, XLSX_MAX_MATERIALIZED_CELLS),
    ownedUtf8Bytes: cappedAdd(
      current.ownedUtf8Bytes,
      addition.ownedUtf8Bytes,
      XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
    ),
  };
}

export function addWorksheetCacheUsage(
  current: WorksheetCacheUsage,
  addition: WorksheetModelUsage & { jsonBytes: number },
  subtraction: Partial<WorksheetCacheUsage> = {},
): WorksheetCacheUsage {
  const baseRows = current.rows - (subtraction.rows ?? 0);
  const baseCells = current.cells - (subtraction.cells ?? 0);
  const baseOwnedUtf8Bytes = current.ownedUtf8Bytes - (subtraction.ownedUtf8Bytes ?? 0);
  const baseJsonBytes = current.jsonBytes - (subtraction.jsonBytes ?? 0);
  if (baseRows < 0 || baseCells < 0 || baseOwnedUtf8Bytes < 0 || baseJsonBytes < 0) {
    throw new Error('worksheet cache accounting underflow');
  }
  return {
    rows: cappedAdd(baseRows, addition.rows, XLSX_MAX_CACHED_ROWS),
    cells: cappedAdd(baseCells, addition.cells, XLSX_MAX_CACHED_CELLS),
    ownedUtf8Bytes: cappedAdd(
      baseOwnedUtf8Bytes,
      addition.ownedUtf8Bytes,
      XLSX_MAX_CACHED_OWNED_UTF8_BYTES,
    ),
    jsonBytes: cappedAdd(baseJsonBytes, addition.jsonBytes, XLSX_MAX_CACHED_JSON_BYTES),
  };
}

export function worksheetLimitError(
  operation: string,
  part: string | undefined,
  resource:
    | 'delimited-text-source'
    | 'worksheet-model'
    | 'worksheet-cell-content'
    | 'worksheet-json'
    | 'worksheet-cache',
  metric: 'rows' | 'cells' | 'owned-utf8-bytes' | 'bytes',
  limit: number,
  observed: number,
  usage?: OoxmlResourceUsageSnapshot,
): OoxmlResourceLimitError {
  const stage = resource === 'worksheet-json' ? 'serialization' : 'parsing';
  return new OoxmlResourceLimitError(
    `OOXML resource limit exceeded${part ? ` for ${part}` : ''}: ${metric} ${observed} > ${limit}`,
    {
      stage,
      violation: {
        format: 'xlsx',
        operation,
        resource,
        metric,
        ...(part === undefined ? {} : { part }),
        limit,
        observed: Math.min(observed, limit + 1),
        configurable: false,
        usage: usage ?? ZERO_RESOURCE_USAGE,
      },
    },
  );
}

export function assertWorksheetModelUsage(
  measured: WorksheetModelUsage,
  operation: string,
  part: string | undefined,
  usage?: OoxmlResourceUsageSnapshot,
): void {
  const checks = [
    ['rows', measured.rows, XLSX_MAX_MATERIALIZED_ROWS],
    ['cells', measured.cells, XLSX_MAX_MATERIALIZED_CELLS],
    ['owned-utf8-bytes', measured.ownedUtf8Bytes, XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES],
  ] as const;
  for (const [metric, observed, limit] of checks) {
    if (observed > limit) {
      throw worksheetLimitError(
        operation,
        part,
        metric === 'owned-utf8-bytes' ? 'worksheet-cell-content' : 'worksheet-model',
        metric,
        limit,
        observed,
        usage,
      );
    }
  }
}

export function assertWorksheetJsonBytes(
  observed: number,
  operation: string,
  part: string | undefined,
  usage?: OoxmlResourceUsageSnapshot,
): void {
  if (observed > XLSX_MAX_MATERIALIZED_JSON_BYTES) {
    throw worksheetLimitError(
      operation,
      part,
      'worksheet-json',
      'bytes',
      XLSX_MAX_MATERIALIZED_JSON_BYTES,
      observed,
      usage,
    );
  }
}

export function assertWorksheetCacheUsage(
  usage: WorksheetCacheUsage,
  operation: string,
  part: string | undefined,
  resourceUsage?: OoxmlResourceUsageSnapshot,
): void {
  if (usage.rows > XLSX_MAX_CACHED_ROWS) {
    throw worksheetLimitError(
      operation, part, 'worksheet-cache', 'rows', XLSX_MAX_CACHED_ROWS, usage.rows, resourceUsage,
    );
  }
  if (usage.cells > XLSX_MAX_CACHED_CELLS) {
    throw worksheetLimitError(
      operation, part, 'worksheet-cache', 'cells', XLSX_MAX_CACHED_CELLS, usage.cells, resourceUsage,
    );
  }
  if (usage.ownedUtf8Bytes > XLSX_MAX_CACHED_OWNED_UTF8_BYTES) {
    throw worksheetLimitError(
      operation,
      part,
      'worksheet-cache',
      'owned-utf8-bytes',
      XLSX_MAX_CACHED_OWNED_UTF8_BYTES,
      usage.ownedUtf8Bytes,
      resourceUsage,
    );
  }
  if (usage.jsonBytes > XLSX_MAX_CACHED_JSON_BYTES) {
    throw worksheetLimitError(
      operation,
      part,
      'worksheet-cache',
      'bytes',
      XLSX_MAX_CACHED_JSON_BYTES,
      usage.jsonBytes,
      resourceUsage,
    );
  }
}
