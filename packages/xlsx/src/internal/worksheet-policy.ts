import type { XlsxWorksheetLimits } from '@silurus/ooxml-core';
import {
  HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES,
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS,
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES,
  HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES,
  HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS,
  HARD_MAX_XLSX_WORKSHEET_CELLS,
  HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
  HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
  HARD_MAX_XLSX_WORKSHEET_ROWS,
} from '@silurus/ooxml-core/worker';

/**
 * Resolved logical XLSX counters: rows, cell records, owned string UTF-8 bytes
 * and structural JSON bytes. These are admission policy values, not physical
 * memory measurements or out-of-memory guarantees.
 */
export interface ResolvedXlsxWorksheetLimits {
  readonly maxRows: number;
  readonly maxCells: number;
  readonly maxOwnedUtf8Bytes: number;
  readonly maxJsonBytes: number;
}

/**
 * Per-session XLSX worksheet policy. `worksheet` applies to one materialized
 * worksheet, `cache` to the aggregate of cached worksheets, and
 * `maxCoordinateIndexEntries` to the renderer coordinate index, which is
 * tracked separately from the aggregate cache.
 */
export interface NormalizedXlsxWorksheetPolicy {
  readonly worksheet: ResolvedXlsxWorksheetLimits;
  readonly cache: ResolvedXlsxWorksheetLimits;
  readonly maxCoordinateIndexEntries: number;
}

const FIELDS = ['maxRows', 'maxCells', 'maxOwnedUtf8Bytes', 'maxJsonBytes'] as const;
type XlsxWorksheetLimitField = (typeof FIELDS)[number];

/** Largest accepted value, so that `value + 1` stays an exact safe integer. */
const MAX_ACCEPTED_LIMIT = Number.MAX_SAFE_INTEGER - 1;

function frozenLimits(limits: ResolvedXlsxWorksheetLimits): ResolvedXlsxWorksheetLimits {
  return Object.freeze({
    maxRows: limits.maxRows,
    maxCells: limits.maxCells,
    maxOwnedUtf8Bytes: limits.maxOwnedUtf8Bytes,
    maxJsonBytes: limits.maxJsonBytes,
  });
}

/**
 * Default XLSX worksheet policy. Sourced from the generated legacy `HARD_*`
 * constants for backward compatibility; these are policy defaults rather than
 * absolute ceilings.
 */
export const DEFAULT_XLSX_WORKSHEET_POLICY: NormalizedXlsxWorksheetPolicy = Object.freeze({
  worksheet: frozenLimits({
    maxRows: HARD_MAX_XLSX_WORKSHEET_ROWS,
    maxCells: HARD_MAX_XLSX_WORKSHEET_CELLS,
    maxOwnedUtf8Bytes: HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
    maxJsonBytes: HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
  }),
  cache: frozenLimits({
    maxRows: HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS,
    maxCells: HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS,
    maxOwnedUtf8Bytes: HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES,
    maxJsonBytes: HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES,
  }),
  maxCoordinateIndexEntries: HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES,
});

function readLimit(limits: Readonly<Record<string, unknown>>, field: XlsxWorksheetLimitField): number {
  const value = limits[field];
  if (value === undefined) return DEFAULT_XLSX_WORKSHEET_POLICY.worksheet[field];
  if (typeof value !== 'number') {
    throw new TypeError(`xlsxWorksheetLimits.${field} must be a number`);
  }
  if (!Number.isSafeInteger(value) || value <= 0 || value > MAX_ACCEPTED_LIMIT) {
    throw new RangeError(
      `xlsxWorksheetLimits.${field} must be a positive safe integer no greater than ${MAX_ACCEPTED_LIMIT}`,
    );
  }
  return value;
}

/**
 * Validate `xlsxWorksheetLimits` and resolve an independent, deeply frozen
 * policy. Neither the input nor the default policy is mutated.
 */
export function normalizeXlsxWorksheetPolicy(
  options: { readonly xlsxWorksheetLimits?: XlsxWorksheetLimits } = {},
): NormalizedXlsxWorksheetPolicy {
  const raw: unknown = options.xlsxWorksheetLimits;
  // ast-grep-ignore: no-plain-object-string-map -- read-only empty default, never written.
  let source: Readonly<Record<string, unknown>> = {};
  if (raw !== undefined) {
    if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
      throw new TypeError('xlsxWorksheetLimits must be a non-null, non-array object');
    }
    source = raw as Readonly<Record<string, unknown>>;
  }
  const worksheet = frozenLimits({
    maxRows: readLimit(source, 'maxRows'),
    maxCells: readLimit(source, 'maxCells'),
    maxOwnedUtf8Bytes: readLimit(source, 'maxOwnedUtf8Bytes'),
    maxJsonBytes: readLimit(source, 'maxJsonBytes'),
  });
  const defaults = DEFAULT_XLSX_WORKSHEET_POLICY;
  const cache = frozenLimits({
    maxRows: Math.max(defaults.cache.maxRows, worksheet.maxRows),
    maxCells: Math.max(defaults.cache.maxCells, worksheet.maxCells),
    maxOwnedUtf8Bytes: Math.max(defaults.cache.maxOwnedUtf8Bytes, worksheet.maxOwnedUtf8Bytes),
    maxJsonBytes: Math.max(defaults.cache.maxJsonBytes, worksheet.maxJsonBytes),
  });
  return Object.freeze({
    worksheet,
    cache,
    maxCoordinateIndexEntries: Math.max(defaults.maxCoordinateIndexEntries, worksheet.maxCells),
  });
}

/**
 * Encode the per-worksheet limits for the XLSX-specific WASM ABI. Values
 * are already validated safe integers by {@link normalizeXlsxWorksheetPolicy}.
 */
export function xlsxWorksheetPolicyForWasm(
  policy: NormalizedXlsxWorksheetPolicy,
): readonly [bigint, bigint, bigint, bigint] {
  const { worksheet } = policy;
  return [
    BigInt(worksheet.maxRows),
    BigInt(worksheet.maxCells),
    BigInt(worksheet.maxOwnedUtf8Bytes),
    BigInt(worksheet.maxJsonBytes),
  ] as const;
}
