import {
  validateLegacyXlsSourceDescriptor,
  type LegacyXlsDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-xls-source';
import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from './direct-source-runtime.js';
import {
  measureXlsFont,
  type LegacyXlsFontMeasurement,
  type LegacyXlsNormalFont,
} from './xls-font-metrics.js';

const MAX_MEASUREMENT_REQUEST_BYTES = 4096;

export interface LegacyXlsNativeArchive {
  free(): void;
  measurement_request(): Uint8Array;
  configure_mdw(mdw?: number): void;
  parse(): Uint8Array;
  open_sheet_cursor(index: number, name: string): void;
  pull_sheet_cursor(rowCredit: number): Uint8Array;
  sheet_cursor_pull_finished(): boolean;
  acknowledge_sheet_cursor_terminal(): void;
  cancel_sheet_cursor(): void;
  close_sheet_cursor(): void;
  resource_usage(): Uint8Array;
  sheet_cursor_resource_usage(): Uint8Array;
  extract_image(key: string): Uint8Array;
  to_markdown(): string;
  close_workbook_session(): void;
  assert_healthy(): void;
}

export type OwnedLegacyXlsSource = OwnedDirectSource<LegacyXlsNativeArchive>;

/** Resolve and apply the one native geometry decision before workbook parsing. */
export async function configureLegacyXlsMeasurement(
  archive: Pick<LegacyXlsNativeArchive, 'measurement_request' | 'configure_mdw'>,
  measure?: LegacyXlsFontMeasurement,
  signal?: AbortSignal,
): Promise<number | undefined> {
  throwIfMeasurementAborted(signal);
  const request = decodeMeasurementRequest(archive.measurement_request());
  if (!request.required) return undefined;
  let width: number | undefined;
  if (measure && request.font) {
    width = await measureXlsFont(
      measure,
      {
        family: request.font.name,
        sizePoints: request.font.sizePoints,
        bold: request.font.bold,
        italic: request.font.italic,
      },
      signal ?? new AbortController().signal,
    );
  }
  throwIfMeasurementAborted(signal);
  archive.configure_mdw(width);
  return width;
}

function decodeMeasurementRequest(bytes: Uint8Array): Readonly<{
  required: boolean;
  font: (LegacyXlsNormalFont & { readonly name: string }) | null;
}> {
  if (bytes.byteLength > MAX_MEASUREMENT_REQUEST_BYTES) {
    throw new RangeError('legacy XLS measurement request byte budget exceeded');
  }
  let value: unknown;
  try {
    value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
  } catch {
    throw new TypeError('invalid legacy XLS measurement request');
  }
  if (!isExactObject(value, ['required', 'font']) || typeof value.required !== 'boolean') {
    throw new TypeError('invalid legacy XLS measurement request');
  }
  if (value.font === null) return { required: value.required, font: null };
  if (!isExactObject(value.font, ['name', 'sizePoints', 'bold', 'italic'])) {
    throw new TypeError('invalid legacy XLS measurement font');
  }
  const { name, sizePoints, bold, italic } = value.font;
  if (
    typeof name !== 'string' || name.length === 0 || name.length > 255
    || typeof sizePoints !== 'number' || !Number.isFinite(sizePoints)
    || sizePoints <= 0 || sizePoints > 65535 / 20
    || typeof bold !== 'boolean' || typeof italic !== 'boolean'
  ) {
    throw new TypeError('invalid legacy XLS measurement font');
  }
  return { required: value.required, font: { name, family: name, sizePoints, bold, italic } };
}

function isExactObject(value: unknown, keys: readonly string[]): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const own = Object.keys(value);
  return own.length === keys.length && keys.every((key) => Object.hasOwn(value, key));
}

function throwIfMeasurementAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('legacy XLS font measurement aborted');
  error.name = 'AbortError';
  throw error;
}

interface LegacyXlsGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  LegacyXlsWorkbook: new (bytes: Uint8Array) => LegacyXlsNativeArchive;
}

export function createLegacyXlsSourceEngine(
  loadGlue: () => Promise<LegacyXlsGlue>,
  resolveWasm: (wasmUrl: string) => Promise<unknown>,
) {
  return createDirectSourceRuntime({
    label: 'legacy XLS',
    maximumSourceBytes: 256 * 1024 * 1024,
    validate: validateLegacyXlsSourceDescriptor,
    loadGlue,
    resolveWasm,
    construct: (glue, bytes) => new glue.LegacyXlsWorkbook(bytes),
    closeNative: (archive) => archive.close_workbook_session(),
  });
}

const defaultEngine = createLegacyXlsSourceEngine(
  () => import('./wasm-direct-xls/legacy_office_converter.js'),
  resolveDirectWasmInput,
);

export function openLegacyXlsSource(
  bytes: Uint8Array,
  descriptor: LegacyXlsDirectSourceDescriptor,
  signal?: AbortSignal,
): Promise<OwnedLegacyXlsSource> {
  return defaultEngine.open(bytes, descriptor, signal);
}
