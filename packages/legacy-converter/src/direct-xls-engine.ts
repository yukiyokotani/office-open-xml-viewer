import {
  validateLegacyXlsSourceDescriptor,
  type LegacyXlsDirectSourceDescriptor,
} from '@silurus/ooxml-core/internal/legacy-xls-source';
import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from './direct-source-runtime.js';

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
