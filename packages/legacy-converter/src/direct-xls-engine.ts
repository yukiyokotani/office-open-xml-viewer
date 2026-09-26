import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from './direct-source-runtime.js';

import { MAX_LEGACY_SOURCE_BYTES as MAX_LEGACY_XLS_SOURCE_BYTES } from './legacy-source-limits.js';

export { MAX_LEGACY_XLS_SOURCE_BYTES };

/**
 * The direct XLS workbook archive. `host_layout_request` / `configure_host_layout`
 * are the XLSX host-layout capability: BIFF drawing anchors are stored in the
 * Normal font's maximum digit width, which only the rendering host can measure.
 */
export interface LegacyXlsNativeArchive {
  free(): void;
  host_layout_request(): Uint8Array;
  configure_host_layout(maximumDigitWidth?: number): void;
  parse(): Uint8Array;
  open_sheet_cursor(index: number, name: string): void;
  pull_sheet_cursor(rowCredit: number): Uint8Array;
  sheet_cursor_pull_finished(): boolean;
  acknowledge_sheet_cursor_terminal(): void;
  cancel_sheet_cursor(): void;
  close_sheet_cursor(): void;
  extract_image(key: string): Uint8Array;
  close_workbook_session(): void;
  assert_healthy(): void;
}

export type OwnedLegacyXlsSource = OwnedDirectSource<LegacyXlsNativeArchive>;

export interface LegacyXlsGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  LegacyXlsWorkbook: new (bytes: Uint8Array) => LegacyXlsNativeArchive;
}

export function createLegacyXlsSourceEngine(
  loadGlue: () => Promise<LegacyXlsGlue>,
  resolveWasm: (wasmUrl: string) => Promise<unknown>,
): Readonly<{
  open(bytes: Uint8Array, wasmUrl: string, signal?: AbortSignal): Promise<OwnedLegacyXlsSource>;
}> {
  return createDirectSourceRuntime({
    label: 'legacy XLS',
    maximumSourceBytes: MAX_LEGACY_XLS_SOURCE_BYTES,
    loadGlue,
    resolveWasm,
    construct: (glue, bytes) => new glue.LegacyXlsWorkbook(bytes),
    closeNative: (archive) => archive.close_workbook_session(),
  });
}

export { resolveDirectWasmInput };
