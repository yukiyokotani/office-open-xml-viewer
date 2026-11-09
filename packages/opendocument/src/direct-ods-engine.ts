import type { ModelSourceOpenContext } from '@silurus/ooxml-core';
import {
  createDirectSourceRuntime,
  resolveDirectWasmInput,
  type OwnedDirectSource,
} from '@silurus/ooxml-core/internal/model-source-runtime';
import { odfNativeRejection } from './native-rejection.js';

/**
 * The XLSX model-source archive the ODS reader exposes: the workbook
 * bootstrap, the worksheet cursor, image reads, and the host-layout
 * capability (`host_layout_request` / `configure_host_layout`) through which
 * the renderer measures the default cell font's maximum digit width.
 */
export interface OdsNativeArchive {
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

export type OwnedOdsSource = OwnedDirectSource<OdsNativeArchive>;

export interface OdsGlue {
  default(input: { module_or_path: unknown }): Promise<unknown>;
  OdsWorkbook: new (
    bytes: Uint8Array,
    maxArchiveEntryBytes?: number,
    maxTotalInflatedBytes?: number,
    maxArchiveEntries?: number,
  ) => OdsNativeArchive;
}

/** Internal injectable engine; generated glue is wired only by the source module. */
export function createOdsSourceEngine(
  loadGlue: () => Promise<OdsGlue>,
  resolveWasm: (wasmUrl: string) => Promise<unknown>,
): Readonly<{
  open(
    bytes: Uint8Array,
    wasmUrl: string,
    context: ModelSourceOpenContext | undefined,
    signal?: AbortSignal,
  ): Promise<OwnedOdsSource>;
}> {
  const runtime = createDirectSourceRuntime<OdsGlue, OdsNativeArchive, ModelSourceOpenContext>({
    label: 'OpenDocument ODS',
    mapNativeError: (error) => odfNativeRejection(error, 'ods'),
    loadGlue,
    resolveWasm,
    // The shared governor's WASM encoding: absent → standard limit, 0 → the
    // caller removed the limit (hard ceilings still apply).
    construct: (glue, bytes, context) => new glue.OdsWorkbook(
      bytes,
      wasmLimit(context?.resourcePolicy.maxArchiveEntryBytes),
      wasmLimit(context?.resourcePolicy.maxTotalInflatedBytes),
      wasmLimit(context?.resourcePolicy.maxArchiveEntries),
    ),
    closeNative: (archive) => archive.close_workbook_session(),
  });
  return Object.freeze({
    open: (bytes, wasmUrl, context, signal) => runtime.open(bytes, wasmUrl, signal, context),
  });
}

function wasmLimit(value: number | null | undefined): number | undefined {
  return value === null ? 0 : value;
}

export { resolveDirectWasmInput };
