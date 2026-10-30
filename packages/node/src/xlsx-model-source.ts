import {
  acquireXlsxNodeSession,
  normalizeXlsxWorksheetPolicy,
  type XlsxNodeAcquisition,
  type XlsxNodeArchive,
} from '@silurus/ooxml-xlsx/internal/session';
import { WorksheetPullWorker as SourcePullWorker } from '@silurus/ooxml-xlsx/internal/source-pull-worker';
import { GridGeometry } from '@silurus/ooxml-xlsx/internal/grid-geometry';
import type { OoxmlNodeSessionOptions } from './session-options.ts';
import { resolveNodeSessionInput } from './model-source.ts';
import { XlsxWorkbookSessionImpl, type OpenXlsxWorkbookOptions, type XlsxWorksheetRowChunk } from './xlsx.ts';

export async function openXlsxSource(
  buffer: ArrayBuffer | Uint8Array,
  options: OpenXlsxWorkbookOptions,
  wasmModule: () => WebAssembly.Module,
): Promise<XlsxWorkbookSessionImpl> {
  const acquired = await acquireXlsxInput(buffer, options, wasmModule);
  return new SourceXlsxWorkbookSession(acquired, options.signal);
}

class SourceXlsxWorkbookSession extends XlsxWorkbookSessionImpl {
  constructor(acquired: XlsxNodeAcquisition, signal: AbortSignal | undefined) {
    super(acquired.closeArchive, acquired.archive as XlsxNodeArchive,
      acquired.workbookIndex, acquired.metrics, acquired.usage, signal, acquired.worksheetPolicy);
    const policy = acquired.worksheetPolicy;
    this.pull = new SourcePullWorker(
      () => acquired.archive,
      undefined,
      undefined,
      undefined,
      undefined,
      () => policy,
    );
  }

  override async *worksheetRows(sheetIndex: number): AsyncGenerator<XlsxWorksheetRowChunk, void, void> {
    for await (const unit of super.worksheetRows(sheetIndex)) {
      if (unit.kind === 'finished') {
        // Preserve the Normal-font width used when the selected source placed
        // anchors; the ordinary OOXML session does no host-layout work.
        const mdw = this.workbookIndex.layoutMetrics?.maximumDigitWidth;
        if (mdw !== undefined) GridGeometry.forWorksheet(unit.worksheet, mdw);
      }
      yield unit;
    }
  }
}

/** Selected-source XLSX acquisition and renderer-owned Normal-font layout. */
export async function acquireXlsxInput(
  buffer: ArrayBuffer | Uint8Array,
  options: OoxmlNodeSessionOptions & Pick<OpenXlsxWorkbookOptions, 'factory' | 'legacyXls' | 'ods'>,
  wasmModule: () => WebAssembly.Module,
): Promise<XlsxNodeAcquisition> {
  const worksheetPolicy = normalizeXlsxWorksheetPolicy(options);
  const snapshot = { ...options, xlsxWorksheetLimits: worksheetPolicy.worksheet };
  const {
    acquireXlsxSessionFromArchive,
    configureHostLayout,
    validateXlsxModelSourceArchive,
    validateXlsxModelSourceViewDefaults,
  } = await import('@silurus/ooxml-xlsx/internal/model-source-session');
  const input = await resolveNodeSessionInput(
    buffer, 'xlsx', { ...snapshot, sources: [snapshot.legacyXls, snapshot.ods], worksheetPolicy }, validateXlsxModelSourceArchive,
  );
  if (input.kind === 'ooxml') {
    return acquireXlsxNodeSession(input.bytes, wasmModule(), snapshot);
  }
  const { opened } = input;
  let maximumDigitWidth: number | undefined;
  try {
    validateXlsxModelSourceViewDefaults(opened.viewDefaults);
    maximumDigitWidth = await configureHostLayout(opened.archive, async (font) => {
      if (!options.factory) return undefined;
      const { measureHostLayoutFont } = await import('@silurus/ooxml-xlsx/internal/host-layout-measure');
      const context = options.factory.createCanvas(1, 1).getContext('2d');
      return context
        ? measureHostLayoutFont(font, context as unknown as CanvasRenderingContext2D)
        : undefined;
    });
    if (options.signal?.aborted) {
      const error = new Error('XLSX workbook session was aborted');
      error.name = 'AbortError';
      throw error;
    }
  } catch (error) {
    try { opened.close(); } catch {}
    throw error;
  }
  return acquireXlsxSessionFromArchive({
    archive: opened.archive,
    sourceByteLength: input.sourceByteLength,
    ...(maximumDigitWidth === undefined ? {} : { layoutMetrics: { maximumDigitWidth } }),
    closeArchive: opened.close,
  }, snapshot);
}
