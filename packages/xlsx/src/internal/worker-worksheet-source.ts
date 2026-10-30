import type { WasmParserHost } from '@silurus/ooxml-core';
import {
  copyModelSourceBytes,
  hasModelSourceCapability,
  openModelSourceModule,
  requireModelSourceArchiveMethods,
  unsupportedModelSourceCapability,
  type ModelSourceModuleDescriptor,
  type ModelSourceOpenContext,
  type OpenedModelSourceModule,
} from '@silurus/ooxml-core/internal/model-source';
import {
  copyingCursorArchive,
  runModelSourceOperation,
} from '@silurus/ooxml-core/internal/worker-model-source';
import type { WorksheetCursorArchive } from '../worksheet-pull-source-worker.js';
import {
  configureHostLayout,
  isHostLayoutResult,
  requestHostLayoutFromPage,
  type HostLayoutArchive,
  type HostLayoutFont,
} from './host-layout.js';

export interface WorkerWorksheetArchive extends WorksheetCursorArchive {
  assert_healthy(): void;
  parse(): Uint8Array;
  extract_image(path: string): Uint8Array;
}

/**
 * The workbook archive an XLSX model source supplies: the XLSX parser
 * archive's workbook bootstrap, worksheet cursor and image reads. Optional
 * capabilities: without `resource_usage` / `sheet_cursor_resource_usage`
 * metrics and worksheet pulls carry no usage snapshot,
 * without `to_markdown` `toMarkdown()` rejects, and `host_layout_request` /
 * `configure_host_layout` ask the host for the Normal font's maximum digit
 * width (see host-layout.ts).
 */
export interface XlsxModelSourceArchive extends WorkerWorksheetArchive, HostLayoutArchive {
  resource_usage?(): Uint8Array;
  to_markdown?(): string;
}

export interface OoxmlWorksheetArchive extends WorkerWorksheetArchive {
  free(): void;
  resource_usage(): Uint8Array;
  sheet_cursor_resource_usage(): Uint8Array;
  to_markdown(): string;
}

const REQUIRED_ARCHIVE_METHODS = [
  'open_sheet_cursor',
  'pull_sheet_cursor',
  'sheet_cursor_pull_finished',
  'acknowledge_sheet_cursor_terminal',
  'cancel_sheet_cursor',
  'close_sheet_cursor',
  'assert_healthy',
  'parse',
  'extract_image',
] as const;

export function validateXlsxModelSourceArchive(value: unknown): XlsxModelSourceArchive {
  requireModelSourceArchiveMethods(value, 'XLSX', REQUIRED_ARCHIVE_METHODS);
  return value as XlsxModelSourceArchive;
}

/** XLSX defines no view defaults; any reported key fails closed. */
export function validateXlsxModelSourceViewDefaults(value: Readonly<Record<string, boolean>>): void {
  const [key] = Object.keys(value);
  if (key !== undefined) throw new TypeError(`unsupported XLSX model source view default: ${key}`);
}

type OpenModelSource = (
  descriptor: ModelSourceModuleDescriptor,
  bytes: Uint8Array,
  transfer: readonly Transferable[],
  context?: ModelSourceOpenContext,
) => Promise<OpenedModelSourceModule<XlsxModelSourceArchive>>;

const openXlsxModelSource: OpenModelSource = (descriptor, bytes, transfer, context) =>
  openModelSourceModule(descriptor, bytes, validateXlsxModelSourceArchive, undefined, transfer, context);

interface OwnedModelSource {
  readonly archive: XlsxModelSourceArchive;
  readonly transferArchive: XlsxModelSourceArchive;
  close(): void;
}

/** Own exactly one worker-local workbook archive: XLSX parser or model source. */
export class WorkerWorksheetSourceOwner<TArchive extends OoxmlWorksheetArchive> {
  private modelSource: OwnedModelSource | undefined;
  private measuredMaximumDigitWidth: number | undefined;

  copyBytes(value: Uint8Array): ArrayBuffer {
    return copyModelSourceBytes(value);
  }

  isHostLayoutResult(value: unknown): boolean {
    return isHostLayoutResult(value);
  }

  requestHostLayoutFromPage(
    scope: Parameters<typeof requestHostLayoutFromPage>[0],
    font: HostLayoutFont,
  ): Promise<number | undefined> {
    return requestHostLayoutFromPage(scope, font);
  }

  extractImage(path: string): ArrayBuffer {
    return this.execute((archive) => this.copyBytes(archive.extract_image(path)));
  }

  constructor(
    private readonly ooxmlHost: WasmParserHost<TArchive>,
    private readonly openSource: OpenModelSource = openXlsxModelSource,
  ) {}

  get kind(): 'ooxml' | 'model-source' {
    return this.modelSource ? 'model-source' : 'ooxml';
  }

  /**
   * Open a model source and settle its host layout before any bootstrap or
   * sheet cursor exists. `measure` is the host's renderer-owned measurement.
   */
  async openModelSource(
    bytes: Uint8Array,
    descriptor: ModelSourceModuleDescriptor,
    measure: (font: HostLayoutFont) => number | undefined | Promise<number | undefined>,
    transfer: readonly Transferable[] = [],
    context?: ModelSourceOpenContext,
  ): Promise<void> {
    if (this.modelSource || this.ooxmlHost.archive) throw new Error('Workbook source already loaded');
    const opened = await this.openSource(descriptor, bytes, transfer, context);
    try {
      validateXlsxModelSourceViewDefaults(opened.viewDefaults);
      if (this.modelSource || this.ooxmlHost.archive) {
        throw new Error('Workbook source open was superseded');
      }
    } catch (error) {
      try { opened.close(); } catch {}
      throw error;
    }
    const transferArchive = copyingCursorArchive(opened.archive, 'pull_sheet_cursor');
    const source = { archive: opened.archive, transferArchive, close: opened.close };
    this.modelSource = source;
    try {
      this.measuredMaximumDigitWidth = await configureHostLayout(
        opened.archive,
        measure,
      );
      if (this.modelSource !== source) throw new Error('Workbook source open was superseded');
    } catch (error) {
      if (this.modelSource === source) {
        try { this.closeModelSource(); } catch {}
      }
      throw error;
    }
  }

  /** The width the host configured for the model source, if any. */
  get maximumDigitWidth(): number | undefined {
    return this.measuredMaximumDigitWidth;
  }

  cursor(): WorkerWorksheetArchive | null {
    return this.modelSource?.archive ?? this.ooxmlHost.archive;
  }

  execute<T>(operation: (archive: WorkerWorksheetArchive) => T): T {
    const source = this.modelSource;
    if (!source) {
      const archive = this.ooxmlHost.archive;
      if (!archive) throw new Error('Workbook not loaded');
      return this.ooxmlHost.run(() => operation(archive));
    }
    return runModelSourceOperation(
      () => operation(source.transferArchive),
      () => { if (this.modelSource === source) this.closeModelSource(); },
    );
  }

  /** ZIP resource accounting; `undefined` when the source lacks the capability. */
  resourceUsage(): Uint8Array | undefined {
    const source = this.modelSource;
    if (source) {
      if (!hasModelSourceCapability(source.archive, 'resource_usage')) return undefined;
      return this.execute(() => (source.archive.resource_usage as () => Uint8Array)());
    }
    const archive = this.ooxml();
    return this.ooxmlHost.run(() => archive.resource_usage());
  }

  toMarkdown(): string {
    const source = this.modelSource;
    if (source) {
      if (!hasModelSourceCapability(source.archive, 'to_markdown')) {
        throw unsupportedModelSourceCapability('Markdown conversion');
      }
      return this.execute(() => (source.archive.to_markdown as () => string)());
    }
    const archive = this.ooxml();
    return this.ooxmlHost.run(() => archive.to_markdown());
  }

  closeModelSource(): void {
    const source = this.modelSource;
    this.modelSource = undefined;
    this.measuredMaximumDigitWidth = undefined;
    source?.close();
  }

  private ooxml(): TArchive {
    const archive = this.ooxmlHost.archive;
    if (!archive) throw new Error('Workbook not loaded');
    return archive;
  }
}
