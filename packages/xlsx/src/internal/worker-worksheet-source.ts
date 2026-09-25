import {
  hasModelSourceCapability,
  isWasmTrap,
  openModelSourceModule,
  requireModelSourceArchiveMethods,
  unsupportedModelSourceCapability,
  type ModelSourceModuleDescriptor,
  type OpenedModelSourceModule,
  type WasmParserHost,
} from '@silurus/ooxml-core';
import type { WorksheetCursorArchive } from '../worksheet-pull-worker.js';
import {
  configureHostLayout,
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
 * capabilities: without `resource_usage` metrics carry no usage snapshot,
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
  to_markdown(): string;
}

const REQUIRED_ARCHIVE_METHODS = [
  'open_sheet_cursor',
  'pull_sheet_cursor',
  'sheet_cursor_pull_finished',
  'sheet_cursor_resource_usage',
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
) => Promise<OpenedModelSourceModule<XlsxModelSourceArchive>>;

const openXlsxModelSource: OpenModelSource = (descriptor, bytes, transfer) =>
  openModelSourceModule(descriptor, bytes, validateXlsxModelSourceArchive, undefined, transfer);

interface OwnedModelSource {
  readonly archive: XlsxModelSourceArchive;
  close(): void;
}

/** Own exactly one worker-local workbook archive: XLSX parser or model source. */
export class WorkerWorksheetSourceOwner<TArchive extends OoxmlWorksheetArchive> {
  private modelSource: OwnedModelSource | undefined;
  private measuredMaximumDigitWidth: number | undefined;

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
  ): Promise<void> {
    if (this.modelSource || this.ooxmlHost.archive) throw new Error('Workbook source already loaded');
    const opened = await this.openSource(descriptor, bytes, transfer);
    try {
      validateXlsxModelSourceViewDefaults(opened.viewDefaults);
      if (this.modelSource || this.ooxmlHost.archive) {
        throw new Error('Workbook source open was superseded');
      }
    } catch (error) {
      try { opened.close(); } catch {}
      throw error;
    }
    const source = { archive: opened.archive, close: opened.close };
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
    try {
      return operation(source.archive);
    } catch (error) {
      if (isWasmTrap(error) && this.modelSource === source) {
        try { this.closeModelSource(); } catch {}
      }
      throw error;
    }
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
