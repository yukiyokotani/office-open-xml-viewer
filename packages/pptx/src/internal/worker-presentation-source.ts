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
import type { PptxSlideCursorArchive } from '../slide-cursor-operation.js';

export interface WorkerCursorArchive extends PptxSlideCursorArchive {
  assert_healthy(): void;
  presentation_bootstrap(): Uint8Array;
  extract_image(path: string): Uint8Array;
}

/**
 * The presentation archive a PPTX model source supplies: the PPTX parser
 * archive's bootstrap, slide cursor and image reads. Optional capabilities:
 * without `extract_media` / `extract_font` those reads reject, without
 * `resource_usage` metrics carry no usage snapshot, and without `to_markdown`
 * `toMarkdown()` rejects.
 */
export interface PptxModelSourceArchive extends WorkerCursorArchive {
  extract_media?(path: string): Uint8Array;
  extract_font?(path: string): Uint8Array;
  resource_usage?(): Uint8Array;
  to_markdown?(): string;
}

export interface OoxmlWorkerArchive extends WorkerCursorArchive {
  free(): void;
  extract_media(path: string): Uint8Array;
  extract_font(path: string): Uint8Array;
  resource_usage(): Uint8Array;
  to_markdown(): string;
}

const REQUIRED_ARCHIVE_METHODS = [
  'pull_slide',
  'slide_cursor_resource_usage',
  'acknowledge_slide',
  'cancel_slide',
  'close_presentation_session',
  'assert_healthy',
  'presentation_bootstrap',
  'extract_image',
] as const;

export function validatePptxModelSourceArchive(value: unknown): PptxModelSourceArchive {
  requireModelSourceArchiveMethods(value, 'PPTX', REQUIRED_ARCHIVE_METHODS);
  return value as PptxModelSourceArchive;
}

/** PPTX defines no view defaults; any reported key fails closed. */
export function validatePptxModelSourceViewDefaults(value: Readonly<Record<string, boolean>>): void {
  const [key] = Object.keys(value);
  if (key !== undefined) throw new TypeError(`unsupported PPTX model source view default: ${key}`);
}

type OpenModelSource = (
  descriptor: ModelSourceModuleDescriptor,
  bytes: Uint8Array,
  transfer: readonly Transferable[],
  context?: ModelSourceOpenContext,
) => Promise<OpenedModelSourceModule<PptxModelSourceArchive>>;

const openPptxModelSource: OpenModelSource = (descriptor, bytes, transfer, context) =>
  openModelSourceModule(descriptor, bytes, validatePptxModelSourceArchive, undefined, transfer, context);

interface OwnedModelSource {
  readonly archive: PptxModelSourceArchive;
  readonly transferArchive: PptxModelSourceArchive;
  close(): void;
}

/** Owns exactly one worker-local presentation archive: PPTX parser or model source. */
export class WorkerPresentationSourceOwner<TArchive extends OoxmlWorkerArchive> {
  private modelSource: OwnedModelSource | undefined;

  copyBytes(value: Uint8Array): ArrayBuffer {
    return copyModelSourceBytes(value);
  }

  extractImage(path: string): ArrayBuffer {
    return this.execute((archive) => this.copyBytes(archive.extract_image(path)));
  }

  constructor(
    private readonly ooxmlHost: WasmParserHost<TArchive>,
    private readonly openSource: OpenModelSource = openPptxModelSource,
  ) {}

  get kind(): 'ooxml' | 'model-source' {
    return this.modelSource ? 'model-source' : 'ooxml';
  }

  async openModelSource(
    bytes: Uint8Array,
    descriptor: ModelSourceModuleDescriptor,
    transfer: readonly Transferable[] = [],
    context?: ModelSourceOpenContext,
  ): Promise<PptxModelSourceArchive> {
    if (this.modelSource || this.ooxmlHost.archive) throw new Error('Presentation source already loaded');
    const opened = await this.openSource(descriptor, bytes, transfer, context);
    try {
      validatePptxModelSourceViewDefaults(opened.viewDefaults);
      if (this.modelSource || this.ooxmlHost.archive) {
        throw new Error('Presentation source open was superseded');
      }
    } catch (error) {
      try { opened.close(); } catch {}
      throw error;
    }
    const transferArchive = copyingCursorArchive(opened.archive, 'pull_slide');
    this.modelSource = { archive: opened.archive, transferArchive, close: opened.close };
    return opened.archive;
  }

  cursor(): WorkerCursorArchive | null {
    return this.modelSource?.archive ?? this.ooxmlHost.archive;
  }

  execute<T>(operation: (archive: WorkerCursorArchive) => T): T {
    const source = this.modelSource;
    if (!source) {
      const archive = this.ooxmlHost.archive;
      if (!archive) throw new Error('Presentation not loaded');
      return this.ooxmlHost.run(() => operation(archive));
    }
    return runModelSourceOperation(
      () => operation(source.transferArchive),
      () => { if (this.modelSource === source) this.closeModelSource(); },
    );
  }

  extractMedia(path: string): Uint8Array {
    return this.optional('extract_media', 'media extraction', (archive) => archive.extract_media(path));
  }

  extractFont(path: string): Uint8Array {
    return this.optional('extract_font', 'font extraction', (archive) => archive.extract_font(path));
  }

  /** ZIP resource accounting; `undefined` when the source lacks the capability. */
  resourceUsage(): Uint8Array | undefined {
    const source = this.modelSource;
    if (source && !hasModelSourceCapability(source.archive, 'resource_usage')) return undefined;
    return this.optional('resource_usage', 'resource usage', (archive) => archive.resource_usage());
  }

  toMarkdown(): string {
    return this.optional('to_markdown', 'Markdown conversion', (archive) => archive.to_markdown());
  }

  closeModelSource(): void {
    const source = this.modelSource;
    this.modelSource = undefined;
    source?.close();
  }

  private optional<T>(
    method: keyof PptxModelSourceArchive & keyof OoxmlWorkerArchive,
    operation: string,
    call: (archive: OoxmlWorkerArchive) => T,
  ): T {
    const source = this.modelSource;
    if (source) {
      if (!hasModelSourceCapability(source.archive, method)) {
        throw unsupportedModelSourceCapability(operation);
      }
      return this.execute(() => call(source.archive as unknown as OoxmlWorkerArchive));
    }
    const archive = this.ooxmlHost.archive;
    if (!archive) throw new Error('Presentation not loaded');
    return this.ooxmlHost.run(() => call(archive));
  }
}
