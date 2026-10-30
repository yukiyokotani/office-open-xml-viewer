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
import type { DocxDocumentCursorArchive } from '../document-pull-worker.js';

export interface WorkerDocumentArchive extends DocxDocumentCursorArchive {
  assert_healthy(): void;
  extract_image(path: string): Uint8Array;
}

/**
 * The document archive a DOCX model source supplies. It exposes the same pull
 * cursor and image reads as the DOCX parser archive; the remaining methods are
 * optional capabilities. Without `resource_usage` a load reports metrics
 * without a usage snapshot, and without `to_markdown` `toMarkdown()` rejects.
 */
export interface DocxModelSourceArchive extends WorkerDocumentArchive {
  image_mime_type?(path: string): string;
  resource_usage?(): Uint8Array;
  to_markdown?(): string;
}

export interface OoxmlWorkerDocumentArchive extends WorkerDocumentArchive {
  free(): void;
  resource_usage(): Uint8Array;
  to_markdown(): string;
}

/** The view preferences a DOCX model source may report. */
export interface DocxModelSourceViewDefaults {
  readonly showTrackedChanges?: boolean;
}

const REQUIRED_ARCHIVE_METHODS = [
  'open_document_cursor',
  'pull_document_chunk',
  'document_chunk_done',
  'acknowledge_document_chunk',
  'cancel_document_cursor',
  'close_document_session',
  'assert_healthy',
  'extract_image',
] as const;

export function validateDocxModelSourceArchive(value: unknown): DocxModelSourceArchive {
  requireModelSourceArchiveMethods(value, 'DOCX', REQUIRED_ARCHIVE_METHODS);
  return value as DocxModelSourceArchive;
}

/** Admit exactly `{ showTrackedChanges?: boolean }`; any other key fails closed. */
export function validateDocxModelSourceViewDefaults(
  value: Readonly<Record<string, boolean>>,
): DocxModelSourceViewDefaults {
  for (const key of Object.keys(value)) {
    if (key !== 'showTrackedChanges') {
      throw new TypeError(`unsupported DOCX model source view default: ${key}`);
    }
  }
  return value.showTrackedChanges === undefined
    ? {}
    : { showTrackedChanges: value.showTrackedChanges };
}

type OpenModelSource = (
  descriptor: ModelSourceModuleDescriptor,
  bytes: Uint8Array,
  transfer: readonly Transferable[],
  context?: ModelSourceOpenContext,
) => Promise<OpenedModelSourceModule<DocxModelSourceArchive>>;

const openDocxModelSource: OpenModelSource = (descriptor, bytes, transfer, context) =>
  openModelSourceModule(descriptor, bytes, validateDocxModelSourceArchive, undefined, transfer, context);

interface OwnedModelSource {
  readonly archive: DocxModelSourceArchive;
  readonly transferArchive: DocxModelSourceArchive;
  readonly viewDefaults: DocxModelSourceViewDefaults;
  close(): void;
}

/**
 * Owns the one worker-local document archive: either the DOCX parser archive
 * held by `ooxmlHost`, or an archive opened by an application-supplied model
 * source. OOXML calls run under the parser host's trap boundary; a trap from
 * a model-source archive closes that source.
 */
export class WorkerDocumentSourceOwner<TArchive extends OoxmlWorkerDocumentArchive> {
  private modelSource: OwnedModelSource | undefined;
  private pending: object | undefined;

  copyBytes(value: Uint8Array): ArrayBuffer {
    return copyModelSourceBytes(value);
  }

  extractImage(path: string): ArrayBuffer {
    return this.execute((archive) => this.copyBytes(archive.extract_image(path)));
  }

  constructor(
    private readonly ooxmlHost: WasmParserHost<TArchive>,
    private readonly openSource: OpenModelSource = openDocxModelSource,
  ) {}

  get kind(): 'ooxml' | 'model-source' {
    return this.modelSource ? 'model-source' : 'ooxml';
  }

  /** Open a model source; resolves to the view defaults it reported. */
  async openModelSource(
    bytes: Uint8Array,
    descriptor: ModelSourceModuleDescriptor,
    transfer: readonly Transferable[] = [],
    context?: ModelSourceOpenContext,
  ): Promise<DocxModelSourceViewDefaults> {
    if (this.modelSource || this.pending || this.ooxmlHost.archive) {
      throw new Error('Document source already loaded or opening');
    }
    const token = {};
    this.pending = token;
    try {
      const opened = await this.openSource(descriptor, bytes, transfer, context);
      let viewDefaults: DocxModelSourceViewDefaults;
      try {
        viewDefaults = validateDocxModelSourceViewDefaults(opened.viewDefaults);
        if (this.pending !== token || this.ooxmlHost.archive) {
          throw new Error('Document source open was superseded');
        }
      } catch (error) {
        try { opened.close(); } catch {}
        throw error;
      }
      // Normalize at the source boundary. The ordinary pull coordinator can
      // keep its original transfer path and never inspect source ownership.
      const archive = copyingCursorArchive(opened.archive, 'pull_document_chunk');
      this.modelSource = { archive: opened.archive, transferArchive: archive, viewDefaults, close: opened.close };
      return viewDefaults;
    } finally {
      if (this.pending === token) this.pending = undefined;
    }
  }

  get viewDefaults(): DocxModelSourceViewDefaults {
    return this.modelSource?.viewDefaults ?? {};
  }

  cursor(): WorkerDocumentArchive | null {
    return this.modelSource?.archive ?? this.ooxmlHost.archive;
  }

  execute<T>(operation: (archive: WorkerDocumentArchive) => T): T {
    const source = this.modelSource;
    if (!source) {
      const archive = this.ooxmlHost.archive;
      if (!archive) throw new Error('Document not loaded');
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
    this.pending = undefined;
    const source = this.modelSource;
    this.modelSource = undefined;
    source?.close();
  }

  private ooxml(): TArchive {
    const archive = this.ooxmlHost.archive;
    if (!archive) throw new Error('Document not loaded');
    return archive;
  }
}
