import type { WasmParserHost } from '@silurus/ooxml-core';
import type { LegacyDocDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-doc-source';
import type {
  LegacyDocNativeDocument,
  OwnedLegacyDocSource,
} from '@silurus/ooxml-legacy-converter/internal/direct-doc-engine';
import type { DocxDocumentCursorArchive } from '../document-pull-worker.js';

export interface WorkerDocumentArchive extends DocxDocumentCursorArchive {
  assert_healthy(): void;
  extract_image(path: string): Uint8Array;
}

export interface OoxmlWorkerDocumentArchive extends WorkerDocumentArchive {
  free(): void;
  resource_usage(): Uint8Array;
  to_markdown(): string;
}

type OpenNativeDocument = (
  bytes: Uint8Array,
  descriptor: LegacyDocDirectSourceDescriptor,
  signal?: AbortSignal,
) => Promise<OwnedLegacyDocSource>;

const openNativeDocument: OpenNativeDocument = async (bytes, descriptor, signal) => {
  const engine = await import('@silurus/ooxml-legacy-converter/internal/direct-doc-engine');
  return engine.openLegacyDocSource(bytes, descriptor, signal);
};

/** Owns one worker-local DOC source while keeping OOXML-only operations explicit. */
export class WorkerDocumentSourceOwner<TArchive extends OoxmlWorkerDocumentArchive> {
  private nativeSource: OwnedLegacyDocSource | undefined;
  private pending: object | undefined;

  constructor(
    private readonly ooxmlHost: WasmParserHost<TArchive>,
    private readonly openDirect: OpenNativeDocument = openNativeDocument,
  ) {}

  get kind(): 'ooxml' | 'legacy-doc' {
    return this.nativeSource ? 'legacy-doc' : 'ooxml';
  }

  async openNative(
    bytes: Uint8Array,
    descriptor: LegacyDocDirectSourceDescriptor,
    signal?: AbortSignal,
  ): Promise<LegacyDocNativeDocument> {
    if (this.nativeSource || this.pending || this.ooxmlHost.archive) {
      throw new Error('Document source already loaded or opening');
    }
    const token = {};
    this.pending = token;
    try {
      const owned = await this.openDirect(bytes, descriptor, signal);
      if (this.pending !== token || this.ooxmlHost.archive || signal?.aborted) {
        try { owned.closeArchive(); } catch {}
        if (signal?.aborted) throw abortError();
        throw new Error('Document source open was superseded');
      }
      this.nativeSource = owned;
      return owned.archive;
    } finally {
      if (this.pending === token) this.pending = undefined;
    }
  }

  cursor(): WorkerDocumentArchive | null {
    return this.nativeSource?.archive ?? this.ooxmlHost.archive;
  }

  execute<T>(operation: (archive: WorkerDocumentArchive) => T): T {
    const archive = this.cursor();
    if (!archive) throw new Error('Document not loaded');
    return this.nativeSource ? operation(archive) : this.ooxmlHost.run(() => operation(archive));
  }

  ooxml(operation: string): TArchive {
    if (this.nativeSource) {
      throw new Error(`${operation} is unsupported for direct legacy DOC sources`);
    }
    const archive = this.ooxmlHost.archive;
    if (!archive) throw new Error('Document not loaded');
    return archive;
  }

  closeNative(): void {
    this.pending = undefined;
    const owned = this.nativeSource;
    this.nativeSource = undefined;
    owned?.closeArchive();
  }
}

function abortError(): Error {
  const error = new Error('legacy DOC direct source was aborted');
  error.name = 'AbortError';
  return error;
}
