import { isWasmTrap, type WasmParserHost } from '@silurus/ooxml-core';
import type { LegacyPptDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-ppt-source';
import type {
  LegacyPptNativeArchive,
  OwnedLegacyPptSource,
} from '@silurus/ooxml-legacy-converter/internal/direct-ppt-engine';
import type { PptxSlideCursorArchive } from '../slide-cursor-operation.js';

export interface WorkerCursorArchive extends PptxSlideCursorArchive {
  assert_healthy(): void;
  presentation_bootstrap(): Uint8Array;
  extract_image(path: string): Uint8Array;
}

export interface OoxmlWorkerArchive extends WorkerCursorArchive {
  free(): void;
  extract_media(path: string): Uint8Array;
  extract_font(path: string): Uint8Array;
  resource_usage(): Uint8Array;
  to_markdown(): string;
}

/** Owns exactly one worker-local source without conflating independent WASM runtimes. */
export class WorkerPresentationSourceOwner<TArchive extends OoxmlWorkerArchive> {
  private legacy: OwnedLegacyPptSource | undefined;

  constructor(private readonly ooxmlHost: WasmParserHost<TArchive>) {}

  get kind(): 'ooxml' | 'legacy-ppt' {
    return this.legacy ? 'legacy-ppt' : 'ooxml';
  }

  async openLegacy(
    bytes: Uint8Array,
    descriptor: LegacyPptDirectSourceDescriptor,
  ): Promise<LegacyPptNativeArchive> {
    if (this.legacy || this.ooxmlHost.archive) throw new Error('Presentation source already loaded');
    const { openLegacyPptSource } = await import(
      '@silurus/ooxml-legacy-converter/internal/direct-ppt-engine'
    );
    const owned = await openLegacyPptSource(bytes, descriptor);
    this.legacy = owned;
    return owned.archive;
  }

  cursor(): WorkerCursorArchive | null {
    return this.legacy?.archive ?? this.ooxmlHost.archive;
  }

  execute<T>(operation: (archive: WorkerCursorArchive) => T): T {
    const archive = this.cursor();
    if (!archive) throw new Error('Presentation not loaded');
    if (!this.legacy) return this.ooxmlHost.run(() => operation(archive));
    try {
      return operation(archive);
    } catch (error) {
      if (isWasmTrap(error)) {
        try { this.closeLegacy(); } catch {}
      }
      throw error;
    }
  }

  ooxml(operation: string): TArchive {
    if (this.legacy) {
      throw new Error(`${operation} is unsupported for direct legacy PPT sources`);
    }
    const archive = this.ooxmlHost.archive;
    if (!archive) throw new Error('Presentation not loaded');
    return archive;
  }

  closeLegacy(): void {
    const owned = this.legacy;
    this.legacy = undefined;
    owned?.closeArchive();
  }
}
