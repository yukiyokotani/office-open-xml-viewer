import type { LegacyPptDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-ppt-source';
import type { PptxNodeArchive } from '@silurus/ooxml-pptx/internal/session';

export interface OwnedLegacyPptSource {
  readonly archive: PptxNodeArchive;
  readonly sourceByteLength: number;
  closeArchive(): void;
}

export declare function openLegacyPptSource(
  bytes: Uint8Array,
  descriptor: LegacyPptDirectSourceDescriptor,
  signal?: AbortSignal,
): Promise<OwnedLegacyPptSource>;
