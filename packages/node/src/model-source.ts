import {
  beginModelSourceLoad,
  openModelSourceModule,
  resolveOoxmlContainer,
  selectModelSource,
  type ModelSource,
  type ModelSourceTarget,
  type OpenedModelSourceModule,
} from '@silurus/ooxml-core';

/** One Node input: an opened model-source archive, or resolved OOXML bytes. */
export type NodeSessionInput<TArchive> =
  | Readonly<{ kind: 'model-source'; opened: OpenedModelSourceModule<TArchive>; sourceByteLength: number }>
  | Readonly<{ kind: 'ooxml'; bytes: Uint8Array }>;

/**
 * Resolve a Node session input. The first configured model source that claims
 * the raw bytes opens them in this realm; otherwise the bytes are resolved as
 * an OOXML container (decrypting an Agile-encrypted package with `password`).
 * Without `modelSources` nothing is imported and the OOXML path is unchanged.
 */
export async function resolveNodeSessionInput<TArchive>(
  buffer: ArrayBuffer | Uint8Array,
  target: ModelSourceTarget,
  options: Readonly<{
    modelSources?: readonly ModelSource[];
    password?: string;
    signal?: AbortSignal;
  }>,
  validateArchive: (archive: unknown) => TArchive,
): Promise<NodeSessionInput<TArchive>> {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const selected = options.modelSources === undefined
    ? undefined
    : selectModelSource(options.modelSources, target, bytes);
  if (!selected) {
    return { kind: 'ooxml', bytes: await resolveOoxmlContainer(bytes, options.password) };
  }
  const load = beginModelSourceLoad(selected, target);
  try {
    const opened = await openModelSourceModule(
      load.module,
      bytes,
      validateArchive,
      options.signal,
      load.transfer,
    );
    return { kind: 'model-source', opened, sourceByteLength: bytes.byteLength };
  } finally {
    load.release();
  }
}
