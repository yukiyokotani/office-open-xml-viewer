import { normalizeLoadResourceOptions, parseTypedParserError } from '@silurus/ooxml-core/worker';
import {
  beginModelSourceLoad,
  effectiveModelSources,
  openModelSourceModule,
  selectModelSource,
  type ModelSource,
  type ModelSourceOpenContext,
  type ModelSourceTarget,
  type OpenedModelSourceModule,
} from '@silurus/ooxml-core/internal/model-source';
import { resolveOoxmlContainer, type OoxmlResourceLimits } from '@silurus/ooxml-core';

/** One Node input: an opened model-source archive, or resolved OOXML bytes. */
export type NodeSessionInput<TArchive> =
  | Readonly<{ kind: 'model-source'; opened: OpenedModelSourceModule<TArchive>; sourceByteLength: number }>
  | Readonly<{ kind: 'ooxml'; bytes: Uint8Array }>;

/**
 * Resolve a Node session input. The first configured model source (the
 * format's named source options, in its fixed order) that claims the raw bytes opens
 * them in this realm, receiving the session's normalized resource policy;
 * otherwise the bytes are resolved as an OOXML container (decrypting an
 * Agile-encrypted package with `password`). Selection happens once.
 */
export async function resolveNodeSessionInput<TArchive>(
  buffer: ArrayBuffer | Uint8Array,
  target: ModelSourceTarget,
  options: Readonly<{
    sources?: readonly (ModelSource | undefined)[];
    password?: string;
    signal?: AbortSignal;
    resourceLimits?: OoxmlResourceLimits;
    maxZipEntryBytes?: number;
    worksheetPolicy?: ModelSourceOpenContext['worksheetPolicy'];
  }>,
  validateArchive: (archive: unknown) => TArchive,
  normalizeError: (error: unknown) => unknown = error => parseTypedParserError(error) ?? error,
): Promise<NodeSessionInput<TArchive>> {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const selected = selectModelSource(effectiveModelSources(...(options.sources ?? [])), target, bytes);
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
      {
        resourcePolicy: normalizeLoadResourceOptions(options).policy,
        ...(options.worksheetPolicy === undefined ? {} : { worksheetPolicy: options.worksheetPolicy }),
      },
    );
    try {
      load.release();
    } catch (error) {
      try { opened.close(); } catch {}
      throw error;
    }
    return {
      kind: 'model-source',
      opened: {
        ...opened,
        archive: normalizeSourceArchive(opened.archive, normalizeError),
        close() {
          try { opened.close(); } catch (error) { throw normalizeError(error); }
        },
      },
      sourceByteLength: bytes.byteLength,
    };
  } catch (error) {
    throw normalizeError(error);
  } finally {
    load.release();
  }
}

/** Browser workers normalize declared failures when serializing their responses.
 * Node has no wire boundary, so guard lazy source calls here, once for every
 * format. A format may supply its parser-private decoder without teaching this
 * transport those reasons. Native receivers and immutable archive methods
 * remain supported. */
function normalizeSourceArchive<T>(raw: T, normalizeError: (error: unknown) => unknown): T {
  const target = raw as object;
  return new Proxy(Object.create(target) as object, {
    get(_facade, key) {
      try {
        const value = Reflect.get(target, key, target);
        if (typeof value !== 'function') return value;
        return (...args: unknown[]) => {
          try { return Reflect.apply(value, target, args); }
          catch (error) { throw normalizeError(error); }
        };
      } catch (error) { throw normalizeError(error); }
    },
  }) as T;
}
