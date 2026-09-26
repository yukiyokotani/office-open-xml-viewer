/**
 * Application-supplied model sources.
 *
 * A model source lets an application open input that is not an OOXML package
 * into a renderer's own model archive: the same pull cursors, image reads and
 * optional capabilities the DOCX / XLSX / PPTX parser WASM archives expose.
 * Core and the format packages know only this contract. They never name,
 * import or special-case a concrete source; which inputs a source accepts, its
 * size limits and its container sniffing belong to the source itself.
 *
 * The split has two realms:
 *
 *  - {@link ModelSource} lives in the calling realm (the page or Node). The
 *    loader asks each configured source, in order, to `claim()` the raw input
 *    bytes. The first source that claims them supplies a
 *    {@link ModelSourceModuleDescriptor} from `beginLoad()`.
 *  - The descriptor is structured-clone safe. The realm that owns the parser
 *    archive (normally a Worker, or Node) imports `moduleUrl` and calls its
 *    `openModelSource()` export ({@link ModelSourceModule}).
 *
 * Security boundary: module URLs come only from application options. Document
 * content never selects, extends or rewrites a module URL or its config.
 */

export type ModelSourceTarget = 'docx' | 'xlsx' | 'pptx';

/** Scalar configuration passed verbatim to a source module. */
export type ModelSourceConfigValue = string | number | boolean | null;

export type ModelSourceConfig = Readonly<Record<string, ModelSourceConfigValue>>;

export const MODEL_SOURCE_MODULE_PROTOCOL = 'ooxml-model-source-module/v1';

/** Structured-clone-safe instructions for loading one source module. */
export interface ModelSourceModuleDescriptor {
  readonly protocol: typeof MODEL_SOURCE_MODULE_PROTOCOL;
  readonly target: ModelSourceTarget;
  /** Absolute http(s), file or blob URL of a self-contained ES module. */
  readonly moduleUrl: string;
  /** Frozen, bounded scalar configuration for the module. */
  readonly config: ModelSourceConfig;
}

/** One admitted load, returned by {@link ModelSource.beginLoad}. */
export interface ModelSourceLoad {
  readonly module: ModelSourceModuleDescriptor;
  /**
   * Objects transferred with the load request and handed to
   * `openModelSource()` as its fourth argument, in order.
   */
  readonly transfer?: readonly Transferable[];
  /** Called once when the load settles, successfully or not. */
  release(): void;
}

/** Calling-realm half of a model source. */
export interface ModelSource<T extends ModelSourceTarget = ModelSourceTarget> {
  readonly target: T;
  /**
   * Synchronous, content-free admission of the raw input bytes. `true` selects
   * this source, `false` leaves the input to the next source or to the
   * ordinary OOXML path. Throwing rejects the load (fail closed).
   */
  claim(bytes: Uint8Array): boolean;
  beginLoad(): ModelSourceLoad;
}

/** What a source module returns for one opened input. */
export interface OpenedModelSource<TArchive> {
  /** The renderer archive (see each format's model-source archive contract). */
  readonly archive: TArchive;
  /**
   * The document's own view preferences, applied only where the caller did
   * not choose explicitly. Each format validates the keys it understands and
   * rejects any other key.
   */
  readonly viewDefaults?: Readonly<Record<string, boolean>>;
  /** Release the archive and every native resource it holds. Idempotent. */
  close(): void;
}

/** Owning-realm half of a model source: the export `moduleUrl` must provide. */
export interface ModelSourceModule<TArchive = unknown> {
  openModelSource(
    bytes: Uint8Array,
    config: ModelSourceConfig,
    signal?: AbortSignal,
    transfer?: readonly Transferable[],
  ): Promise<OpenedModelSource<TArchive>>;
}

const TARGETS: ReadonlySet<string> = new Set<ModelSourceTarget>(['docx', 'xlsx', 'pptx']);
const MODULE_URL_PROTOCOLS: ReadonlySet<string> = new Set(['http:', 'https:', 'file:', 'blob:']);
const DESCRIPTOR_FIELDS = ['protocol', 'target', 'moduleUrl', 'config'] as const;
/** Implementation bounds, not format limits: a descriptor is small metadata. */
const MAX_MODEL_SOURCES = 16;
const MAX_URL_LENGTH = 8192;
const MAX_CONFIG_ENTRIES = 32;
const MAX_CONFIG_STRING_LENGTH = 8192;
const MAX_VIEW_DEFAULT_ENTRIES = 16;
const CONFIG_KEY = /^[A-Za-z][A-Za-z0-9_]{0,63}$/;

/**
 * Pick the first configured source that claims `bytes`, or `undefined` when
 * none is configured or none claims them. A source for another target is a
 * configuration error.
 */
export function selectModelSource<T extends ModelSourceTarget>(
  sources: readonly ModelSource[] | undefined,
  target: T,
  bytes: Uint8Array,
): ModelSource<T> | undefined {
  if (sources === undefined) return undefined;
  if (!Array.isArray(sources)) throw new TypeError('modelSources must be an array');
  if (sources.length > MAX_MODEL_SOURCES) {
    throw new RangeError(`modelSources accepts at most ${MAX_MODEL_SOURCES} sources`);
  }
  for (const source of sources as readonly unknown[]) {
    if (typeof source !== 'object' || source === null) {
      throw new TypeError('each model source must be an object');
    }
    const candidate = source as Partial<ModelSource>;
    if (candidate.target !== target) {
      throw new TypeError(
        `a ${String(candidate.target)} model source cannot load ${target} input`,
      );
    }
    if (typeof candidate.claim !== 'function' || typeof candidate.beginLoad !== 'function') {
      throw new TypeError('a model source must implement claim() and beginLoad()');
    }
    const claimed: unknown = candidate.claim(bytes);
    if (typeof claimed !== 'boolean') {
      throw new TypeError('model source claim() must return a boolean');
    }
    if (claimed) return candidate as ModelSource<T>;
  }
  return undefined;
}

/** Validated, detached result of {@link ModelSource.beginLoad}. */
export interface AdmittedModelSourceLoad {
  readonly module: ModelSourceModuleDescriptor;
  readonly transfer: readonly Transferable[];
  release(): void;
}

/**
 * Start the selected source's load and validate what it returned. The caller
 * must call `release()` exactly when the load settles.
 */
export function beginModelSourceLoad(
  source: ModelSource,
  target: ModelSourceTarget,
): AdmittedModelSourceLoad {
  const load: unknown = source.beginLoad();
  if (typeof load !== 'object' || load === null) {
    throw new TypeError('model source beginLoad() must return an object');
  }
  const record = load as Partial<ModelSourceLoad>;
  const release = record.release;
  if (typeof release !== 'function') {
    throw new TypeError('model source load must provide release()');
  }
  let released = false;
  const releaseOnce = (): void => {
    if (released) return;
    released = true;
    release.call(load);
  };
  try {
    const module = validateModelSourceModuleDescriptor(record.module, target);
    const transfer = record.transfer === undefined ? [] : record.transfer;
    if (!Array.isArray(transfer)) {
      throw new TypeError('model source load transfer must be an array');
    }
    return Object.freeze({
      module,
      transfer: Object.freeze([...transfer]),
      release: releaseOnce,
    });
  } catch (error) {
    try { releaseOnce(); } catch {}
    throw error;
  }
}

/**
 * Validate and detach a descriptor. `expectedTarget` rejects a descriptor for
 * another renderer.
 */
export function validateModelSourceModuleDescriptor(
  value: unknown,
  expectedTarget?: ModelSourceTarget,
): ModelSourceModuleDescriptor {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError('model source module descriptor must be an object');
  }
  const keys = Reflect.ownKeys(value);
  if (
    keys.length !== DESCRIPTOR_FIELDS.length
    || DESCRIPTOR_FIELDS.some((field) => !keys.includes(field))
  ) {
    throw new TypeError('model source module descriptor has unknown or missing fields');
  }
  const fields = Object.getOwnPropertyDescriptors(value);
  if (DESCRIPTOR_FIELDS.some((field) => !fields[field]?.enumerable || !('value' in fields[field]))) {
    throw new TypeError('model source module descriptor fields must be data properties');
  }
  const protocol: unknown = fields.protocol.value;
  const target: unknown = fields.target.value;
  const moduleUrl: unknown = fields.moduleUrl.value;
  if (protocol !== MODEL_SOURCE_MODULE_PROTOCOL) {
    throw new TypeError('unsupported model source module protocol');
  }
  if (typeof target !== 'string' || !TARGETS.has(target)) {
    throw new TypeError('unsupported model source target');
  }
  if (expectedTarget !== undefined && target !== expectedTarget) {
    throw new TypeError(`a ${target} model source module cannot load ${expectedTarget} input`);
  }
  if (
    typeof moduleUrl !== 'string'
    || moduleUrl.length === 0
    || moduleUrl.length > MAX_URL_LENGTH
    || moduleUrl.trim() !== moduleUrl
  ) {
    throw new TypeError('model source moduleUrl must be a nonempty absolute URL');
  }
  let url: URL;
  try {
    url = new URL(moduleUrl);
  } catch {
    throw new TypeError('model source moduleUrl must be a nonempty absolute URL');
  }
  if (!MODULE_URL_PROTOCOLS.has(url.protocol)) {
    throw new TypeError('model source moduleUrl protocol is unsupported');
  }
  return Object.freeze({
    protocol: MODEL_SOURCE_MODULE_PROTOCOL,
    target: target as ModelSourceTarget,
    moduleUrl,
    config: validateConfig(fields.config.value),
  });
}

function validateConfig(value: unknown): ModelSourceConfig {
  if (!isPlainRecord(value)) {
    throw new TypeError('model source config must be a plain object');
  }
  const keys = Reflect.ownKeys(value);
  if (keys.length > MAX_CONFIG_ENTRIES) {
    throw new RangeError(`model source config accepts at most ${MAX_CONFIG_ENTRIES} entries`);
  }
  const config: Record<string, ModelSourceConfigValue> = {};
  for (const key of keys) {
    if (typeof key !== 'string' || !CONFIG_KEY.test(key)) {
      throw new TypeError('model source config keys must be short identifiers');
    }
    const field = Object.getOwnPropertyDescriptor(value, key);
    if (!field?.enumerable || !('value' in field)) {
      throw new TypeError('model source config entries must be data properties');
    }
    const entry: unknown = field.value;
    if (typeof entry === 'string') {
      if (entry.length > MAX_CONFIG_STRING_LENGTH) {
        throw new RangeError('model source config string is too long');
      }
    } else if (typeof entry === 'number') {
      if (!Number.isFinite(entry)) throw new TypeError('model source config numbers must be finite');
    } else if (typeof entry !== 'boolean' && entry !== null) {
      throw new TypeError('model source config values must be strings, numbers, booleans or null');
    }
    config[key] = entry as ModelSourceConfigValue;
  }
  return Object.freeze(config);
}

/** Validated, detached result of a source module's `openModelSource()`. */
export interface OpenedModelSourceModule<TArchive> {
  readonly archive: TArchive;
  readonly viewDefaults: Readonly<Record<string, boolean>>;
  close(): void;
}

/**
 * Import `descriptor.moduleUrl` in the calling realm and open `bytes` with it.
 * `validateArchive` checks the format's archive contract. Any validation
 * failure closes what the module opened before rethrowing.
 */
export async function openModelSourceModule<TArchive>(
  descriptor: ModelSourceModuleDescriptor,
  bytes: Uint8Array,
  validateArchive: (archive: unknown) => TArchive,
  signal?: AbortSignal,
  transfer: readonly Transferable[] = [],
): Promise<OpenedModelSourceModule<TArchive>> {
  const module = validateModelSourceModuleDescriptor(descriptor);
  throwIfAborted(signal);
  const namespace: unknown = await import(/* @vite-ignore */ module.moduleUrl);
  const open = typeof namespace === 'object' && namespace !== null
    ? (namespace as Partial<ModelSourceModule>).openModelSource
    : undefined;
  if (typeof open !== 'function') {
    throw new TypeError('model source module must export openModelSource()');
  }
  throwIfAborted(signal);
  const opened: unknown = await open(bytes, module.config, signal, transfer);
  if (typeof opened !== 'object' || opened === null) {
    throw new TypeError('openModelSource() must return an object');
  }
  const record = opened as Partial<OpenedModelSource<unknown>>;
  const close = record.close;
  if (typeof close !== 'function') {
    throw new TypeError('an opened model source must provide close()');
  }
  let closed = false;
  const closeOnce = (): void => {
    if (closed) return;
    closed = true;
    close.call(opened);
  };
  try {
    const archive = validateArchive(record.archive);
    const viewDefaults = validateViewDefaults(record.viewDefaults);
    throwIfAborted(signal);
    return Object.freeze({ archive, viewDefaults, close: closeOnce });
  } catch (error) {
    try { closeOnce(); } catch {}
    throw error;
  }
}

function validateViewDefaults(value: unknown): Readonly<Record<string, boolean>> {
  if (value === undefined) return Object.freeze({});
  if (!isPlainRecord(value)) {
    throw new TypeError('model source viewDefaults must be a plain object');
  }
  const keys = Reflect.ownKeys(value);
  if (keys.length > MAX_VIEW_DEFAULT_ENTRIES) {
    throw new RangeError('model source viewDefaults has too many entries');
  }
  const result: Record<string, boolean> = {};
  for (const key of keys) {
    const field = typeof key === 'string' ? Object.getOwnPropertyDescriptor(value, key) : undefined;
    if (
      typeof key !== 'string'
      || !CONFIG_KEY.test(key)
      || !field?.enumerable
      || !('value' in field)
      || typeof field.value !== 'boolean'
    ) {
      throw new TypeError('model source viewDefaults entries must be boolean data properties');
    }
    result[key] = field.value;
  }
  return Object.freeze(result);
}

/**
 * Require the archive methods a format needs. Optional capabilities are
 * checked by the caller at their use site.
 */
export function requireModelSourceArchiveMethods(
  archive: unknown,
  label: string,
  methods: readonly string[],
): void {
  if ((typeof archive !== 'object' && typeof archive !== 'function') || archive === null) {
    throw new TypeError(`${label} model source archive must be an object`);
  }
  for (const method of methods) {
    if (typeof (archive as Record<string, unknown>)[method] !== 'function') {
      throw new TypeError(`${label} model source archive must implement ${method}()`);
    }
  }
}

/** Whether an archive offers an optional capability. */
export function hasModelSourceCapability(archive: unknown, method: string): boolean {
  return (typeof archive === 'object' || typeof archive === 'function')
    && archive !== null
    && typeof (archive as Record<string, unknown>)[method] === 'function';
}

/** The error an operation reports when the loaded source lacks a capability. */
export function unsupportedModelSourceCapability(operation: string): Error {
  return new Error(`${operation} is unsupported for this source`);
}

function isPlainRecord(value: unknown): value is object {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const prototype: unknown = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('model source load was aborted');
  error.name = 'AbortError';
  throw error;
}
