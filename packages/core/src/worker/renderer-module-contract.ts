/** Internal structured-clone-safe contract used to reconstruct first-party
 * optional renderers inside a dedicated render worker. Functions never cross
 * the worker boundary. This transport detail intentionally stays out of the
 * public renderer interfaces. */
const WORKER_RENDERER_MODULE_PROTOCOL = 'ooxml-worker-renderer-module/v1' as const;

export type WorkerBuiltinRendererName = 'math' | 'threeD' | 'regionMap' | 'chartEx';

interface WorkerBuiltinRendererDescriptorBase {
  readonly protocol: typeof WORKER_RENDERER_MODULE_PROTOCOL;
  /** Stable first-party renderer identity resolved by a worker-local lazy import. */
  readonly builtin: WorkerBuiltinRendererName;
}

interface WorkerMathRendererDescriptor extends WorkerBuiltinRendererDescriptorBase {
  readonly builtin: 'math';
  /** Absolute URL resolved by the consumer bundler in the main realm. */
  readonly engineAssetUrl: string;
}

interface WorkerChartRendererDescriptor extends WorkerBuiltinRendererDescriptorBase {
  readonly builtin: 'threeD' | 'regionMap' | 'chartEx';
}

export type WorkerRendererDescriptor =
  | WorkerMathRendererDescriptor
  | WorkerChartRendererDescriptor;

export interface WorkerRendererDescriptors {
  readonly math?: WorkerRendererDescriptor;
  readonly threeD?: WorkerRendererDescriptor;
  readonly regionMap?: WorkerRendererDescriptor;
  readonly chartEx?: WorkerRendererDescriptor;
}

export interface WorkerRendererSources {
  readonly math?: object;
  readonly threeD?: object;
  readonly regionMap?: object;
  readonly chartEx?: object;
}

const workerRendererRegistry = new WeakMap<object, WorkerRendererDescriptor>();

/** Create a bundler-stable descriptor for a built-in optional renderer. */
function createBuiltinWorkerRendererDescriptor(
  builtin: 'math',
  engineAssetUrl: string,
): WorkerMathRendererDescriptor;
function createBuiltinWorkerRendererDescriptor(
  builtin: 'threeD' | 'regionMap' | 'chartEx',
): WorkerChartRendererDescriptor;
function createBuiltinWorkerRendererDescriptor(
  builtin: WorkerBuiltinRendererName,
  engineAssetUrl?: string,
): WorkerRendererDescriptor {
  if (builtin === 'math') {
    if (!engineAssetUrl) throw new TypeError('Math worker renderer requires an engine asset URL');
    return Object.freeze({
      protocol: WORKER_RENDERER_MODULE_PROTOCOL,
      builtin,
      engineAssetUrl,
    });
  }
  return Object.freeze({ protocol: WORKER_RENDERER_MODULE_PROTOCOL, builtin });
}

export function registerBuiltinWorkerRenderer<T extends object>(
  renderer: T,
  builtin: 'math',
  options: { readonly engineAssetUrl: string },
): T;
export function registerBuiltinWorkerRenderer<T extends object>(
  renderer: T,
  builtin: 'threeD' | 'regionMap' | 'chartEx',
): T;
export function registerBuiltinWorkerRenderer<T extends object>(
  renderer: T,
  builtin: WorkerBuiltinRendererName,
  options?: { readonly engineAssetUrl: string },
): T {
  const descriptor = builtin === 'math'
    ? createBuiltinWorkerRendererDescriptor(builtin, options?.engineAssetUrl ?? '')
    : createBuiltinWorkerRendererDescriptor(builtin);
  workerRendererRegistry.set(renderer, descriptor);
  return renderer;
}

export function assertWorkerRendererDescriptor(
  descriptor: WorkerRendererDescriptor,
): WorkerRendererDescriptor {
  if (descriptor.protocol !== WORKER_RENDERER_MODULE_PROTOCOL) {
    throw new TypeError(`Unsupported worker renderer protocol: ${String(descriptor.protocol)}`);
  }
  if (descriptor.builtin !== 'math'
    && descriptor.builtin !== 'threeD'
    && descriptor.builtin !== 'regionMap'
    && descriptor.builtin !== 'chartEx') {
    throw new TypeError(`Unsupported built-in worker renderer: ${String(descriptor.builtin)}`);
  }
  if (descriptor.builtin === 'math' && typeof descriptor.engineAssetUrl !== 'string') {
    throw new TypeError('Math worker renderer requires an engine asset URL');
  }
  return descriptor;
}

/** Strip direct function implementations before a load request crosses to a
 * render worker. Returns undefined when no supplied renderer advertises worker
 * support, keeping the ordinary request payload minimal. */
export function workerRendererDescriptors(
  sources: WorkerRendererSources,
): WorkerRendererDescriptors | undefined {
  const math = sources.math ? workerRendererRegistry.get(sources.math) : undefined;
  const threeD = sources.threeD ? workerRendererRegistry.get(sources.threeD) : undefined;
  const regionMap = sources.regionMap ? workerRendererRegistry.get(sources.regionMap) : undefined;
  const chartEx = sources.chartEx ? workerRendererRegistry.get(sources.chartEx) : undefined;
  const descriptors: WorkerRendererDescriptors = {
    ...(math ? { math } : {}),
    ...(threeD ? { threeD } : {}),
    ...(regionMap ? { regionMap } : {}),
    ...(chartEx ? { chartEx } : {}),
  };
  return Object.keys(descriptors).length > 0 ? Object.freeze(descriptors) : undefined;
}
