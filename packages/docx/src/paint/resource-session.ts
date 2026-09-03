import type {
  DeepReadonly,
  PaintResourceDescriptor,
  PaintResourceDescriptorKind,
  PaintResourceRegistry,
} from '../layout/types.js';

export interface OpaquePaintResourceHandle {
  readonly resourceKey: string;
  readonly kind: PaintResourceDescriptorKind;
  readonly handle: unknown;
}

/** A descriptor exists, but this realm deliberately has no drawable. Missing
 * entries remain contract failures so acquisition bugs cannot look optional. */
export type UnavailablePaintResourceHandle = Readonly<{
  status: 'unavailable';
  reason: string;
  placeholder?: 'tiff';
}>;

export type ResolvedPaintResource<K extends PaintResourceDescriptorKind> = Readonly<{
  descriptor: DeepReadonly<Extract<PaintResourceDescriptor, { kind: K }>>;
  handle: unknown;
}>;

export interface PaintResourceSession {
  readonly keys: readonly string[];
  resolve<K extends PaintResourceDescriptorKind>(
    resourceKey: string,
    expectedKind: K,
  ): ResolvedPaintResource<K>;
}

export type PaintResourceHandleResolver = (
  descriptor: DeepReadonly<PaintResourceDescriptor>,
) => unknown | undefined;

function assertNonEmptyString(value: unknown, path: string): asserts value is string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new TypeError(`${path} must be a non-empty string`);
  }
}

export function unavailablePaintResourceHandle(
  reason: string,
  options: Readonly<{ placeholder?: 'tiff' }> = {},
): UnavailablePaintResourceHandle {
  assertNonEmptyString(reason, 'unavailable paint resource reason');
  return Object.freeze({ status: 'unavailable', reason, ...options });
}

export function isUnavailablePaintResourceHandle(
  handle: unknown,
): handle is UnavailablePaintResourceHandle {
  return typeof handle === 'object' && handle !== null
    && (handle as { status?: unknown }).status === 'unavailable'
    && typeof (handle as { reason?: unknown }).reason === 'string'
    && (handle as { reason: string }).reason.trim().length > 0
    && ((handle as { placeholder?: unknown }).placeholder === undefined
      || (handle as { placeholder?: unknown }).placeholder === 'tiff');
}

function assertValidOpaqueHandle(handle: unknown): void {
  if (typeof handle !== 'object' || handle === null
    || (handle as { status?: unknown }).status !== 'unavailable') return;
  // The discriminant is reserved so malformed lookalikes cannot be mistaken
  // for either a valid drawable or an intentional no-draw decision.
  assertNonEmptyString(
    (handle as { reason?: unknown }).reason,
    'unavailable paint resource reason',
  );
  const placeholder = (handle as { placeholder?: unknown }).placeholder;
  if (placeholder !== undefined && placeholder !== 'tiff') {
    throw new TypeError('unavailable paint resource placeholder must be tiff when supplied');
  }
}

/** Browser-owned handles are render-session state. Keeping them here prevents
 * ImageBitmap/CanvasImageSource identity from entering cloneable layout data. */
export function createPaintResourceSession(
  registry: PaintResourceRegistry,
  entries: readonly OpaquePaintResourceHandle[],
): PaintResourceSession {
  const handles = new Map<string, Readonly<{
    kind: PaintResourceDescriptorKind;
    handle: unknown;
  }>>();
  for (const entry of entries) {
    if (handles.has(entry.resourceKey)) {
      throw new Error(`Duplicate paint resource handle: ${entry.resourceKey}`);
    }
    assertValidOpaqueHandle(entry.handle);
    registry.resolve(entry.resourceKey, entry.kind);
    handles.set(entry.resourceKey, Object.freeze({ kind: entry.kind, handle: entry.handle }));
  }
  const keys = Object.freeze([...handles.keys()].sort());
  return Object.freeze({
    keys,
    resolve<K extends PaintResourceDescriptorKind>(
      resourceKey: string,
      expectedKind: K,
    ): ResolvedPaintResource<K> {
      const descriptor = registry.resolve(resourceKey, expectedKind);
      const entry = handles.get(resourceKey);
      if (!entry) {
        throw new Error(
          `Missing paint resource handle for ${resourceKey}: expected ${expectedKind}`,
        );
      }
      if (entry.kind !== expectedKind) {
        throw new Error(
          `Paint resource kind mismatch for ${resourceKey}: expected ${expectedKind}, got ${entry.kind}`,
        );
      }
      return Object.freeze({ descriptor, handle: entry.handle });
    },
  });
}

export function createProductionPaintResourceSession(
  registry: PaintResourceRegistry,
  resolveHandle: PaintResourceHandleResolver,
): PaintResourceSession {
  const entries: OpaquePaintResourceHandle[] = registry.descriptors.map((descriptor) => {
    if (descriptor.kind === 'chart') {
      return { resourceKey: descriptor.resourceKey, kind: descriptor.kind, handle: null };
    }
    const handle = resolveHandle(descriptor);
    if (handle === undefined || handle === null) {
      throw new Error(`Missing ${descriptor.kind} paint handle for ${descriptor.resourceKey}`);
    }
    return { resourceKey: descriptor.resourceKey, kind: descriptor.kind, handle };
  });
  return createPaintResourceSession(registry, entries);
}
