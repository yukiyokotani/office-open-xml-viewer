import type {
  DeepReadonly,
  PaintResourceDescriptor,
  PaintResourceDescriptorKind,
  PaintResourceRegistry,
} from './types.js';
import { isDeepFrozenPlainDataRoot, sealPlainData, snapshotPlainData } from './plain-data.js';

function assertNonEmptyString(value: string, path: string): void {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new TypeError(`${path} must be a non-empty string`);
  }
}

function assertFiniteNonNegative(value: number, path: string): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new TypeError(`${path} must be finite and non-negative`);
  }
}

function assertFinite(value: number, path: string): void {
  if (!Number.isFinite(value)) {
    throw new TypeError(`${path} must be finite`);
  }
}

function assertUnitInterval(value: number, path: string): void {
  if (!Number.isFinite(value) || value < 0 || value > 1) {
    throw new TypeError(`${path} must be between 0 and 1`);
  }
}

function validateIntrinsicSize(
  size: Readonly<{ widthPt: number; heightPt: number }>,
  path: string,
): void {
  assertFiniteNonNegative(size.widthPt, `${path}.widthPt`);
  assertFiniteNonNegative(size.heightPt, `${path}.heightPt`);
}

function validateDescriptor(descriptor: DeepReadonly<PaintResourceDescriptor>): void {
  assertNonEmptyString(descriptor.resourceKey, 'resourceKey');
  switch (descriptor.kind) {
    case 'image':
    case 'picture-bullet': {
      if (descriptor.documentOrder !== undefined
        && (!Number.isSafeInteger(descriptor.documentOrder) || descriptor.documentOrder < 0)) {
        throw new TypeError('documentOrder must be a non-negative safe integer');
      }
      assertNonEmptyString(descriptor.partPath, 'partPath');
      assertNonEmptyString(descriptor.mimeType, 'mimeType');
      if (descriptor.svgImagePath !== undefined) {
        assertNonEmptyString(descriptor.svgImagePath, 'svgImagePath');
      }
      validateIntrinsicSize(descriptor.intrinsicSize, 'intrinsicSize');
      if (descriptor.alpha !== undefined) assertUnitInterval(descriptor.alpha, 'alpha');
      if (descriptor.rotation !== undefined && !Number.isFinite(descriptor.rotation)) {
        throw new TypeError('rotation must be finite');
      }
      if (descriptor.srcRect !== undefined) {
        // CT_RelativeRect edges are ST_Percentage, which is signed and unbounded.
        assertFinite(descriptor.srcRect.l, 'srcRect.l');
        assertFinite(descriptor.srcRect.t, 'srcRect.t');
        assertFinite(descriptor.srcRect.r, 'srcRect.r');
        assertFinite(descriptor.srcRect.b, 'srcRect.b');
      }
      break;
    }
    case 'chart':
      validateIntrinsicSize(descriptor.intrinsicSize, 'intrinsicSize');
      break;
    case 'math':
      break;
    default: {
      const exhaustive: never = descriptor;
      throw new TypeError(`Unknown paint resource kind: ${String(exhaustive)}`);
    }
  }
}

function snapshotDescriptor(
  descriptor: PaintResourceDescriptor,
): DeepReadonly<PaintResourceDescriptor> {
  validateDescriptor(descriptor);
  return snapshotPlainData(descriptor, `paint resource ${descriptor.resourceKey}`);
}

function kindMismatch(
  resourceKey: string,
  expectedKind: PaintResourceDescriptorKind,
  actualKind: PaintResourceDescriptorKind,
): Error {
  return new Error(
    `Paint resource kind mismatch for ${resourceKey}: expected ${expectedKind}, got ${actualKind}`,
  );
}

export function createPaintResourceRegistry(
  descriptors: readonly PaintResourceDescriptor[],
): PaintResourceRegistry {
  const seen = new Set<string>();
  for (const descriptor of descriptors) {
    if (seen.has(descriptor.resourceKey)) {
      throw new Error(`Duplicate paint resource key: ${descriptor.resourceKey}`);
    }
    seen.add(descriptor.resourceKey);
  }
  const snapshot = descriptors
    .map(snapshotDescriptor)
    .sort((left, right) => left.resourceKey.localeCompare(right.resourceKey));
  const frozenDescriptors = Object.freeze(snapshot);
  return indexSealedPaintResourceDescriptors(frozenDescriptors);
}

/** Index descriptors already validated/sealed by their owning source store.
 * This path intentionally performs no payload clone. */
export function indexSealedPaintResourceDescriptors(
  descriptors: readonly DeepReadonly<PaintResourceDescriptor>[],
): PaintResourceRegistry {
  if (!Object.isFrozen(descriptors)) {
    throw new TypeError('Owned paint descriptors must be sealed');
  }
  let previousKey: string | null = null;
  for (const descriptor of descriptors) {
    validateDescriptor(descriptor);
    if (!Object.isFrozen(descriptor)) {
      throw new TypeError(`Owned paint descriptor must be sealed: ${descriptor.resourceKey}`);
    }
    if ((descriptor.kind === 'image' || descriptor.kind === 'picture-bullet') && !isDeepFrozenPlainDataRoot(descriptor)) {
      // This immutable graph may have arrived under a cloned/array root.
      // Re-establish descriptor-root authority with the existing full walk;
      // shallow freezing is not a replacement ownership proof.
      sealPlainData(descriptor, `owned image descriptor ${descriptor.resourceKey}`);
    }
    if (previousKey !== null && descriptor.resourceKey.localeCompare(previousKey) <= 0) {
      throw new Error(`Owned paint descriptors must have unique sorted keys: ${descriptor.resourceKey}`);
    }
    previousKey = descriptor.resourceKey;
  }
  const frozenDescriptors = descriptors;
  const byKey = new Map(frozenDescriptors.map((descriptor) => [descriptor.resourceKey, descriptor]));
  const keys = Object.freeze(frozenDescriptors.map((descriptor) => descriptor.resourceKey));
  return Object.freeze({
    keys,
    descriptors: frozenDescriptors,
    resolve<K extends PaintResourceDescriptorKind>(resourceKey: string, expectedKind: K) {
      const descriptor = byKey.get(resourceKey);
      if (!descriptor) throw new Error(`Unknown paint resource: ${resourceKey}`);
      if (descriptor.kind !== expectedKind) {
        throw kindMismatch(resourceKey, expectedKind, descriptor.kind);
      }
      return descriptor as DeepReadonly<Extract<PaintResourceDescriptor, { kind: K }>>;
    },
  });
}
