import type { DeepReadonly } from './types.js';
import { documentLayoutValidationEnabled } from './validation-policy.js';

/** A plain-data contract violation detected inside the unconditional clone /
 *  freeze walks. Carries the reason without a property path — the path-precise
 *  report is `assertPlainData`'s job, and that pre-pass is development-only.
 *  Fatal state (a non-finite number in retained data) must be detected whether
 *  or not the development pre-pass ran, per the layout engine's error
 *  contract, so these checks are fused into the walks that always run. */
class PlainDataContractError extends TypeError {}

/** Graph roots snapshotted by snapshotPlainData, plus every node sealed in
 * place by sealPlainData. Registered graphs are engine-owned and immutable, so
 * later layout boundaries can reuse them by reference instead of reprocessing
 * the same multi-megabyte structures on every call. */
const processedPlainData = new WeakSet<object>();

/** Object graphs already deeply frozen by deepFreezePlainData. Freeze walks
 * use this only to skip re-walking; it never relaxes validation, because
 * deepFreezePlainData alone does not prove a graph is plain data. */
const frozenPlainData = new WeakSet<object>();

function assertPlainData(
  value: unknown,
  path: string,
  visiting = new WeakSet<object>(),
  completed = new WeakSet<object>(),
): void {
  if (
    value === null
    || value === undefined
    || typeof value === 'string'
    || typeof value === 'boolean'
  ) return;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new TypeError(`${path} must contain finite numbers`);
    return;
  }
  if (typeof value !== 'object') {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  if (visiting.has(value)) {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  if (completed.has(value) || processedPlainData.has(value)) return;
  const prototype = Object.getPrototypeOf(value);
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  if (Object.getOwnPropertySymbols(value).length !== 0) {
    throw new TypeError(`${path} must contain only enumerable string data properties`);
  }
  visiting.add(value);
  try {
    for (const key of Object.getOwnPropertyNames(value)) {
      if (Array.isArray(value) && key === 'length') continue;
      // Plain-data arrays carry index properties only. Enforced here (rather
      // than merely assumed) because `deepFreezePlainData` walks arrays by
      // index: an array with an extra own property would otherwise have that
      // property's subgraph left unfrozen.
      if (Array.isArray(value) && String(Number(key)) !== key) {
        throw new TypeError(`${path}.${key} must be an array index`);
      }
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (!descriptor || !descriptor.enumerable || !('value' in descriptor)) {
        throw new TypeError(`${path}.${key} must be an enumerable data property`);
      }
      assertPlainData(descriptor.value, `${path}.${key}`, visiting, completed);
    }
  } finally {
    visiting.delete(value);
  }
  completed.add(value);
}

export function deepFreezePlainData<T>(
  value: T,
  seen = new WeakSet<object>(),
): DeepReadonly<T> {
  if (value === null || typeof value !== 'object' || seen.has(value)) {
    // Non-finite geometry is fatal state; the check rides the walk that always
    // runs so it cannot be disabled with the development-only pre-pass.
    if (typeof value === 'number' && !Number.isFinite(value)) {
      throw new PlainDataContractError('must contain finite numbers');
    }
    return value as DeepReadonly<T>;
  }
  if (processedPlainData.has(value) || frozenPlainData.has(value)) {
    return value as DeepReadonly<T>;
  }
  seen.add(value);
  // Walked without `Object.values`, which allocates a fresh array for every
  // node: retained geometry is a deep graph of small objects, so that
  // array-per-node is pure garbage on a hot path.
  if (Array.isArray(value)) {
    for (let index = 0; index < value.length; index += 1) {
      deepFreezePlainData(value[index], seen);
    }
    // Plain-data arrays carry index properties only; walking any stray extra
    // property too (rather than assuming the contract) means its subgraph can
    // never be left unfrozen when the development pre-pass did not run.
    for (const key in value) {
      if (String(Number(key)) !== key && Object.prototype.hasOwnProperty.call(value, key)) {
        deepFreezePlainData((value as unknown as Record<string, unknown>)[key], seen);
      }
    }
  } else {
    for (const key in value) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        deepFreezePlainData((value as Record<string, unknown>)[key], seen);
      }
    }
  }
  Object.freeze(value);
  frozenPlainData.add(value);
  return value as DeepReadonly<T>;
}

/**
 * Deep-copy and freeze in ONE traversal.
 *
 * The previous `deepFreezePlainData(structuredClone(value))` walked the graph
 * twice — once inside the structured-clone serialize/deserialize round trip,
 * then again to freeze the result — and left a whole intermediate unfrozen copy
 * for the collector in between. Pagination snapshots every accepted block, on
 * every convergence pass, so that second walk and its garbage are a hot-path
 * cost rather than a one-off.
 *
 * Semantics match `structuredClone` on the plain-data subset this module
 * admits: the `seen` map preserves internal aliasing (an object referenced
 * twice yields the same clone twice) and terminates on cycles, exactly as the
 * structured-clone algorithm does.
 */
function cloneAndFreezePlainData<T>(value: T, seen: Map<object, unknown>): DeepReadonly<T> {
  if (value === null || typeof value !== 'object') {
    // structuredClone rejected these outright; keep that backstop so a genuine
    // violation is still reported when validation is off, rather than silently
    // smuggling a non-plain value into the retained graph.
    if (typeof value === 'function' || typeof value === 'symbol') {
      throw new TypeError('value must be structured-clone-safe plain data');
    }
    // Fatal state stays fatal without the development pre-pass: a non-finite
    // number in retained data must throw here, not surface as a paint defect.
    if (typeof value === 'number' && !Number.isFinite(value)) {
      throw new PlainDataContractError('must contain finite numbers');
    }
    return value as DeepReadonly<T>;
  }
  // Every processed graph is frozen. The cheap brand check avoids a WeakSet
  // lookup for the overwhelmingly common fresh mutable nodes.
  if (Object.isFrozen(value) && processedPlainData.has(value)) {
    return value as DeepReadonly<T>;
  }
  const prior = seen.get(value);
  if (prior !== undefined) return prior as DeepReadonly<T>;
  if (Array.isArray(value)) {
    const copy = new Array(value.length);
    seen.set(value, copy);
    // `new Array(length)` preserves a completely sparse array. Copy only own
    // indices so individual holes remain holes instead of becoming explicit
    // `undefined` entries.
    for (let index = 0; index < value.length; index += 1) {
      if (Object.prototype.hasOwnProperty.call(value, index)) {
        copy[index] = cloneAndFreezePlainData(value[index], seen);
      }
    }
    Object.freeze(copy);
    return copy as DeepReadonly<T>;
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('value must be structured-clone-safe plain data');
  }
  const copy: Record<string, unknown> = {};
  seen.set(value, copy);
  for (const key in value) {
    if (Object.prototype.hasOwnProperty.call(value, key)) {
      const child = cloneAndFreezePlainData((value as Record<string, unknown>)[key], seen);
      if (key === '__proto__') {
        // Assignment would mutate the prototype instead of creating the own
        // data property that structuredClone produces.
        Object.defineProperty(copy, key, {
          value: child,
          enumerable: true,
          writable: true,
          configurable: true,
        });
      } else {
        copy[key] = child;
      }
    }
  }
  Object.freeze(copy);
  return copy as DeepReadonly<T>;
}

export function snapshotPlainData<T>(value: T, label: string): DeepReadonly<T> {
  if (typeof value === 'object' && value !== null && processedPlainData.has(value)) {
    return value as DeepReadonly<T>;
  }
  // Path-precise contract check on engine-produced data — see
  // validation-policy.ts. The clone below still rejects the fatal structural
  // violations without this development pass; the native preflight additionally
  // pins the platform's Proxy brand check while validation is enabled.
  if (documentLayoutValidationEnabled()) {
    validatePlainData(value, label);
  }
  try {
    const snapshot = cloneAndFreezePlainData(value, new Map<object, unknown>());
    if (typeof snapshot === 'object' && snapshot !== null) processedPlainData.add(snapshot);
    return snapshot;
  } catch (error) {
    const reason = error instanceof PlainDataContractError
      ? error.message
      : 'must be structured-clone-safe plain data';
    throw new TypeError(`${label} ${reason}`);
  }
}

/** Validate and recursively seal builder-owned plain data in place. Unlike
 * snapshotPlainData this has no second structured-clone peak; callers must own
 * the supplied graph and must not expose it for later mutation. */
export function sealPlainData<T>(value: T, label: string): DeepReadonly<T> {
  if (documentLayoutValidationEnabled()) validatePlainData(value, label);
  return deepFreezeAndRegister(value, new WeakSet()) as DeepReadonly<T>;
}

function validatePlainData(value: unknown, label: string): void {
  // A transparent Proxy is intentionally indistinguishable from its target
  // through reflection. Native structuredClone provides the platform brand
  // check and rejects it before our descriptor walk can invoke user traps.
  // This extra pass is development-only; production receives engine-owned
  // retained data and keeps the single clone/freeze walk.
  try {
    structuredClone(value);
  } catch {
    throw new TypeError(`${label} must be structured-clone-safe plain data`);
  }
  assertPlainData(value, label);
}

function deepFreezeAndRegister(value: unknown, seen: WeakSet<object>): unknown {
  if (value === null || typeof value !== 'object' || seen.has(value)) {
    if (typeof value === 'number' && !Number.isFinite(value)) {
      throw new PlainDataContractError('must contain finite numbers');
    }
    return value;
  }
  if (processedPlainData.has(value)) return value;
  seen.add(value);
  if (Array.isArray(value)) {
    for (let index = 0; index < value.length; index += 1) {
      deepFreezeAndRegister(value[index], seen);
    }
    for (const key in value) {
      if (String(Number(key)) !== key && Object.prototype.hasOwnProperty.call(value, key)) {
        deepFreezeAndRegister((value as unknown as Record<string, unknown>)[key], seen);
      }
    }
  } else {
    for (const key in value) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        deepFreezeAndRegister((value as Record<string, unknown>)[key], seen);
      }
    }
  }
  Object.freeze(value);
  processedPlainData.add(value);
  return value;
}
