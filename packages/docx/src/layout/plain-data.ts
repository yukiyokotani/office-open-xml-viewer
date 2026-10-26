import type { DeepReadonly } from './types.js';
import { documentLayoutValidationEnabled } from './validation-policy.js';

/** A plain-data contract violation detected inside the unconditional clone /
 *  freeze walks. Carries the reason without a property path — the path-precise
 *  report is `assertPlainData`'s job, and that pre-pass is development-only.
 *  Fatal state (a non-finite number in retained data) must be detected whether
 *  or not the development pre-pass ran, per the layout engine's error
 *  contract, so these checks are fused into the walks that always run. */
class PlainDataContractError extends TypeError {}

/** Verified graph roots from snapshotPlainData and sealPlainData. A root brand
 * proves its entire graph is plain and immutable; branding each descendant
 * would keep millions of WeakSet entries alive in large layouts. */
const processedPlainData = new WeakSet<object>();

/** Roots already deeply frozen by deepFreezePlainData. Descendants are tracked
 * only by the traversal-local `seen` set; this brand never relaxes validation,
 * because freezing alone does not prove a graph is plain data. */
const frozenPlainData = new WeakSet<object>();

/** Internal immutable-acquisition identity, stronger than shallow Object.freeze.
 * This permits scalar memoization only for roots sealed by our complete walks;
 * it does not waive the separate structured-clone/plain-data validation. */
export function isDeepFrozenPlainDataRoot(value: object): boolean {
  return frozenPlainData.has(value) || processedPlainData.has(value);
}

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
  const frozen = freezePlainDataGraph(value, seen, [], false);
  if (value !== null && typeof value === 'object' && Object.isFrozen(value)) frozenPlainData.add(value);
  return frozen;
}

/** The caller must construct a fresh root whose frozen aliases come only from
 * the supplied previously deep-frozen source (or from separately sealed
 * roots). This lets continuation builders skip revisiting immutable line
 * payloads while keeping ordinary deepFreezePlainData conservative. */
export function deepFreezePlainDataWithFrozenAliases<T>(
  value: T,
  source: object,
  seen = new WeakSet<object>(),
): DeepReadonly<T> {
  if (!frozenPlainData.has(source) && !processedPlainData.has(source)) {
    return deepFreezePlainData(value, seen);
  }
  const frozen = freezePlainDataGraph(value, seen, [], true);
  if (value !== null && typeof value === 'object' && Object.isFrozen(value)) frozenPlainData.add(value);
  return frozen;
}

/** Whether the index loop above already walked `key`: a canonical array
 * index below the length. A numeric-looking key such as "-1" or "1.5" is an
 * ordinary property and must be walked separately. */
function isWalkedArrayIndex(key: string, length: number): boolean {
  const index = Number(key);
  return Number.isInteger(index) && index >= 0 && index < length && String(index) === key;
}

function freezePlainDataGraph<T>(
  value: T,
  seen: WeakSet<object>,
  frozenAncestors: object[],
  skipFrozenAliases: boolean,
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
  if (skipFrozenAliases && Object.isFrozen(value)) return value as DeepReadonly<T>;
  // Already-frozen descendants may be shared by many independent roots. A
  // traversal-local ancestry stack still terminates frozen cycles without
  // allocating a WeakSet entry on every repeat visit. Mutable descendants
  // retain `seen`'s DAG and cycle handling while they are sealed in place.
  const alreadyFrozen = Object.isFrozen(value);
  if (alreadyFrozen) {
    if (frozenAncestors.includes(value)) return value as DeepReadonly<T>;
    frozenAncestors.push(value);
  } else {
    seen.add(value);
  }
  // Walked without `Object.values`, which allocates a fresh array for every
  // node: retained geometry is a deep graph of small objects, so that
  // array-per-node is pure garbage on a hot path.
  if (Array.isArray(value)) {
    for (let index = 0; index < value.length; index += 1) {
      freezePlainDataGraph(value[index], seen, frozenAncestors, skipFrozenAliases);
    }
    // Plain-data arrays carry index properties only; walking any stray extra
    // property too (rather than assuming the contract) means its subgraph can
    // never be left unfrozen when the development pre-pass did not run.
    for (const key in value) {
      if (!isWalkedArrayIndex(key, value.length) && Object.prototype.hasOwnProperty.call(value, key)) {
        freezePlainDataGraph((value as unknown as Record<string, unknown>)[key], seen, frozenAncestors, skipFrozenAliases);
      }
    }
  } else {
    for (const key in value) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        freezePlainDataGraph((value as Record<string, unknown>)[key], seen, frozenAncestors, skipFrozenAliases);
      }
    }
  }
  if (alreadyFrozen) frozenAncestors.pop();
  else Object.freeze(value);
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
  // ast-grep-ignore: no-plain-object-string-map -- mirrors structuredClone (ordinary prototype); `__proto__` is defined as an own property below.
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
  // Every node of a clone is verified plain data by construction, as the
  // snapshot root is. Block layouts are branded as roots too: a table result
  // or occurrence snapshot that holds them again (each layout pass, page
  // fragment and re-layout of the same cells) then shares them instead of
  // copying their lines once more. Only block layouts are branded, so the
  // brand set grows with the number of blocks, not of all retained nodes.
  if (copy.kind === 'paragraph' || copy.kind === 'table') processedPlainData.add(copy);
  return copy as DeepReadonly<T>;
}

export function snapshotPlainData<T>(value: T, label: string, ownedProjectionOf?: object): DeepReadonly<T> {
  if (typeof value === 'object' && value !== null && processedPlainData.has(value)) {
    return value as DeepReadonly<T>;
  }
  // Path-precise contract check on engine-produced data — see
  // validation-policy.ts. The clone below still rejects non-finite numbers,
  // functions, symbols, and foreign prototypes without this diagnostic pass;
  // only the pre-pass promises strict descriptor, cycle, and Proxy rejection.
  if (documentLayoutValidationEnabled()) {
    validatePlainData(value, label);
  }
  try {
    // Projection owns every newly allocated node. With a deeply frozen source
    // root, its remaining aliases are immutable source facts and the new
    // nodes can be frozen in place. This avoids copying the entire occurrence
    // after translation/re-keying already made its distinct nodes, so a
    // paragraph acquisition and the pages that place it share one copy of
    // its unchanged payload (glyph clusters, paint operations, typography).
    // A deepFreezePlainData root qualifies as well as a processed one: the
    // sharing needs immutability, which both brands prove for the whole
    // graph. Plain-data validity of the shared payload is still enforced —
    // by the development pre-pass above, and fatally by the unconditional
    // assertDocumentLayout walk over every finished layout. Mutable sources
    // still take the deep-copy path.
    // The projection allocates only mutable nodes; every frozen node it holds
    // is an alias into the deeply frozen source root, so the freeze walk need
    // not re-enter those aliases (deepFreezePlainDataWithFrozenAliases).
    const snapshot = ownedProjectionOf && (
      processedPlainData.has(ownedProjectionOf) || frozenPlainData.has(ownedProjectionOf)
    )
      ? deepFreezePlainDataWithFrozenAliases(value, ownedProjectionOf)
      : cloneAndFreezePlainData(value, new Map<object, unknown>());
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
  const sealed = deepFreezePlainData(value);
  if (value !== null && typeof value === 'object') processedPlainData.add(value);
  return sealed;
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
