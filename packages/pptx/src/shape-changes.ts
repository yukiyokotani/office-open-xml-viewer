import type { ShapeElement } from './types';

/** One segment in a path rooted at a slide-owned {@link ShapeElement}. */
export type PptxShapeChangePathSegment = string | number;

/**
 * Serializable JSON-Patch-style change rooted at a slide-owned
 * {@link ShapeElement}. Object properties use string segments and array items
 * use zero-based numeric segments.
 */
export type PptxShapeChange =
  | {
      op: 'add' | 'replace';
      path: readonly PptxShapeChangePathSegment[];
      /** Must be JSON-serializable. This is validated before the batch commits. */
      value: unknown;
    }
  | {
      op: 'remove';
      path: readonly PptxShapeChangePathSegment[];
    };

export interface AppliedPptxShapeChanges {
  /** Detached, normalized copies of the changes that were applied. */
  applied: PptxShapeChange[];
  /**
   * Detached changes that restore the previous shape. They are already in the
   * reverse order required for an undo call.
   */
  inverse: PptxShapeChange[];
}

type ChangeContainer = Record<string, unknown> | unknown[];

const forbiddenPathSegments = new Set(['__proto__', 'prototype', 'constructor']);

function fail(changeIndex: number, message: string): never {
  throw new Error(`Invalid shape change at index ${changeIndex}: ${message}`);
}

function validatePath(
  path: readonly PptxShapeChangePathSegment[],
  changeIndex: number,
): void {
  if (path.length === 0) fail(changeIndex, 'path must not be empty');
  if (path[0] === 'type' || path[0] === 'id') {
    fail(changeIndex, 'shape identity fields "type" and "id" are immutable');
  }
  for (const segment of path) {
    if (typeof segment === 'number') {
      if (!Number.isSafeInteger(segment) || segment < 0) {
        fail(changeIndex, `array index must be a non-negative safe integer: ${segment}`);
      }
      continue;
    }
    if (forbiddenPathSegments.has(segment)) {
      fail(changeIndex, `unsafe path segment: ${segment}`);
    }
  }
}

function assertJsonValue(value: unknown, changeIndex: number, ancestors = new Set<object>()): void {
  if (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean'
  ) {
    return;
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) fail(changeIndex, 'value contains a non-finite number');
    return;
  }
  if (typeof value !== 'object') {
    fail(changeIndex, `value contains unsupported type "${typeof value}"`);
  }
  if (ancestors.has(value)) fail(changeIndex, 'value contains a cycle');

  const prototype = Object.getPrototypeOf(value);
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) {
    fail(changeIndex, 'value must contain only JSON objects and arrays');
  }

  ancestors.add(value);
  const entries = Array.isArray(value) ? value.entries() : Object.entries(value);
  for (const [, child] of entries) assertJsonValue(child, changeIndex, ancestors);
  ancestors.delete(value);
}

function cloneChangeValue(value: unknown, changeIndex: number): unknown {
  assertJsonValue(value, changeIndex);
  return structuredClone(value);
}

function clonePath(
  path: readonly PptxShapeChangePathSegment[],
): PptxShapeChangePathSegment[] {
  return [...path];
}

function isContainer(value: unknown): value is ChangeContainer {
  return typeof value === 'object' && value !== null;
}

function resolveParent(
  root: ShapeElement,
  path: readonly PptxShapeChangePathSegment[],
  changeIndex: number,
): { parent: ChangeContainer; key: PptxShapeChangePathSegment } {
  let current: unknown = root;
  for (let depth = 0; depth < path.length - 1; depth += 1) {
    if (!isContainer(current)) {
      fail(changeIndex, `path does not resolve to a container at segment ${depth}`);
    }
    const segment = path[depth]!;
    if (Array.isArray(current)) {
      if (typeof segment !== 'number' || segment >= current.length) {
        fail(changeIndex, `array index ${String(segment)} is out of range at segment ${depth}`);
      }
      current = current[segment];
    } else {
      if (typeof segment !== 'string' || !Object.hasOwn(current, segment)) {
        fail(changeIndex, `object property "${String(segment)}" does not exist at segment ${depth}`);
      }
      current = current[segment];
    }
  }
  if (!isContainer(current)) {
    fail(changeIndex, 'path parent is not an object or array');
  }
  return { parent: current, key: path[path.length - 1]! };
}

function readObjectProperty(
  parent: Record<string, unknown>,
  key: PptxShapeChangePathSegment,
  changeIndex: number,
): { exists: boolean; value: unknown; key: string } {
  if (typeof key !== 'string') {
    fail(changeIndex, `object property segment must be a string: ${String(key)}`);
  }
  return {
    exists: Object.hasOwn(parent, key),
    value: parent[key],
    key,
  };
}

function readArrayItem(
  parent: unknown[],
  key: PptxShapeChangePathSegment,
  changeIndex: number,
  allowEnd: boolean,
): { exists: boolean; value: unknown; index: number } {
  if (typeof key !== 'number') {
    fail(changeIndex, `array item segment must be a number: ${String(key)}`);
  }
  const upperBound = allowEnd ? parent.length : parent.length - 1;
  if (key > upperBound) {
    fail(changeIndex, `array index ${key} is out of range`);
  }
  return { exists: key < parent.length, value: parent[key], index: key };
}

/**
 * Apply a batch to an isolated shape draft. The caller owns committing the
 * draft after any additional shape-level invariants have been checked.
 */
export function applyPptxShapeChanges(
  draft: ShapeElement,
  changes: readonly PptxShapeChange[],
): AppliedPptxShapeChanges {
  const applied: PptxShapeChange[] = [];
  const inverse: PptxShapeChange[] = [];

  changes.forEach((change, changeIndex) => {
    validatePath(change.path, changeIndex);
    const path = clonePath(change.path);
    const { parent, key } = resolveParent(draft, path, changeIndex);

    if (change.op === 'add') {
      const value = cloneChangeValue(change.value, changeIndex);
      if (Array.isArray(parent)) {
        const current = readArrayItem(parent, key, changeIndex, true);
        parent.splice(current.index, 0, value);
        inverse.unshift({ op: 'remove', path: clonePath(path) });
      } else {
        const current = readObjectProperty(parent, key, changeIndex);
        parent[current.key] = value;
        inverse.unshift(
          current.exists
            ? { op: 'replace', path: clonePath(path), value: structuredClone(current.value) }
            : { op: 'remove', path: clonePath(path) },
        );
      }
      applied.push({ op: 'add', path, value: structuredClone(value) });
      return;
    }

    if (change.op === 'replace') {
      const value = cloneChangeValue(change.value, changeIndex);
      if (Array.isArray(parent)) {
        const current = readArrayItem(parent, key, changeIndex, false);
        parent[current.index] = value;
        inverse.unshift({
          op: 'replace',
          path: clonePath(path),
          value: structuredClone(current.value),
        });
      } else {
        const current = readObjectProperty(parent, key, changeIndex);
        if (!current.exists) fail(changeIndex, `object property "${current.key}" does not exist`);
        parent[current.key] = value;
        inverse.unshift({
          op: 'replace',
          path: clonePath(path),
          value: structuredClone(current.value),
        });
      }
      applied.push({ op: 'replace', path, value: structuredClone(value) });
      return;
    }

    if (Array.isArray(parent)) {
      const current = readArrayItem(parent, key, changeIndex, false);
      parent.splice(current.index, 1);
      inverse.unshift({
        op: 'add',
        path: clonePath(path),
        value: structuredClone(current.value),
      });
    } else {
      const current = readObjectProperty(parent, key, changeIndex);
      if (!current.exists) fail(changeIndex, `object property "${current.key}" does not exist`);
      delete parent[current.key];
      inverse.unshift({
        op: 'add',
        path: clonePath(path),
        value: structuredClone(current.value),
      });
    }
    applied.push({ op: 'remove', path });
  });

  return { applied, inverse };
}
