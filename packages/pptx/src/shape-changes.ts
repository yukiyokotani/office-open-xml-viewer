import type { ShapeElement } from './types';

type DefinedPartial<T> = Partial<{
  [Key in keyof T]-?: Exclude<T[Key], undefined>;
}>;

type OptionalKey<T> = {
  [Key in keyof T]-?: {} extends Pick<T, Key> ? Key : never;
}[keyof T];

type EditableShape = Omit<ShapeElement, 'type' | 'id'>;

/**
 * Shape-level delta. Top-level properties are optional; nested objects and
 * arrays, including `textBody`, are replaced as complete values.
 */
export type PptxShapeProperties = DefinedPartial<EditableShape>;
export type PptxOptionalShapeProperty = OptionalKey<EditableShape>;

/** Public operation names accepted by {@link PptxShapeChange}. */
export enum PptxShapeChangeType {
  Update = 'shape.update',
}

export interface PptxApplyShapeChangesRequest {
  /** Zero-based slide index. */
  slideIndex: number;
  /** Slide-local DrawingML `cNvPr@id`. */
  shapeId: string;
  /** Ordered atomic batch of shape-level deltas. */
  changes: readonly PptxShapeChange[];
}

/**
 * A serializable delta for one slide-owned shape.
 *
 * `patch` updates only the supplied top-level properties. Nested values are
 * replaced whole rather than deep-merged. Use `unset` for optional properties
 * that must be removed; the SDK also uses it when generating an exact inverse.
 */
export interface PptxShapeChange {
  type: PptxShapeChangeType.Update;
  patch: PptxShapeProperties;
  unset?: readonly PptxOptionalShapeProperty[];
}

export interface AppliedPptxShapeChanges {
  /** Detached copies of the shape deltas that were applied. */
  applied: PptxShapeChange[];
  /**
   * Detached deltas that restore the previous shape. They are already in the
   * reverse order required for an undo call.
   */
  inverse: PptxShapeChange[];
}

const forbiddenPropertyNames = new Set([
  '__proto__',
  'prototype',
  'constructor',
  'type',
  'id',
]);

function fail(changeIndex: number, message: string): never {
  throw new Error(`Invalid shape change at index ${changeIndex}: ${message}`);
}

function assertJsonValue(value: unknown, changeIndex: number, ancestors = new Set<object>()): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) fail(changeIndex, 'change contains a non-finite number');
    return;
  }
  if (typeof value !== 'object') {
    fail(changeIndex, `change contains unsupported type "${typeof value}"`);
  }
  if (ancestors.has(value)) fail(changeIndex, 'change contains a cycle');
  const prototype = Object.getPrototypeOf(value);
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) {
    fail(changeIndex, 'change must contain only JSON objects and arrays');
  }
  ancestors.add(value);
  const entries = Array.isArray(value) ? value.entries() : Object.entries(value);
  for (const [, child] of entries) assertJsonValue(child, changeIndex, ancestors);
  ancestors.delete(value);
}

function applyShapeProperties(
  target: ShapeElement,
  patch: PptxShapeProperties,
  unset: readonly PptxOptionalShapeProperty[] | undefined,
  changeIndex: number,
): Pick<PptxShapeChange, 'patch' | 'unset'> {
  const inversePatch: Record<string, unknown> = {};
  const inverseUnset: PptxOptionalShapeProperty[] = [];
  const patchEntries = Object.entries(patch);
  const setKeys = new Set(patchEntries.map(([key]) => key));

  for (const [key, value] of patchEntries) {
    if (forbiddenPropertyNames.has(key)) fail(changeIndex, `property "${key}" is immutable`);
    if (Object.hasOwn(target, key)) {
      inversePatch[key] = structuredClone(Reflect.get(target, key));
    } else {
      inverseUnset.push(key as PptxOptionalShapeProperty);
    }
    Reflect.set(target, key, structuredClone(value));
  }

  for (const key of unset ?? []) {
    if (typeof key !== 'string') fail(changeIndex, 'unset property names must be strings');
    if (forbiddenPropertyNames.has(key)) fail(changeIndex, `property "${key}" is immutable`);
    if (setKeys.has(key)) fail(changeIndex, `property "${key}" cannot be set and unset together`);
    if (!Object.hasOwn(target, key)) continue;
    inversePatch[key] = structuredClone(Reflect.get(target, key));
    Reflect.deleteProperty(target, key);
  }

  return {
    patch: inversePatch as PptxShapeProperties,
    ...(inverseUnset.length > 0 ? { unset: inverseUnset } : {}),
  };
}

/** Apply shape-level deltas to an isolated draft. */
export function applyPptxShapeChanges(
  draft: ShapeElement,
  changes: readonly PptxShapeChange[],
): AppliedPptxShapeChanges {
  const applied: PptxShapeChange[] = [];
  const inverse: PptxShapeChange[] = [];

  changes.forEach((input, changeIndex) => {
    assertJsonValue(input, changeIndex);
    const change = structuredClone(input);
    if (change.type !== PptxShapeChangeType.Update) {
      fail(changeIndex, `unsupported type "${String(change.type)}"`);
    }

    const undo = applyShapeProperties(
      draft,
      change.patch,
      change.unset,
      changeIndex,
    );
    inverse.unshift({ type: PptxShapeChangeType.Update, ...undo });
    applied.push(change);
  });

  return { applied, inverse };
}
