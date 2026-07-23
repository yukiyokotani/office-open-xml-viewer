import type {
  Paragraph,
  ShapeElement,
  TextBody,
  TextRun,
  TextRunData,
} from './types';

type DefinedPartial<T> = Partial<{
  [Key in keyof T]-?: Exclude<T[Key], undefined>;
}>;

type OptionalKey<T> = {
  [Key in keyof T]-?: {} extends Pick<T, Key> ? Key : never;
}[keyof T];

type EditableShape = Omit<ShapeElement, 'type' | 'id' | 'textBody'>;
type EditableTextBody = Omit<TextBody, 'paragraphs'>;
type EditableParagraph = Omit<Paragraph, 'runs'>;
type EditableTextRun = Omit<TextRunData, 'type'>;

export type PptxShapeProperties = DefinedPartial<EditableShape>;
export type PptxTextBodyProperties = DefinedPartial<EditableTextBody>;
export type PptxParagraphProperties = DefinedPartial<EditableParagraph>;
export type PptxTextRunProperties = DefinedPartial<EditableTextRun>;
export type PptxOptionalShapeProperty = OptionalKey<EditableShape>;
export type PptxOptionalTextBodyProperty = OptionalKey<EditableTextBody>;
export type PptxOptionalParagraphProperty = OptionalKey<EditableParagraph>;
export type PptxOptionalTextRunProperty = OptionalKey<EditableTextRun>;

export interface PptxApplyShapeChangesRequest {
  /** Zero-based slide index. */
  slideIndex: number;
  /** Slide-local DrawingML `cNvPr@id`. */
  shapeId: string;
  /** Ordered atomic batch. */
  changes: readonly PptxShapeChange[];
}

/**
 * Serializable, model-aware edit operations for one slide-owned shape.
 *
 * Each variant exposes the target model and its legal property names directly
 * to TypeScript/IntelliSense. Callers never construct paths into the private
 * presentation model.
 */
export type PptxShapeChange =
  | {
      type: 'shape.update';
      patch: PptxShapeProperties;
      unset?: readonly PptxOptionalShapeProperty[];
    }
  | {
      type: 'textBody.replace';
      textBody: TextBody | null;
    }
  | {
      type: 'textBody.update';
      patch: PptxTextBodyProperties;
      unset?: readonly PptxOptionalTextBodyProperty[];
    }
  | {
      type: 'paragraph.update';
      paragraphIndex: number;
      patch: PptxParagraphProperties;
      unset?: readonly PptxOptionalParagraphProperty[];
    }
  | {
      type: 'paragraph.insert';
      paragraphIndex: number;
      paragraph: Paragraph;
    }
  | {
      type: 'paragraph.remove';
      paragraphIndex: number;
    }
  | {
      type: 'textRun.update';
      paragraphIndex: number;
      runIndex: number;
      patch: PptxTextRunProperties;
      unset?: readonly PptxOptionalTextRunProperty[];
    }
  | {
      type: 'run.insert' | 'run.replace';
      paragraphIndex: number;
      runIndex: number;
      run: TextRun;
    }
  | {
      type: 'run.remove';
      paragraphIndex: number;
      runIndex: number;
    };

export interface AppliedPptxShapeChanges {
  /** Detached copies of the semantic changes that were applied. */
  applied: PptxShapeChange[];
  /**
   * Detached semantic changes that restore the previous shape. They are
   * already in the reverse order required for an undo call.
   */
  inverse: PptxShapeChange[];
}

const forbiddenPropertyNames = new Set([
  '__proto__',
  'prototype',
  'constructor',
  'type',
  'id',
  'textBody',
  'paragraphs',
  'runs',
]);

function fail(changeIndex: number, message: string): never {
  throw new Error(`Invalid shape change at index ${changeIndex}: ${message}`);
}

function assertIndex(
  index: number,
  length: number,
  changeIndex: number,
  label: string,
  allowEnd = false,
): void {
  const upperBound = allowEnd ? length : length - 1;
  if (!Number.isSafeInteger(index) || index < 0 || index > upperBound) {
    fail(changeIndex, `${label} ${index} is out of range`);
  }
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

function applyProperties<T extends object>(
  target: T,
  patch: DefinedPartial<T>,
  unset: readonly OptionalKey<T>[] | undefined,
  changeIndex: number,
): { patch: DefinedPartial<T>; unset?: OptionalKey<T>[] } {
  const targetRecord = target as Record<string, unknown>;
  const patchRecord = patch as Record<string, unknown>;
  const inversePatch: Record<string, unknown> = {};
  const inverseUnset: string[] = [];
  const setKeys = new Set(Object.keys(patchRecord));

  for (const key of setKeys) {
    if (forbiddenPropertyNames.has(key)) fail(changeIndex, `property "${key}" is immutable`);
    if (Object.hasOwn(targetRecord, key)) {
      inversePatch[key] = structuredClone(targetRecord[key]);
    } else {
      inverseUnset.push(key);
    }
    targetRecord[key] = structuredClone(patchRecord[key]);
  }

  for (const key of unset ?? []) {
    if (typeof key !== 'string') fail(changeIndex, 'unset property names must be strings');
    if (forbiddenPropertyNames.has(key)) fail(changeIndex, `property "${key}" is immutable`);
    if (setKeys.has(key)) fail(changeIndex, `property "${key}" cannot be set and unset together`);
    if (!Object.hasOwn(targetRecord, key)) continue;
    inversePatch[key] = structuredClone(targetRecord[key]);
    delete targetRecord[key];
  }

  return {
    patch: inversePatch as DefinedPartial<T>,
    ...(inverseUnset.length > 0 ? { unset: inverseUnset as OptionalKey<T>[] } : {}),
  };
}

function requireTextBody(
  shape: ShapeElement,
  changeIndex: number,
): TextBody {
  if (!shape.textBody) fail(changeIndex, 'shape has no text body');
  return shape.textBody;
}

function requireParagraph(
  shape: ShapeElement,
  paragraphIndex: number,
  changeIndex: number,
): Paragraph {
  const textBody = requireTextBody(shape, changeIndex);
  assertIndex(paragraphIndex, textBody.paragraphs.length, changeIndex, 'paragraph index');
  return textBody.paragraphs[paragraphIndex]!;
}

function requireRun(
  paragraph: Paragraph,
  runIndex: number,
  changeIndex: number,
): TextRun {
  assertIndex(runIndex, paragraph.runs.length, changeIndex, 'run index');
  return paragraph.runs[runIndex]!;
}

/** Apply semantic changes to an isolated shape draft. */
export function applyPptxShapeChanges(
  draft: ShapeElement,
  changes: readonly PptxShapeChange[],
): AppliedPptxShapeChanges {
  const applied: PptxShapeChange[] = [];
  const inverse: PptxShapeChange[] = [];

  changes.forEach((input, changeIndex) => {
    assertJsonValue(input, changeIndex);
    const change = structuredClone(input);

    switch (change.type) {
      case 'shape.update': {
        const undo = applyProperties<EditableShape>(
          draft,
          change.patch,
          change.unset,
          changeIndex,
        );
        inverse.unshift({ type: 'shape.update', ...undo });
        break;
      }
      case 'textBody.replace': {
        inverse.unshift({
          type: 'textBody.replace',
          textBody: structuredClone(draft.textBody),
        });
        draft.textBody = structuredClone(change.textBody);
        break;
      }
      case 'textBody.update': {
        const textBody = requireTextBody(draft, changeIndex);
        const undo = applyProperties<EditableTextBody>(
          textBody,
          change.patch,
          change.unset,
          changeIndex,
        );
        inverse.unshift({ type: 'textBody.update', ...undo });
        break;
      }
      case 'paragraph.update': {
        const paragraph = requireParagraph(draft, change.paragraphIndex, changeIndex);
        const undo = applyProperties<EditableParagraph>(
          paragraph,
          change.patch,
          change.unset,
          changeIndex,
        );
        inverse.unshift({
          type: 'paragraph.update',
          paragraphIndex: change.paragraphIndex,
          ...undo,
        });
        break;
      }
      case 'paragraph.insert': {
        const textBody = requireTextBody(draft, changeIndex);
        assertIndex(
          change.paragraphIndex,
          textBody.paragraphs.length,
          changeIndex,
          'paragraph index',
          true,
        );
        textBody.paragraphs.splice(
          change.paragraphIndex,
          0,
          structuredClone(change.paragraph),
        );
        inverse.unshift({
          type: 'paragraph.remove',
          paragraphIndex: change.paragraphIndex,
        });
        break;
      }
      case 'paragraph.remove': {
        const textBody = requireTextBody(draft, changeIndex);
        assertIndex(
          change.paragraphIndex,
          textBody.paragraphs.length,
          changeIndex,
          'paragraph index',
        );
        const [paragraph] = textBody.paragraphs.splice(change.paragraphIndex, 1);
        inverse.unshift({
          type: 'paragraph.insert',
          paragraphIndex: change.paragraphIndex,
          paragraph: structuredClone(paragraph!),
        });
        break;
      }
      case 'textRun.update': {
        const paragraph = requireParagraph(draft, change.paragraphIndex, changeIndex);
        const run = requireRun(paragraph, change.runIndex, changeIndex);
        if (run.type !== 'text') fail(changeIndex, 'target run is not a text run');
        const undo = applyProperties<EditableTextRun>(
          run,
          change.patch,
          change.unset,
          changeIndex,
        );
        inverse.unshift({
          type: 'textRun.update',
          paragraphIndex: change.paragraphIndex,
          runIndex: change.runIndex,
          ...undo,
        });
        break;
      }
      case 'run.insert': {
        const paragraph = requireParagraph(draft, change.paragraphIndex, changeIndex);
        assertIndex(change.runIndex, paragraph.runs.length, changeIndex, 'run index', true);
        paragraph.runs.splice(change.runIndex, 0, structuredClone(change.run));
        inverse.unshift({
          type: 'run.remove',
          paragraphIndex: change.paragraphIndex,
          runIndex: change.runIndex,
        });
        break;
      }
      case 'run.replace': {
        const paragraph = requireParagraph(draft, change.paragraphIndex, changeIndex);
        const previous = requireRun(paragraph, change.runIndex, changeIndex);
        paragraph.runs[change.runIndex] = structuredClone(change.run);
        inverse.unshift({
          type: 'run.replace',
          paragraphIndex: change.paragraphIndex,
          runIndex: change.runIndex,
          run: structuredClone(previous),
        });
        break;
      }
      case 'run.remove': {
        const paragraph = requireParagraph(draft, change.paragraphIndex, changeIndex);
        const previous = requireRun(paragraph, change.runIndex, changeIndex);
        paragraph.runs.splice(change.runIndex, 1);
        inverse.unshift({
          type: 'run.insert',
          paragraphIndex: change.paragraphIndex,
          runIndex: change.runIndex,
          run: structuredClone(previous),
        });
        break;
      }
    }

    applied.push(change);
  });

  return { applied, inverse };
}
