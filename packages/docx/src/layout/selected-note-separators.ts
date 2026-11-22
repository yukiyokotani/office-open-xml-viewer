import type {
  NoteSeparatorInput,
  SelectedNoteSeparatorDefinitionInput,
  SelectedNoteSeparatorParagraphsInput,
  SelectedNoteSeparatorRole,
} from './body-layout-input.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type { DeepReadonly, SourceRef } from './types.js';
import type { DocParagraph } from '../types.js';

/**
 * Ordinary DOCX selected separator stories: document-listed (ECMA-376
 * §17.11.3/.9) footnote separator/continuation stories and endnote separators.
 * One paragraph holds a §17.11.23/§17.11.1 mark run or is bare and empty.
 *
 * Identity: one canonical `selected:*` root per role and note kind. These are library identities, distinct from producer-supplied
 * `reserved:*` stories and never DOCX note ids. A root is shared by every
 * page occurrence and never joins note numbering.
 *
 * Geometry: the story paragraph's acquired advance replaces only the scalar
 * band height of its role, including inherited §17.3.1.33 line/spacing facts
 * when no pPr is authored. None suppresses rule ink while retaining its line
 * box. Missing stories keep the fallback. The parser admits only a mark run
 * whose rPr equals its paragraph
 * mark's, so the run-free paragraph line box is the mark's line box.
 */

const SELECTED_INSTANCES: Readonly<Record<SelectedNoteSeparatorRole, string>> = Object.freeze({
  separator: 'selected:separator',
  continuationSeparator: 'selected:continuation-separator',
});
const ROLES = Object.freeze(Object.keys(SELECTED_INSTANCES) as SelectedNoteSeparatorRole[]);
const SELECTED_ROLE_BY_INSTANCE: ReadonlyMap<string, SelectedNoteSeparatorRole> = new Map(
  ROLES.map((role) => [SELECTED_INSTANCES[role], role]),
);

export type SelectedNoteSeparatorParagraphAcquirer = (
  paragraph: DeepReadonly<DocParagraph>,
  source: SourceRef,
  mark: 'short' | 'full' | 'none',
) => ParagraphAcquisitionInput;

export interface NormalizedSelectedNoteSeparators {
  readonly definitions: Readonly<Partial<Record<SelectedNoteSeparatorRole, SelectedNoteSeparatorDefinitionInput>>>;
  readonly stories: readonly Readonly<{
    source: SourceRef;
    body: readonly ParagraphAcquisitionInput[];
  }>[];
}

/** The one identity factory for selected roots and their paragraph paths. */
export function selectedNoteSeparatorSource(
  role: SelectedNoteSeparatorRole,
  path: readonly number[] = [],
  kind: 'footnote' | 'endnote' = 'footnote',
): SourceRef {
  return Object.freeze({
    story: kind,
    storyInstance: SELECTED_INSTANCES[role],
    path: Object.freeze([...path]),
  });
}

export function selectedNoteSeparatorRole(
  source: Pick<SourceRef, 'story' | 'storyInstance'>,
): SelectedNoteSeparatorRole | undefined {
  return source.story === 'footnote' || source.story === 'endnote'
    ? SELECTED_ROLE_BY_INSTANCE.get(source.storyInstance) : undefined;
}

export function isSelectedNoteSeparatorInstance(storyInstance: string): boolean {
  return SELECTED_ROLE_BY_INSTANCE.has(storyInstance);
}

function reject(reason: string): never {
  throw new Error(`Unsupported selected note separator story: ${reason}`);
}

/** Validate and normalize the parser-boundary paragraphs once, before the
 * layout source is sealed. Returns at most two definitions. */
export function normalizeSelectedNoteSeparators(
  raw: SelectedNoteSeparatorParagraphsInput | undefined,
  marks: Readonly<Record<SelectedNoteSeparatorRole, NoteSeparatorInput | undefined>>,
  acquire: SelectedNoteSeparatorParagraphAcquirer,
  kind: 'footnote' | 'endnote' = 'footnote',
): NormalizedSelectedNoteSeparators | undefined {
  if (raw === undefined) return undefined;
  const definitions: Partial<Record<SelectedNoteSeparatorRole, SelectedNoteSeparatorDefinitionInput>> = {};
  const stories: { source: SourceRef; body: readonly ParagraphAcquisitionInput[] }[] = [];
  for (const key of Object.keys(raw)) {
    if (!(ROLES as readonly string[]).includes(key)) reject(`unknown role ${key}`);
  }
  for (const role of ROLES) {
    const paragraph = raw[role];
    if (paragraph === undefined) continue;
    const mark = marks[role];
    if (mark !== 'short' && mark !== 'full' && mark !== 'none') reject(`${role} paragraph has no selected mark`);
    // The closed shape the parser admits: no runs besides the omitted mark,
    // or text runs around the marker sentinel, and no paragraph ink,
    // numbering, frame or hidden mark.
    if ((paragraph.runs ?? []).length !== 0 && !isNoteSeparatorMarkerRuns(paragraph.runs)) {
      reject(`${role} paragraph carries runs`);
    }
    if (paragraph.numbering != null || paragraph.framePr != null || paragraph.borders != null
      || paragraph.shading != null || paragraph.markVanish === true) {
      reject(`${role} paragraph carries unrepresented properties`);
    }
    const root = selectedNoteSeparatorSource(role, [], kind);
    const paragraphSource = selectedNoteSeparatorSource(role, [0], kind);
    definitions[role] = Object.freeze({ role, root, paragraph: paragraphSource });
    stories.push(Object.freeze({ source: root, body: Object.freeze([acquire(paragraph, paragraphSource, mark)]) }));
  }
  return Object.freeze({ definitions: Object.freeze(definitions), stories: Object.freeze(stories) });
}

/** The parser's marker sentinel of a separator story with text or with a
 * marker run formatted unlike its paragraph mark (one U+FFFC text run). */
export const NOTE_SEPARATOR_MARKER_TEXT = '\ufffc';

/** Parser wire of a separator story paragraph holding its marker in place:
 * text runs and line breaks only, exactly one text run the marker sentinel. */
export function isNoteSeparatorMarkerRuns(runs: readonly Readonly<{ type: string }>[]): boolean {
  return runs.length > 0
    && runs.every((run) => run.type === 'text'
      || (run.type === 'break' && (run as Readonly<{ breakType?: string }>).breakType === 'line'))
    && runs.filter((run) => (run as Readonly<{ text?: string }>).text === NOTE_SEPARATOR_MARKER_TEXT).length === 1;
}

/** Acquired separator story paragraph: text runs around the one inline
 * object that lays the marker out. */
export function isAcquiredNoteSeparatorMarkerRuns(runs: readonly Readonly<{ type: string }>[]): boolean {
  return runs.length > 0 && runs.filter((run) => run.type === 'shape').length === 1
    && runs.every((run) => run.type === 'text' || run.type === 'shape' || run.type === 'break');
}

/** Inline advance of the separator marker object: the ordinary mark's 144pt
 * (min(144pt, area width) is applied to the drawn rule). The continuation
 * mark fills its own line instead: it advances nothing and line breaks
 * separate it from text before and after it (WORD_NOTE_SEPARATOR_GEOMETRY). */
export function noteSeparatorMarkerAdvancePt(mark: string | undefined): number {
  return mark === 'full' ? 0 : 144;
}
