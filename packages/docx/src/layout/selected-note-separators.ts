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
 * Their paragraphs hold text around at most one §17.11.23/§17.11.1 marker
 * run, or a bare empty paragraph.
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

/** Validate and normalize the parser-boundary stories once, before the
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
    const paragraphs = raw[role];
    if (paragraphs === undefined) continue;
    const mark = marks[role];
    if (mark !== 'short' && mark !== 'full' && mark !== 'none') reject(`${role} paragraph has no selected mark`);
    if (!Array.isArray(paragraphs) || paragraphs.length === 0) reject(`${role} story has no paragraph`);
    // The closed shape the parser admits: text runs and line breaks with at
    // most one marker run in the whole story, and no paragraph ink,
    // numbering, frame or hidden mark.
    let markerParagraph = -1;
    paragraphs.forEach((paragraph, index) => {
      if (!isNoteSeparatorStoryRuns(paragraph.runs ?? [])) reject(`${role} paragraph carries unrepresented runs`);
      if ((paragraph.runs ?? []).some(isNoteSeparatorMarkerRun)) {
        if (markerParagraph >= 0) reject(`${role} story holds two markers`);
        markerParagraph = index;
      }
      if (paragraph.numbering != null || paragraph.framePr != null || paragraph.borders != null
        || paragraph.shading != null || paragraph.markVanish === true) {
        reject(`${role} paragraph carries unrepresented properties`);
      }
    });
    const root = selectedNoteSeparatorSource(role, [], kind);
    const paragraphSource = selectedNoteSeparatorSource(role, [Math.max(0, markerParagraph)], kind);
    definitions[role] = Object.freeze({ role, root, paragraph: paragraphSource, paragraphCount: paragraphs.length });
    stories.push(Object.freeze({
      source: root,
      body: Object.freeze(paragraphs.map((paragraph, index) =>
        acquire(paragraph, selectedNoteSeparatorSource(role, [index], kind), mark))),
    }));
  }
  return Object.freeze({ definitions: Object.freeze(definitions), stories: Object.freeze(stories) });
}

/** The parser's marker run of a separator story: an empty text run flagged
 * `__noteSeparatorMarker` that carries the marker's run properties. Authored
 * text is never mistaken for it. */
export function isNoteSeparatorMarkerRun(run: Readonly<{ type: string }>): boolean {
  return run.type === 'text' && (run as Readonly<{ __noteSeparatorMarker?: boolean }>).__noteSeparatorMarker === true;
}

/** Parser wire of one separator story paragraph: text and field runs and
 * line breaks only, at most one of them the marker run. */
export function isNoteSeparatorStoryRuns(runs: readonly Readonly<{ type: string }>[]): boolean {
  return runs.every((run) => run.type === 'text' || run.type === 'field'
      || (run.type === 'break' && (run as Readonly<{ breakType?: string }>).breakType === 'line'))
    && runs.filter(isNoteSeparatorMarkerRun).length <= 1;
}

/** Acquired separator story paragraph: text runs and line breaks, around at
 * most one inline object that lays the marker out. */
export function isAcquiredNoteSeparatorStoryRuns(runs: readonly Readonly<{ type: string }>[]): boolean {
  return runs.filter((run) => run.type === 'shape').length <= 1
    && runs.every((run) => run.type === 'text' || run.type === 'field' || run.type === 'shape' || run.type === 'break');
}

/** Inline advance of the separator marker object: the ordinary mark's 144pt
 * (min(144pt, area width) is applied to the drawn rule). The continuation
 * mark fills its own line instead: it advances nothing and line breaks
 * separate it from text before and after it (WORD_NOTE_SEPARATOR_GEOMETRY). */
export function noteSeparatorMarkerAdvancePt(mark: string | undefined): number {
  return mark === 'full' ? 0 : 144;
}
