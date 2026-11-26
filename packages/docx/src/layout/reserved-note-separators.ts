import type {
  ReservedNoteSeparatorCharacterInput,
  ReservedNoteSeparatorDefinitionInput,
  ReservedNoteSeparatorDefinitionsInput,
  ReservedNoteSeparatorRole,
  ReservedNoteSeparatorRolesInput,
  ReservedNoteSeparatorStoryInput,
  ReservedNoteSeparatorsInput,
} from './body-layout-input.js';
import { projectBodyOccurrence } from './occurrence-projection.js';
import { translateBorder, translateRect } from './retained-geometry-translation.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type {
  DeepReadonly,
  NoteLayout,
  NoteSeparatorLayout,
  PaintNode,
  SourceRef,
  StoryLayout,
} from './types.js';
import type { DocParagraph, DocxTextRun } from '../types.js';

/**
 * Reserved note separator stories supplied by a producer (ECMA-376 §17.18.33
 * `separator`, `continuationSeparator` and `continuationNotice` of each note
 * kind). The WordprocessingML parser does not emit them; its formatted
 * footnote separators use selected-note-separators.ts.
 *
 * Identity: each definition is one canonical source root per note kind and
 * role, `reserved:*` inside the existing `footnote`/`endnote` story kinds.
 * These are library identities, never DOCX special note ids or producer
 * addresses; a numbered note using one of them is rejected rather than
 * shadowed. A root is shared by every page occurrence and never joins note
 * numbering.
 *
 * Validation covers only what the wire represents. The producer owns the
 * control's rule meaning and its source story shape; this consumer cannot
 * reread source bytes. It checks the closed story shape and refuses facts it
 * cannot represent (presence-only frame or numbering properties,
 * revision-marked characters, authored runs) before layout, instead of
 * approximating them.
 *
 * Page dependence: an admitted story holds no runs other than its control and
 * content-mark metric participants, so its occurrences carry no fields,
 * drawings or text boxes. That closed shape is why note reuse walks need not
 * visit separator occurrences for page-dependent content.
 */

const RESERVED_INSTANCES: Readonly<Record<ReservedNoteSeparatorRole, string>> = Object.freeze({
  separator: 'reserved:separator',
  continuationSeparator: 'reserved:continuation-separator',
  continuationNotice: 'reserved:continuation-notice',
});
const ROLES = Object.freeze(Object.keys(RESERVED_INSTANCES) as ReservedNoteSeparatorRole[]);
const RESERVED_ROLE_BY_INSTANCE: ReadonlyMap<string, ReservedNoteSeparatorRole> = new Map(
  ROLES.map((role) => [RESERVED_INSTANCES[role], role]),
);

export type NoteSeparatorParticipantRole = 'rule-control' | 'paragraph-mark';

export interface NoteSeparatorParticipant {
  readonly role: NoteSeparatorParticipantRole;
  readonly run: DeepReadonly<DocxTextRun>;
}

/** Builds the immutable paragraph acquisition input through the parser-model
 * boundary; this module never imports the parser projection itself. */
export type ReservedNoteSeparatorParagraphAcquirer = (
  paragraph: DeepReadonly<DocParagraph>,
  participants: readonly NoteSeparatorParticipant[],
  source: SourceRef,
) => ParagraphAcquisitionInput;

export interface NormalizedReservedNoteSeparators {
  readonly roles: ReservedNoteSeparatorRolesInput;
  readonly stories: readonly Readonly<{
    source: SourceRef;
    body: readonly ParagraphAcquisitionInput[];
  }>[];
}

/** The one identity factory for reserved roots and their structural paths. */
export function reservedNoteSeparatorSource(
  kind: 'footnote' | 'endnote',
  role: ReservedNoteSeparatorRole,
  path: readonly number[] = [],
): SourceRef {
  return Object.freeze({
    story: kind,
    storyInstance: RESERVED_INSTANCES[role],
    path: Object.freeze([...path]),
  });
}

export function reservedNoteSeparatorRole(
  source: Pick<SourceRef, 'story' | 'storyInstance'>,
): ReservedNoteSeparatorRole | undefined {
  if (source.story !== 'footnote' && source.story !== 'endnote') return undefined;
  return RESERVED_ROLE_BY_INSTANCE.get(source.storyInstance);
}

export function isReservedNoteSeparatorInstance(storyInstance: string): boolean {
  return RESERVED_ROLE_BY_INSTANCE.has(storyInstance);
}

function reject(reason: string): never {
  throw new Error(`Unsupported reserved note separator story: ${reason}`);
}

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

/** A closed wire record: a plain object whose keys are all known. Absent
 * optional fields are `undefined` (the producer omits them); any other value,
 * including `null`, must have the field's own shape. */
function requireRecord(value: unknown, keys: readonly string[], label: string): Readonly<Record<string, unknown>> {
  if (!isRecord(value)) reject(`${label} is not a record`);
  for (const key of Object.keys(value)) {
    if (!keys.includes(key)) reject(`${label} has unknown field ${key}`);
  }
  return value;
}

function optionalRecord(
  value: unknown, keys: readonly string[], label: string,
): Readonly<Record<string, unknown>> | undefined {
  return value === undefined ? undefined : requireRecord(value, keys, label);
}

function optionalBoolean(value: unknown, label: string): void {
  if (value !== undefined && typeof value !== 'boolean') reject(`${label} is not a boolean`);
}

function validateCharacterShape(value: unknown, label: string): void {
  const character = requireRecord(value, ['run', 'insertion'], label);
  if (character.run !== undefined && !isRecord(character.run)) reject(`${label} run is not a record`);
  optionalBoolean(character.insertion, `${label} insertion`);
}

/** Closed shape of one story record, before its semantic validation. */
function validateStoryShape(value: unknown, label: string): void {
  const story = requireRecord(value, ['rule', 'paragraph'], label);
  const rule = optionalRecord(story.rule, ['mark', 'control'], `${label} rule`);
  if (rule) {
    if (rule.mark !== undefined && typeof rule.mark !== 'string') reject(`${label} rule mark is not a string`);
    validateCharacterShape(rule.control, `${label} rule control`);
  }
  const paragraph = optionalRecord(
    story.paragraph, ['paragraph', 'contentMark', 'numbered', 'framed'], `${label} paragraph`,
  );
  if (paragraph) {
    if (!isRecord(paragraph.paragraph)) reject(`${label} paragraph properties are not a record`);
    if (paragraph.contentMark !== undefined) {
      validateCharacterShape(paragraph.contentMark, `${label} content mark`);
    }
    optionalBoolean(paragraph.numbered, `${label} numbered flag`);
    optionalBoolean(paragraph.framed, `${label} framed flag`);
  }
}

function validateCharacter(
  character: ReservedNoteSeparatorCharacterInput | undefined,
  label: string,
): void {
  if (!character) reject(`${label} is missing`);
  // Revision metadata is not retained by the wire.
  if (character.insertion === true) reject(`${label} carries an unretained tracked insertion`);
  // A producer's character properties can turn an otherwise empty control
  // into a glyph (for example a symbol substitution). This closed rule/mark
  // contract has only metric participants; authored glyphs need a separate
  // content contract. Reject them before normalization can erase text or
  // omit a notice mark.
  if (character.run && character.run.text !== '') {
    reject(`${label} carries unsupported authored text`);
  }
}

function validateParagraph(story: ReservedNoteSeparatorStoryInput): void {
  const authored = story.paragraph;
  if (!authored) reject('rule story has no paragraph');
  const paragraph = authored.paragraph;
  // Presence-only flags: the owned frame/numbering properties are absent.
  if (authored.numbered === true || authored.framed === true
    || paragraph.numbering != null || paragraph.framePr != null) {
    reject('numbered or framed paragraph cascades are not represented');
  }
  if ((paragraph.runs ?? []).length !== 0) reject('authored paragraph runs are not represented');
  validateCharacter(authored.contentMark, 'content paragraph mark');
  // No mark run is an effective vanish of the same properties as the paragraph mark.
  if ((authored.contentMark!.run === undefined) !== (paragraph.markVanish === true)) {
    reject('content mark visibility disagrees with its paragraph');
  }
}

/** Closed structural validation of one retained story: empty (no paragraph),
 * paragraph-only, or a whole rule (control plus its paragraph). */
export function validateReservedNoteSeparatorStory(story: ReservedNoteSeparatorStoryInput): void {
  const rule = story.rule;
  if (!rule) {
    if (story.paragraph) validateParagraph(story);
    return;
  }
  if (rule.mark !== 'short' && rule.mark !== 'full') reject('rule control lacks a special-character mark');
  validateCharacter(rule.control, 'rule control');
  if (!rule.control.run) reject('rule control is hidden');
  validateParagraph(story);
}

function explicitColor(run: DeepReadonly<DocxTextRun>): string | undefined {
  // ECMA-376 §17.3.2.6: an explicit hex colour; auto/absent keeps the
  // existing library rule ink colour.
  return typeof run.color === 'string' && /^[0-9A-Fa-f]{6}$/.test(run.color)
    ? `#${run.color}` : undefined;
}

function normalizeStory(
  kind: 'footnote' | 'endnote',
  role: ReservedNoteSeparatorRole,
  story: ReservedNoteSeparatorStoryInput,
  acquire: ReservedNoteSeparatorParagraphAcquirer,
): Readonly<{ definition: ReservedNoteSeparatorDefinitionInput; body: readonly ParagraphAcquisitionInput[] }> {
  validateReservedNoteSeparatorStory(story);
  const root = reservedNoteSeparatorSource(kind, role);
  const authored = story.paragraph;
  if (!authored) return Object.freeze({ definition: Object.freeze({ role, root }), body: Object.freeze([]) });
  const paragraphSource = reservedNoteSeparatorSource(kind, role, [0]);
  // Story order: the rule control precedes the content paragraph mark. A
  // visible content mark is a participant in every class, so its complete
  // effective character properties (including §17.3.2.24 position and
  // §17.3.2.42 vertAlign, which paragraph-mark font facts do not carry)
  // reach its line box. An
  // effectively vanished mark has no run: a paragraph-only story is then
  // the shared fully hidden paragraph, and a rule keeps only its control.
  const participants: NoteSeparatorParticipant[] = [];
  if (story.rule?.control.run) participants.push({ role: 'rule-control', run: story.rule.control.run });
  if (authored.contentMark?.run) {
    participants.push({ role: 'paragraph-mark', run: authored.contentMark.run });
  }
  const input = acquire(authored.paragraph, Object.freeze(participants), paragraphSource);
  const color = story.rule?.control.run ? explicitColor(story.rule.control.run) : undefined;
  const definition: ReservedNoteSeparatorDefinitionInput = Object.freeze({
    role,
    root,
    paragraph: Object.freeze({
      source: paragraphSource,
      hidden: story.rule === undefined && authored.paragraph.markVanish === true,
    }),
    ...(story.rule ? {
      rule: Object.freeze({
        mark: story.rule.mark!,
        control: reservedNoteSeparatorSource(kind, role, [0, 0]),
        ...(color === undefined ? {} : { color }),
      }),
    } : {}),
  });
  return Object.freeze({ definition, body: Object.freeze([input]) });
}

/** Validate and normalize the raw wire once, before the layout source is
 * sealed. Returns at most six definitions. */
export function normalizeReservedNoteSeparators(
  raw: ReservedNoteSeparatorsInput | undefined,
  acquire: ReservedNoteSeparatorParagraphAcquirer,
): NormalizedReservedNoteSeparators | undefined {
  if (raw === undefined) return undefined;
  // Fail closed on any producer's malformed wire: a non-record or unknown
  // field must never seal as an empty story, which would suppress the scalar
  // separator of its note kind.
  requireRecord(raw, ['footnote', 'endnote'], 'separator set');
  const stories: { source: SourceRef; body: readonly ParagraphAcquisitionInput[] }[] = [];
  const roles: { footnote?: ReservedNoteSeparatorDefinitionsInput; endnote?: ReservedNoteSeparatorDefinitionsInput } = {};
  for (const kind of ['footnote', 'endnote'] as const) {
    const kindStories = raw[kind];
    if (kindStories === undefined) continue;
    requireRecord(kindStories, ROLES, `${kind} stories`);
    const definitions = {} as Record<ReservedNoteSeparatorRole, ReservedNoteSeparatorDefinitionInput>;
    for (const role of ROLES) {
      validateStoryShape(kindStories[role], `${kind} ${role}`);
      const normalized = normalizeStory(kind, role, kindStories[role], acquire);
      definitions[role] = normalized.definition;
      stories.push(Object.freeze({ source: normalized.definition.root, body: normalized.body }));
    }
    roles[kind] = Object.freeze(definitions);
  }
  return Object.freeze({ roles: Object.freeze(roles), stories: Object.freeze(stories) });
}

/**
 * Build one occurrence from the acquired story in its acquisition frame.
 *
 * Flow: the story paragraph's acquired advance (lines plus authored spacing)
 * is the occurrence's whole reservation; no scalar band is added beside it.
 *
 * Rule ink, library policy separate from paragraph geometry:
 * - span: Full covers the main-story text extent (ECMA-376 §17.11.1; the
 *   U+0004 control). Short keeps the existing one-third library span; §17.11.23
 *   specifies a partial rule but no fraction. Neither span follows paragraph
 *   indentation or alignment: no Office alignment/indent rule is settled.
 * - placement: centred on the control's acquired line box, as the scalar
 *   path centres its rule in its band. Not an Office baseline/gap observation.
 * - stroke: the existing 0.5pt single rule in the control's effective explicit
 *   colour (character properties of the displayed control), else black.
 */
export function noteSeparatorOccurrence(
  definition: ReservedNoteSeparatorDefinitionInput,
  story: StoryLayout,
  inline: Readonly<{ xPt: number; widthPt: number }>,
): NoteSeparatorLayout {
  const [paragraph, ...rest] = story.blocks;
  if (!paragraph || paragraph.kind !== 'paragraph' || rest.length !== 0) {
    throw new Error('A reserved note separator story must acquire exactly one paragraph');
  }
  const flowBounds = Object.freeze({
    xPt: inline.xPt, yPt: story.flowBounds.yPt, widthPt: inline.widthPt, heightPt: story.advancePt,
  });
  let rule: NoteSeparatorLayout['rule'];
  if (definition.rule) {
    const line = paragraph.lines[0];
    if (!line) throw new Error('A reserved note separator rule has no acquired control line');
    const yPt = line.bounds.yPt + line.bounds.heightPt / 2;
    rule = Object.freeze({
      mark: definition.rule.mark,
      source: definition.rule.control,
      segment: Object.freeze({
        edge: 'top' as const,
        from: Object.freeze({ xPt: inline.xPt, yPt }),
        to: Object.freeze({
          xPt: inline.xPt + inline.widthPt * (definition.rule.mark === 'full' ? 1 : 1 / 3),
          yPt,
        }),
        color: definition.rule.color ?? '#000000',
        widthPt: 0.5,
        authoredStyle: 'single',
        style: 'solid' as const,
      }),
    });
  }
  return Object.freeze({
    role: definition.role,
    source: definition.root,
    flowBounds,
    advancePt: story.advancePt,
    paragraph,
    ...(rule ? { rule } : {}),
  });
}

/** Give an occurrence its destination page identity, domain and block offset.
 * The canonical definition source is unchanged. */
export function placeNoteSeparatorOccurrence(
  occurrence: NoteSeparatorLayout,
  destination: Readonly<{ occurrenceId: string; flowDomainId: string; yPt: number }>,
): NoteSeparatorLayout {
  const delta = { xPt: 0, yPt: destination.yPt - occurrence.flowBounds.yPt };
  return Object.freeze({
    ...occurrence,
    flowBounds: translateRect(occurrence.flowBounds, delta),
    ...(occurrence.paragraph ? {
      paragraph: projectBodyOccurrence(occurrence.paragraph, {
        occurrenceId: destination.occurrenceId,
        destination: {
          coordinateSpace: 'logical-page-points',
          flowDomainId: destination.flowDomainId,
          translation: delta,
        },
      }),
    } : {}),
    ...(occurrence.followingParagraphs ? {
      followingParagraphs: Object.freeze(occurrence.followingParagraphs.map((block, index) =>
        projectBodyOccurrence(block, {
          occurrenceId: `${destination.occurrenceId}:${index + 1}`,
          destination: {
            coordinateSpace: 'logical-page-points',
            flowDomainId: destination.flowDomainId,
            translation: delta,
          },
        }))),
    } : {}),
    ...(occurrence.rule ? {
      rule: Object.freeze({ ...occurrence.rule, segment: translateBorder(occurrence.rule.segment, delta) }),
    } : {}),
  });
}

/** Retained blocks a note owns, in paint and reading order. */
export function noteOwnedBlocks(note: NoteLayout): readonly PaintNode[] {
  return [
    ...(note.leading?.paragraph ? [note.leading.paragraph] : []),
    ...(note.leading?.followingParagraphs ?? []),
    ...note.story.blocks,
    ...(note.trailing?.paragraph ? [note.trailing.paragraph] : []),
  ];
}
