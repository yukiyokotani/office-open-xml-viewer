import type { SectionLayoutContext } from '../layout-context.js';
import type {
  DocParagraph,
  DocxTextRun,
  HeadersFooters,
  PageBorders,
  SectionProps,
} from '../types.js';
import {
  resolveAcquiredSectionLayoutContext,
  type BodySectionIndexInput,
  type BodySectionOccurrence,
  type SectionPageLayoutPolicy,
} from './context.js';
import type { SectionStartType, AuthoredBreak } from './paginator.js';
import { snapshotPlainData } from './plain-data.js';
import type { DeepReadonly, LayoutDiagnostic, SourceRef } from './types.js';
import { wordContinuousSectionRole } from './body-pagination-compatibility.js';

/** ECMA-376 §17.11.1/.23 define the mark kind independently of the separator
 * story's ordinary/continuation role. Default is an absent producer fact, not
 * an explicitly authored mark. A producer must validate control
 * semantics/formatting before emitting a mode. */
export type NoteSeparatorMark = 'short' | 'full' | 'none';
export type NoteSeparatorInput = NoteSeparatorMark | 'default';

/** Effective character properties of one character. No run means an
 * effectively hidden character, which differs from an absent character.
 * Parser-private run keys (for example typography acquisition) are retained
 * verbatim. */
export interface ReservedNoteSeparatorCharacterInput {
  readonly run?: DeepReadonly<DocxTextRun>;
  /** A tracked insertion; revision metadata is not retained. */
  readonly insertion?: boolean;
}

/** One reserved separator story (ECMA-376 §17.18.33 `separator`,
 * `continuationSeparator` or `continuationNotice`) in its closed shape: no
 * paragraph is an empty story, a paragraph without a rule is a
 * paragraph-only story, and a rule always comes with its paragraph. */
export interface ReservedNoteSeparatorStoryInput {
  readonly rule?: Readonly<{
    /** Present only when the producer established the control's rule meaning. */
    mark?: Exclude<NoteSeparatorMark, 'none'>;
    control: ReservedNoteSeparatorCharacterInput;
  }>;
  readonly paragraph?: Readonly<{
    paragraph: DeepReadonly<DocParagraph>;
    /** Absent when the story authors no paragraph mark. */
    contentMark?: ReservedNoteSeparatorCharacterInput;
    /** Presence only; numbering/frame properties are not retained. */
    numbered?: boolean;
    framed?: boolean;
  }>;
}

export interface ReservedNoteSeparatorStoriesInput {
  readonly separator: ReservedNoteSeparatorStoryInput;
  readonly continuationSeparator: ReservedNoteSeparatorStoryInput;
  readonly continuationNotice: ReservedNoteSeparatorStoryInput;
}

/** Parser-boundary facts retained once per document for the reserved
 * separator stories a producer admitted; the WordprocessingML parser does not
 * emit them. The source-model adapter replaces them with
 * {@link ReservedNoteSeparatorRolesInput} before the layout source is sealed,
 * so layout never reads the raw wire. */
export interface ReservedNoteSeparatorsInput {
  readonly footnote?: ReservedNoteSeparatorStoriesInput;
  readonly endnote?: ReservedNoteSeparatorStoriesInput;
}

export type ReservedNoteSeparatorRole = keyof ReservedNoteSeparatorStoriesInput;

/** Canonical definition of one reserved separator story: a source definition
 * shared by every page occurrence, never a numbered note. */
export interface ReservedNoteSeparatorDefinitionInput {
  readonly role: ReservedNoteSeparatorRole;
  /** Closed reserved story root (`footnote`/`endnote` + `reserved:*`). */
  readonly root: SourceRef;
  /** The story's authored paragraph `[0]`. `hidden` follows the shared
   * §17.3.1.29/§17.3.2.41 policy for an inkless paragraph whose mark vanishes:
   * it stays source-addressable but owns neither flow nor ink. */
  readonly paragraph?: Readonly<{ source: SourceRef; hidden: boolean }>;
  /** The U+0003/U+0004 control at `[0, 0]`; its mark selects Short/Full. */
  readonly rule?: Readonly<{
    mark: Exclude<NoteSeparatorMark, 'none'>;
    control: SourceRef;
    /** Effective explicit control colour (`#RRGGBB`), absent for auto. */
    color?: string;
  }>;
}

export type ReservedNoteSeparatorDefinitionsInput = Readonly<
  Record<ReservedNoteSeparatorRole, ReservedNoteSeparatorDefinitionInput>
>;

/** At most six canonical definitions per document (two note kinds, three roles). */
export interface ReservedNoteSeparatorRolesInput {
  readonly footnote?: ReservedNoteSeparatorDefinitionsInput;
  readonly endnote?: ReservedNoteSeparatorDefinitionsInput;
}

/** Ordinary DOCX listed separator story roles that can retain an effective
 * paragraph. DOCX continuation notices keep no story geometry. */
export type SelectedNoteSeparatorRole = 'separator' | 'continuationSeparator';

/** Raw parser-boundary effective paragraphs of document-listed (§17.11.3/.9)
 * separator stories. The source-model adapter replaces them
 * with {@link SelectedNoteSeparatorDefinitionsInput} before sealing. */
export type SelectedNoteSeparatorParagraphsInput = Readonly<
  Partial<Record<SelectedNoteSeparatorRole, DeepReadonly<DocParagraph>>>
>;

/** Canonical definition of one ordinary DOCX selected story: a `selected:*`
 * root (never native `reserved:*` provenance or a DOCX note id) shared by
 * every page occurrence, and its single paragraph `[0]`. Its mark stays the
 * role's existing Short/Full separator mode. */
export interface SelectedNoteSeparatorDefinitionInput {
  readonly role: SelectedNoteSeparatorRole;
  readonly root: SourceRef;
  readonly paragraph: SourceRef;
}

export type SelectedNoteSeparatorDefinitionsInput = Readonly<
  Partial<Record<SelectedNoteSeparatorRole, SelectedNoteSeparatorDefinitionInput>>
>;

export interface BodyParagraphSourceInput {
  readonly kind: 'paragraph';
  readonly source: SourceRef;
  readonly pageBreakBefore: boolean;
  readonly keepLines: boolean;
  readonly keepNext: boolean;
  readonly widowControl: boolean;
  /** WORD_TEXT_FRAME_PRECEDING_TEXT_PUSH: detached from paragraph-gap flow. */
  readonly textFrameMember?: true;
  readonly spaceBeforePt: number;
  readonly spaceAfterPt: number;
  readonly contextualSpacing: boolean;
  readonly styleId: string | null;
  /** Source-level visibility used only for pagination look-ahead across unmeasured blocks. */
  readonly inkless?: boolean;
  /** Parser-established text/line-break-only content, including an empty mark; excludes objects. */
  readonly onlyTextualContent?: boolean;
  /** Mutually exclusive Word/LibreOffice section-mark spacing interop role. */
  readonly continuousSectionRole?:
    | 'suppress-before'
    | 'collapse-mark'
    | 'drop-previous-after';
  readonly pageOwnedAnchorOccurrenceIds?: readonly string[];
}

export interface BodyTableSourceInput {
  readonly kind: 'table';
  readonly source: SourceRef;
  readonly rowCount?: number;
  /** §17.4.57 page/margin-positioned table whose exclusion can affect earlier text. */
  readonly pageOwnedFloatingTable?: boolean;
}

export interface BodyAdjacentTableGroupInput {
  readonly kind: 'adjacent-table-group';
  readonly logicalSequenceId: string;
  readonly source: SourceRef;
  readonly tables: readonly BodyTableSourceInput[];
}

export type BodyStoryReferenceSet = Readonly<{
  default: SourceRef | null;
  first: SourceRef | null;
  even: SourceRef | null;
}>;

export interface BodySectionLayoutInput {
  readonly sectionOccurrenceId: string;
  readonly source: SourceRef;
  readonly startType: SectionStartType;
  readonly context: DeepReadonly<SectionLayoutContext>;
  readonly pageNumbering: Readonly<{ start: number | null; format: string | null }>;
  readonly titlePage: boolean;
  readonly evenAndOddHeaders: boolean;
  readonly headers: BodyStoryReferenceSet;
  readonly footers: BodyStoryReferenceSet;
  readonly pageBordersAuthored: boolean;
  readonly pageBorders: DeepReadonly<PageBorders> | null;
  readonly pageLayout: DeepReadonly<SectionPageLayoutPolicy>;
}

export type BodyLayoutSequenceEntryFor<TSection> =
  | Readonly<{ kind: 'body-block'; block: BodyParagraphSourceInput | BodyTableSourceInput }>
  | BodyAdjacentTableGroupInput
  | Readonly<{
      kind: 'authored-break';
      source: SourceRef;
      break: AuthoredBreak;
      origin?: 'authored' | 'coverPageSynthetic';
      parity?: 'odd' | 'even';
      sameSourceParagraphAsPrevious?: boolean;
    }>
  | Readonly<{ kind: 'begin-section'; source: SourceRef; section: TSection }>
  | Readonly<{ kind: 'consume-source'; source: SourceRef; reason: 'hidden-paragraph' }>;

export type BodyLayoutSequenceEntry = BodyLayoutSequenceEntryFor<BodySectionLayoutInput>;

export interface BodyLayoutInput {
  /** `w:compatSetting` compatibilityMode value from settings (§17.15.3.4). */
  readonly compatibilityMode?: number;
  readonly footnoteLayoutLikeWW8?: boolean;
  readonly source: SourceRef;
  readonly initialSection: BodySectionLayoutInput;
  readonly sequence: readonly BodyLayoutSequenceEntry[];
  /** Immutable parse facts; attached only after geometry convergence. */
  readonly parserDiagnostics?: readonly LayoutDiagnostic[];
  /** §17.11 document-end note stories, in authored numbering order. */
  readonly endnoteIds?: readonly string[];
  readonly noteLayoutSettings?: Readonly<{
    footnotePosition: string;
    endnotePosition: string;
    footnoteNumbering?: NoteNumberingInput;
    endnoteNumbering?: NoteNumberingInput;
    footnoteSeparator?: NoteSeparatorInput;
    endnoteSeparator?: NoteSeparatorInput;
    endnoteContinuationSeparator?: NoteSeparatorInput;
    footnoteContinuationSeparator?: NoteSeparatorInput;
    /** Parser-boundary only; rejected by the sealed layout source. */
    footnoteSeparatorParagraphs?: SelectedNoteSeparatorParagraphsInput;
    endnoteSeparatorParagraphs?: SelectedNoteSeparatorParagraphsInput;
    /** Canonical ordinary DOCX formatted story definitions; each replaces
     * only the scalar band height of its role, never its mark. */
    footnoteSeparatorStories?: SelectedNoteSeparatorDefinitionsInput;
    endnoteSeparatorStories?: SelectedNoteSeparatorDefinitionsInput;
    /** Parser-boundary only; rejected by the sealed layout source. */
    reservedSeparators?: ReservedNoteSeparatorsInput;
    /** Canonical reserved separator definitions; they replace the scalar
     * separator modes for their note kind. */
    reservedSeparatorRoles?: ReservedNoteSeparatorRolesInput;
  }>;
}

/** ECMA-376 §17.11.17/.18 numFmt and §17.11.20 numStart for one note kind. */
export interface NoteNumberingInput {
  /** ST_NumberFormat (§17.18.59); `decimal` when not authored. */
  readonly format: string;
  /** First automatic note number; 1 when not authored. */
  readonly start: number;
}

export interface BodyLayoutAcquisitionInput {
  readonly compatibilityMode?: number;
  readonly footnoteLayoutLikeWW8?: boolean;
  readonly sectionIndex: BodySectionIndexInput;
  readonly evenAndOddHeaders: boolean;
  readonly parserDiagnostics?: readonly LayoutDiagnostic[];
  readonly endnoteIds?: readonly string[];
  readonly noteLayoutSettings: Readonly<{
    footnotePosition: string;
    endnotePosition: string;
    footnoteNumbering?: NoteNumberingInput;
    endnoteNumbering?: NoteNumberingInput;
    footnoteSeparator?: NoteSeparatorInput;
    endnoteSeparator?: NoteSeparatorInput;
    endnoteContinuationSeparator?: NoteSeparatorInput;
    footnoteContinuationSeparator?: NoteSeparatorInput;
    footnoteSeparatorParagraphs?: SelectedNoteSeparatorParagraphsInput;
    endnoteSeparatorParagraphs?: SelectedNoteSeparatorParagraphsInput;
    reservedSeparators?: ReservedNoteSeparatorsInput;
  }>;
  readonly pageLayoutSettings: Readonly<{
    mirrorMargins: boolean;
    gutterAtTop: boolean;
    bookFoldPrinting: boolean;
    bookFoldRevPrinting: boolean;
    printTwoOnOne: boolean;
  }>;
  readonly sequence: readonly BodyLayoutSequenceEntryFor<Readonly<{
    sectionOccurrenceId: string;
    startType: string;
  }>>[];
}

export function normalizeSectionStartType(value: string | null | undefined): SectionStartType {
  switch (value) {
    case 'continuous':
    case 'nextColumn':
    case 'nextPage':
    case 'oddPage':
    case 'evenPage':
      return value;
    default:
      return 'nextPage';
  }
}

export function bodyStoryReferences(
  stories: HeadersFooters,
  story: 'header' | 'footer',
  markerBodyIndex: number | null,
): BodyStoryReferenceSet {
  const prefix = markerBodyIndex === null ? null : `section:${markerBodyIndex}`;
  const reference = (kind: 'default' | 'first' | 'even'): SourceRef | null => (
    stories[kind] === null
      ? null
      : {
          story,
          storyInstance: prefix === null ? kind : `${prefix}:${kind}`,
          path: [],
        }
  );
  return Object.freeze({
    default: reference('default'),
    first: reference('first'),
    even: reference('even'),
  });
}

function sectionProps(occurrence: BodySectionOccurrence): SectionProps {
  return {
    ...occurrence.geometry,
    titlePage: occurrence.titlePage,
    evenAndOddHeaders: false,
    sectionStart: occurrence.startType,
    columns: occurrence.columns,
    textDirection: occurrence.textDirection,
    docGridType: occurrence.docGridType,
    docGridLinePitch: occurrence.docGridLinePitch,
    docGridCharSpace: occurrence.docGridCharSpace,
    pageNumType: occurrence.pageNumType,
    vAlign: occurrence.vAlign,
    lineNumbering: occurrence.lineNumbering,
  };
}

function sectionInput(
  occurrence: BodySectionOccurrence,
  acquired: BodyLayoutAcquisitionInput,
): BodySectionLayoutInput {
  const markerBodyIndex = occurrence.markerBodyIndex;
  return Object.freeze({
    sectionOccurrenceId: occurrence.sectionOccurrenceId,
    source: markerBodyIndex === null
      ? Object.freeze({ story: 'body' as const, storyInstance: 'body', path: Object.freeze([]) })
      : Object.freeze({
          story: 'body' as const,
          storyInstance: 'body',
          path: Object.freeze([markerBodyIndex]),
        }),
    startType: normalizeSectionStartType(occurrence.startType),
    context: Object.freeze(resolveAcquiredSectionLayoutContext(
      sectionProps(occurrence),
      occurrence.sectionBidi,
      occurrence.sectionFlow,
    )),
    pageNumbering: Object.freeze({
      start: occurrence.pageNumType?.start ?? null,
      format: occurrence.pageNumType?.fmt ?? null,
    }),
    titlePage: occurrence.titlePage,
    evenAndOddHeaders: acquired.evenAndOddHeaders,
    headers: bodyStoryReferences(occurrence.headers, 'header', markerBodyIndex),
    footers: bodyStoryReferences(occurrence.footers, 'footer', markerBodyIndex),
    pageBordersAuthored: occurrence.pageBordersAuthored,
    pageBorders: occurrence.pageBorders,
    pageLayout: Object.freeze({
      physicalGeometry: Object.freeze({ ...occurrence.geometry }),
      columns: occurrence.columns,
      textDirection: occurrence.textDirection ?? 'lrTb',
      gutterPt: occurrence.gutterPt,
      rtlGutter: occurrence.rtlGutter,
      ...acquired.pageLayoutSettings,
    }),
  });
}

export function projectBodyLayoutInput(acquired: BodyLayoutAcquisitionInput): BodyLayoutInput {
  const sections = new Map(acquired.sectionIndex.occurrences.map((occurrence) => [
    occurrence.sectionOccurrenceId,
    sectionInput(occurrence, acquired),
  ]));
  const initialOccurrence = acquired.sectionIndex.occurrences[0];
  if (!initialOccurrence) throw new Error('DOCX body requires a final section owner');
  const initialSection = sections.get(initialOccurrence.sectionOccurrenceId)!;
  // The spacing role depends only on the authored section start type, already
  // present in the acquisition sequence. Resolve section owners and roles in
  // one pass so no second body-sized sequence is retained transiently.
  const sequence: BodyLayoutSequenceEntry[] = acquired.sequence.map((entry, index) => {
    if (entry.kind === 'begin-section') {
      const section = sections.get(entry.section.sectionOccurrenceId);
      if (!section) throw new Error(`Missing body section owner: ${entry.section.sectionOccurrenceId}`);
      return Object.freeze({ ...entry, section });
    }
    if (entry.kind !== 'body-block' || entry.block.kind !== 'paragraph') return entry;
    const continuousSectionRole = wordContinuousSectionRole(acquired.sequence, index);
    if (continuousSectionRole === undefined) return entry;
    return Object.freeze({
      ...entry,
      block: Object.freeze({
        ...entry.block,
        continuousSectionRole,
      }),
    });
  });
  return snapshotPlainData({
    source: { story: 'body', storyInstance: 'body', path: [] },
    compatibilityMode: acquired.compatibilityMode,
    footnoteLayoutLikeWW8: acquired.footnoteLayoutLikeWW8,
    initialSection,
    sequence,
    parserDiagnostics: acquired.parserDiagnostics ?? [],
    endnoteIds: acquired.endnoteIds ?? [],
    noteLayoutSettings: acquired.noteLayoutSettings,
  }, 'DOCX body layout input') as BodyLayoutInput;
}
