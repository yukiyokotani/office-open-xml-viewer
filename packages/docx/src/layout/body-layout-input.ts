import type { SectionLayoutContext } from '../layout-context.js';
import type { HeadersFooters, PageBorders, SectionProps } from '../types.js';
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

export interface BodyParagraphSourceInput {
  readonly kind: 'paragraph';
  readonly source: SourceRef;
  readonly pageBreakBefore: boolean;
  readonly keepLines: boolean;
  readonly keepNext: boolean;
  readonly widowControl: boolean;
  readonly spaceBeforePt: number;
  readonly spaceAfterPt: number;
  readonly contextualSpacing: boolean;
  readonly styleId: string | null;
  /** Source-level visibility used only for pagination look-ahead across unmeasured blocks. */
  readonly inkless?: boolean;
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
  }>;
}

export interface BodyLayoutAcquisitionInput {
  readonly sectionIndex: BodySectionIndexInput;
  readonly evenAndOddHeaders: boolean;
  readonly parserDiagnostics?: readonly LayoutDiagnostic[];
  readonly endnoteIds?: readonly string[];
  readonly noteLayoutSettings: Readonly<{
    footnotePosition: string;
    endnotePosition: string;
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
  const resolved: BodyLayoutSequenceEntry[] = acquired.sequence.map((entry) => {
    if (entry.kind !== 'begin-section') return entry;
    const section = sections.get(entry.section.sectionOccurrenceId);
    if (!section) throw new Error(`Missing body section owner: ${entry.section.sectionOccurrenceId}`);
    return Object.freeze({ ...entry, section });
  });
  const sequence: BodyLayoutSequenceEntry[] = resolved.map((entry, index) => {
    if (entry.kind !== 'body-block' || entry.block.kind !== 'paragraph') return entry;
    const continuousSectionRole = wordContinuousSectionRole(resolved, index);
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
    initialSection,
    sequence,
    parserDiagnostics: acquired.parserDiagnostics ?? [],
    endnoteIds: acquired.endnoteIds ?? [],
    noteLayoutSettings: acquired.noteLayoutSettings,
  }, 'DOCX body layout input') as BodyLayoutInput;
}
