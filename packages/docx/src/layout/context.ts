import type {
  ColumnGeom,
  ColumnsSpec,
  HeadersFooters,
  LineNumbering,
  PageNumType,
  SectionGeom,
  SectionProps,
} from '../types.js';

/**
 * Section facts that must change atomically at a page-flow boundary. Keeping the
 * occurrence identity beside geometry prevents two equal-looking consecutive
 * sections from being mistaken for the same section by page-number/header logic.
 */
export interface PageFlowSectionContext {
  readonly sectionOccurrenceId: string;
  readonly geometry: Readonly<SectionGeom>;
  readonly columns: readonly Readonly<ColumnGeom>[];
  readonly textDirection: string;
}

export function createPageFlowSectionContext(input: Readonly<{
  sectionOccurrenceId: string;
  geometry: SectionGeom;
  columns: readonly Readonly<ColumnGeom>[];
  textDirection: string;
}>): PageFlowSectionContext {
  if (input.sectionOccurrenceId.length === 0) {
    throw new RangeError('Section occurrence id must not be empty');
  }
  if (input.columns.length === 0) {
    throw new RangeError('A page-flow section requires at least one column');
  }
  return Object.freeze({
    sectionOccurrenceId: input.sectionOccurrenceId,
    geometry: Object.freeze({ ...input.geometry }),
    columns: Object.freeze(input.columns.map((column) => Object.freeze({ ...column }))),
    textDirection: input.textDirection,
  });
}

/** §17.6.11 permits signed top/bottom margins, but body flow uses their distance
 * from the page edge; the sign controls header/footer overlap separately. */
export function sectionContentStartBlockPt(section: PageFlowSectionContext): number {
  return sectionBodyInsetPt(section.geometry.marginTop);
}

/** Signed top/bottom margins retain overlap policy; body placement uses distance. */
export function sectionBodyInsetPt(marginPt: number): number {
  return Math.abs(marginPt);
}

/** Physical-to-logical quarter turn for vertical section body layout. */
export function logicalSectionGeometry(physical: SectionGeom): SectionGeom {
  return {
    pageWidth: physical.pageHeight,
    pageHeight: physical.pageWidth,
    marginLeft: physical.marginTop,
    marginTop: physical.marginRight,
    marginRight: physical.marginBottom,
    marginBottom: physical.marginLeft,
    headerDistance: physical.headerDistance,
    footerDistance: physical.footerDistance,
  };
}

/** Inverse logical-to-physical quarter turn for a vertical section page box. */
export function physicalSectionGeometry(logical: SectionGeom): SectionGeom {
  return {
    pageWidth: logical.pageHeight,
    pageHeight: logical.pageWidth,
    marginTop: logical.marginLeft,
    marginRight: logical.marginTop,
    marginBottom: logical.marginRight,
    marginLeft: logical.marginBottom,
    headerDistance: logical.headerDistance,
    footerDistance: logical.footerDistance,
  };
}

export interface SectionPlacementFacts {
  readonly sectionId: string;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
}

/**
 * One lexical section occurrence in body order. Equal section properties do not
 * make two occurrences interchangeable: page numbering, title-page selection,
 * and line-number restart rules are occurrence-sensitive.
 */
export interface BodySectionOccurrence {
  readonly sectionOccurrenceId: string;
  readonly ordinal: number;
  /** First body item owned by this occurrence. */
  readonly startBodyIndex: number;
  /** Last body item owned by this occurrence, inclusive. */
  readonly endBodyIndex: number;
  /** The paragraph-owned sectPr marker which terminates this occurrence. */
  readonly markerBodyIndex: number | null;
  readonly final: boolean;
  /** ECMA-376 §17.6.22: how this section starts relative to its predecessor. */
  readonly startType: string;
  readonly columns: ColumnsSpec | null;
  /** Physical §17.6.13/§17.6.11 page box; writing-mode transforms happen later. */
  readonly geometry: SectionGeom;
  readonly textDirection: string | null;
  readonly pageNumType: PageNumType | null;
  readonly headers: HeadersFooters;
  readonly footers: HeadersFooters;
  readonly titlePage: boolean;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
  readonly placement: SectionPlacementFacts;
}

/** Complete parser-boundary projection; layout does not scan document nodes. */
export interface BodySectionIndexInput {
  readonly bodyLength: number;
  readonly occurrences: readonly BodySectionOccurrence[];
}

export interface BodySectionIndex {
  readonly occurrences: readonly BodySectionOccurrence[];
  /** Accepts body.length as the insertion point owned by the final section. */
  sectionAtBodyIndex(bodyIndex: number): BodySectionOccurrence;
}

export function sectionGeometry(section: SectionProps): SectionGeom {
  return {
    pageWidth: section.pageWidth,
    pageHeight: section.pageHeight,
    marginTop: section.marginTop,
    marginRight: section.marginRight,
    marginBottom: section.marginBottom,
    marginLeft: section.marginLeft,
    headerDistance: section.headerDistance,
    footerDistance: section.footerDistance,
  };
}

/**
 * Index a complete §17.6.18/§17.6.17 occurrence projection. Construction is
 * O(occurrences + body), and subsequent source-index lookup is one array access.
 */
export function createBodySectionIndex(input: BodySectionIndexInput): BodySectionIndex {
  if (!Number.isInteger(input.bodyLength) || input.bodyLength < 0 || input.occurrences.length === 0) {
    throw new RangeError('A body section index requires a non-negative length and occurrences');
  }
  const occurrenceOrdinalByBodyIndex = new Array<number>(input.bodyLength + 1);
  let expectedStart = 0;
  input.occurrences.forEach((occurrence, ordinal) => {
    const last = ordinal === input.occurrences.length - 1;
    if (
      occurrence.ordinal !== ordinal
      || occurrence.startBodyIndex !== expectedStart
      || occurrence.endBodyIndex !== (last ? input.bodyLength - 1 : occurrence.markerBodyIndex)
      || occurrence.final !== last
      || (last ? occurrence.markerBodyIndex !== null : occurrence.markerBodyIndex === null)
    ) {
      throw new RangeError(`Invalid section occurrence ${ordinal}`);
    }
    for (
      let ownedIndex = occurrence.startBodyIndex;
      ownedIndex <= occurrence.endBodyIndex;
      ownedIndex += 1
    ) {
      occurrenceOrdinalByBodyIndex[ownedIndex] = ordinal;
    }
    expectedStart = occurrence.endBodyIndex + 1;
  });
  const finalOrdinal = input.occurrences.length - 1;
  occurrenceOrdinalByBodyIndex[input.bodyLength] = finalOrdinal;
  const retainedOccurrences = Object.freeze([...input.occurrences]);
  const retainedOrdinals = Object.freeze(occurrenceOrdinalByBodyIndex);
  return Object.freeze({
    occurrences: retainedOccurrences,
    sectionAtBodyIndex(bodyIndex: number): BodySectionOccurrence {
      if (!Number.isInteger(bodyIndex) || bodyIndex < 0 || bodyIndex >= retainedOrdinals.length) {
        throw new RangeError(`Body index ${bodyIndex} is outside the retained section index`);
      }
      return retainedOccurrences[retainedOrdinals[bodyIndex]!]!;
    },
  });
}
