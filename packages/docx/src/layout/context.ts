import {
  sectionPlacementInputFromBody,
  type SectionPlacementInput,
} from '../parser-model.js';
import type {
  BodyElement,
  ColumnGeom,
  ColumnsSpec,
  DocxDocumentModel,
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
  return Math.abs(section.geometry.marginTop);
}

type SectionBreak = Extract<BodyElement, { type: 'sectionBreak' }>;

const EMPTY_HEADERS_FOOTERS: HeadersFooters = Object.freeze({
  default: null,
  first: null,
  even: null,
});

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
  readonly placement: SectionPlacementInput;
}

export interface BodySectionIndex {
  readonly occurrences: readonly BodySectionOccurrence[];
  /** Accepts body.length as the insertion point owned by the final section. */
  sectionAtBodyIndex(bodyIndex: number): BodySectionOccurrence;
}

function sectionGeometry(section: SectionProps): SectionGeom {
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

function endingSectionOccurrence(
  marker: SectionBreak,
  markerBodyIndex: number,
  startBodyIndex: number,
  ordinal: number,
  placement: SectionPlacementInput,
  inheritedGeometry: SectionGeom,
): BodySectionOccurrence {
  return Object.freeze({
    sectionOccurrenceId: placement.sectionId,
    ordinal,
    startBodyIndex,
    endBodyIndex: markerBodyIndex,
    markerBodyIndex,
    final: false,
    startType: marker.kind ?? 'nextPage',
    columns: marker.columns ?? null,
    // A paragraph-owned sectPr may omit pgSz/pgMar. The parser deliberately
    // preserves that absence; the existing renderer inherits the final physical
    // page box rather than manufacturing spec defaults at this later boundary.
    geometry: marker.geom ?? inheritedGeometry,
    textDirection: marker.textDirection ?? null,
    pageNumType: marker.pageNumType ?? null,
    headers: marker.headers ?? EMPTY_HEADERS_FOOTERS,
    footers: marker.footers ?? EMPTY_HEADERS_FOOTERS,
    titlePage: marker.titlePage ?? false,
    vAlign: placement.vAlign,
    lineNumbering: placement.lineNumbering,
    placement,
  });
}

function finalSectionOccurrence(
  doc: DocxDocumentModel,
  startBodyIndex: number,
  ordinal: number,
  placement: SectionPlacementInput,
  geometry: SectionGeom,
): BodySectionOccurrence {
  return Object.freeze({
    sectionOccurrenceId: placement.sectionId,
    ordinal,
    startBodyIndex,
    endBodyIndex: doc.body.length - 1,
    markerBodyIndex: null,
    final: true,
    startType: doc.section.sectionStart ?? 'nextPage',
    columns: doc.section.columns ?? null,
    geometry,
    textDirection: doc.section.textDirection ?? null,
    pageNumType: doc.section.pageNumType ?? null,
    headers: doc.headers ?? EMPTY_HEADERS_FOOTERS,
    footers: doc.footers ?? EMPTY_HEADERS_FOOTERS,
    titlePage: doc.section.titlePage,
    vAlign: placement.vAlign,
    lineNumbering: placement.lineNumbering,
    placement,
  });
}

/**
 * Pre-index ECMA-376 §17.6.18 paragraph-owned sectPr markers and the final
 * body-level §17.6.17 sectPr. A marker belongs to the section it terminates;
 * ownership switches at the following body index. Construction is O(body), and
 * every subsequent source-index lookup is a single array access.
 */
export function createBodySectionIndex(doc: DocxDocumentModel): BodySectionIndex {
  const inheritedGeometry = Object.freeze(sectionGeometry(doc.section));
  const occurrences: BodySectionOccurrence[] = [];
  const occurrenceOrdinalByBodyIndex = new Array<number>(doc.body.length + 1);
  let startBodyIndex = 0;

  for (let bodyIndex = 0; bodyIndex < doc.body.length; bodyIndex += 1) {
    const element = doc.body[bodyIndex];
    if (element?.type !== 'sectionBreak') continue;

    const ordinal = occurrences.length;
    const placement = sectionPlacementInputFromBody(doc.body, doc.section, bodyIndex);
    occurrences.push(endingSectionOccurrence(
      element,
      bodyIndex,
      startBodyIndex,
      ordinal,
      placement,
      inheritedGeometry,
    ));
    for (let ownedIndex = startBodyIndex; ownedIndex <= bodyIndex; ownedIndex += 1) {
      occurrenceOrdinalByBodyIndex[ownedIndex] = ordinal;
    }
    startBodyIndex = bodyIndex + 1;
  }

  const finalOrdinal = occurrences.length;
  const finalPlacement = sectionPlacementInputFromBody(
    doc.body,
    doc.section,
    doc.body.length,
  );
  occurrences.push(finalSectionOccurrence(
    doc,
    startBodyIndex,
    finalOrdinal,
    finalPlacement,
    inheritedGeometry,
  ));
  for (let ownedIndex = startBodyIndex; ownedIndex <= doc.body.length; ownedIndex += 1) {
    occurrenceOrdinalByBodyIndex[ownedIndex] = finalOrdinal;
  }

  const retainedOccurrences = Object.freeze(occurrences);
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
