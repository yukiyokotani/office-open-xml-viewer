import type { ColumnGeom, SectionGeom } from '../types.js';

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
