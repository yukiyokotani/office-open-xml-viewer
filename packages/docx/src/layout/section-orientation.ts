import type { DocxDocumentModel, SectionProps } from '../types.js';
import type { SectionFlow } from './types.js';
import {
  isVerticalSectionDirection,
  logicalSectionGeometry,
  physicalSectionGeometry,
} from './context.js';
import { sectionWritingMode } from './coordinate-space.js';

/** True when a section flows vertically. The Transitional ST_TextDirection
 * values `tbRl`, `tbRlV`, `tbLrV`, and `btLr` share the rotated page frame (a
 * bottomToTop section flow is nominally `btLr`; its own frame is selected by its
 * canonical `sectionFlow`). */
export function isVerticalSection(section: SectionProps): boolean {
  return isVerticalTextDirection(section.textDirection);
}

/** Raw-token predicate for section-region carriers that do not own a complete
 * SectionProps value. Unknown and absent values remain horizontal. */
export function isVerticalTextDirection(
  textDirection: string | null | undefined,
): boolean {
  return typeof textDirection === 'string' && isVerticalSectionDirection(textDirection);
}

/** `btLr` uses the rotated page frame without upright-CJK counter-rotation. */
export function isAllRotatedVerticalTextDirection(
  textDirection: string | null | undefined,
): boolean {
  return textDirection === 'btLr';
}

/** Writing mode of a vertical section's own frame (its text direction, or
 * the canonical frame its section flow selects). */
function verticalSectionWritingMode(section: SectionProps, sectionFlow?: SectionFlow) {
  if (!isVerticalSection(section)) {
    throw new RangeError(`Section text direction ${JSON.stringify(section.textDirection)} is not vertical`);
  }
  return sectionWritingMode({ textDirection: section.textDirection as string, sectionFlow });
}

/** Map physical vertical-section geometry into the logical horizontal frame
 * used by retained acquisition. ECMA-376 §17.6.11 defines the page margins. */
export function verticalLayoutSection(
  physical: SectionProps,
  sectionFlow?: SectionFlow,
): SectionProps {
  return {
    ...physical,
    ...logicalSectionGeometry(physical, verticalSectionWritingMode(physical, sectionFlow)),
  };
}

/** Replace only the body-level section with its logical vertical frame.
 * Per-body section-break geometry remains physical until the paginator selects
 * each section's own frame. Horizontal documents preserve object identity. */
export function verticalLayoutDoc(document: DocxDocumentModel): DocxDocumentModel {
  if (!isVerticalSection(document.section)) return document;
  return { ...document, section: verticalLayoutSection(document.section) };
}

/** Map a logical vertical frame back to physical page geometry. */
export function physicalLayoutSection(
  logical: SectionProps,
  sectionFlow?: SectionFlow,
): SectionProps {
  return {
    ...logical,
    ...physicalSectionGeometry(logical, verticalSectionWritingMode(logical, sectionFlow)),
  };
}
