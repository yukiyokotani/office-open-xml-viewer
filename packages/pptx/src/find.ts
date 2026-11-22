import { UnitFindController } from '@silurus/ooxml-core/internal/find-controller';
import type { TextMatch } from '@silurus/ooxml-core';
import type { PptxTextRunInfo } from './renderer';

export interface PptxMatchLocation { slide: number; }

function sameSearchContainer(left: PptxTextRunInfo, right: PptxTextRunInfo): boolean {
  const leftCell = left.tableCell;
  const rightCell = right.tableCell;
  if (!leftCell && !rightCell) return true;
  if (!leftCell || !rightCell) return false;
  // Runtime MCE previews have paint identities but no source-model locators.
  // Keep cells from separate preview tables distinct in both render modes.
  return (left.renderElementIndex ?? left.elementIndex) === (right.renderElementIndex ?? right.elementIndex) &&
    left.origin === right.origin &&
    left.shapeId === right.shapeId &&
    leftCell.row === rightCell.row &&
    leftCell.column === rightCell.column;
}

/** Drawing runs may join within a text body, but table cells are independent
 * text containers. Keep this format boundary in the adapter. */
function matchStaysWithinSearchContainer(runs: PptxTextRunInfo[], slices: TextMatch['slices']): boolean {
  for (let index = 1; index < slices.length; index++) {
    const left = runs[slices[index - 1].runIndex];
    const right = runs[slices[index].runIndex];
    if (!left || !right || !sameSearchContainer(left, right)) return false;
  }
  return true;
}

export class PptxFindController extends UnitFindController<PptxTextRunInfo, PptxMatchLocation> {
  constructor(
    count: () => number,
    collectRuns: (slide: number) => Promise<PptxTextRunInfo[]>,
    includeSlide?: (slide: number) => boolean,
  ) {
    super(count, collectRuns, (slide) => ({ slide }), matchStaysWithinSearchContainer, includeSlide);
  }

  slideRuns(slide: number): PptxTextRunInfo[] | undefined { return this.unitRuns(slide); }
  setSlideRuns(slide: number, runs: PptxTextRunInfo[]): void { this.setUnitRuns(slide, runs); }
  slideHighlights(slide: number): ReturnType<PptxFindController['unitHighlights']> {
    return this.unitHighlights(slide);
  }
  activeSlide(): number | null { return this.activeUnit(); }
}
