import { UnitFindController } from '@silurus/ooxml-core/internal/find-controller';
import { runTextVisible } from './overlay-geometry.js';
import type { DocxTextRunInfo } from './renderer';

export interface DocxMatchLocation { page: number; }

/** DOCX names each unit as a page; search state and cancellation live in core. */
export class DocxFindController extends UnitFindController<DocxTextRunInfo, DocxMatchLocation> {
  constructor(count: () => number, collectRuns: (page: number) => Promise<DocxTextRunInfo[]>) {
    // A match is text a reader can see: one whose glyphs its clips hide
    // whole (overlay-geometry.ts) could never be highlighted, and is none.
    super(count, collectRuns, (page) => ({ page }), (runs, slices) => slices.some((slice) => {
      const run = runs[slice.runIndex];
      return run !== undefined && (run.clips === undefined || runTextVisible(run, slice.start, slice.end));
    }));
  }

  pageRuns(page: number): DocxTextRunInfo[] | undefined { return this.unitRuns(page); }
  setPageRuns(page: number, runs: DocxTextRunInfo[]): void { this.setUnitRuns(page, runs); }
  pageHighlights(page: number): ReturnType<DocxFindController['unitHighlights']> {
    return this.unitHighlights(page);
  }
  activePage(): number | null { return this.activeUnit(); }
}
