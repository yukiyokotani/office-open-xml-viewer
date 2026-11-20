import { UnitFindController } from '@silurus/ooxml-core/internal/find-controller';
import type { DocxTextRunInfo } from './renderer';

export interface DocxMatchLocation { page: number; }

/** DOCX names each unit as a page; search state and cancellation live in core. */
export class DocxFindController extends UnitFindController<DocxTextRunInfo, DocxMatchLocation> {
  constructor(count: () => number, collectRuns: (page: number) => Promise<DocxTextRunInfo[]>) {
    super(
      count, collectRuns, (page) => ({ page }), undefined, undefined,
      // Whole-word matching only; kept out of the static bundle.
      () => import('./find-run-break.js'),
    );
  }

  pageRuns(page: number): DocxTextRunInfo[] | undefined { return this.unitRuns(page); }
  setPageRuns(page: number, runs: DocxTextRunInfo[]): void { this.setUnitRuns(page, runs); }
  pageHighlights(page: number): ReturnType<DocxFindController['unitHighlights']> {
    return this.unitHighlights(page);
  }
}
