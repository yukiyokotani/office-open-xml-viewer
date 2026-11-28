import { UnitFindController } from '@silurus/ooxml-core/internal/find-controller';
import { runGlyphRects, runRectVisible } from './overlay-geometry.js';
import type { DocxTextRunInfo } from './renderer';

export interface DocxMatchLocation { page: number; }

/** DOCX names each unit as a page; search state and cancellation live in core. */
export class DocxFindController extends UnitFindController<DocxTextRunInfo, DocxMatchLocation> {
  constructor(count: () => number, collectRuns: (page: number) => Promise<DocxTextRunInfo[]>) {
    // A match is text a reader can see (overlay-geometry.ts): every glyph
    // it paints lies, at least in part, inside its clips, and it paints one.
    // An occurrence that takes in a hidden character is none, so the hidden
    // character cannot take a visible one away from the occurrence a reader
    // sees beside it.
    super(count, collectRuns, (page) => ({ page }), (runs, slices) => {
      let seen = false;
      for (const slice of slices) {
        const run = runs[slice.runIndex];
        if (run === undefined) return false;
        if (run.clips === undefined) {
          seen = true;
          continue;
        }
        const glyphs = runGlyphRects(run, slice.start, slice.end);
        if (!glyphs.every((rect) => runRectVisible(run, rect))) return false;
        seen ||= glyphs.length > 0;
      }
      return seen;
    });
  }

  pageRuns(page: number): DocxTextRunInfo[] | undefined { return this.unitRuns(page); }
  setPageRuns(page: number, runs: DocxTextRunInfo[]): void { this.setUnitRuns(page, runs); }
  pageHighlights(page: number): ReturnType<DocxFindController['unitHighlights']> {
    return this.unitHighlights(page);
  }
  activePage(): number | null { return this.activeUnit(); }
}
