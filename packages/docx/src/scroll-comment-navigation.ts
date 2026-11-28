import type { DocxDocument } from './document';
import type { DocxTextRunInfo } from './renderer';
import { commentAnchorPlacement, resolveCommentAnchorRunsOnPage } from './comments';
import { overlayBoxes } from './overlay-geometry';
import { visiblePoint } from './run-clip';

type ScrollBehavior = { behavior?: 'auto' | 'smooth' };

interface CommentNavigationHooks {
  document(): DocxDocument | null;
  destroyed(): boolean;
  scale(): number;
  pageWidth(page: number): number;
  currentDate(): Date | number | undefined;
  showTrackedChanges(): boolean;
  waitForLayout(doc: DocxDocument): Promise<void>;
  /** `point` (page CSS px) lies in the part of the comment that is seen. */
  select(commentId: string, page: number, point: Readonly<{ x: number; y: number }>, options?: ScrollBehavior): void;
}

/** A point of what the page shows of an anchor's `runs`: one that every clip
 * of the first visible box of its tint shows (overlay-geometry.ts
 * overlayBoxes, run-clip.ts visiblePoint). The boxes are those of the tint,
 * made from the runs together: a tab is marked beside the text it leads,
 * not over the whole height of its line. Undefined when the clips hide
 * everything. */
function shownPoint(runs: readonly Readonly<DocxTextRunInfo>[]): Readonly<{ x: number; y: number }> | undefined {
  for (const box of overlayBoxes(runs)) {
    const point = visiblePoint(box, box.clips);
    if (point) return point;
  }
  return undefined;
}

/** DOCX anchor-to-page index for application-owned comment list navigation.
 * A comment is reached where it is seen: on the page that carries its card
 * (CommentAnchorPlacement.cardPageIndex), else on the first page that shows
 * something of an anchor of it, at the first run with a visible part, and
 * at a point of that part: the middle of the run's own box may lie where a
 * clip hides it (a picture wider than its fixed cell). A run its clips hide
 * whole is not a destination. Pages are indexed in order
 * because repeated header anchors must resolve to their earliest occurrence;
 * run collection is shared by concurrent requests. */
export class DocxScrollCommentNavigation {
  private readonly pageById = new Map<string, number>();
  private readonly runsByPage = new Map<number, {
    readonly scale: number;
    readonly runs: Promise<readonly Readonly<DocxTextRunInfo>[]>;
  }>();
  private readonly indexedPages = new Set<number>();
  private frontier = 0;
  private generation = 0;

  constructor(private readonly hooks: CommentNavigationHooks) {}

  reset(): void {
    this.generation++;
    this.pageById.clear();
    this.runsByPage.clear();
    this.indexedPages.clear();
    this.frontier = 0;
  }

  commitRuns(page: number, runs: readonly Readonly<DocxTextRunInfo>[]): void {
    this.runsByPage.set(page, { scale: this.hooks.scale(), runs: Promise.resolve(runs) });
    const doc = this.hooks.document();
    if (doc && page === this.frontier) this.indexPage(page, runs, doc.commentAnchorRanges());
  }

  private indexPage(page: number, runs: readonly Readonly<DocxTextRunInfo>[],
    anchors: ReturnType<DocxDocument['commentAnchorRanges']>): void {
    if (this.indexedPages.has(page)) return;
    for (const anchor of anchors) {
      if (!this.pageById.has(anchor.commentId)
        && overlayBoxes(resolveCommentAnchorRunsOnPage(anchor, runs, page)).length > 0) {
        this.pageById.set(anchor.commentId, page);
      }
    }
    this.indexedPages.add(page);
    while (this.indexedPages.has(this.frontier)) this.frontier++;
  }

  private async runsForPage(page: number, doc: DocxDocument):
    Promise<readonly Readonly<DocxTextRunInfo>[] | null> {
    while (!this.hooks.destroyed() && this.hooks.document() === doc) {
      const scale = this.hooks.scale();
      let entry = this.runsByPage.get(page);
      if (!entry || entry.scale !== scale) {
        const runs = doc.collectPageRuns(page, {
          width: this.hooks.pageWidth(page),
          currentDate: this.hooks.currentDate(),
          ...(this.hooks.showTrackedChanges() ? { showTrackedChanges: true } : {}),
        });
        entry = { scale, runs };
        this.runsByPage.set(page, entry);
      }
      try {
        const runs = await entry.runs;
        if (this.hooks.destroyed() || this.hooks.document() !== doc) return null;
        if (this.hooks.scale() !== scale) continue;
        return runs;
      } catch (error) {
        if (this.runsByPage.get(page) === entry) this.runsByPage.delete(page);
        throw error;
      }
    }
    return null;
  }

  async goToComment(commentId: string,
    opts?: { pageIndex?: number; behavior?: 'auto' | 'smooth' }): Promise<boolean> {
    if (this.hooks.destroyed()) throw new Error('DocxScrollViewer is destroyed');
    const doc = this.hooks.document();
    if (!doc || !doc.comments.some((comment) =>
      comment.id === commentId && comment.parentId === undefined)) return false;
    const generation = ++this.generation;
    const provisional = !doc.layoutComplete;
    let anchors = doc.commentAnchorRanges().filter((anchor) => anchor.commentId === commentId);
    const requestedPage = opts?.pageIndex;
    if (requestedPage !== undefined && (!Number.isInteger(requestedPage) || requestedPage < 0)) {
      return false;
    }
    const cardPage = (): number | undefined => {
      for (const anchor of anchors) {
        const page = commentAnchorPlacement(anchor)?.cardPageIndex;
        if (page !== undefined) return page;
      }
      return undefined;
    };
    let page = requestedPage ?? cardPage() ?? this.pageById.get(commentId);
    const scanAvailablePages = async (): Promise<void> => {
      const allAnchors = doc.commentAnchorRanges();
      while (page === undefined && this.frontier < doc.pageCount) {
        const index = this.frontier;
        const runs = await this.runsForPage(index, doc);
        if (this.hooks.destroyed()) throw new Error('DocxScrollViewer is destroyed');
        if (this.hooks.document() !== doc || generation !== this.generation || !runs) return;
        this.indexPage(index, runs, allAnchors);
        page = this.pageById.get(commentId);
      }
    };
    if (requestedPage === undefined && page === undefined && anchors.length > 0) {
      await scanAvailablePages();
    }
    if (provisional && ((requestedPage !== undefined && requestedPage >= doc.pageCount) ||
      (requestedPage === undefined && page === undefined))) {
      await this.hooks.waitForLayout(doc);
      if (this.hooks.destroyed()) throw new Error('DocxScrollViewer is destroyed');
      if (this.hooks.document() !== doc || generation !== this.generation) return false;
      anchors = doc.commentAnchorRanges().filter((anchor) => anchor.commentId === commentId);
      this.pageById.clear();
      this.runsByPage.clear();
      this.indexedPages.clear();
      this.frontier = 0;
      page = requestedPage ?? cardPage();
      if (page === undefined && anchors.length > 0) await scanAvailablePages();
    }
    if (anchors.length === 0 || page === undefined ||
        (requestedPage !== undefined && requestedPage >= doc.pageCount)) return false;
    const runs = await this.runsForPage(page, doc);
    if (this.hooks.destroyed()) throw new Error('DocxScrollViewer is destroyed');
    if (this.hooks.document() !== doc || generation !== this.generation || !runs) return false;
    const shownPage = page;
    let target: Readonly<{ x: number; y: number }> | undefined;
    for (const anchor of anchors) {
      target = shownPoint(resolveCommentAnchorRunsOnPage(anchor, runs, shownPage));
      if (target) break;
    }
    if (!target) return false;
    this.hooks.select(commentId, page, target, opts);
    return true;
  }

  destroy(): void { this.reset(); }
}
