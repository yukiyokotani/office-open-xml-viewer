import { subscribeDocxLayout, type DocxLayoutPublication } from './document-layout-events.js';
import { subscribeDocxLayoutView, type DocxLayoutViewPublication } from './document-layout-view.js';
import type { DocxDocument } from './document';

/** Owns the two publication subscriptions and the paintable page prefix. */
export class DocxScrollLayoutController {
  private unsubscribe: (() => void) | null = null;
  private viewGeneration = 0;
  private prefix = 0;

  constructor(private readonly hooks: {
    current: () => DocxDocument | null;
    destroyed: () => boolean;
    report: (error: unknown) => void;
    reportBackground: (error: unknown) => void;
    invalidateFind: () => void;
    refreshComments: () => void;
    adoptView: (publication: DocxLayoutViewPublication) => void;
    relayout: () => void;
    invalidateRender: () => void;
    mounted: () => Iterable<readonly [number, unknown]>;
    stillMounted: (page: number, slot: unknown) => boolean;
    refreshSlot: (page: number, slot: unknown) => void;
  }) {}

  get presentedPageCount(): number { return this.prefix; }
  set presentedPageCount(count: number) { this.prefix = count; }

  bind(doc: DocxDocument): void {
    this.unbind();
    this.viewGeneration = 0;
    this.prefix = doc.pageCount;
    const unsubscribeView = subscribeDocxLayoutView(
      doc,
      (publication) => this.onViewPublication(doc, publication),
      this.hooks.report,
    );
    let initial = true;
    const unsubscribeLayout = subscribeDocxLayout(
      doc,
      () => ({ pageCount: doc.pageCount, exact: doc.layoutComplete, complete: doc.layoutComplete }),
      (publication) => {
        if (initial) { initial = false; return; }
        this.onPublication(doc, publication);
      },
      this.hooks.report,
    );
    this.unsubscribe = () => { unsubscribeLayout(); unsubscribeView(); };
  }

  unbind(): void {
    this.unsubscribe?.();
    this.unsubscribe = null;
    this.prefix = 0;
  }

  private onPublication(doc: DocxDocument, publication: DocxLayoutPublication): void {
    if (this.hooks.destroyed() || doc !== this.hooks.current()) return;
    if (publication.error !== undefined) {
      if (publication.pageCount === 0) this.prefix = 0;
      this.hooks.reportBackground(publication.error);
      return;
    }
    this.hooks.invalidateFind();
    this.hooks.refreshComments();
    // Publish each paintable prefix. A mounted page keeps its old canvas until
    // the replacement commits, so progressive growth does not blank the view.
    this.apply(publication);
  }

  private onViewPublication(doc: DocxDocument, publication: DocxLayoutViewPublication): void {
    if (this.hooks.destroyed() || doc !== this.hooks.current() ||
        publication.generation <= this.viewGeneration) return;
    this.viewGeneration = publication.generation;
    this.hooks.adoptView(publication);
  }

  apply(publication: DocxLayoutPublication): void {
    const mounted = [...this.hooks.mounted()];
    const presented = this.prefix;
    this.prefix = publication.pageCount;
    // A publication that keeps every presented page identical (typically one
    // that only appends pages) leaves their canvases, painted or still in
    // flight, valid. Discarding them would delay the first visible paint by a
    // whole render per progressive checkpoint.
    if ((publication.unchangedPages ?? 0) >= Math.min(presented, publication.pageCount)) {
      this.hooks.relayout();
      return;
    }
    this.hooks.invalidateRender();
    this.hooks.relayout();
    for (const [page, slot] of mounted) {
      if (page >= this.prefix || !this.hooks.stillMounted(page, slot)) continue;
      this.hooks.refreshSlot(page, slot);
    }
  }

  destroy(): void { this.unbind(); }
}
