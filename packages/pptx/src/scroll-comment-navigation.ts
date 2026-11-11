import { EMU_PER_PX } from '@silurus/ooxml-core';
import type { PptxPresentation } from './presentation';
import type { PptxComment } from './types';
import type { PptxElementBounds } from './element-selection';
import { pptxCommentOccurrenceKey } from './comment-occurrence';

type Bounds = Readonly<{ x: number; y: number; width: number; height: number }>;
type Options = { behavior?: 'auto' | 'smooth' };

interface NavigationSlot { commentElementBounds: readonly PptxElementBounds[]; }

interface NavigationHooks<Slot extends NavigationSlot> {
  presentation(): PptxPresentation | null;
  destroyed(): boolean;
  slots(): ReadonlyMap<number, Slot>;
  scale(): number;
  scrollPoint(slide: number, x: number, y: number, options?: Options): void;
  scrollToSlide(slide: number, options?: Options): void;
  /** Whether `slide` is laid out (`hiddenSlideMode: 'skip'` leaves hidden
   *  slides out). Asked once the slide's metadata is available. */
  reachable?(slide: number): boolean;
  select(commentId: string, slide: number): void;
  ownBackground<T>(operation: () => Promise<T>): Promise<T>;
}

/** PPTX-only comment occurrence navigation and progressive slide availability. */
export class PptxScrollCommentNavigation<Slot extends NavigationSlot> {
  private generation = 0;
  private readonly layoutWaiters = new Set<() => void>();
  private layoutFailed = false;

  constructor(private readonly hooks: NavigationHooks<Slot>) {}

  resetLayout(): void {
    this.layoutFailed = false;
    this.wake();
  }

  failLayout(): void {
    this.layoutFailed = true;
    this.wake();
  }

  wake(): void {
    for (const resolve of this.layoutWaiters) resolve();
    this.layoutWaiters.clear();
  }

  begin(): number {
    const generation = ++this.generation;
    this.wake();
    return generation;
  }

  private async waitForSlideMetadata(presentation: PptxPresentation,
    slide: number, generation: number): Promise<boolean> {
    return await this.hooks.ownBackground(async () => {
      while (!this.hooks.destroyed() && generation === this.generation &&
        presentation === this.hooks.presentation() &&
        slide >= presentation.availableSlideCount && !presentation.layoutComplete &&
        !this.layoutFailed) {
        await new Promise<void>((resolve) => this.layoutWaiters.add(resolve));
      }
      if (this.hooks.destroyed() || presentation !== this.hooks.presentation()) return false;
      if (presentation.layoutComplete || this.layoutFailed) {
        await presentation.waitUntilLayoutComplete?.();
      }
      if (generation !== this.generation) return false;
      return slide < presentation.availableSlideCount;
    });
  }

  private async resolveBounds(slide: number, comment: Readonly<PptxComment>):
    Promise<Bounds | undefined> {
    const presentation = this.hooks.presentation();
    if (!presentation) return undefined;
    const elementIds = (comment.anchors ?? []).flatMap((anchor) =>
      (anchor.type === 'drawingElement' || anchor.type === 'textRange') && anchor.elementId
        ? [anchor.elementId] : []);
    if (elementIds.length === 0) return undefined;
    const cached = new Map((this.hooks.slots().get(slide)?.commentElementBounds ?? [])
      .map((entry) => [entry.elementId, entry.bounds]));
    const cachedTarget = elementIds.flatMap((id) => {
      const bounds = cached.get(id);
      return bounds ? [bounds] : [];
    })[0];
    if (cachedTarget) return cachedTarget;
    const bounds = await presentation.getElementBoundsByIds(slide, elementIds);
    return elementIds.flatMap((id) => {
      const entry = bounds.find((candidate) => candidate.elementId === id);
      return entry ? [entry.bounds] : [];
    })[0];
  }

  scrollToTarget(slide: number, comment: Readonly<PptxComment>,
    options?: Options, resolvedBounds?: Bounds): boolean {
    if (!this.hooks.presentation()) return false;
    const slot = this.hooks.slots().get(slide);
    const boundsById = new Map((slot?.commentElementBounds ?? [])
      .map((entry) => [entry.elementId, entry.bounds]));
    const anchored = resolvedBounds ?? (comment.anchors ?? []).flatMap((anchor) => {
      if ((anchor.type !== 'drawingElement' && anchor.type !== 'textRange') || !anchor.elementId) {
        return [];
      }
      const bounds = boundsById.get(anchor.elementId);
      return bounds ? [bounds] : [];
    })[0];
    const anchors = comment.anchors ?? [];
    const hasPosition = Number.isFinite(comment.x) && Number.isFinite(comment.y) &&
      (anchors.length === 0 || anchors.some((anchor) => anchor.type === 'slide'));
    if (!anchored && !hasPosition) return false;
    const x = anchored
      ? anchored.x + (hasPosition ? comment.x as number : anchored.width)
      : comment.x as number;
    const y = anchored
      ? anchored.y + (hasPosition ? comment.y as number : 0)
      : comment.y as number;
    const targetX = x / EMU_PER_PX * this.hooks.scale();
    const targetY = y / EMU_PER_PX * this.hooks.scale();
    this.hooks.scrollPoint(slide, targetX, targetY, options);
    return true;
  }

  async goToComment(slide: number, commentIndex: number, options?: Options): Promise<boolean> {
    if (this.hooks.destroyed()) throw new Error('PptxScrollViewer is destroyed');
    const presentation = this.hooks.presentation();
    if (!presentation || !Number.isInteger(slide) || !Number.isInteger(commentIndex) ||
      slide < 0 || slide >= presentation.slideCount || commentIndex < 0) return false;
    const generation = this.begin();
    if (slide >= presentation.availableSlideCount && !presentation.layoutComplete) {
      if (!await this.waitForSlideMetadata(presentation, slide, generation)) return false;
    }
    if (this.hooks.destroyed()) throw new Error('PptxScrollViewer is destroyed');
    if (generation !== this.generation || presentation !== this.hooks.presentation()) return false;
    if (this.hooks.reachable && !this.hooks.reachable(slide)) return false;
    const comment = presentation.getComments(slide)[commentIndex];
    if (!comment) return false;
    const bounds = await this.resolveBounds(slide, comment);
    if (this.hooks.destroyed()) throw new Error('PptxScrollViewer is destroyed');
    if (generation !== this.generation || presentation !== this.hooks.presentation()) return false;
    if (this.hooks.reachable && !this.hooks.reachable(slide)) return false;
    const anchors = comment.anchors ?? [];
    const hasSlidePoint = Number.isFinite(comment.x) && Number.isFinite(comment.y) &&
      (anchors.length === 0 || anchors.some((anchor) => anchor.type === 'slide'));
    if (!bounds && !hasSlidePoint) return false;
    this.hooks.scrollToSlide(slide, options);
    if (!this.scrollToTarget(slide, comment, options, bounds)) return false;
    this.hooks.select(pptxCommentOccurrenceKey(comment, commentIndex, slide), slide);
    return true;
  }

  destroy(): void { this.begin(); }
}
