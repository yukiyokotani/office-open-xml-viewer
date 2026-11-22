import type { LayoutSeg } from './model.js';

interface PrependedSegment {
  readonly segment: LayoutSeg;
  readonly next: PrependedSegment | null;
  readonly length: number;
}

/** A pass-local position in the prepared array plus persistent split/requeue
 * prefixes. Only queue membership is immutable; segment measurement caches
 * retain the same shared identity as the former shallow array snapshots. */
export interface SegmentQueueCursor {
  readonly index: number;
  readonly front: PrependedSegment | null;
}

/** Gap rollback must never copy or replay the remaining paragraph. The
 * prepared array stays fixed, consumption advances its index, and prepends
 * share immutable nodes. Taking/restoring a cursor and shift/unshift are O(1);
 * lookahead visits only the requested suffix, without materializing it.
 * Old prefixes are collectible once neither the live cursor nor a gap
 * snapshot retains them. No history is kept beyond the current transaction. */
export class SegmentQueue implements Iterable<LayoutSeg> {
  private cursor: SegmentQueueCursor = { index: 0, front: null };

  constructor(private readonly segments: readonly LayoutSeg[] = []) {}

  get length(): number {
    return this.segments.length - this.cursor.index + (this.cursor.front?.length ?? 0);
  }

  peek(): LayoutSeg | undefined {
    return this.cursor.front?.segment ?? this.segments[this.cursor.index];
  }

  /** The first two queued segments, without iterating further. */
  peekPair(): readonly [LayoutSeg | undefined, LayoutSeg | undefined] {
    const { index, front } = this.cursor;
    if (!front) return [this.segments[index], this.segments[index + 1]];
    return [front.segment, front.next?.segment ?? this.segments[index]];
  }

  shift(): LayoutSeg | undefined {
    const { index, front } = this.cursor;
    if (front) {
      this.cursor = { index, front: front.next };
      return front.segment;
    }
    if (index >= this.segments.length) return undefined;
    this.cursor = { index: index + 1, front: null };
    return this.segments[index];
  }

  unshift(segment: LayoutSeg): void {
    this.cursor = this.snapshot(segment);
  }

  /** The iterator may already hold the first item when a flush opens a gap.
   * Include it in the rollback cursor without changing the live queue. */
  snapshot(inHand?: LayoutSeg): SegmentQueueCursor {
    if (!inHand) return this.cursor;
    return { index: this.cursor.index, front: {
      segment: inHand, next: this.cursor.front, length: (this.cursor.front?.length ?? 0) + 1,
    } };
  }

  restore(cursor: SegmentQueueCursor): void {
    this.cursor = cursor;
  }

  *[Symbol.iterator](): IterableIterator<LayoutSeg> {
    const { index, front } = this.cursor;
    for (let node = front; node; node = node.next) yield node.segment;
    for (let offset = index; offset < this.segments.length; offset += 1) yield this.segments[offset];
  }
}
