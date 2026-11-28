/**
 * Painted table extensions for hit testing. A table whose rows grow to their
 * content (renderTable) paints below its authored frame; hit testing extends
 * the frame to what was painted. Only the latest successful display render
 * of a slide counts:
 *
 * - a render records into its own snapshot and publishes it, replacing the
 *   slide's previous one, only after it completes; a hit during the render
 *   keeps the previous snapshot, and a failed render publishes nothing;
 * - text collection and other non-display paints never record;
 * - a slide not rendered (not on screen) is hit tested by its authored frame;
 * - only extensions are kept (tables painted taller than authored), for the
 *   `limit` most recently rendered slides.
 */
export class TableHitExtents {
  private readonly bySlide = new Map<number, ReadonlyMap<number, number>>();

  constructor(private readonly limit: number) {}

  /** Recorder for one display render of `slideIndex`. */
  begin(slideIndex: number): {
    record: (elementIndex: number, heightEmu: number) => void;
    publish: () => void;
  } {
    const snapshot = new Map<number, number>();
    return {
      record: (elementIndex, heightEmu) => { snapshot.set(elementIndex, heightEmu); },
      publish: () => {
        this.bySlide.delete(slideIndex);
        if (snapshot.size === 0) return;
        this.bySlide.set(slideIndex, snapshot);
        for (const oldest of this.bySlide.keys()) {
          if (this.bySlide.size <= this.limit) break;
          this.bySlide.delete(oldest);
        }
      },
    };
  }

  /** Painted heights (EMU) of the slide's extended tables, by element index. */
  get(slideIndex: number): ReadonlyMap<number, number> | undefined {
    return this.bySlide.get(slideIndex);
  }

  get size(): number {
    return this.bySlide.size;
  }

  clear(): void {
    this.bySlide.clear();
  }
}
