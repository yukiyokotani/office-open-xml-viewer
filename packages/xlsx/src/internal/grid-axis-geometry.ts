/** Sparse cumulative geometry for one worksheet row/column axis. */
export class GridAxisGeometry {
  private readonly indices: number[];
  private readonly cumulativeDelta: number[];
  private readonly customPx: number[];
  private readonly defaultPx: number;
  private readonly spans?: Readonly<{ starts: number[]; pixels: number[]; offsets: number[] }>;

  constructor(
    customs: Record<number, number>,
    defaultPx: number,
    toPx: (raw: number) => number,
    private readonly maxIndex: number,
    sortedPixels?: readonly { index: number; px: number }[],
    ranges?: readonly { first: number; last: number; px: number }[],
  ) {
    this.defaultPx = Number.isFinite(defaultPx) && defaultPx >= 0 ? defaultPx : 0;
    this.indices = sortedPixels
      ? sortedPixels.map((entry) => entry.index)
      : Object.keys(customs)
          .map(Number)
          .filter((value) => value >= 1 && value <= maxIndex)
          .sort((a, b) => a - b);
    this.cumulativeDelta = new Array(this.indices.length);
    this.customPx = new Array(this.indices.length);
    let accumulated = 0;
    for (let index = 0; index < this.indices.length; index++) {
      const measured = sortedPixels?.[index]?.px ?? toPx(customs[this.indices[index]]);
      const px = Number.isFinite(measured) && measured >= 0 ? measured : this.defaultPx;
      this.customPx[index] = px;
      accumulated += px - this.defaultPx;
      this.cumulativeDelta[index] = accumulated;
    }
    if (ranges?.length) {
      // Compact view-only ranges use prefix sums over runs. Point overrides
      // win; the row caller removes positive authored points covered by a resize
      // but retains zero points so outline collapse cannot reveal hidden rows.
      const starts = [...new Set([1, maxIndex + 1,
        ...ranges.flatMap(r => [r.first, r.last + 1]),
        ...this.indices.flatMap(i => [i, i + 1])])].sort((a, b) => a - b);
      const points = new Map(this.indices.map((index, i) => [index, this.customPx[i]]));
      const pixels: number[] = [], offsets: number[] = [];
      let rangeIndex = 0, offset = 0;
      for (let i = 0; i < starts.length; i++) {
        const first = starts[i];
        while (rangeIndex < ranges.length && ranges[rangeIndex].last < first) rangeIndex++;
        const range = ranges[rangeIndex];
        const px = points.get(first) ?? (range && range.first <= first ? range.px : this.defaultPx);
        offsets.push(offset); pixels.push(px);
        if (i + 1 < starts.length) offset += (starts[i + 1] - first) * px;
      }
      this.spans = { starts, pixels, offsets };
    }
  }

  private deltaBefore(index: number): number {
    let low = 0;
    let high = this.indices.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (this.indices[middle] < index) low = middle + 1;
      else high = middle;
    }
    return low === 0 ? 0 : this.cumulativeDelta[low - 1];
  }

  offsetOf(index: number): number {
    if (this.spans) {
      let low = 0, high = this.spans.starts.length;
      while (low < high) {
        const middle = (low + high) >>> 1;
        if (this.spans.starts[middle] <= index) low = middle + 1;
        else high = middle;
      }
      const i = Math.max(0, low - 1);
      return this.spans.offsets[i] + (index - this.spans.starts[i]) * this.spans.pixels[i];
    }
    return (index - 1) * this.defaultPx + this.deltaBefore(index);
  }

  indexAt(offset: number): { index: number; partial: number } {
    if (offset < 0) return { index: 1, partial: 0 };
    let low = 1;
    let high = this.maxIndex;
    while (low < high) {
      const middle = (low + high + 1) >> 1;
      if (this.offsetOf(middle) <= offset) low = middle;
      else high = middle - 1;
    }
    return { index: low, partial: offset - this.offsetOf(low) };
  }

  scrollableIndexAt(content: number, firstScrollable: number): number | null {
    const absoluteOffset = content + this.offsetOf(firstScrollable);
    if (absoluteOffset >= this.offsetOf(this.maxIndex) + this.sizeOf(this.maxIndex)) return null;
    return this.indexAt(absoluteOffset).index;
  }

  sizeOf(index: number): number {
    return this.offsetOf(index + 1) - this.offsetOf(index);
  }

  scaled(scale: number): GridAxisGeometry {
    if (this.spans) {
      return new GridAxisGeometry({}, Math.round(this.defaultPx * scale), v => v, this.maxIndex, [],
        this.spans.starts.slice(0, -1).map((first, i) => ({ first,
          last: this.spans!.starts[i + 1] - 1, px: Math.round(this.spans!.pixels[i] * scale) })));
    }
    return new GridAxisGeometry(
      {},
      Math.round(this.defaultPx * scale),
      (value) => value,
      this.maxIndex,
      this.indices.map((index, position) => ({
        index,
        px: Math.round(this.customPx[position] * scale),
      })),
    );
  }

  /** Positive runs for resize capture. Work scales with sparse points/runs,
   * never the million-row ordinal range; differing positive sizes coalesce. */
  positiveRanges(): ReadonlyArray<{ first: number; last: number }> {
    const result: Array<{ first: number; last: number }> = [];
    const add = (first: number, last: number, px: number) => {
      if (px <= 0 || first > last) return;
      const previous = result.at(-1);
      if (previous && previous.last + 1 === first) previous.last = last;
      else result.push({ first, last });
    };
    if (this.spans) {
      for (let i = 0; i + 1 < this.spans.starts.length; i++) {
        add(this.spans.starts[i], this.spans.starts[i + 1] - 1, this.spans.pixels[i]);
      }
    } else {
      let start = 1;
      for (let i = 0; i < this.indices.length; i++) {
        add(start, this.indices[i] - 1, this.defaultPx);
        add(this.indices[i], this.indices[i], this.customPx[i]);
        start = this.indices[i] + 1;
      }
      add(start, this.maxIndex, this.defaultPx);
    }
    return result;
  }

  countToCover(index: number, distance: number): number {
    if (index > this.maxIndex || distance <= 0) return 0;
    const target = this.offsetOf(index) + distance;
    const end = this.offsetOf(this.maxIndex) + this.sizeOf(this.maxIndex);
    if (target >= end) return this.maxIndex - index + 1;
    const located = this.indexAt(target);
    return located.index - index + (located.partial > 0 ? 1 : 0);
  }

  /** Materialize only positive-width bands, jumping across zero-sized runs by
   * cumulative offset. Work is bounded by visible pixels, not sheet ordinals. */
  bandsToCover(
    startIndex: number,
    endIndex: number,
    distance = Number.POSITIVE_INFINITY,
  ): ReadonlyArray<{ index: number; size: number }> {
    const start = Math.max(1, startIndex);
    const end = Math.min(this.maxIndex, endIndex);
    if (start > end || distance <= 0) return [];
    const bands: Array<{ index: number; size: number }> = [];
    let covered = 0;
    let index = Math.max(start, this.indexAt(this.offsetOf(start)).index);
    while (index <= end && covered < distance) {
      const size = this.sizeOf(index);
      if (Number.isFinite(size) && size > 0) {
        bands.push({ index, size });
        covered += size;
      }
      if (index >= end) break;
      const nextOffset = this.offsetOf(index + 1);
      const located = this.indexAt(nextOffset).index;
      index = Math.max(index + 1, located);
    }
    return bands;
  }
}
