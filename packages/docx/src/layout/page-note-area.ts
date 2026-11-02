import type { FootnoteCursor, FootnotePartition } from './footnote-fragmentation.js';
import type { NoteLayout } from './types.js';

/** One source-ordered note queue. A cursor always names the unpainted suffix. */
export interface PageNoteSource {
  readonly id: string;
  readonly cursor: FootnoteCursor | null;
  /** Single-column owners retain the existing reference-line permission for
   * physical-page splitting; a column change alone does not grant it. */
  readonly allowPartial?: boolean;
}

export interface PageNoteChunk {
  readonly columnIndex: number;
  readonly offsetPt: number;
  readonly hostShiftPt: number;
  readonly layout: NoteLayout;
}

export interface PageNoteFlow {
  readonly chunks: readonly PageNoteChunk[];
  readonly heightPt: number;
  readonly remaining: readonly PageNoteSource[];
  readonly progress: Readonly<{ sourceIndex: number; cursor: FootnoteCursor | null }>;
}

export interface PageNoteAreaInput {
  readonly sources: readonly PageNoteSource[];
  readonly columnCount: number;
  /** A common height, including the separator's blank line box in each column. */
  readonly capacityPt: number;
  /** Canonical acquisition and source-cut partitioning; never a paint probe. */
  readonly acquire: (source: PageNoteSource, columnIndex: number, firstInColumn: boolean,
    capacityPt: number, offsetPt: number) => FootnotePartition | null;
}

/** Flow one page's selected note queue through authored columns. The caller
 * owns the section of the first reference; it must not substitute the section
 * active at the end of the page. A separator occupies the same top line box
 * in every column, while only its first occurrence paints ink. */
export function flowPageNoteArea(input: PageNoteAreaInput): PageNoteFlow {
  if (!Number.isSafeInteger(input.columnCount) || input.columnCount < 1
    || !Number.isFinite(input.capacityPt) || input.capacityPt < 0) {
    throw new RangeError('Invalid page note area extent');
  }
  const chunks: PageNoteChunk[] = [];
  let sourceIndex = 0;
  let cursor = input.sources[0]?.cursor ?? null;
  let heightPt = 0;
  for (let columnIndex = 0; columnIndex < input.columnCount && sourceIndex < input.sources.length; columnIndex++) {
    let offsetPt = 0;
    let firstInColumn = true;
    while (sourceIndex < input.sources.length) {
      const source = input.sources[sourceIndex]!;
      const partition = input.acquire({ ...source, cursor }, columnIndex,
        firstInColumn, Math.max(0, input.capacityPt - offsetPt), offsetPt);
      if (!partition) break;
      const note = partition.fragment;
      // A common blank separator line is geometry, not an extra painted
      // separator story. Keep its advance without duplicating its text/rule.
      const layout = columnIndex > 0 && firstInColumn ? Object.freeze({
        ...note, separator: Object.freeze([]),
        ...(note.leading ? { leading: Object.freeze({
          role: note.leading.role, source: note.leading.source,
          flowBounds: note.leading.flowBounds, advancePt: note.leading.advancePt,
        }) } : {}),
      }) : note;
      chunks.push(Object.freeze({ columnIndex, offsetPt, layout, hostShiftPt: partition.hostShiftPt }));
      offsetPt += note.advancePt;
      if (!Number.isFinite(offsetPt)) throw new RangeError('Invalid page note area advance');
      heightPt = Math.max(heightPt, offsetPt);
      firstInColumn = false;
      cursor = partition.nextCursor;
      if (cursor) break;
      sourceIndex++;
      cursor = input.sources[sourceIndex]?.cursor ?? null;
    }
  }
  const remaining = input.sources.slice(sourceIndex).map((source, index) => index === 0
    ? Object.freeze({ ...source, cursor }) : source);
  return Object.freeze({ chunks: Object.freeze(chunks), heightPt,
    remaining: Object.freeze(remaining), progress: Object.freeze({ sourceIndex, cursor }) });
}

/** Source ordering is independent of the widths that shaped each epoch. */
function compareProgress(left: PageNoteFlow['progress'], right: PageNoteFlow['progress']): number {
  if (left.sourceIndex !== right.sourceIndex) return left.sourceIndex - right.sourceIndex;
  const a = left.cursor, b = right.cursor;
  if (a === null || b === null) return a === b ? 0 : a === null ? -1 : 1;
  const x = a.sourceContinuation, y = b.sourceContinuation;
  if (x && y) {
    const xb = x.paragraph.boundary, yb = y.paragraph.boundary;
    return x.blockIndex - y.blockIndex || (xb?.segIndex ?? 0) - (yb?.segIndex ?? 0)
      || (xb?.charOffset ?? 0) - (yb?.charOffset ?? 0);
  }
  // Synthetic same-width acquisitions can omit source boundaries. Production
  // width-changing continuation requires them before reaching this helper.
  return a.blockIndex - b.blockIndex || a.lineIndex - b.lineIndex;
}

/** Balance the largest source prefix admitted by capacity, using exactly the
 * same acquisition/partition algorithm as final placement. Search accepts a
 * measured frontier, never a fitted point threshold. The evaluation ceiling is
 * resource governance; exhaustion rejects layout instead of returning a guess.
 * This helper retains only the best plan and the current trial, not all suffix
 * acquisitions or all candidate layouts. */
export function balancePageNoteArea(input: PageNoteAreaInput): PageNoteFlow {
  const target = flowPageNoteArea(input);
  if (target.chunks.length === 0 || target.heightPt === 0) return target;
  let best = target;
  let lowerPt = 0;
  let upperPt = target.heightPt;
  for (let evaluation = 0; evaluation < 128; evaluation++) {
    const middlePt = lowerPt + (upperPt - lowerPt) / 2;
    if (middlePt === lowerPt || middlePt === upperPt) return best;
    const trial = flowPageNoteArea({ ...input, capacityPt: middlePt });
    const progress = compareProgress(trial.progress, target.progress);
    if (progress > 0) throw new Error('Page note area source frontier is not monotone');
    if (trial.heightPt <= middlePt && progress === 0) {
      best = trial;
      upperPt = trial.heightPt;
    } else lowerPt = middlePt;
  }
  throw new Error('Page note area balancing exceeded its operational evaluation limit');
}
