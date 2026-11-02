import type { FootnoteCursor } from './footnote-fragmentation.js';
import type { TableFragmentCursor } from './table-pagination.js';
import type { NoteLayout, PaintNode, SourceRef } from './types.js';

type Frontier = readonly [segment: number, offset: number];
const after = (a: Frontier, b: Frontier) => a[0] > b[0] || (a[0] === b[0] && a[1] > b[1]);

function compareTable(start: TableFragmentCursor | undefined, end: TableFragmentCursor | undefined): number {
  const row = (end?.rowIndex ?? 0) - (start?.rowIndex ?? 0);
  if (row !== 0) return row;
  // Fragment ordinals describe geometry, not consumed source. Only row/cell
  // block/line positions (recursively for nested tables) earn progress. Every
  // parallel cell must be monotone; alternating forward/backward cells must
  // not manufacture progress forever.
  let advanced = 0;
  for (let index = 0; index < Math.max(start?.cells.length ?? 0, end?.cells.length ?? 0); index++) {
    const previous = start?.cells[index], cell = end?.cells[index];
    const delta = (cell?.blockIndex ?? 0) - (previous?.blockIndex ?? 0)
      || (cell?.paragraphLineStart ?? 0) - (previous?.paragraphLineStart ?? 0)
      || compareTable(previous?.nestedCursor ?? undefined, cell?.nestedCursor ?? undefined);
    if (delta < 0) return -1;
    if (delta > 0) advanced = 1;
  }
  return advanced;
}

/** A successful partition must consume source, independently of line wrapping,
 * field display lengths and fragment/page ordinals. Completion consumes the
 * remaining suffix. Synthetic paragraph acquisitions can use same-width line
 * cursors; production reflow uses the canonical segment/character boundary. */
export function noteSourceAdvanced(start: FootnoteCursor | null, end: FootnoteCursor | null): boolean {
  if (!end) return true;
  const block = end.sourceContinuation?.blockIndex ?? end.blockIndex;
  const previousBlock = start?.sourceContinuation?.blockIndex ?? start?.blockIndex ?? 0;
  if (block !== previousBlock) return block > previousBlock;
  if (end.tableCursor) return compareTable(start?.tableCursor, end.tableCursor) > 0;
  const boundary = end.sourceContinuation?.paragraph.boundary;
  const previous = start?.sourceContinuation?.paragraph.boundary;
  if (boundary) return after([boundary.segIndex, boundary.charOffset], [previous?.segIndex ?? 0, previous?.charOffset ?? 0]);
  return end.lineIndex > (start?.lineIndex ?? 0);
}

/** Library recovery policy, independent of geometry: a recovered page must
 * commit new source content. Trial reserves, repeated headers, separators and
 * translated copies of already consumed content cannot justify another page.
 * Keep a sealed prefix plus the live page's journal so replay rolls back only
 * that page, without copying or scanning the completed document per opening.
 */
export class NoteRecoveryProgress {
  private readonly prefix = new Map<string, Frontier>();
  private readonly live = new Map<string, Frontier>();
  private failure: Error | undefined;
  private bandFailure: Error | undefined;
  private revision = 0;
  private bandStart = 0;
  private noteCommitted = false;
  private workSource: SourceRef | undefined;

  constructor(private readonly recoveryWork?: Readonly<{
    allows(source: SourceRef): boolean;
    charge(source: SourceRef, units: number): void;
  }>) {}

  /** Page-sequence copies visit the completed prefix too. Attribute them to
   * the note retained on this live page, before copying, so a long tail cannot
   * perform quadratic prefix work outside its note-local recovery allowance.
   * Ordinary notes have no recovery ledger and keep their existing behavior. */
  chargePageSequenceCopy(pageCount: number): void {
    if (this.workSource) this.recoveryWork?.charge(this.workSource, pageCount);
  }

  authorize(failure: Error | undefined): void {
    this.failure ??= failure;
    this.bandFailure ??= failure;
  }

  private commit(source: SourceRef, frontier: Frontier): void {
    const key = JSON.stringify([source.story, source.storyInstance, source.path]);
    const previous = this.live.get(key) ?? this.prefix.get(key);
    if (!previous || after(frontier, previous)) {
      this.live.set(key, frontier);
      this.revision++;
    }
  }

  retain(block: PaintNode): void {
    if (block.kind === 'paragraph') {
      const boundary = block.noteLineEndBoundaries?.at(-1);
      this.commit(block.source, boundary ? [boundary.segIndex, boundary.charOffset]
        : [0, block.lines.at(-1)?.range.end ?? 0]);
    } else if (block.kind === 'table') {
      for (const row of block.rows) {
        if (('ownership' in row && row.ownership === 'repeated-header') || row.repeatedHeader) continue;
        // Empty source rows/cells still consume an authored unit.
        if (row.cells.length === 0) this.commit(row.source, [0, 0]);
        for (const cell of row.cells) {
          if (cell.blocks.length === 0) this.commit(cell.source, [0, 0]);
          for (const child of cell.blocks) this.retain(child.layout);
        }
      }
    }
  }

  /** Called only at commit, after the acquisition/partition seam verified the
   * note's source cursor. A plan, separator or repeated decoration earns none. */
  retainNote(note: NoteLayout): void {
    if (note.story.blocks.length === 0) return;
    // A later ordinary note cannot erase this page's recovery work owner.
    if (this.recoveryWork?.allows(note.source)) this.workSource = note.source;
    this.noteCommitted = true;
    this.revision++;
  }

  assertProgress(): void {
    if (this.failure && this.live.size === 0 && !this.noteCommitted) throw this.failure;
  }

  openBand(): void {
    if (this.bandFailure && this.revision === this.bandStart) throw this.bandFailure;
    this.bandStart = this.revision;
    this.bandFailure = undefined;
  }

  closePage(): void {
    this.assertProgress();
    for (const [key, frontier] of this.live) this.prefix.set(key, frontier);
    this.resetPage();
  }

  resetPage(): void {
    this.live.clear();
    this.failure = this.bandFailure = undefined;
    this.revision = this.bandStart = 0;
    this.noteCommitted = false;
    this.workSource = undefined;
  }

  fail(error: unknown): never { throw this.failure ?? error; }
}
