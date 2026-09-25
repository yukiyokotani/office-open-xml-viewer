/**
 * ECMA-376 §17.13.5 tracked-change author colouring for the markup view.
 *
 * The palette is the evidence-backed compatibility fact
 * (`WORD_TRACK_CHANGE_AUTHOR_PALETTE` / `WORD_TRACK_CHANGE_AUTHOR_COLORS` in
 * paint-compatibility.ts). The author → palette-index policy is deliberately
 * outside that claim, and is defined HERE as first appearance in document run
 * order over the main story: deterministic for one document across rebuilds,
 * variants, and the main/worker split (both walk the same projected body).
 * An author first seen outside the body walk (a header/footer/text-box story)
 * is appended on first lookup — still deterministic within a layout build,
 * because story acquisition order is itself deterministic.
 */

import { WORD_TRACK_CHANGE_AUTHOR_COLORS } from './paint-compatibility.js';
import type {
  ChangeBarLayout,
  DocumentLayout,
  LineLayout,
  ParagraphLayout,
  TableLayout,
} from './types.js';

interface RevisionRunShape {
  readonly type?: string;
  readonly revision?: Readonly<{ kind?: string; author?: string }>;
}

interface RevisionBlockShape {
  readonly type?: string;
  readonly runs?: readonly RevisionRunShape[];
  readonly rows?: readonly Readonly<{
    cells: readonly Readonly<{ content: readonly RevisionBlockShape[] }>[];
  }>[];
}

/** Resolve an author (undefined = authorless bucket) to a stable palette
 * colour. See module doc for the index policy. */
export type RevisionAuthorColorResolver = (author?: string) => string;

export function createRevisionAuthorColorResolver(
  body: readonly RevisionBlockShape[],
): RevisionAuthorColorResolver {
  const indexByAuthor = new Map<string, number>();
  const record = (author: string) => {
    if (!indexByAuthor.has(author)) indexByAuthor.set(author, indexByAuthor.size);
  };
  const walk = (elements: readonly RevisionBlockShape[]) => {
    for (const element of elements) {
      if (element.type === 'paragraph' && element.runs) {
        for (const run of element.runs) {
          if (run.revision?.kind) record(run.revision.author ?? '');
        }
      } else if (element.type === 'table' && element.rows) {
        for (const row of element.rows) {
          for (const cell of row.cells) walk(cell.content);
        }
      }
    }
  };
  walk(body);
  return (author) => {
    const key = author ?? '';
    record(key);
    const index = indexByAuthor.get(key) ?? 0;
    return WORD_TRACK_CHANGE_AUTHOR_COLORS[index % WORD_TRACK_CHANGE_AUTHOR_COLORS.length]!;
  };
}

/** ECMA-376 defines no bar geometry; the hairline weight is the fixed
 * convention claimed by the `word-track-change-bar` compatibility rule. */
const CHANGE_BAR_WIDTH_PT = 0.75;

function lineHasRevisionText(line: LineLayout): boolean {
  return line.placements.some(
    (placement) => placement.kind === 'text' && placement.revision !== undefined,
  );
}

function sliceRevisionLines(slice: ParagraphLayout | TableLayout): readonly LineLayout[] {
  if (slice.kind === 'paragraph') return slice.lines.filter(lineHasRevisionText);
  return slice.rows.flatMap((row) => row.cells.flatMap((cell) =>
    cell.blocks.flatMap((block) => sliceRevisionLines(block.layout))));
}

/**
 * Markup-view post-pass (`word-track-change-bar`): attach one margin bar per
 * line that retains revision content, centered in the left page margin. Pure
 * geometry translation over the completed body layers — no measurement, no
 * repagination — mirroring the §17.6.8 line-number composition pass. Pages
 * without revision lines are returned unchanged, and the default final-view
 * variant never runs this (its layouts carry no `changeBars` at all).
 */
export function attachTrackChangeBars(layout: DocumentLayout): DocumentLayout {
  let attached = false;
  const pages = layout.pages.map((page) => {
    const lines = page.layers.body.flatMap((node) => (
      node.kind === 'paragraph' || node.kind === 'table' ? sliceRevisionLines(node) : []
    ));
    if (lines.length === 0) return page;
    attached = true;
    const xPt = Math.max(0, page.section.geometry.marginLeft / 2 - CHANGE_BAR_WIDTH_PT / 2);
    const changeBars: readonly ChangeBarLayout[] = Object.freeze(lines.map((line) => Object.freeze({
      bounds: Object.freeze({
        xPt,
        yPt: line.bounds.yPt,
        widthPt: CHANGE_BAR_WIDTH_PT,
        heightPt: line.bounds.heightPt,
      }),
    })));
    return Object.freeze({ ...page, changeBars });
  });
  return attached ? Object.freeze({ ...layout, pages: Object.freeze(pages) }) : layout;
}
