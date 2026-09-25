import type { CellElement } from '../types.js';
import type { FlowFragment } from './flow-fragment.js';
import type { ParagraphLayout } from './types.js';
import { paragraphGapPt } from './paragraph-spacing.js';
import { paragraphFragmentAdvancePt } from './flow-fragment.js';
import {
  resolveParagraphBorderEdges,
  type ParagraphBorderEdges,
} from './paragraph-border-adjacency.js';
import type { ParagraphLayoutSource } from './text.js';
import type { TableLayoutSource } from './table-source-acquisition.js';

type TableCellLayoutSource = TableLayoutSource['rows'][number]['cells'][number];

export type AcquireNestedCellBlocks<State> = (
  cell: TableCellLayoutSource,
  table: TableLayoutSource,
  cellTotalWidthPt: number,
  outerState: State,
  sourcePath: readonly number[],
) => readonly FlowFragment[];

export interface TableCellBlockAcquisitionDependencies<State> {
  resolveContentWidthPt(cell: TableCellLayoutSource, table: TableLayoutSource, totalWidthPt: number): number;
  createCellState(outerState: State, contentWidthPt: number, cell: TableCellLayoutSource): State;
  acquireParagraph(
    state: State,
    paragraph: ParagraphLayoutSource,
    contentWidthPt: number,
    sourcePath: readonly number[],
    paragraphBorderEdges: ParagraphBorderEdges,
  ): ParagraphLayout;
  acquireNestedTable(
    state: State,
    table: TableLayoutSource,
    contentWidthPt: number,
    sourcePath: readonly number[],
    continuation: Readonly<{ fromPrevious: boolean; onNext: boolean }>,
    acquireNestedCellBlocks: AcquireNestedCellBlocks<State>,
  ): FlowFragment;
  advanceState(state: State, advancePt: number): void;
}

export interface AcquireTableCellBlocksInput<State> {
  readonly cell: TableCellLayoutSource;
  readonly table: TableLayoutSource;
  readonly cellTotalWidthPt: number;
  readonly outerState: State;
  readonly sourcePath: readonly number[];
}

export interface RetainedCellBlockPlacement {
  readonly blockPlacements: readonly Readonly<{ offsetPt: number; advancePt: number }>[];
  readonly contentTranslationPt: number;
  readonly inkBlock: Readonly<{ topPt: number; heightPt: number }>;
}

/**
 * Whether a cell's final empty paragraph owns no row height. That holds for
 * the required paragraph after a nested table, and, in a cell with
 * `hideMark` (ECMA-376 §17.4.21: the end-of-cell mark is ignored for the
 * row height), for a final paragraph that holds only that mark, including
 * the sole paragraph of an empty cell. The rule applies to each hideMark
 * cell independently of the other cells in the row.
 */
export function isStructuralTrailingParagraph(
  content: TableCellLayoutSource['content'],
  index: number,
  hideMark = false,
): boolean {
  if (index !== content.length - 1) return false;
  const current = content[index];
  if (current?.type !== 'paragraph') return false;
  if (!hideMark && (index === 0 || content[index - 1]?.type !== 'table')) return false;
  return current.runs.length === 0;
}

/**
 * Resolve the immutable point-space placement of a cell's retained block tree.
 * Paragraph edge spacing participates in top flow, but center/bottom align the
 * ink block between the resolved cell margins (ECMA-376 §17.4.83). Nested
 * tables use the same document-order fold rather than a table-specific branch.
 */
export function resolveRetainedCellBlockPlacement(
  cell: TableCellLayoutSource,
  table: TableLayoutSource,
  blocks: readonly FlowFragment[],
  boxHeightPt: number,
): RetainedCellBlockPlacement {
  const blockPlacements: Array<{ offsetPt: number; advancePt: number }> = [];
  let cursorPt = 0;
  let previousParagraph: ParagraphLayoutSource | null = null;
  let previousAfterPt = 0;
  let firstInkTopPt: number | undefined;
  let lastInkBottomPt = 0;

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]!;
    const element = cell.content[index];
    const structural = isStructuralTrailingParagraph(
      cell.content,
      index,
      cell.hideMark === true,
    );
    if (block.kind === 'paragraph' && element?.type === 'paragraph') {
      const paragraph: ParagraphLayoutSource = element;
      const blockBeforePt = block.spacing?.beforePt ?? 0;
      const blockAfterPt = block.spacing?.afterPt ?? 0;
      const gapPt = previousParagraph
        ? paragraphGapPt(
            previousParagraph,
            paragraph,
            previousAfterPt,
            blockBeforePt,
          )
        : index === 0 || cell.content[index - 1]?.type === 'table'
          ? blockBeforePt
          : 0;
      const lineBlockPt = Math.max(
        0,
        paragraphFragmentAdvancePt(block) - blockBeforePt - blockAfterPt,
      );
      const offsetPt = cursorPt + gapPt;
      blockPlacements.push({ offsetPt, advancePt: lineBlockPt });
      cursorPt = offsetPt + lineBlockPt;
      if (!structural) {
        firstInkTopPt ??= offsetPt;
        lastInkBottomPt = cursorPt;
      }
      previousParagraph = paragraph;
      previousAfterPt = blockAfterPt;
      continue;
    }

    if (previousParagraph) cursorPt += previousAfterPt;
    const advancePt = block.kind === 'table'
      ? block.advancePt
      : paragraphFragmentAdvancePt(block);
    blockPlacements.push({ offsetPt: cursorPt, advancePt });
    firstInkTopPt ??= cursorPt;
    cursorPt += advancePt;
    lastInkBottomPt = cursorPt;
    previousParagraph = null;
    previousAfterPt = 0;
  }

  const inkTopPt = firstInkTopPt ?? 0;
  const inkHeightPt = Math.max(0, lastInkBottomPt - inkTopPt);
  const marginTopPt = cell.marginTop ?? table.cellMarginTop ?? 0;
  const marginBottomPt = cell.marginBottom ?? table.cellMarginBottom ?? 0;
  const contentHeightPt = boxHeightPt - marginTopPt - marginBottomPt;
  const alignedInkTopPt = cell.vAlign === 'center'
    ? marginTopPt + (contentHeightPt - inkHeightPt) / 2
    : cell.vAlign === 'bottom'
      ? boxHeightPt - marginBottomPt - inkHeightPt
      : marginTopPt + inkTopPt;
  const contentTranslationPt = cell.vAlign === 'top'
    ? marginTopPt
    : alignedInkTopPt - inkTopPt;

  return {
    blockPlacements: Object.freeze(blockPlacements.map((placement) => Object.freeze(placement))),
    contentTranslationPt,
    inkBlock: Object.freeze({ topPt: inkTopPt, heightPt: inkHeightPt }),
  };
}

/**
 * Acquire a cell's recursive paragraph/table block tree. Table pagination owns
 * the injected geometry callbacks; this module owns document-order recursion and
 * the invariant that every paragraph becomes one self-contained ParagraphLayout.
 */
export function acquireTableCellBlocks<State>(
  input: AcquireTableCellBlocksInput<State>,
  dependencies: TableCellBlockAcquisitionDependencies<State>,
): readonly FlowFragment[] {
  const { cell, table, cellTotalWidthPt, outerState, sourcePath } = input;
  const contentWidthPt = dependencies.resolveContentWidthPt(cell, table, cellTotalWidthPt);
  const cellState = dependencies.createCellState(outerState, contentWidthPt, cell);
  const blocks: FlowFragment[] = [];

  for (let cellElementIndex = 0; cellElementIndex < cell.content.length; cellElementIndex += 1) {
    const element = cell.content[cellElementIndex];
    if (!element) continue;
    const elementPath = [...sourcePath, cellElementIndex];
    if (element.type === 'paragraph') {
      const previousElement = cell.content[cellElementIndex - 1];
      const nextElement = cell.content[cellElementIndex + 1];
      const paragraph: ParagraphLayoutSource = element;
      const block = dependencies.acquireParagraph(
        cellState,
        paragraph,
        contentWidthPt,
        elementPath,
        resolveParagraphBorderEdges(
          previousElement?.type === 'paragraph'
            ? previousElement : null,
          paragraph,
          nextElement?.type === 'paragraph'
            ? nextElement : null,
        ),
      );
      blocks.push(block);
      dependencies.advanceState(cellState, block.advancePt);
      continue;
    }

    const inner: TableLayoutSource = element;
    const nestedSlice = element as typeof element & {
      nestedSliceContinuesFromPrevious?: boolean;
      nestedSliceContinuesOnNext?: boolean;
    };
    blocks.push(dependencies.acquireNestedTable(
      cellState,
      inner,
      contentWidthPt,
      elementPath,
      {
        fromPrevious: nestedSlice.nestedSliceContinuesFromPrevious ?? false,
        onNext: nestedSlice.nestedSliceContinuesOnNext ?? false,
      },
      (nestedCell, nestedTable, widthPt, nestedOuterState, nestedPath) =>
        acquireTableCellBlocks({
          cell: nestedCell,
          table: nestedTable,
          cellTotalWidthPt: widthPt,
          outerState: nestedOuterState,
          sourcePath: nestedPath,
        }, dependencies),
    ));
  }

  return blocks;
}
