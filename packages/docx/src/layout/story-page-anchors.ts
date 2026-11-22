import type { DrawingLayout, PaintNode, ParagraphLayout } from './types.js';

/** A DOCX text-box story's VML shape placed from the page origin
 * (WORD_STORY_VML_PAGE_POSITION). Other producers' page-owned drawings keep
 * their text box's transform and clip. */
export function isStoryPagePositionedDrawing(drawing: DrawingLayout): boolean {
  return drawing.anchorLayer?.storyPagePosition === true;
}

/** A text-box story's page-positioned VML drawings with their owning
 * paragraphs, table cells included, in stacking order. Paint and the element
 * index place them in the box's host frame (paint/canvas-text.ts keeps its
 * own copy: paint imports only layout types). */
export function storyPagePositionedAnchors(
  blocks: readonly PaintNode[],
): Array<Readonly<{ owner: ParagraphLayout; drawing: DrawingLayout }>> {
  const found: Array<Readonly<{ owner: ParagraphLayout; drawing: DrawingLayout }>> = [];
  const visit = (block: PaintNode): void => {
    if (block.kind === 'paragraph') {
      for (const drawing of block.drawings) {
        if (isStoryPagePositionedDrawing(drawing)) found.push({ owner: block, drawing });
      }
    } else if (block.kind === 'table') {
      for (const row of block.rows) for (const cell of row.cells) for (const child of cell.blocks) visit(child.layout);
    }
  };
  blocks.forEach(visit);
  return found.sort((a, b) => a.drawing.anchorLayer!.relativeHeight - b.drawing.anchorLayer!.relativeHeight
    || a.drawing.anchorLayer!.sourceOrder - b.drawing.anchorLayer!.sourceOrder);
}
