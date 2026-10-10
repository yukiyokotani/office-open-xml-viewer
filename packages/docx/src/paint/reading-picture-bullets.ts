import type { DeepReadonly, LayoutPage, PaintNode } from '../layout/types.js';

/** Traverse retained graph identities, including table cells and child stories.
 * Each immutable node is visited once; never rediscover geometry or source. */
export function readingPictureBulletKeys(page: LayoutPage | DeepReadonly<LayoutPage>): readonly string[] {
  const pending: DeepReadonly<PaintNode>[] = page.layers.roots.map(root => root.node);
  for (const entry of page.layers.paintOrder) if (entry.kind === 'drawing') pending.push(...entry.textBoxes);
  const seen = new Set<DeepReadonly<PaintNode>>();
  const keys = new Set<string>();
  while (pending.length) {
    const node = pending.pop()!;
    if (seen.has(node)) continue;
    seen.add(node);
    if (node.kind === 'paragraph') {
      if (node.nativeReadingPictureBullet) for (const line of node.lines) for (const placement of line.placements) {
        if (placement.kind === 'resource' && placement.resourceKind === 'picture-bullet') keys.add(placement.resourceKey);
      }
      pending.push(...node.textBoxes);
    } else if (node.kind === 'table') {
      for (const row of node.rows) for (const cell of row.cells) pending.push(...cell.blocks.map(block => block.layout));
      pending.push(...(node.resolvedFloatingTables?.map(table => table.child) ?? []));
    } else if (node.kind === 'textbox' || node.kind === 'note') {
      pending.push(...node.story.blocks);
      if (node.kind === 'note') {
        if (node.leading?.paragraph) pending.push(node.leading.paragraph);
        if (node.trailing?.paragraph) pending.push(node.trailing.paragraph);
      }
    }
  }
  return Object.freeze([...keys]);
}
