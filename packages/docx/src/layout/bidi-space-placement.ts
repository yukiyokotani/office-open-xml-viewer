import { resolvedBidiLevels } from '../character-spacing.js';
import { bidiPlacedOffsets } from './bidi-cluster-placement.js';
import { retainedWavePath } from './retained-typography.js';
import { sliceSemanticSlotSpans } from './text.js';
import type { TextClusterLayout, TextDecorationLayout, TextPlacement, TextPaintOp } from './types.js';

function decorationsAt(placement: TextPlacement, xPt: number, widthPt: number): TextPlacement['decorations'] {
  return placement.decorations?.map(decoration => {
    const from = { ...decoration.from, xPt };
    const to = { ...decoration.to, xPt: xPt + widthPt };
    const { path: _oldPath, ...base } = decoration;
    return { ...base, from, to,
      ...(base.style === 'wavy' ? { path: retainedWavePath(from, to, base.widthPt) } : {}),
    } satisfies TextDecorationLayout;
  });
}

function ownedRange(placement: TextPlacement, start: number, end: number) {
  const from = start - placement.range.start;
  return {
    text: placement.text.slice(from, from + end - start),
    range: { start, end },
    ...(placement.sourceRuns ? { sourceRuns: placement.sourceRuns.flatMap(owner => {
      const a = Math.max(start, owner.range.start), b = Math.min(end, owner.range.end);
      return b > a ? [{ ...owner, range: { start: a, end: b } }] : [];
    }) } : {}),
    ...(placement.semanticSlotSpans ? { semanticSlotSpans: sliceSemanticSlotSpans(
      placement.semanticSlotSpans, from, from + end - start,
    ) } : {}),
  };
}

/** A source-addressed visible word. The source trailing whitespace can be
 * elsewhere in the same level block, so it cannot share this hit rectangle.
 * Glyph shaping stays contextual: trim only the already-retained zero-ink
 * U+0020 suffix, never shape the word again. */
export function rotatedWordView(
  template: TextPlacement,
  visibleEnd: number,
  xPt: number,
  widthPt: number,
  logicalClusters: readonly TextClusterLayout[],
): TextPlacement | undefined {
  if (visibleEnd === template.range.start) return undefined;
  const { trailingSpaceCompressionPt: _space, ...base } = template;
  const source = ownedRange(template, template.range.start, visibleEnd);
  const paintOps = template.paintOps.flatMap(operation => {
    if (operation.range.start >= visibleEnd) return [];
    if (operation.range.end <= visibleEnd) return [operation];
    let text = operation.text;
    if (operation.sourceMapping === 'kashida') {
      let end = text.length;
      while (end > 0 && text[end - 1] === ' ') end--;
      text = text.slice(0, end);
    } else text = text.slice(0, visibleEnd - operation.range.start);
    return [{ ...operation, text, range: { ...operation.range, end: visibleEnd } }];
  });
  const logical = logicalClusters.filter(cluster => cluster.range.start < visibleEnd).map(cluster => ({
    ...cluster, range: { ...cluster.range, end: Math.min(cluster.range.end, visibleEnd) },
  }));
  const levels = resolvedBidiLevels(source.text, template.direction);
  const xs = bidiPlacedOffsets(logical.map(cluster => ({
    xPt: cluster.offset.xPt, advancePt: cluster.advancePt,
    level: levels[cluster.range.start - source.range.start] ?? (template.direction === 'rtl' ? 1 : 0),
  })));
  return {
    ...base, ...source,
    origin: template.origin,
    bounds: { ...template.bounds, xPt, widthPt }, advancePt: widthPt,
    clusters: logical.map((cluster, index) => ({ ...cluster, offset: { ...cluster.offset, xPt: xs[index]! } })),
    paintOps,
    decorations: decorationsAt(template, xPt, widthPt + (template.ownedTrailingSlackPt ?? 0)),
    ...(template.highlightBounds ? { highlightBounds: { ...template.highlightBounds, xPt, widthPt } } : {}),
    ...(template.highlightFragments ? { highlightFragments: template.highlightFragments.map(fragment => ({
      ...fragment, rect: { ...fragment.rect, xPt, widthPt },
    })) } : {}),
  };
}

/** Zero-ink whitespace with its original source/style and actual physical
 * slot. Keeping it separate prevents a discontinuous source run from owning
 * another word's hit box, highlight, or underline. */
export function rotatedSpaceView(
  template: TextPlacement,
  start: number,
  xPt: number,
  widthPt: number,
  naturalWidthPt: number,
  side: 'left' | 'right',
): TextPlacement | undefined {
  if (start === template.range.end) return undefined;
  const source = ownedRange(template, start, template.range.end);
  const direction = side === 'left' ? 'rtl' : 'ltr';
  const { ruby: _ruby, emphasis: _emphasis, ownedTrailingSlackPt: _slack,
    trailingSpaceCompressionPt: _oldSpace, ...base } = template;
  let penPt = 0;
  const clusters: TextClusterLayout[] = template.clusters.filter(c => c.range.start >= start).map(cluster => {
    const advancePt = widthPt * (cluster.range.end - cluster.range.start) / source.text.length;
    const offset = direction === 'rtl' ? widthPt - penPt - advancePt : penPt;
    penPt += advancePt;
    return { ...cluster, advancePt, offset: { ...cluster.offset, xPt: offset } };
  });
  const first = template.paintOps[0];
  if (!first) throw new Error('Retained source whitespace is missing its font operation');
  const { inkBounds: _ink, blockAxisInkBounds: _blockInk, naturalAdvancePt: _natural,
    glyphOffsetPt: _glyphOffset, sourceMapping: _mapping, ...zeroInk } = first;
  const paintOps: TextPaintOp[] = [{ ...zeroInk, text: source.text, range: source.range,
    offset: { ...zeroInk.offset, xPt: 0 }, direction }];
  return {
    ...base, ...source, direction,
    origin: { ...template.origin, xPt },
    bounds: { ...template.bounds, xPt, widthPt }, advancePt: widthPt,
    trailingSpaceCompressionPt: template.trailingSpaceCompressionPt ?? naturalWidthPt - widthPt,
    clusters, paintOps,
    decorations: decorationsAt(template, xPt, widthPt),
    ...(template.highlightBounds ? { highlightBounds: { ...template.highlightBounds, xPt, widthPt } } : {}),
    ...(template.highlightFragments ? { highlightFragments: template.highlightFragments.map(fragment => ({
      ...fragment, rect: { ...fragment.rect, xPt, widthPt },
    })) } : {}),
  };
}
