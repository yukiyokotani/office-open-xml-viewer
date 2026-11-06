import type { LayoutLine, LayoutTextSeg } from './model.js';

/** Prefix and tail pieces can inherit cluster ranges for the unsplit source.
 * Refresh retained geometry through the same measurement authority. Visible
 * spaces-only admission also defers face metadata to this publication boundary. */
export function finalizeRetainedLineShapes(
  lines: readonly LayoutLine[],
  widthPolicy: 'bounded' | 'intrinsic' | 'unwrapped',
  measureText: (segment: LayoutTextSeg, clusterGeometry?: boolean) => TextMetrics,
): void {
  if (widthPolicy === 'intrinsic') return;
  for (const line of lines) {
    for (const segment of line.segments) {
      if (!('text' in segment) || segment.metricOnly || segment.text.length === 0) continue;
      segment.shapedClusters = undefined;
      if (segment.textLayoutService && segment.textShapeRequest) measureText(segment, true);
    }
  }
}
