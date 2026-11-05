import type { LayoutSeg, LineBoundary, LineLayoutEnvironment } from './model.js';
import type { ParagraphLayoutRun } from '../layout/text.js';

/** Optional native acquisition can decline after a supported engine setter.
 * Continuations of this fixed-band single-source class therefore own UTF-16
 * source positions, not ordinals in whichever representation was selected.
 * This is source ownership only: it performs no measurement or native admission. */
export function ownSourceCursor(
  segments: LayoutSeg[], runs: readonly ParagraphLayoutRun[], environment: LineLayoutEnvironment,
): LayoutSeg[] {
  if (environment.nativeUnitMinimumBandPt === undefined || runs.length !== 1 || runs[0]?.type !== 'text') return segments;
  const text = runs[0].text;
  let cursor = 0;
  for (const segment of segments) {
    if (!('text' in segment) || segment.metricOnly || segment.sourceRunIndex !== 0
      || segment.text.length === 0
      || (segment.sourceTextOffset !== undefined && segment.sourceTextOffset !== cursor)
      || segment.textShapeRequest?.substituteContext?.text !== text
      || segment.textShapeRequest.substituteContext.offset !== cursor
      || text.slice(cursor, cursor + segment.text.length) !== segment.text) return segments;
    cursor += segment.text.length;
  }
  if (cursor !== text.length || segments.length === 0) return segments;
  cursor = 0;
  for (const segment of segments) {
    segment.sourceCursorStart = cursor;
    cursor += ('text' in segment ? segment.text.length : 0);
  }
  return segments;
}

export function advanceLineBoundary(source: LineBoundary, count: number): LineBoundary {
  return { segIndex: source.segIndex, charOffset: source.charOffset + count,
    ...(source.sourceTextStart !== undefined ? { sourceTextStart: source.sourceTextStart } : {}) };
}

export function sourceCursorEnd(segments: readonly LayoutSeg[]): number | undefined {
  if (segments[0]?.sourceCursorStart === undefined) return undefined;
  let cursor = 0;
  for (const segment of segments) {
    if (!('text' in segment) || segment.sourceCursorStart !== cursor || !segment.text.length) {
      throw new Error('Invalid authoritative source cursor partition');
    }
    cursor += segment.text.length;
  }
  return cursor;
}

export function remapSourceBoundary(segments: readonly LayoutSeg[], boundary: LineBoundary): LineBoundary {
  if (boundary.sourceTextStart === undefined) return boundary;
  const end = sourceCursorEnd(segments);
  const absolute = boundary.sourceTextStart + boundary.charOffset;
  if (end === undefined || !Number.isSafeInteger(boundary.sourceTextStart)
    || boundary.sourceTextStart < 0 || !Number.isSafeInteger(boundary.charOffset)
    || boundary.charOffset < 0 || !Number.isSafeInteger(absolute) || absolute > end) {
    throw new Error('Invalid authoritative source resume boundary');
  }
  if (absolute === end) return { ...boundary, segIndex: segments.length, charOffset: 0, sourceTextStart: end };
  for (let index = 0; index < segments.length; index++) {
    const segment = segments[index]!;
    const start = segment.sourceCursorStart!;
    if ('text' in segment && absolute < start + segment.text.length) {
      return { ...boundary, segIndex: index, charOffset: absolute - start, sourceTextStart: start };
    }
  }
  throw new Error('Unowned authoritative source resume boundary');
}
