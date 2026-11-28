import { normalizeWrapSide, type FloatRect } from './float-wrap.js';
import type {
  DrawingMLCollisionEntryPt,
  WrapExclusion,
} from './types.js';

/** Project the compatibility float registry into retained text-wrap authority.
 * A DrawingML wrapNone object remains absent here because §20.4.2.3 gives it
 * object-collision semantics without adding text exclusion geometry. */
export function paragraphWrapExclusions(
  floats: readonly FloatRect[],
  flowDomainId: string,
): readonly WrapExclusion[] {
  return floats.flatMap((float, index): WrapExclusion[] => {
    if (
      float.kind === 'shape'
      && float.anchorOccurrenceId
      && float.authoredWrap === undefined
    ) return [];
    return [{
      id: float.imageKey || `${flowDomainId}:float:${index}`,
      wrap: float.authoredWrap
        ?? (float.mode === 'topAndBottom' ? 'topAndBottom' : 'square'),
      wrapSide: normalizeWrapSide(float.side),
      bounds: {
        xPt: float.xLeft,
        yPt: float.yTop,
        widthPt: Math.max(0, float.xRight - float.xLeft),
        heightPt: Math.max(0, float.yBottom - float.yTop),
      },
      polygon: float.wrapPolygon ?? [
        { xPt: float.xLeft, yPt: float.yTop },
        { xPt: float.xRight, yPt: float.yTop },
        { xPt: float.xRight, yPt: float.yBottom },
        { xPt: float.xLeft, yPt: float.yBottom },
      ],
      ...(float.kind === 'table' && !float.anchorOccurrenceId
        ? { verticalOwnership: 'page' as const }
        : {}),
      ...(float.anchorOccurrenceId
        ? {
            anchorOccurrenceId: float.anchorOccurrenceId,
            verticalOwnership: 'page' as const,
          }
        : {}),
      ...(float.topEdgeInclusiveFromYPt === undefined
        ? {} : { topEdgeInclusiveFromYPt: float.topEdgeInclusiveFromYPt }),
      ...(float.exemptLineTopPt === undefined
        ? {} : { anchorLineExemptTopPt: float.exemptLineTopPt }),
      ...(float.bandStartInclusive === true ? { bandStartInclusive: true as const } : {}),
      ...(float.leastBoundedGapPt === undefined ? {} : { leastBoundedGapPt: float.leastBoundedGapPt }),
    }];
  });
}

/** Collision authority is intentionally independent from text-wrap authority:
 * DrawingML wrapNone objects still participate in §20.4.2.3 avoidance. */
export function paragraphAnchorCollisions(
  floats: readonly FloatRect[],
): readonly DrawingMLCollisionEntryPt[] {
  return floats.flatMap((float): DrawingMLCollisionEntryPt[] => {
    if (float.kind !== 'shape' || !float.anchorOccurrenceId) return [];
    return [{
      occurrenceId: float.anchorOccurrenceId,
      bounds: {
        xPt: float.imageX,
        yPt: float.imageY,
        widthPt: float.imageW,
        heightPt: float.imageH,
      },
      horizontalOwnership: 'page',
      verticalOwnership: 'page',
    }];
  });
}
