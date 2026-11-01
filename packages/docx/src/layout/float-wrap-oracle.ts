import {
  prepareFloatWrap,
  computePreparedLineFloatWindow,
  skipPastTopAndBottom,
  type FloatRect,
  type LineFloatReference,
} from './float-wrap.js';

export interface WrapOracle {
  readonly hasExclusions?: boolean;
  lineWindow(input: {
    readonly topYPt: number;
    readonly minimumStartWidthPt: number;
    readonly squareMinimumStartWidthPt?: number;
    readonly probeHeightPt: number;
    readonly paragraphXPt: number;
    readonly maximumWidthPt: number;
    /** The paragraph's COLUMN band, scoping the topAndBottom gate (§20.4.2.20 /
     *  §17.6.4) to the column the float is anchored in — NOT the indented text
     *  band `paragraphXPt`/`maximumWidthPt` the square side-gap math uses. */
    readonly columnXPt: number;
    readonly columnWidthPt: number;
  }): {
    readonly topYPt: number;
    readonly xOffsetPt: number;
    readonly maximumWidthPt: number;
  };
  skipTopAndBottomBands(input: {
    readonly yPt: number;
    /** The paragraph's COLUMN band (colX()/colW()), used to scope a topAndBottom
     *  float to the column it is anchored in (§20.4.2.20 / §17.6.4) — NOT the
     *  indented text band `lineWindow` uses. */
    readonly columnXPt: number;
    readonly columnWidthPt: number;
  }): number;
}

/** Adapt immutable scale-1 exclusion geometry to the paragraph placement
 * boundary once, so each line query shares one compiled float-wrap authority. */
export function createFloatWrapOracle(
  floats: readonly FloatRect[],
  reference?: LineFloatReference,
): WrapOracle {
  const activeFloats = floats.map((float) => Object.freeze({ ...float }));
  const prepared = prepareFloatWrap(activeFloats);
  return {
    hasExclusions: activeFloats.length > 0,
    lineWindow: ({
      topYPt,
      minimumStartWidthPt,
      squareMinimumStartWidthPt,
      probeHeightPt,
      paragraphXPt,
      maximumWidthPt,
      columnXPt,
      columnWidthPt,
    }) => {
      const window = computePreparedLineFloatWindow(
        topYPt,
        minimumStartWidthPt,
        probeHeightPt,
        paragraphXPt,
        maximumWidthPt,
        prepared,
        columnXPt,
        columnXPt + columnWidthPt,
        reference ?? {
          xLeftPt: paragraphXPt,
          xRightPt: paragraphXPt + maximumWidthPt,
          readingDirection: 'ltr',
        },
        squareMinimumStartWidthPt ?? minimumStartWidthPt,
      );
      return {
        topYPt: window.topY,
        xOffsetPt: window.xOffset,
        maximumWidthPt: window.maxWidth,
      };
    },
    skipTopAndBottomBands: ({ yPt, columnXPt, columnWidthPt }) =>
      skipPastTopAndBottom(yPt, activeFloats, columnXPt, columnXPt + columnWidthPt),
  };
}
