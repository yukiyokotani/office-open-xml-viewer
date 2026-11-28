import { wordPageLevelAnchorY } from './anchor-compatibility.js';
import { isWrapFloat } from './float-wrap.js';

/** Whether vertical anchor placement uses a page-owned coordinate origin under
 * ECMA-376 §20.4.3.5. */
export function isPageLevelAnchorY(
  relativeFrom: string | null | undefined,
  fromParagraph: boolean,
): boolean {
  return wordPageLevelAnchorY(relativeFrom, fromParagraph);
}

export function isPageLevelWrapFloat(
  run: DeepReadonly<ImageRun | Omit<ChartRun, 'chart'> | ShapeRun>,
): boolean {
  return isWrapFloat(run.wrapMode)
    && isPageLevelAnchorY(
      run.anchorYRelativeFrom ?? null,
      run.anchorYFromPara ?? false,
    );
}
import type { ChartRun, ImageRun, ShapeRun } from '../types.js';
import type { DeepReadonly } from './types.js';

/** The single test for a floating drawing payload (ECMA-376 §20.4.2.3
 * wp:anchor and its VML/legacy/ODF equivalents) in any run shape the layout
 * receives: a parser anchor (anchor acquisition input, public or parser
 * spelling), a public picture or chart with `anchor: true`, a shape or group
 * that is not inline, and an unavailable drawing's anchored placeholder.
 * Layout routes anchors and classifies anchor-only paragraphs through it. */
export function isFloatingPayloadRun(run: Readonly<{
  type: string;
  anchor?: boolean;
  inline?: boolean;
  anchorAcquisitionInput?: unknown;
  __anchorAcquisition?: unknown;
}>): boolean {
  const acquired = run.anchorAcquisitionInput !== undefined || run.__anchorAcquisition !== undefined;
  switch (run.type) {
    case 'image':
    case 'chart':
      return run.anchor === true || acquired;
    case 'shape':
      return run.inline !== true || acquired;
    case 'unavailableDrawing':
      return acquired;
    default:
      return false;
  }
}
