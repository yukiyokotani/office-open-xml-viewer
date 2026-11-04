import type { ParagraphLayoutContext, StoryContext } from '../layout-context.js';
import {
  enterTableCellStoryContext,
  resolveParagraphLayoutContext,
} from '../layout-context.js';
import { getDefaultFontSize } from '../line-layout.js';
import type { ParagraphLayoutSource } from './text.js';
import type { BodyAcquisitionState, RetainedTableRecord } from './acquisition-context.js';
import type { AnchorReferenceFramesInput } from './anchor-frame.js';
import { applyNumberingBodyOffset } from './numbering-marker.js';

export const BODY_STORY_CONTEXT: StoryContext = Object.freeze({
  story: 'body',
  containers: Object.freeze([]),
  lineNumberingEligible: true,
});

/** The main body's root story position, which carries
 * StoryContext.closingMarkHangAdmitted (as do header and footnote story roots,
 * production-body-layout.ts layoutBodyStory). BODY_STORY_CONTEXT, the
 * fallback for states without a story, never admits. */
export const MAIN_BODY_STORY_CONTEXT: StoryContext = Object.freeze({
  ...BODY_STORY_CONTEXT,
  closingMarkHangAdmitted: true,
});

type ParagraphContextState = Pick<
  BodyAcquisitionState,
  'layoutSettings' | 'sectionLayout' | 'acquisitionInputs'
> & Partial<Pick<BodyAcquisitionState, 'layoutServices' | 'defaultTabPt'>>;

type BodyAnchorFrameState = Pick<
  BodyAcquisitionState,
  | 'pageIndex'
  | 'pageWidth'
  | 'pageH'
  | 'marginLeft'
  | 'marginRight'
  | 'marginTop'
  | 'marginBottom'
  | 'contentX'
  | 'contentW'
>;

/** Projects the body page, margin, and column reference frames without
 * crossing out of retained layout's canonical point coordinate space. */
export function bodyAnchorReferenceFrames(
  state: BodyAnchorFrameState,
): Readonly<Pick<
  AnchorReferenceFramesInput,
  'page' | 'margin' | 'column' | 'pageParity'
>> {
  const blockExtentPt = Math.max(
    0,
    state.pageH - state.marginTop - state.marginBottom,
  );

  return {
    page: {
      xPt: 0,
      yPt: 0,
      widthPt: state.pageWidth,
      heightPt: state.pageH,
    },
    margin: {
      xPt: state.marginLeft,
      yPt: state.marginTop,
      widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
      heightPt: blockExtentPt,
    },
    column: {
      xPt: state.contentX,
      yPt: state.marginTop,
      widthPt: state.contentW,
      heightPt: blockExtentPt,
    },
    pageParity: state.pageIndex % 2 === 0 ? 'odd' : 'even',
  };
}

function applyNumberingContext(
  state: ParagraphContextState,
  paragraph: ParagraphLayoutSource,
  context: ParagraphLayoutContext,
): ParagraphLayoutContext {
  return applyNumberingBodyOffset(context, {
    numbering: paragraph.numbering,
    ...(paragraph.numbering ? {
      markerInput: state.acquisitionInputs.numberingMarkerShapeInput(
        paragraph.numbering,
        getDefaultFontSize(paragraph),
      ),
    } : {}),
    authoredFirstIndentPt: paragraph.indentFirst,
    tabStops: paragraph.tabStops,
    defaultTabPt: state.defaultTabPt,
    compatibilityMode: state.layoutSettings.compat.compatibilityMode,
    service: state.layoutServices?.text,
  });
}

export function resolveBodyParagraphLayoutContext(
  state: ParagraphContextState & Partial<Pick<BodyAcquisitionState, 'storyContext'>>,
  paragraph: ParagraphLayoutSource,
): ParagraphLayoutContext {
  return applyNumberingContext(
    state,
    paragraph,
    resolveParagraphLayoutContext(
      state.layoutSettings,
      state.sectionLayout,
      // A top-level paragraph keeps the main body's hang admission only
      // when the state is the main body root itself (frame paragraphs
      // reaching here stay outside the class by their framePr).
      state.storyContext === MAIN_BODY_STORY_CONTEXT ? MAIN_BODY_STORY_CONTEXT : BODY_STORY_CONTEXT,
      paragraph,
    ),
  );
}

export function resolveStateParagraphLayoutContext(
  state: ParagraphContextState & Partial<Pick<BodyAcquisitionState, 'storyContext'>>,
  paragraph: ParagraphLayoutSource,
): ParagraphLayoutContext {
  return applyNumberingContext(
    state,
    paragraph,
    resolveParagraphLayoutContext(
      state.layoutSettings,
      state.sectionLayout,
      state.storyContext ?? BODY_STORY_CONTEXT,
      paragraph,
    ),
  );
}

export function withTableCellStory(
  state: BodyAcquisitionState,
  cellFlowsHorizontally: boolean,
): BodyAcquisitionState {
  return {
    ...state,
    storyContext: enterTableCellStoryContext(
      state.storyContext ?? BODY_STORY_CONTEXT,
      cellFlowsHorizontally,
    ),
  };
}

export function retainedTableRecord(
  state: BodyAcquisitionState,
  sourceIndex: number,
): RetainedTableRecord {
  const record = state.retainedTablesBySourceIndex?.get(sourceIndex);
  if (!record) throw new Error('Table placement requires retained table acquisition');
  return record;
}
