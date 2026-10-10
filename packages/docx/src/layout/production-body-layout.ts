import { acquireRequestedNativeReadingScenes } from './native-reading-paragraph.js';
import { MAX_BODY_LAYOUT_PAGES } from './resource-budgets.js';
import { deepFreezePlainData, deepFreezePlainDataWithFrozenAliases } from './plain-data.js';
import { quarterTurnMathMetadataService } from './resources.js';
import { pageOwnedAnchorKeysByLine } from './anchor-line-deferral.js';
import type { CjkLang } from '@silurus/ooxml-core';
import type {
  BodyElement,
  DocParagraph,
  DocTableCell,
  FramePr,
  ImageRun,
  ChartRun,
  ShapeRun,
  SectionProps,
} from '../types';
import type { ResolvedFontMetric } from '@silurus/ooxml-core';
import { type FloatRect, isWrapFloat } from '../float-layout.js';
import {
  type FrameBox,
  computeFrameBox,
  frameWrapExclusionMode,
  frameXContainer,
  pushFloatRect,
  registerFrameFloat,
} from '../frame-geometry.js';
import { resolveFloatingTableBoxPt } from '../float-table-geometry.js';
import { xContainer, yContainer, resolveAnchorX, resolveAnchorY } from '../anchor-geometry.js';
import { resolveParagraphLayoutContext, resolveSectionLayoutContext, type DocumentLayoutSettings, type SectionLayoutContext } from '../layout-context.js';
import type {
  BlockLayoutAlgorithms,
  BlockLayoutResult,
  DeepReadonly,
  FlowBlockPlacement,
  DrawingMLCollisionRegistrySnapshotPt,
  LayoutServices,
  FloatRegistryEntryPt,
  FloatRegistrySnapshotPt,
  DrawingMLCollisionEntryPt,
  NativeSectionFlow,
  NoteLayout,
  NoteSeparatorLayout,
  ParagraphLayout,
  SourceRef,
  StoryBlockInput,
  StoryLayout,
  TableLayout,
  TableLayoutInput,
  TableColumnLayoutInput,
  TablePreferredWidthConstraint,
} from './types.js';
import {
  floatingTableRegistryDelta,
  validateFloatingTableRegistryDelta,
} from './floating-table-transaction.js';
import { FLOAT_OVERLAP_EPS, floatRectParticipant, resolveBlockFlowAdmission } from './floats.js';
import type { LayoutOptions } from './options.js';
import { createStoryLayoutCache, type StoryLayoutCache } from './story-layout-cache.js';
import {
  createLayoutServicesRuntimeView,
  fieldAcquisitionContextOf,
  footnoteAcquisitionWorkBudgetOf,
  paintResourceRegistryOf,
  paragraphAcquisitionCacheOf,
  verticalGlyphMeasurementServiceOf,
} from './runtime-state.js';
import { attachStoryBlockLayoutAlgorithms, layoutStory as layoutSharedStory } from './stories.js';
import { buildNoteNumberMap, footnoteIdsInRetainedLines, footnoteIdsInRetainedSlice, indexNotes, noteReferenceIdsInDocumentOrder } from './note-reference-ownership.js';
import type {
  BodyAcquisitionLocation,
  BodyLayoutKernel,
  BodyLayoutSession,
  BodyParagraphAcquisitionInput,
} from './body-layout-kernel.js';
import { NoteCapacityExceededError } from './body-layout-kernel.js';
import { FlowCapacityExceededError } from './flow.js';
import { projectBodyOccurrence } from './occurrence-projection.js';
import {
  noteSeparatorOccurrence,
  placeNoteSeparatorOccurrence,
  reservedNoteSeparatorRole,
} from './native-note-separators.js';
import { selectedNoteSeparatorRole } from './selected-note-separators.js';
import { sourceKey } from './source-key.js';
import type {
  NativeNoteSeparatorDefinitionInput,
  SelectedNoteSeparatorDefinitionInput,
} from './body-layout-input.js';
import {
  sectionBodyInsetPt as bodyMarginInsetPt,
  physicalSectionGeometry,
} from './context.js';
import { isAllRotatedVerticalTextDirection, isVerticalTextDirection, physicalLayoutSection, verticalLayoutSection } from './section-orientation.js';
import { gridForParagraphContext, paragraphMeasurementEnvironment } from './measurement-environment.js';
import { createRevisionAuthorColorResolver } from './track-changes.js';
import { BODY_STORY_CONTEXT, bodyAnchorReferenceFrames, retainedTableRecord, resolveBodyParagraphLayoutContext, resolveStateParagraphLayoutContext, withTableCellStory } from './acquisition-state.js';
import { applyNumberingBodyOffset, resolveNumberingMarkerGeometry } from './numbering-marker.js';
import { projectTableColumnLayoutInput, type TableSourceAcquisitionInput } from './table-source-acquisition.js';
import { measureTableIntrinsicWidths, resolveTableColumnWidths } from './table-columns.js';
import { decideLogicalTable, type LogicalTableDecision, type TableMemberDecision } from './table-layout-decision.js';
import {
  bodyTableAcquisitionState,
  measureBodyTableEntry,
  ownerSegmentedFlowExtents,
} from './body-table-measurement.js';
import {
  finishOwnerHostLayout,
  frameBoxAt,
  ownerCarrierAnchorsToPage,
  ownerHostOccupiedBox,
  ownerHostPageAxes,
  ownerHostPageFrameBox,
  ownerSegmentInput,
  tableOwnerSegments,
} from './table-owner-runs.js';
import { solveExactTranslation, storyAnchorPageFrames, type StoryPageFrames } from './story-page-frames.js';
import { ExactConvergenceError } from './convergence.js';
import { LayoutInvariantError } from './diagnostics.js';
import { measureParagraphIntrinsicWidths, measureTableCellIntrinsicWidths } from './intrinsic-width.js';
// ── Line-layout engine (segmentation + line-breaking + measurement) ──────────
// Body acquisition drives the pure root line-layout kernel through this
// one-directional dependency, with mutable acquisition state owned under layout/.
import { buildFont, fontClassesWithPitches, getDefaultFontSize, paragraphMarkLineHeight } from '../line-layout.js';
import type { DocGridCtx } from '../line-layout.js';
import { measureParagraph } from '../paragraph-measure.js';
import {
  acquireRetainedTable,
  type RetainedTableAcquisition,
  retainedTableAcquisitionIsReusableAcrossPages,
} from './table-acquisition.js';
import { combineAdjacentTableLayoutInputs } from './adjacent-table-layout-input.js';
import { layoutTable as layoutRetainedTableInput } from './table.js';
import {
  hasPagePlacedTableContent,
  layoutWholeTableOnPage,
  type PageDependentTableBlockRequest,
} from './table-pagination.js';
import { paragraphGapAdjustment } from './paragraph-spacing.js';
import { bottomBorderExtentPt, resolveParagraphBorderEdges, topBorderExtentPt, type ParagraphBorderEdges } from './paragraph-border-adjacency.js';
import { acquireParagraphResult, acquireRetainedFrameGroup, bodyFrameGroupFor, bodyParagraphBorderEdgesFor, projectPhysicalAnchorResult, retainedFrameMaximumBaselineLoweringPt, type BodyFrameGroup } from './paragraph.js';
import { wordLoweredDropCapAnchorLeadingPt } from './body-pagination-compatibility.js';
import { wordGridFrameCarrierSourceOwnsNoFlow } from './table-compatibility.js';
import type { CompleteTextBoxStoryAcquirer } from './paragraph.js';
import type { AnchorFloatRegistrationState, BodyAcquisitionState, BodyMeasurementContext, CompleteTextBoxStoryOwner, PhysicalAnchorFrame, RetainedTableRecord } from './acquisition-context.js';
import { ownedParagraphAnchorCollisions, inheritedParagraphAuthorityForReacquisition, TRANSIENT_TABLE_FINAL_FRAME_EXCLUSION_PREFIX } from './paragraph-wrap-registry.js';
import { acquireRegisteredParagraph } from './registered-paragraph-acquisition.js';
import { paragraphAnchorCollisions, paragraphWrapExclusions } from './paragraph-float-authority.js';
import { applyDrawingMLCollisionRegistryDelta, createDrawingMLCollisionRegistry, drawingMLCollisionRegistryDelta, validateDrawingMLCollisionRegistryDelta } from './drawingml-collision-registry.js';
import { resolveAnchorFrame } from './anchor-frame.js';
import { isPageLevelWrapFloat } from './anchor-classification.js';
import { physicalToLogicalAnchorBox } from '../vertical-text.js';
import type { MeasurementTextContext } from './measurement-capabilities.js';
import type {
  LayoutFlowBlock,
  LayoutParagraphBlock,
  LayoutSourceStore,
  LayoutStoryBlock,
  LayoutTableBlock,
} from './layout-source-store.js';
import type { ParagraphLayoutSource } from './text.js';
import type { TableLayoutSource } from './table-source-acquisition.js';
import { collectBodyFrameGroups, prepareBodyFrameMetadata } from './frame.js';
import {
  physicalToLogicalMatrix,
  sectionWritingMode,
  transformRect,
  uprightPhysicalExtent,
} from './coordinate-space.js';

export function createProductionBodyLayoutRuntime(
  source: LayoutSourceStore,
  measureContext: MeasurementTextContext | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>>,
  cjkFallback?: CjkLang,
) {
  prepareBodyFrameMetadata(source.blocks.body);
  const model = source.acquisition;
  const bodyAcquisitionInputProjections = model.acquisitionInputs;
  const effectiveTablePositioning = model.effectiveTablePositioning;
  const publicAnchorBridge = model.publicAnchorBridge;
  const documentFontFamilyClasses = fontClassesWithPitches(
    source.fonts.familyClasses,
    source.fonts.familyPitches,
  );
  const concreteBodyKernelContext = Object.freeze({
    source, measureContext, resolvedLocalFonts, documentFontFamilyClasses,
    publicAnchorBridge, effectiveTablePositioning,
  });
  type ConcreteBodyKernelContext = typeof concreteBodyKernelContext;
  const anchoredImageCollisionKey = (
    imagePath: string,
    colorReplaceFrom?: string,
    duotone?: { readonly clr1: string; readonly clr2: string },
  ): string => `${imagePath}${colorReplaceFrom ? `|clr:${colorReplaceFrom}` : ''}`
    + `${duotone ? `|duo:${duotone.clr1}:${duotone.clr2}` : ''}`;
/** Retained default separator band of the shared note story layout. It applies
 * only where no native reserved separator definition owns the note kind. */
const FOOTNOTE_SEPARATOR_GAP_PT = 6;
const BODY_HOST_FLOW_PAGE_TRANSLATION = Object.freeze({ xPt: 0, yPt: 0 });

/** A visible §17.3.1.42 top border owns space above the first line in every
 * paragraph container. Page/cell-start suppression removes authored w:before,
 * never the border's own spacing or outer half-stroke. */
function paragraphContextWithTopBorder<T extends { readonly spaceBeforePt: number }>(
  context: T,
  paragraph: { readonly borders?: DocParagraph['borders'] },
  topEdge: ParagraphBorderEdges['top'],
  suppressSpaceBefore: boolean,
  continuing = false,
): { context: T; suppressSpaceBefore: boolean } {
  const reservePt = continuing ? 0 : topBorderExtentPt(paragraph.borders, topEdge);
  if (reservePt === 0) return { context, suppressSpaceBefore };
  return {
    context: {
      ...context,
      spaceBeforePt: (suppressSpaceBefore ? 0 : context.spaceBeforePt) + reservePt,
    },
    suppressSpaceBefore: false,
  };
}

function buildMeasureState(
  ctx: MeasurementTextContext,
  section: SectionProps,
  fontFamilyClasses: Record<string, string> = {},
  layoutSettings: DocumentLayoutSettings,
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>> = {},
  layoutServices: LayoutServices,
  layoutOptions?: LayoutOptions,
  nativeSectionFlow?: NativeSectionFlow,
): BodyAcquisitionState {
  const sectionLayout = resolveSectionLayoutContext(layoutSettings, section, nativeSectionFlow);
  // Acquisition always uses the document-scoped service owner supplied by the
  // private body kernel, so its text and vertical measurement capabilities have
  // one auditable lineage and fingerprint.
  const effectiveLayoutServices = layoutServices;
  return {
    ctx,
    verticalGlyphMeasurement: verticalGlyphMeasurementServiceOf(effectiveLayoutServices),
    acquisitionInputs: bodyAcquisitionInputProjections,
    // contentX/contentW carry the canonical point-space
    // current text column, and §20.4.3.4 `relativeFrom="column"` anchors
    // resolve against them (xContainer). Seeding 0 previously placed body-level
    // column anchors a full marginLeft left of their retained point-space
    // placement, so floats entered or left the wrap band during pagination
    // (PR #844 review F1; pinned by paginate-column-anchor.test.ts).
    contentX: section.marginLeft,
    contentW: section.pageWidth - section.marginLeft - section.marginRight,
    y: 0,
    pageH: section.pageHeight,
    pageIndex: 0,
    totalPages: fieldAcquisitionContextOf(effectiveLayoutServices).totalPages,
    marginLeft: section.marginLeft,
    marginRight: section.marginRight,
    // §17.6.11: the measure state's marginTop is the BODY-LEVEL body inset (|margin|).
    // Canonical per-section regions no longer read this field directly: the split
    // functions derive the region top from the threaded `tagSectionGeom` closure
    // (`bodyMarginInsetPt(tagSectionGeom().marginTop)`), matching pushTagged's
    // `bodyTopPt()` per-section convention. This body-level value is only the
    // single-section-equivalent fallback (identical when there is one section) and
    // still seeds contentW/pageH below. Never the raw sign. Identity for non-negative.
    marginTop: bodyMarginInsetPt(section.marginTop),
    marginBottom: bodyMarginInsetPt(section.marginBottom),
    pageWidth: section.pageWidth,
    floats: [],
    floatParaSeq: 0,
    layoutSettings,
    sectionLayout,
    storyContext: BODY_STORY_CONTEXT,
    docEastAsian: layoutSettings.documentHasEastAsianText,
    fontFamilyClasses,
    resolvedLocalFonts,
    layoutServices: effectiveLayoutServices,
    retainedTableAcquisition: {
      layoutServices: (state) => state.layoutServices,
      tableFormat: bodyAcquisitionInputProjections.tableFormatInput,
      tableDecision: singleTableDecision,
      resolveColumns: resolveColumnWidths,
      createCellState: (state, contentWidthPt, cell) => ({
        ...withTableCellStory(tableCellOwnerState(state)),
        contentX: 0,
        contentW: contentWidthPt,
        y: 0,
        containerShading: cell.background ?? state.containerShading,
        floats: [],
        floatParaSeq: 0,
        pageAnchorPrescanned: new Set<ParagraphLayoutSource>(),
      }),
      acquireParagraph: (
        cellState,
        paragraph,
        paragraphWidthPt,
        paragraphPath,
        flowDomainId,
        paragraphBorderEdges,
        inheritedAuthority,
        sourceRef,
      ) => {
        const source = sourceRef ?? {
          story: 'body' as const,
          storyInstance: 'body',
          path: [...paragraphPath],
        };
        const publicRuns = paragraph.runs.filter((run, runIndex) =>
          publicAnchorBridge(source, runIndex) !== null);
        if (publicRuns.length > 0) {
          // Hand-built compatibility runs have no parser anchor/host acquisition
          // facts, so their paragraph-top projection stays outside the parser
          // fixed point until the public bridge is removed.
          registerAnchorFloats(
            { ...paragraph, runs: publicRuns },
            cellState,
            cellState.y,
          );
        }
        const context = resolveStateParagraphLayoutContext(cellState, paragraph);
        const topBorder = paragraphContextWithTopBorder(
          context, paragraph, paragraphBorderEdges?.top ?? 'top', true,
        );
        const layout = acquireRegisteredParagraph(
          cellState,
          cellState.acquisitionInputs.paragraphAcquisitionInput(paragraph, source),
          {
            id: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
            source,
            flowDomainId,
            ordinaryFlow: true,
            context: topBorder.context,
            placement: {
              startYPt: cellState.y,
              paragraphXPt: 0,
              availableWidthPt: paragraphWidthPt,
              maximumYPt: cellState.pageH,
              suppressSpaceBefore: topBorder.suppressSpaceBefore,
            },
            measurer: {
              context: cellState.ctx,
              fontFamilyClasses: cellState.fontFamilyClasses,
            },
            environment: paragraphMeasurementEnvironment(cellState),
            exclusions: paragraphWrapExclusions(cellState.floats, flowDomainId),
            anchorCollisions: paragraphAnchorCollisions(cellState.floats),
            anchorCellBounds: {
              xPt: 0,
              yPt: 0,
              widthPt: paragraphWidthPt,
              heightPt: cellState.pageH,
            },
            containerShading: cellState.containerShading,
            ...(paragraphBorderEdges ? { paragraphBorderEdges } : {}),
            trailingExtentPt: Math.max(
              context.spaceAfterPt,
              paragraphBorderEdges?.bottom === 'none'
                ? 0
                : bottomBorderExtentPt(paragraph.borders),
            ),
            continuesFromPrevious: false,
            anchorFrames: bodyAnchorReferenceFrames(cellState),
            // Bound to the cell's own owner (its section, page and, for an
            // upright physical table, that table's physical frame).
            acquireCompleteStory: completeTextBoxStoryAcquirerFor(cellState),
            ...(cellState.cellHostFlowPageTranslationPt
              ? { hostFlowPageTranslationPt: cellState.cellHostFlowPageTranslationPt }
              : {}),
            ...textBoxStoryHostOptions(cellState, cellState.cellHostFlowPageTranslationPt),
            ...(cellState.cellParagraphAnchorReferenceDeltaPt === undefined ? {} : {
              paragraphAnchorReferenceDeltaPt: cellState.cellParagraphAnchorReferenceDeltaPt,
            }),
          },
          inheritedAuthority,
        ).layout;
        if (paragraph.spaceBefore === 0) return layout;
        return Object.freeze({
          ...layout,
          flowBounds: Object.freeze({
            ...layout.flowBounds,
            heightPt: layout.flowBounds.heightPt + paragraph.spaceBefore,
          }),
          advancePt: layout.advancePt + paragraph.spaceBefore,
          spacing: Object.freeze({
            ...layout.spacing,
            beforePt: (layout.spacing?.beforePt ?? 0) + paragraph.spaceBefore,
          }),
        });
      },
      registerFloatingTable: (state, request) => {
        const usesTextX = !request.positioning.horzSpecified
          || (request.positioning.horzAnchor !== 'page'
            && request.positioning.horzAnchor !== 'margin');
        const usesTextY = request.positioning.vertAnchor !== 'page'
          && request.positioning.vertAnchor !== 'margin';
        // Page/margin coordinates are not final until the containing table is
        // paginated. Registering them in this cell-local acquisition state would
        // reserve a different rectangle from the later page-local paint box.
        if (!usesTextX || !usesTextY) return null;
        const pageHeightPt = state.pageH;
        const textFrame = {
          xPt: state.contentX,
          yPt: state.y,
          widthPt: state.contentW,
          heightPt: request.child.advancePt,
        };
        const box = resolveFloatingTableBoxPt(
          request.positioning,
          {
            page: {
              xPt: 0,
              yPt: 0,
              widthPt: state.pageWidth,
              heightPt: pageHeightPt,
            },
            margin: {
              xPt: state.marginLeft,
              yPt: state.marginTop,
              widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
              heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
            },
            text: textFrame,
          },
          request.child.columnWidthsPt.reduce((sum, width) => sum + width, 0),
          request.child.advancePt,
        );
        const registered = pushFloatRect(state, {
          x: box.x,
          y: box.y,
          w: box.w,
          h: box.h,
          dl: request.positioning.leftFromTextPt,
          dr: request.positioning.rightFromTextPt,
          dt: request.positioning.topFromTextPt,
          db: request.positioning.bottomFromTextPt,
          kind: 'table',
          mode: 'square',
          side: 'bothSides',
          imageKey: '',
          paraId: state.floatParaSeq++,
          avoidOverlap: true,
          tableOverlap: request.overlap,
        });
        return Object.freeze({
          xPt: registered.imageX - textFrame.xPt,
          yPt: registered.imageY - textFrame.yPt,
        });
      },
      advanceState: (state, advancePt) => {
        state.y += advancePt;
      },
    },
    retainedTablesBySourceIndex: new Map<number, RetainedTableRecord>(),
    currentDateMs: layoutOptions?.currentDateMs,
    showTrackedChanges: layoutOptions?.showTrackedChanges,
    kinsoku: layoutSettings.kinsoku,
    defaultTabPt: layoutSettings.defaultTabPt,
    // ECMA-376 §17.6.20 + §20.4.3.x (issue #988 ②, Codex review F1): for a
    // vertical (tbRl) section — `section` is the SWAPPED logical geometry — the
    // acquisition must resolve DrawingML anchors against the same physical page
    // retained paint uses (`resolveAnchorBox`/`resolveShapeBox` key their
    // physical branch on `verticalPhys`), otherwise a wrapped shape's exclusion
    // band is reserved at the raw logical rectangle during pagination while the
    // retained paint uses the physical projection — diverging page assignment.
    // `physicalPageWidthPt` is the physical page width in canonical points.
    // The direction flags are getters on the current `sectionLayout`, so they
    // follow each owner section. An all-rotated `btLr` section (including a
    // native BtoT section) keeps horizontal glyph metrics through
    // `verticalAllRotated`; only upright-vertical sections plan vertical
    // glyphs. The physical anchor frame is seeded here from the section this
    // state is built from and rebuilt from each acquisition location's own
    // section (`applyBodyAcquisitionLocationTo`, issue #1000), so a mid-body
    // section's anchors resolve against ITS OWN physical frame.
    get verticalCJK() {
      return isVerticalTextDirection(this.sectionLayout.textDirection);
    },
    get verticalAllRotated() {
      return isVerticalTextDirection(this.sectionLayout.textDirection)
        && isAllRotatedVerticalTextDirection(this.sectionLayout.textDirection);
    },
    verticalPhys: physicalAnchorFrameOf(sectionLayout),
  };
}

/** Physical page frame of one section context for DrawingML anchors and
 * upright tables, derived only from that context: its logical page box is
 * un-swapped with its own frame (the established clockwise mapping, or a
 * native flow's canonical matrix). Horizontal contexts have no frame. */
function physicalAnchorFrameOf(
  section: DeepReadonly<SectionLayoutContext>,
): PhysicalAnchorFrame | undefined {
  if (!isVerticalTextDirection(section.textDirection)) return undefined;
  const phys = physicalSectionGeometry(section.geometry, section.nativeSectionFlow);
  return {
    pageWidth: phys.pageWidth,
    pageHeight: phys.pageHeight,
    marginLeft: phys.marginLeft,
    marginRight: phys.marginRight,
    marginTop: bodyMarginInsetPt(phys.marginTop),
    marginBottom: bodyMarginInsetPt(phys.marginBottom),
    physicalPageWidthPt: phys.pageWidth,
    ...(section.nativeSectionFlow ? { nativeSectionFlow: section.nativeSectionFlow } : {}),
  };
}

function ordinaryAcquisitionInputForAdjacentGroup(
  group: ReturnType<typeof combineAdjacentTableLayoutInputs>,
): TableLayoutInput {
  const noEdges = Object.freeze({
    top: null,
    right: null,
    bottom: null,
    left: null,
    insideH: null,
    insideV: null,
  });
  return Object.freeze({
    kind: 'table',
    id: group.id,
    source: group.source,
    flowDomainId: group.flowDomainId,
    ordinaryFlow: true,
    alignment: group.alignment,
    indentPt: group.indentPt,
    bidiVisual: group.bidiVisual,
    columnWidthsPt: group.columnWidthsPt,
    columnWidthKeys: group.columnWidthKeys,
    borders: noEdges,
    rows: Object.freeze(
      group.rows.map((row) =>
        Object.freeze({
          ...row,
          // §17.4.37 gives every authored member table ownership of its own outer
          // border layer; the union grid carries that folded layer per source row.
          exceptionBorders: row.sourceTableEdges,
        }),
      ),
    ),
  });
}
function sourceElement(sourceStore: LayoutSourceStore, ref: SourceRef): LayoutFlowBlock {
  if (ref.story !== 'body' || ref.storyInstance !== 'body' || ref.path.length !== 1) {
    throw new Error('Body acquisition requires a top-level body source');
  }
  const element = sourceStore.blocks.resolve(ref);
  if (!element || (element.type !== 'paragraph' && element.type !== 'table')) {
    throw new Error(`Body source does not identify a flow block: ${ref.path.join('.')}`);
  }
  return element;
}
function nestedSourceElement(sourceStore: LayoutSourceStore, ref: SourceRef): LayoutFlowBlock {
  return sourceStore.blocks.resolve(ref);
}
  /** Body acquisition stays at the kernel adapter because it resolves renderer
   * state into retained paragraph inputs; layout-owned projections below it
   * receive only immutable structural values. */
function acquireBodyParagraphAtLocation(
  state: BodyAcquisitionState,
  paragraph: LayoutParagraphBlock,
  source: SourceRef,
  location: BodyAcquisitionLocation,
  availableInlineExtentPt: number,
  suppressSpaceBefore: boolean,
  continuation: BodyParagraphAcquisitionInput['continuation'] = Object.freeze({
    boundary: null,
  }),
  retainedAnchorCollisions?: readonly DrawingMLCollisionEntryPt[],
) {
  const edges = bodyParagraphBorderEdgesFor(paragraph) ?? {
    top: 'top' as const,
    bottom: 'bottom' as const,
  };
  const context = resolveBodyParagraphLayoutContext(state, paragraph);
  const topBorder = paragraphContextWithTopBorder(
    context,
    paragraph,
    edges.top,
    suppressSpaceBefore,
    continuation.boundary !== null,
  );
  return acquireParagraphResult(
    paragraph,
    {
      id: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
      source,
      flowDomainId: location.flowDomainId,
      ordinaryFlow: true,
      context: topBorder.context,
      placement: {
        startYPt: state.y,
        paragraphXPt: location.availableBounds.xPt,
        availableWidthPt: availableInlineExtentPt,
        maximumYPt: state.pageH,
        suppressSpaceBefore: topBorder.suppressSpaceBefore,
      },
      measurer: { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses },
      environment: paragraphMeasurementEnvironment(state),
      exclusions: paragraphWrapExclusions(state.floats, location.flowDomainId),
      anchorCollisions: retainedAnchorCollisions ?? paragraphAnchorCollisions(state.floats),
      containerShading: state.containerShading,
      paragraphBorderEdges: edges,
      trailingExtentPt: Math.max(
        context.spaceAfterPt,
        edges.bottom === 'none' ? 0 : bottomBorderExtentPt(paragraph.borders),
      ),
      continuesFromPrevious: continuation.boundary !== null,
      ...(continuation.sourceRangeStart === undefined
        ? {}
        : {
            sourceRangeStart: continuation.sourceRangeStart,
          }),
      anchorFrames: bodyAnchorReferenceFrames(state),
      acquireCompleteStory: completeTextBoxStoryAcquirerFor(state),
      // Body paragraphs are acquired at their page position.
      hostFlowPageTranslationPt: BODY_HOST_FLOW_PAGE_TRANSLATION,
      ...(state.frozenAnchorFrames && state.frozenAnchorFrames.size > 0
        ? { frozenAnchorFrames: state.frozenAnchorFrames }
        : {}),
    },
    continuation.boundary === null
      ? undefined
      : {
          boundary: continuation.boundary,
          ...(continuation.uniformRubyAdvancePt === undefined
            ? {}
            : {
                uniformRubyAdvancePt: continuation.uniformRubyAdvancePt,
              }),
        },
  );
}

function bodyStoryRoot(source: LayoutSourceStore, ref: SourceRef): readonly LayoutStoryBlock[] {
  if (ref.path.length !== 0) {
    throw new Error('Story acquisition requires a story-root source');
  }
  return source.blocks.storyRoot(ref);
}

function bodyStoryElement(source: LayoutSourceStore, sourceRef: SourceRef): LayoutFlowBlock {
  if (sourceRef.path.length === 0 || (sourceRef.path.length - 1) % 3 !== 0) {
    throw new Error('Story block acquisition requires a canonical source path');
  }
  return source.blocks.resolve(sourceRef);
}

interface BodyStoryAcquisitionContext {
  readonly state: BodyAcquisitionState;
  readonly services: LayoutServices;
  /** The session's story layouts under explicit admission
   * (acquireBodyStoryLayout). A reserved separator root keeps one active
   * acquisition context (at most six roots); its page occurrences are
   * projections, never per-page copies. */
  readonly storyLayoutCache: StoryLayoutCache;
  /** Whether an immutable continued-footnote root is a proved
   * page-independent (text-only) class; computed once per root. */
  readonly noteSourceReuse: WeakMap<object, boolean>;
  readonly source: LayoutSourceStore;
  readonly publicAnchorBridge: typeof publicAnchorBridge;
}

/**
 * Host-chain options of a paragraph in a text box story or in a cell of one
 * of its tables (story-page-frames.ts): the page frames the story's
 * page-owned anchor axes keep, reached by the paragraph's translation into
 * the story (none at its root; a cell's position, once its table's
 * pagination states it) plus the story's own flow shift. Without either, its
 * drawings' text box stories have no page frames, so their page-placed
 * content keeps its acquisition.
 */
function textBoxStoryHostOptions(
  state: Pick<BodyAcquisitionState, 'textBoxStoryHostFrames'>,
  inStoryPt: Readonly<{ xPt: number; yPt: number }> | undefined,
): Readonly<{
  hostPageFrames?: StoryPageFrames | null;
  hostFlowPageTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
}> {
  const chain = state.textBoxStoryHostFrames;
  if (chain === undefined) return {};
  if (chain === null || !inStoryPt) return { hostPageFrames: null };
  return {
    hostPageFrames: chain.frames,
    hostFlowPageTranslationPt: Object.freeze({
      xPt: inStoryPt.xPt + chain.flowPt.xPt,
      yPt: inStoryPt.yPt + chain.flowPt.yPt,
    }),
  };
}

/** Acquire one retained story from immutable source and session-owned context. */
function acquireBodyStoryLayout(
  dependencies: BodyStoryAcquisitionContext,
  request: import('./body-layout-kernel.js').StoryLayoutAcquisitionInput,
): StoryLayout {
  const { source, services } = dependencies;
  const root = bodyStoryRoot(source, request.source);
  // A native reserved or DOCX selected separator root is a shared source
  // definition, not a numbered (continued) note: it has no number and no note
  // cursor. Its closed run-free shape has no page-dependent content, so one
  // acquisition per width/section/font context serves every page occurrence.
  const reservedRole = reservedNoteSeparatorRole(request.source)
    ?? selectedNoteSeparatorRole(request.source);
  // Ordinary text-only paragraph notes have no destination-page fields or
  // anchors. Their immutable acquisition is shared across equal-width pages;
  // partitionFootnote projects only the admitted slice into its page domain.
  // Other continued stories are page-local and must not retain a full copy
  // for every destination page in this session cache.
  const continuedNote = reservedRole === undefined
    && services.allowFootnoteContinuation === true
    && request.source.story === 'footnote';
  const workBudget = continuedNote || (reservedRole !== undefined && services.allowFootnoteContinuation === true)
    ? footnoteAcquisitionWorkBudgetOf(services) : undefined;
  if (continuedNote && !workBudget) throw new Error('Footnote acquisition requires a pagination work budget');
  let reusableNote = continuedNote && dependencies.noteSourceReuse.get(root);
  if (continuedNote && reusableNote === undefined) {
    workBudget?.sourceUnits(root);
    reusableNote = root.every(element => element.type === 'paragraph'
      && !element.framePr && element.runs.every(run => run.type === 'text'));
    dependencies.noteSourceReuse.set(root, reusableNote);
  }
  // Both proved page-independent classes hold no table, frame, text box or
  // anchor, i.e. no page-placed content, so neither a band translation nor
  // page frames can change their layout: their key keeps only the facts the
  // layout depends on (section and container extent), and they are never a
  // convergence trial. Every other story keeps its page and placement facts.
  const pageIndependent = reusableNote || reservedRole !== undefined;
  const occurrence = JSON.stringify([request.source, pageIndependent ? null : request.pageIndex]);
  const placement = JSON.stringify({
    section: request.section,
    container: pageIndependent ? { ...request.container, id: null } : request.container,
    band: pageIndependent ? null : request.bandTranslationPt ?? null,
    pageFrames: pageIndependent ? null : request.pageFrames ?? null,
  });
  return dependencies.storyLayoutCache.layout({
    occurrence,
    placement,
    trial: !pageIndependent
      && (request.bandTranslationPt !== undefined || request.pageFrames !== undefined),
    // Replace, never accumulate, a separator root's context; never retain a
    // destination-dependent full continued note once per page.
    retention: reservedRole !== undefined
      ? 'latest'
      : continuedNote && !reusableNote ? 'none' : 'session',
  }, () => {
    // This ledger is shared by all passes/service views of this pagination,
    // rather than this shorter-lived concrete acquisition session. Debit
    // every miss before destination-field resolution, shaping, geometry or
    // resource acquisition; cache hits are free.
    workBudget?.charge(root);
    return layoutBodyStory(dependencies, request, root, reservedRole !== undefined);
  });
}

/** The one story layout algorithm behind {@link acquireBodyStoryLayout}'s
 * cache admission; it never reads or writes a cache itself. */
function layoutBodyStory(
  dependencies: BodyStoryAcquisitionContext,
  request: import('./body-layout-kernel.js').StoryLayoutAcquisitionInput,
  root: readonly LayoutStoryBlock[],
  separatorRoot: boolean,
): StoryLayout {
  const { source, state, services, publicAnchorBridge } = dependencies;
  const noteReferenceNumber =
    !separatorRoot
      && (request.source.story === 'footnote' || request.source.story === 'endnote')
      ? state.noteNumbers?.get(`${request.source.story}:${request.source.storyInstance}`)
      : undefined;
  const fieldContext = fieldAcquisitionContextOf(services);
  const pageFieldContext = fieldContext.resolveDestinationPage?.(request.pageIndex);
  const storyVertical = isVerticalTextDirection(request.section.textDirection);
  const candidate: BodyAcquisitionState = {
    ...state,
    sectionLayout: request.section as SectionLayoutContext,
    pageIndex: request.pageIndex,
    totalPages: fieldContext.totalPages,
    displayPageNumber: pageFieldContext?.displayPageNumber ?? request.pageIndex + 1,
    pageNumberFormat: pageFieldContext?.pageNumberFormat ?? state.pageNumberFormat,
    pageWidth: request.section.geometry.pageWidth,
    pageH:
      request.container.capacity === 'unbounded'
        ? Number.MAX_SAFE_INTEGER
        : request.section.geometry.pageHeight,
    marginLeft: request.section.geometry.marginLeft,
    marginRight: request.section.geometry.marginRight,
    marginTop: bodyMarginInsetPt(request.section.geometry.marginTop),
    marginBottom: bodyMarginInsetPt(request.section.geometry.marginBottom),
    contentX: request.container.bounds.xPt,
    contentW: request.container.bounds.widthPt,
    y: request.container.bounds.yPt,
    floats: [],
    floatParaSeq: 0,
    retainedTablesBySourceIndex: new Map(),
    pageAnchorPrescanned: new Set<ParagraphLayoutSource>(),
    noteReferenceNumber,
    verticalCJK: storyVertical,
    verticalAllRotated:
      storyVertical && isAllRotatedVerticalTextDirection(request.section.textDirection),
    // The story's own section supplies its frame (none when horizontal).
    verticalPhys: physicalAnchorFrameOf(request.section),
    storyContext: {
      story: request.source.story,
      containers: [],
      lineNumberingEligible: false,
    },
  };
  preRegisterPageFloats(root, 0, candidate);
  const storyServices = createLayoutServicesRuntimeView(services, request.container.quarterTurnMath
    ? { math: quarterTurnMathMetadataService(services.math) } : {});
  candidate.layoutServices = storyServices;
  // Content only a page position places (header/footer root hosts, the
  // positioned tables of story tables) resolves against the destination
  // page's bands. With the composer's band translation known before layout,
  // it is placed through that translation below. A story given its page
  // frames in its own coordinates (a text box) places it there: it is story
  // content for every later transform, so the band is the identity.
  const storyFramed = request.pageFrames !== undefined;
  const band = request.bandTranslationPt ?? (storyFramed ? Object.freeze({ xPt: 0, yPt: 0 }) : undefined);
  const textBoxHostFrames = request.container.kind !== 'textbox'
    ? undefined
    : request.pageFrames ? storyAnchorPageFrames(request.pageFrames) : null;
  // Its table cells' paragraphs reach those frames through their cell's
  // position in the story (textBoxStoryHostOptions).
  if (textBoxHostFrames !== undefined) candidate.textBoxStoryHostFrames = textBoxHostFrames;
  // Those bands in this story's page frame: the section's physical page and
  // margin box carried into it by the section's own physical-to-logical
  // transform — the inverse of the transform that paints the story's layer
  // (a vertical section's note), the identity for a horizontal or upright
  // physical story. Margin facts are the ones this story's state holds. A
  // native section flow selects its own (counter-clockwise) frame through
  // sectionWritingMode and the same native physical geometry its anchors use
  // (physicalAnchorFrameOf); the text-direction token alone never decides it.
  const storyPageFrames = request.pageFrames ?? ((): StoryPageFrames => {
    const writingMode = sectionWritingMode(request.section);
    const held = {
      ...request.section.geometry,
      marginLeft: candidate.marginLeft,
      marginRight: candidate.marginRight,
      marginTop: candidate.marginTop,
      marginBottom: candidate.marginBottom,
    };
    const physical = writingMode === 'horizontal-tb'
      ? held
      : physicalSectionGeometry(held, request.section.nativeSectionFlow);
    const toStory = physicalToLogicalMatrix(
      writingMode,
      { widthPt: physical.pageWidth, heightPt: physical.pageHeight },
    );
    return Object.freeze({
      page: transformRect(toStory, { xPt: 0, yPt: 0, widthPt: physical.pageWidth, heightPt: physical.pageHeight }),
      margin: transformRect(toStory, {
        xPt: physical.marginLeft,
        yPt: physical.marginTop,
        widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
        heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
      }),
    });
  })();
  const storyTableAcquisitions = new Map<string, RetainedTableAcquisition>();
  let bandDependent = false;
  const blockInputs: StoryBlockInput[] = root.flatMap((element, index): StoryBlockInput[] => {
    const source: SourceRef = {
      story: request.source.story,
      storyInstance: request.source.storyInstance,
      path: [index],
    };
    if (element.type === 'unsupportedTextBoxBlock') {
      return [
        {
          type: 'unsupportedTextBoxBlock',
          qName: element.qName,
          sourcePath: element.sourcePath,
        },
      ];
    }
    if (element.type === 'paragraph') return [{ kind: 'paragraph', source }];
    if (element.type !== 'table') {
      throw new Error(`Unsupported ${request.source.story} story block: ${element.type}`);
    }
    const dependencies = candidate.retainedTableAcquisition;
    const table: LayoutTableBlock = element;
    const decision = singleTableDecision(table, request.container.bounds.widthPt, candidate);
    const columns = resolveColumnWidths(table, request.container.bounds.widthPt, candidate, decision);
    const acquisition = acquireRetainedTable(
      table,
      columns,
      request.container.bounds.widthPt,
      candidate,
      source,
      dependencies,
      decision.logical,
      // Rows elect carriers only where the owner context promotes them
      // (header/footer roots; tableRowsElectCarriers).
      { kind: 'story-root', story: request.source.story },
    );
    const input = acquisition.input;
    if (ownerCarrierAnchorsToPage(input) || hasPagePlacedTableContent(acquisition)) {
      bandDependent = true;
    }
    // Owner runs (table-owner-runs.ts) of a header/footer root table become
    // one story block per segment, in source row order; host segments are
    // placed by layoutTable below.
    const segments = tableOwnerSegments(input);
    if (!segments) {
      storyTableAcquisitions.set(input.id, acquisition);
      return [input];
    }
    return segments.map((segment) => {
      const projected = ownerSegmentInput(input, segment);
      storyTableAcquisitions.set(projected.id, Object.freeze({ ...acquisition, input: projected }));
      return projected;
    });
  });
  let previousParagraph: LayoutParagraphBlock | null = null;
  // ECMA-376 §17.3.1.11 frames are story-local: adjacent paragraphs with
  // identical framePr form one frame anchored to the next non-frame
  // paragraph of the same story (headers and footers included).
  const storyFrameGroups = collectBodyFrameGroups(root);
  const storyFrameAcquisitions = new Map<
    BodyFrameGroup<LayoutParagraphBlock>,
    Readonly<{
      box: FrameBox;
      members: ReadonlyMap<
        ParagraphLayoutSource,
        ReturnType<typeof acquireRetainedFrameGroup>['members'][number]
      >;
    }>
  >();
  const algorithms: BlockLayoutAlgorithms = {
    layoutParagraph(block, placement) {
      const paragraph = bodyStoryElement(source, block.source);
      if (paragraph.type !== 'paragraph') throw new Error('Story paragraph source kind mismatch');
      const sourceIndex = block.source.path[0]!;
      const frameGroup = paragraph.framePr
        ? (storyFrameGroups.get(paragraph) as BodyFrameGroup<LayoutParagraphBlock> | undefined)
        : undefined;
      // Page stories are laid out at the top of a page-sized container and
      // then translated into their header/footer band, so only frames whose
      // vertical anchor moves with the story text (vAnchor="text") keep
      // their authored geometry. Page/margin-anchored frames and drop caps
      // retain the historical in-flow fallback.
      if (
        frameGroup &&
        frameGroup.framePr.vAnchor === 'text' &&
        frameGroup.framePr.dropCap === 'none'
      ) {
        // Acquire each group once, at its owner (the first member, whose
        // cursor all members share because frames add no story advance),
        // and reuse it for the remaining members: re-acquiring per member
        // would re-fingerprint the whole group each time.
        let acquisition = storyFrameAcquisitions.get(frameGroup);
        if (!acquisition) {
          candidate.y = placement.cursor.yPt;
          candidate.contentX = placement.container.bounds.xPt;
          candidate.contentW = placement.container.bounds.widthPt;
          let acquiredGroup: ReturnType<typeof acquireRetainedFrameGroup> | undefined;
          const box = resolveFrameBox(
            frameGroup.owner,
            frameGroup,
            candidate,
            frameAnchorLineHeightPx(root, frameGroup.owner, candidate),
            {
              onAcquired: (acquired) => {
                acquiredGroup = acquired;
              },
              story: { story: block.source.story, storyInstance: block.source.storyInstance },
              borderEdgesFor: (member) => storyFrameBorderEdges(frameGroup, member),
            },
          );
          if (!acquiredGroup) throw new Error('Story frame acquisition omitted its retained group');
          acquisition = {
            box,
            members: new Map(acquiredGroup.members.map((entry) => [entry.paragraph, entry])),
          };
          storyFrameAcquisitions.set(frameGroup, acquisition);
          registerFrameFloat(box, frameGroup.framePr, candidate);
        }
        const member = acquisition.members.get(paragraph);
        if (!member) throw new Error('Story frame acquisition omitted its retained member');
        // The frame occupies no ordinary story flow; the anchor paragraph
        // that follows starts at the same cursor and wraps around it.
        // Story flow owns every retained root it returns (layoutFlowBlocks
        // invariant); the positioned geometry itself is unchanged.
        return {
          layout: Object.freeze({ ...member.fragment, flowDomainId: placement.container.id }),
          nextCursor: placement.cursor,
        };
      }
      const previousCandidate = sourceIndex > 0 ? root[sourceIndex - 1] : undefined;
      const previous: LayoutParagraphBlock | null =
        previousCandidate?.type === 'paragraph' ? previousCandidate : null;
      const nextCandidate = root[sourceIndex + 1];
      const next: LayoutParagraphBlock | null =
        nextCandidate?.type === 'paragraph' ? nextCandidate : null;
      const previousAfterPt = previousParagraph?.spaceAfter ?? 0;
      const spacing = paragraphGapAdjustment(
        previousParagraph,
        paragraph,
        previousAfterPt,
        paragraph.spaceBefore,
      );
      const startYPt = Math.max(
        placement.container.bounds.yPt,
        placement.cursor.yPt - spacing.overlap,
      );
      candidate.y = startYPt;
      candidate.contentX = placement.container.bounds.xPt;
      candidate.contentW = placement.container.bounds.widthPt;
      const publicRuns = paragraph.runs.filter(
        (run, runIndex) => publicAnchorBridge(block.source, runIndex) !== null,
      );
      if (publicRuns.length > 0) {
        registerAnchorFloats(
          Object.freeze({ ...paragraph, runs: Object.freeze(publicRuns) }),
          candidate,
          candidate.y,
        );
      }
      const context = resolveStateParagraphLayoutContext(candidate, paragraph);
      const borderEdges = resolveParagraphBorderEdges(previous, paragraph, next);
      const topBorder = paragraphContextWithTopBorder(
        context,
        paragraph,
        borderEdges.top,
        spacing.suppressBefore,
      );
      const result = acquireRegisteredParagraph(candidate, paragraph, {
        id: `${block.source.story}:${block.source.storyInstance}:${block.source.path.join('.')}`,
        source: block.source,
        flowDomainId: placement.container.id,
        ordinaryFlow: true,
        context: topBorder.context,
        placement: {
          startYPt,
          paragraphXPt: placement.container.bounds.xPt,
          availableWidthPt: placement.container.bounds.widthPt,
          ...(placement.container.noWrap ? { noWrap: true } : {}),
          maximumYPt: placement.availableBounds.yPt + placement.availableBounds.heightPt,
          suppressSpaceBefore: topBorder.suppressSpaceBefore,
        },
        measurer: {
          context: candidate.ctx,
          fontFamilyClasses: candidate.fontFamilyClasses,
        },
        environment: paragraphMeasurementEnvironment(candidate),
        exclusions: paragraphWrapExclusions(candidate.floats, placement.container.id),
        anchorCollisions: paragraphAnchorCollisions(candidate.floats),
        containerShading: candidate.containerShading,
        paragraphBorderEdges: borderEdges,
        trailingExtentPt: Math.max(
          context.spaceAfterPt,
          borderEdges.bottom === 'none' ? 0 : bottomBorderExtentPt(paragraph.borders),
        ),
        continuesFromPrevious: false,
        anchorFrames: bodyAnchorReferenceFrames(candidate),
        // Nested text boxes in this story take the story's own section,
        // page and frame, not the body's current location.
        acquireCompleteStory: completeTextBoxStoryAcquirerFor(candidate),
        // A page story's host flow receives its band. A text box story's
        // anchor frames are not its page: its drawings' text box stories
        // reach the page frames its box carried into it, which page-owned
        // axes keep and its flow reaches by the box's shift; without them
        // (no page band reaches the story) they get no page frames.
        ...(request.bandTranslationPt ? { hostFlowPageTranslationPt: request.bandTranslationPt } : {}),
        ...textBoxStoryHostOptions(candidate, { xPt: 0, yPt: 0 }),
      });
      previousParagraph = paragraph;
      const nextCursor = {
        xPt: placement.cursor.xPt,
        yPt: startYPt + result.layout.advancePt,
      };
      candidate.y = nextCursor.yPt;
      // Trailing space-after is allocation, not occupied flow (as in body
      // placement). A contextualSpacing fold starts the next paragraph inside
      // the previous space-after, which would otherwise report a false
      // FLOW_OVERLAP. The origin stays at the allocation start so story
      // positioning and anchor extents keep their leading-spacing arithmetic.
      const contentOwned = result.layout.ordinaryFlow
        ? deepFreezePlainDataWithFrozenAliases({
            ...result.layout,
            flowBounds: Object.freeze({
              ...result.layout.flowBounds,
              heightPt: Math.max(
                0,
                result.layout.flowBounds.heightPt - result.layout.spacing.afterPt,
              ),
            }),
          }, result.layout)
        : result.layout;
      return { layout: contentOwned, nextCursor };
    },
    layoutTable(block, placement) {
      const normalizedInput: TableLayoutInput = {
        ...block,
        flowDomainId: placement.container.id,
      };
      if (block.ownerHost) {
        return layoutStoryOwnerHost(normalizedInput, block.ownerHost.framePr, placement);
      }
      previousParagraph = null;
      const result = normalizedInput.ordinaryFlow
        ? layoutStoryFlowTable(normalizedInput, placement)
        : layoutStoryTable(normalizedInput, placement, band);
      candidate.y = result.nextCursor.yPt;
      return result;
    },
  };
  /**
   * An ordinary-flow story table, admitted below the story's preceding text
   * exclusions it may not sit beside by the body's resolver
   * (resolveBlockFlowAdmission). The exclusions are the story's registered
   * floats, in story coordinates as its paragraphs consume them (a story-root
   * host's page-owned axes are registered through the inverse of the band), so
   * the table is admitted in its own coordinate space. The cleared start moves
   * the cursor and the available bounds' top, keeping their x, width and
   * bottom, and the next cursor follows the table laid out there. A table is
   * laid out again at each move, since its extent may depend on where its
   * page-placed content lands; the start is monotone and each move clears at
   * least one exclusion for good, so there are at most `blockers.length`
   * moves. Host segments never come here (layoutStoryOwnerHost): a host
   * avoids neither earlier hosts nor its own exclusion.
   */
  function layoutStoryFlowTable(
    input: TableLayoutInput,
    placement: FlowBlockPlacement,
  ): BlockLayoutResult<TableLayout> {
    const blockers = candidate.floats.map(floatRectParticipant);
    const bottomPt = placement.availableBounds.yPt + placement.availableBounds.heightPt;
    let at = placement;
    for (let moveCount = 0; ; moveCount += 1) {
      const result = layoutStoryTable(input, at, band);
      if (blockers.length === 0) return result;
      const { flowBounds } = result.layout;
      const admittedYPt = resolveBlockFlowAdmission({
        inlineStartPt: flowBounds.xPt,
        inlineEndPt: flowBounds.xPt + flowBounds.widthPt,
        flowBandStartPt: at.availableBounds.xPt,
        flowBandEndPt: at.availableBounds.xPt + at.availableBounds.widthPt,
        blockStartPt: at.cursor.yPt,
        blockExtentPt: result.layout.advancePt,
        blockers,
        overlapEpsilonPt: FLOAT_OVERLAP_EPS,
      }).blockStartPt;
      if (admittedYPt <= at.cursor.yPt) return result;
      if (moveCount === blockers.length) {
        throw new LayoutInvariantError('NON_CONVERGENCE', `story table ${input.id} admission did not converge`);
      }
      at = {
        container: at.container,
        cursor: { xPt: at.cursor.xPt, yPt: admittedYPt },
        availableBounds: {
          ...at.availableBounds,
          yPt: admittedYPt,
          heightPt: Math.max(0, bottomPt - admittedYPt),
        },
      };
    }
  }
  /**
   * A story table at `placement`. With `toPage` (story → page translation of
   * the placement's coordinates) the page-placed content below its cells is
   * placed through it, as a paginated table places it through its page
   * translation.
   */
  function layoutStoryTable(
    input: TableLayoutInput,
    placement: FlowBlockPlacement,
    toPage: Readonly<{ xPt: number; yPt: number }> | undefined,
  ): BlockLayoutResult<TableLayout> {
    const acquisition = storyTableAcquisitions.get(input.id);
    if (!toPage || !acquisition || !hasPagePlacedTableContent(acquisition)) {
      return layoutRetainedTableInput(input, placement, storyServices);
    }
    // Laid out whole at its page position, as a nested table placed by its
    // parent's pagination is: §17.4.57 positioned tables at any depth given
    // final frames on the page (a story table resolves no other way; before
    // this they were never placed). Their resolutions avoid only each other:
    // a story keeps no page float registry.
    const layout = layoutWholeTableOnPage({ ...acquisition, input }, {
      availableHeightPt: placement.availableBounds.heightPt,
      freshPageHeightPt: placement.availableBounds.heightPt,
      placement,
      services: storyServices,
      compatibility: 'word',
      page: {
        physicalPageIndex: request.pageIndex,
        displayPageNumber: candidate.displayPageNumber ?? request.pageIndex + 1,
        occurrenceId: `${request.container.id}:${input.id}`,
      },
      pagePlacement: {
        frames: storyPageFrames,
        translationPt: toPage,
        reacquireParagraph: (blockRequest) => reacquireBodyTableBlock(candidate, source, blockRequest),
      },
      floatingTableFrames: {
        page: storyPageFrames.page,
        margin: storyPageFrames.margin,
        column: storyPageFrames.margin,
      },
      floatingTableRegistry: Object.freeze({
        coordinateSpace: 'logical-page-points' as const,
        flowDomainId: request.container.id,
        entries: Object.freeze([]),
        nextParagraphId: 0,
      }),
      reacquirePageDependentBlock: (blockRequest) => reacquireBodyTableBlock(candidate, source, blockRequest),
    });
    return {
      layout,
      nextCursor: { xPt: placement.cursor.xPt, yPt: placement.cursor.yPt + layout.advancePt },
    };
  }
  /**
   * A header/footer story-root cell-owner host (library policy,
   * table-owner-runs.ts; only those roots elect, tableRowsElectCarriers).
   * Story-domain owner frame: page and margin bands are the destination
   * page's on both axes; text bands are the story column and the story
   * cursor. A header/footer is laid out before its band translation, so a
   * page/margin result is stated in page coordinates and the layout owns
   * those axes (`ownerHostPageAxes`): like a page-owned anchor layer, the
   * translation leaves them in place and they do not form the story's flow
   * extent. The host advances no story flow and is never split (stories are
   * not paginated). Its exclusion wraps the following story paragraphs, which
   * are laid out in story coordinates: page-owned axes of the exclusion are
   * stated through the inverse of the band, so the wrap lands where the host
   * is painted. A page/margin carrier makes the story band dependent
   * (body-paginator.ts layoutBandStory), so its accepted layout always has
   * its band; an unbanded trial keeps those axes in page coordinates, as
   * page-owned story anchors' exclusions do.
   */
  function layoutStoryOwnerHost(
    input: TableLayoutInput,
    framePr: FramePr,
    placement: FlowBlockPlacement,
  ): BlockLayoutResult<TableLayout> {
    const container = placement.container.bounds;
    const gridWidthPt = input.columnWidthsPt.reduce((sum, width) => sum + width, 0);
    const viewportWidthPt = Math.max(framePr.w ?? gridWidthPt, gridWidthPt);
    const pageAxes = ownerHostPageAxes(framePr);
    // Page translation of the host layout's coordinates: its page-owned axes
    // already are page coordinates.
    const hostToPage = band ? {
      xPt: pageAxes.horizontal ? 0 : band.xPt,
      yPt: pageAxes.vertical ? 0 : band.yPt,
    } : undefined;
    const layoutAt = (xPt: number, yPt: number) => layoutStoryTable(input, {
      container: placement.container,
      cursor: { xPt, yPt },
      availableBounds: { xPt, yPt, widthPt: viewportWidthPt, heightPt: placement.availableBounds.heightPt },
    }, hostToPage).layout;
    // The page, not the (possibly unbounded) story capacity.
    const frameAt = (extentPt: number) => ownerHostPageFrameBox(
      framePr, storyPageFrames, container.xPt, container.widthPt, placement.cursor.yPt, gridWidthPt, extentPt,
    );
    // The frame box depends on the host extent (yAlign, the page clamp) and,
    // through page-placed content below its cells (a positioned child its
    // anchor paragraph wraps around), the extent on where the host is placed.
    // The host is painted, framed and excluded by ONE placement: the exact
    // fixed point of box.y = frameAt(extent of the host laid out at box.y)
    // (library consistency policy, not a Word guarantee). The frame x does not
    // depend on the extent. A host whose extent does not move with it is
    // accepted by its first evaluation, so it is laid out twice, as before.
    // The limit is a resource guard only; no fixed point (a residual that
    // changes sign by a jump) fails closed. The error is restated as a plain
    // invariant so no exact-state recovery around the story can absorb it.
    const startBox = frameAt(layoutAt(container.xPt, placement.cursor.yPt).advancePt);
    let solved: Readonly<{ box: FrameBox; placed: TableLayout }>;
    try {
      solved = solveExactTranslation(startBox.y, (yPt) => {
        const laidOut = layoutAt(startBox.x, yPt);
        const next = frameAt(laidOut.advancePt);
        return { value: { box: next, placed: laidOut }, implied: next.y };
      }, 16);
    } catch (error) {
      if (error instanceof ExactConvergenceError) {
        throw new LayoutInvariantError(
          'NON_CONVERGENCE',
          `story cell-owner host placement did not converge (${error.reason}; ${error.states.length} states)`,
        );
      }
      throw error;
    }
    const { box, placed } = solved;
    // Laid out before its own exclusion is registered with the story.
    const occupied = ownerHostOccupiedBox(box, framePr, placed.advancePt);
    registerFrameFloat(band ? frameBoxAt(
      occupied,
      occupied.x - (pageAxes.horizontal ? band.xPt : 0),
      occupied.y - (pageAxes.vertical ? band.yPt : 0),
    ) : occupied, framePr, candidate);
    const finished = finishOwnerHostLayout(placed, framePr, gridWidthPt, box.x);
    const pageOwned = pageAxes.horizontal || pageAxes.vertical;
    return {
      layout: pageOwned ? Object.freeze({ ...finished, ownerHostPageAxes: pageAxes }) : finished,
      nextCursor: placement.cursor,
    };
  }
  attachStoryBlockLayoutAlgorithms(storyServices, algorithms);
  const acquired = layoutSharedStory(
    {
      source: request.source,
      container: request.container,
      blocks: Object.freeze(blockInputs),
    },
    storyServices,
  );
  // A text box of a story paragraph places its own story's page-placed
  // content through the band that paragraph receives
  // (hostFlowPageTranslationPt above).
  const textBoxBandDependent = acquired.blocks.some((block) => block.kind === 'paragraph'
    && block.textBoxes.some((textBox) => textBox.story.bandDependent === true));
  const retained = deepFreezePlainData({
    ...acquired,
    ...(bandDependent || textBoxBandDependent ? { bandDependent: true as const } : {}),
    blocks: Object.freeze(
      acquired.blocks.map((block, index) => {
        if (block.kind !== 'paragraph' && block.kind !== 'table') {
          throw new Error(`Shared story emitted unsupported node: ${block.kind}`);
        }
        return projectBodyOccurrence(block, {
          occurrenceId: `${request.container.id}:block:${index}`,
          destination: {
            coordinateSpace: 'logical-page-points',
            flowDomainId: request.container.id,
            translation: { xPt: 0, yPt: 0 },
          },
        });
      }),
    ),
  });
  return retained;
}

function applyBodyAcquisitionLocationTo(
  services: LayoutServices,
  target: BodyAcquisitionState,
  next: BodyAcquisitionLocation,
): void {
  const geometry = next.section.geometry;
  target.sectionLayout = next.section as SectionLayoutContext;
  // Each retained occurrence owns its physical anchor frame; a horizontal
  // section clears it. Two native sections can share the nominal `btLr`
  // token and still differ in frame.
  target.verticalPhys = physicalAnchorFrameOf(next.section);
  target.pageIndex = next.pageIndex;
  const page = fieldAcquisitionContextOf(services).resolveDestinationPage?.(next.pageIndex);
  target.displayPageNumber = page?.displayPageNumber ?? next.pageIndex + 1;
  target.pageNumberFormat = page?.pageNumberFormat ?? target.pageNumberFormat;
  target.pageWidth = geometry.pageWidth;
  target.pageH = geometry.pageHeight;
  target.marginLeft = geometry.marginLeft;
  target.marginRight = geometry.marginRight;
  target.marginTop = bodyMarginInsetPt(geometry.marginTop);
  target.marginBottom = bodyMarginInsetPt(geometry.marginBottom);
  target.contentX = next.availableBounds.xPt;
  target.contentW = next.availableBounds.widthPt;
  target.y = next.cursorPt.yPt;
}

function setBodyAcquisitionLocation(
  services: LayoutServices,
  state: BodyAcquisitionState,
  sessionState: BodySessionDependencies['sessionState'],
  next: BodyAcquisitionLocation,
): void {
  sessionState.location = next;
  applyBodyAcquisitionLocationTo(services, state, next);
}

interface BodySessionDependencies {
  readonly source: LayoutSourceStore;
  readonly dependencies: ConcreteBodyKernelContext;
  readonly services: LayoutServices;
  readonly state: BodyAcquisitionState;
  readonly sessionState: {
    location: BodyAcquisitionLocation;
    floatRegistry: FloatRegistrySnapshotPt;
    drawingCollisionRegistry: DrawingMLCollisionRegistrySnapshotPt;
  };
  readonly footnotesById: ReturnType<typeof indexNotes>;
  readonly endnotesById: ReturnType<typeof indexNotes>;
  readonly storyAcquisitionContext: BodyStoryAcquisitionContext;
  readonly pageRegistryFlowDomainId: (pageIndex: number) => string;
  readonly measureContext: MeasurementTextContext;
  readonly publicAnchorBridge: typeof publicAnchorBridge;
  readonly effectiveTablePositioning: typeof effectiveTablePositioning;
}

function measureBodyParagraphEntry(
  context: BodySessionDependencies,
  request: Parameters<NonNullable<BodyLayoutSession['measureParagraph']>>[0],
): ReturnType<NonNullable<BodyLayoutSession['measureParagraph']>> {
  const {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = context;
  setBodyAcquisitionLocation(services, state, sessionState, request.location);
  const paragraph = sourceElement(dependencies.source, request.input.source);
  if (paragraph.type !== 'paragraph') throw new Error('Paragraph source kind mismatch');
  if (paragraph.framePr) {
    if (request.continuation.boundary !== null) {
      throw new Error('Body frame acquisition cannot continue across flow regions');
    }
    let acquiredGroup: ReturnType<typeof acquireRetainedFrameGroup> | undefined;
    const frameGroup = bodyFrameGroupFor(paragraph);
    if (!frameGroup) {
      throw new Error('Body frame acquisition requires an indexed adjacency group');
    }
    const box = resolveFrameBox(
      paragraph,
      frameGroup,
      state,
      frameAnchorLineHeightPx(source.blocks.body, paragraph, state),
      {
        onAcquired: (acquired) => {
          acquiredGroup = acquired;
        },
      },
    );
    if (!acquiredGroup) throw new Error('Body frame acquisition omitted its retained group');
    const member = acquiredGroup.members.find((candidate) => candidate.paragraph === paragraph);
    if (!member) throw new Error('Body frame acquisition omitted its retained member');
    const dropCapAnchorLeadingPt =
      paragraph === frameGroup.members.at(-1) && frameGroup.framePr.dropCap !== 'none'
        ? wordLoweredDropCapAnchorLeadingPt(retainedFrameMaximumBaselineLoweringPt(acquiredGroup))
        : 0;
    const absoluteVertical =
      paragraph.framePr.vAnchor === 'page' || paragraph.framePr.vAnchor === 'margin';
    const frameOccurrenceId = box.exclusionId ?? `frame:${request.input.source.path.join(':')}`;
    const frameEntry: FloatRegistryEntryPt = Object.freeze({
      kind: 'frame',
      occurrenceId: frameOccurrenceId,
      exclusionId: frameOccurrenceId,
      paragraphId: sessionState.floatRegistry.nextParagraphId,
      exclusionMode: frameWrapExclusionMode(frameGroup.framePr),
      bounds: Object.freeze({
        xPt: box.x,
        yPt: box.y,
        widthPt: box.w,
        heightPt: box.h,
      }),
      exclusionBounds: Object.freeze({
        xPt: box.exLeft,
        yPt: box.exTop,
        widthPt: box.exRight - box.exLeft,
        heightPt: box.exBottom - box.exTop,
      }),
    });
    return Object.freeze({
      layout: member.fragment,
      // The catalogued projection preserves the §17.3.1.11 authored
      // exclusion height; see WORD_LOWERED_DROP_CAP_ANCHOR_LEADING.
      blockExtentPt: dropCapAnchorLeadingPt,
      fragmentation: Object.freeze({ kind: 'indivisible' as const }),
      placement: Object.freeze({
        coordinateSpace: 'logical-body' as const,
        xPt: member.fragment.flowBounds.xPt,
        yPt: member.fragment.flowBounds.yPt,
        sectionFlowOwnership: absoluteVertical ? ('page' as const) : ('host-flow' as const),
      }),
      ...(paragraph === frameGroup.owner
        ? {
            // §17.3.1.11 makes identical adjacent framePr paragraphs one frame,
            // so page admission belongs to the owner before any member is painted.
            retainedFootnoteReferenceIds: Object.freeze([
              ...new Set(
                acquiredGroup.members.flatMap((candidate) =>
                  footnoteIdsInRetainedSlice(candidate.fragment),
                ),
              ),
            ]),
          }
        : {}),
      ...(!absoluteVertical
        ? {
            relocationBlockExtentPt: Math.max(0, box.y + box.h - request.location.cursorPt.yPt),
          }
        : {}),
      ...(box.registerExclusion === false
        ? {}
        : {
            flowRegistryDelta: Object.freeze({
              floats: floatingTableRegistryDelta(
                sessionState.floatRegistry,
                Object.freeze([frameEntry]),
                sessionState.floatRegistry.nextParagraphId + 1,
              ),
            }),
          }),
    });
  }
  const candidate: BodyAcquisitionState = {
    ...state,
    floats: [...state.floats],
    pageAnchorPrescanned: new Set(state.pageAnchorPrescanned),
  };
  applyBodyAcquisitionLocationTo(services, candidate, request.location);
  const publicFloats =
    request.continuation.boundary === null
      ? acquirePublicParagraphFloats(
          sessionState,
          publicAnchorBridge,
          paragraph,
          request.input.source,
          candidate,
        )
      : Object.freeze([]);
  const acquired = acquireBodyParagraphAtLocation(
    candidate,
    paragraph,
    request.input.source,
    request.location,
    request.availableInlineExtentPt,
    request.suppressSpaceBefore,
    request.continuation,
    sessionState.drawingCollisionRegistry.entries,
  );
  const { measured, layout } = acquired;
  const markOnLineGrid =
    measured.markOnly && resolveBodyParagraphLayoutContext(candidate, paragraph).lineGrid.active;
  const allBoundaries = measured.lines.flatMap((line, index) => {
    if (line.layout.physicalLineIndex !== undefined
      && measured.lines[index + 1]?.layout.physicalLineIndex === line.layout.physicalLineIndex) return [];
    const boundary = line.layout.consumedEnd;
    if (!boundary) throw new Error('Measured line omitted its source boundary');
    return [boundary];
  });
  const retainedFloats = retainedBodyParagraphFloatEntries(sessionState, layout);
  const floatEntries = Object.freeze([...publicFloats, ...retainedFloats]);
  // This accepted-collision path is intentionally parser-owned. Hand-built
  // public-model anchors still use the compatibility float bridge (and a
  // public wrapNone run therefore has no collision entry) until the Series
  // B/C bridge removal migrates those runs to the retained OOXML contract.
  const collisionEntries = ownedParagraphAnchorCollisions(layout);
  return Object.freeze({
    layout,
    blockExtentPt: layout.advancePt,
    fragmentation: measured.markOnly || (layout.nativeReadingRelocations?.length ?? 0) > 0
      ? Object.freeze({ kind: 'indivisible' as const })
      : Object.freeze({
          kind: 'splittable' as const,
          lineEndBoundaries: Object.freeze(allBoundaries),
        }),
    ...(measured.markOnly
      ? {
          markBelowBaselinePt: measured.lastLineBelowBaselinePt,
          markOnLineGrid,
        }
      : {}),
    ...(measured.uniformRubyAdvancePt == null
      ? {}
      : { uniformRubyAdvancePt: measured.uniformRubyAdvancePt }),
    ...(floatEntries.length === 0 && collisionEntries.length === 0
      ? {}
      : {
          flowRegistryDelta: Object.freeze({
            ...(floatEntries.length === 0
              ? {}
              : {
                  floats: floatingTableRegistryDelta(
                    sessionState.floatRegistry,
                    floatEntries,
                    sessionState.floatRegistry.nextParagraphId + floatEntries.length,
                  ),
                }),
            ...(collisionEntries.length === 0
              ? {}
              : {
                  drawingCollisions: drawingMLCollisionRegistryDelta(
                    sessionState.drawingCollisionRegistry,
                    collisionEntries,
                  ),
                }),
          }),
        }),
  });
}

function layoutBodyNotes(
  context: BodySessionDependencies,
  request: Parameters<NonNullable<BodyLayoutSession['layoutNotes']>>[0],
): ReturnType<NonNullable<BodyLayoutSession['layoutNotes']>> {
  const {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = context;
  const notes: NoteLayout[] = [];
  let cursorYPt = request.container.bounds.yPt;
  let first = request.firstOnPage;
  for (const id of request.referenceIds) {
    const sourceNotes = request.kind === 'footnote' ? footnotesById : endnotesById;
    if (!sourceNotes.has(id)) continue;
    const source: SourceRef = {
      story: request.kind,
      storyInstance: id,
      path: [],
    };
    const noteSettings = context.source.bodyLayoutInput.noteLayoutSettings;
    // A native producer's reserved stories replace the scalar library band
    // for their note kind. First-on-page ownership is singular; the role
    // (ordinary versus continuing) selects the definition, whose own mark
    // selects Short/Full.
    const nativeDefinitions = noteSettings?.nativeSeparatorRoles?.[request.kind];
    const leading = first && nativeDefinitions
      ? acquireNativeNoteSeparator(context, request, nativeDefinitions[
        request.continuing === true ? 'continuationSeparator' : 'separator'
      ], cursorYPt, `${request.kind}:${id}:page:${request.pageIndex}:separator`)
      : undefined;
    // An ordinary DOCX formatted listed story (§17.11.9) replaces only the
    // scalar band height of its role with its own acquired paragraph advance
    // (authored spacing and line rule). Its mark kind and rule width keep the
    // scalar policy below; the rule sits at the midpoint of the story's mark
    // line box, not in its spacing. Bare and missing stories keep the scalar
    // band, whose rule stays at the band midpoint.
    const selectedDefinition = first && !nativeDefinitions && request.kind === 'footnote'
      ? noteSettings?.footnoteSeparatorStories?.[
        request.continuing === true ? 'continuationSeparator' : 'separator'
      ]
      : undefined;
    const selectedBand = selectedDefinition
      ? acquireSelectedNoteSeparatorBand(context, request, selectedDefinition, cursorYPt)
      : undefined;
    const separatorHeightPt = nativeDefinitions
      ? leading?.advancePt ?? 0
      : selectedBand?.advancePt ?? (first ? FOOTNOTE_SEPARATOR_GAP_PT : 0);
    const ruleYPt = cursorYPt + (selectedBand?.ruleOffsetPt ?? separatorHeightPt / 2);
    const storyContainer = {
      ...request.container,
      // Acquire the source once before page partitioning; this is a resource
      // acquisition container, not permission for retained note ink to overflow.
      ...(services.allowFootnoteContinuation === true && request.kind === 'footnote'
        ? { capacity: 'unbounded' as const } : {}),
      id: `${request.container.id}:${request.kind}:${id}`,
      bounds: {
        ...request.container.bounds,
        yPt: cursorYPt + separatorHeightPt,
        heightPt: Math.max(
          0,
          request.container.bounds.yPt +
            request.container.bounds.heightPt -
            cursorYPt -
            separatorHeightPt,
        ),
      },
    };
    // A planned page-final top moves this note's flow top (cursorYPt) there.
    const plannedTopPt = request.plannedTopsPt?.[id];
    const bandTranslationPt = request.bandTranslationPt
      ?? (plannedTopPt === undefined ? undefined : { xPt: 0, yPt: plannedTopPt - cursorYPt });
    let story: StoryLayout;
    try {
      story = acquireBodyStoryLayout(storyAcquisitionContext, {
        source,
        pageIndex: request.pageIndex,
        section: request.section,
        container: storyContainer,
        ...(bandTranslationPt ? { bandTranslationPt } : {}),
      });
    } catch (error) {
      if (error instanceof FlowCapacityExceededError && error.containerId === storyContainer.id) {
        throw new NoteCapacityExceededError(request.kind, request.pageIndex, request.container.id);
      }
      throw error;
    }
    if (services.allowFootnoteContinuation === true && request.kind === 'footnote'
      && story.advancePt > request.section.geometry.pageHeight * MAX_BODY_LAYOUT_PAGES) {
      // Match the paginator's physical-page budget before retaining a source
      // that cannot finish within it. This is resource policy, not Office fit.
      throw new Error('Footnote source exceeds the document page budget');
    }
    // ECMA-376 §17.11 reserved stories: the observed bare empty story
    // suppresses rule ink while retaining the existing note band gap.
    const separatorMode = request.kind === 'footnote'
      ? request.continuing ? noteSettings?.footnoteContinuationSeparator : noteSettings?.footnoteSeparator
      : noteSettings?.endnoteSeparator;
    const separator = first && !nativeDefinitions && separatorMode !== 'none'
      ? Object.freeze([
          Object.freeze({
            edge: 'top' as const,
            from: Object.freeze({
              xPt: request.container.bounds.xPt,
              yPt: ruleYPt,
            }),
            to: Object.freeze({
              // ECMA-376 §17.11.1/.23: marker kind wins over story role.
              // Word 16.113.3 compatibility controls (300pt main width,
              // 80 plain single-spaced paragraphs, printed Times New Roman
              // 12pt) retain a selected Short mark in a continuation story.
              // A missing continuation story prints a full-width rule and
              // gains a listed Full story on save. Preserve source absence;
              // this fallback does not synthesize a parser fact or change
              // the public continuation option's default. These controls
              // establish neither universal Short dimensions nor gap/leading.
              // §17.11.1 defines the continuation mark as full main-story width.
              // §17.11.23 defines a partial ordinary mark; one third is the
              // existing library policy, not a normative numeric fraction.
              // It differs from the 144pt Short rule in those Office controls.
              xPt: request.container.bounds.xPt + request.container.bounds.widthPt
                * (separatorMode === 'full'
                  || ((separatorMode === undefined || separatorMode === 'default') && request.continuing)
                  ? 1 : 1 / 3),
              yPt: ruleYPt,
            }),
            color: '#000000',
            widthPt: 0.5,
            authoredStyle: 'single',
            style: 'solid' as const,
          }),
        ])
      : Object.freeze([]);
    const advancePt = separatorHeightPt + story.advancePt;
    const flowBounds = Object.freeze({
      xPt: request.container.bounds.xPt,
      yPt: cursorYPt,
      widthPt: request.container.bounds.widthPt,
      heightPt: advancePt,
    });
    const note: NoteLayout = Object.freeze({
      kind: 'note',
      id: `${request.kind}:${id}:page:${request.pageIndex}`,
      source,
      flowDomainId: request.container.id,
      ordinaryFlow: true,
      flowBounds,
      inkBounds: Object.freeze({
        xPt: Math.min(flowBounds.xPt, story.inkBounds.xPt),
        yPt: Math.min(flowBounds.yPt, story.inkBounds.yPt),
        widthPt:
          Math.max(
            flowBounds.xPt + flowBounds.widthPt,
            story.inkBounds.xPt + story.inkBounds.widthPt,
          ) - Math.min(flowBounds.xPt, story.inkBounds.xPt),
        heightPt:
          Math.max(
            flowBounds.yPt + flowBounds.heightPt,
            story.inkBounds.yPt + story.inkBounds.heightPt,
          ) - Math.min(flowBounds.yPt, story.inkBounds.yPt),
      }),
      clipBounds: request.container.bounds,
      advancePt,
      separator,
      ...(leading ? { leading } : {}),
      story,
    });
    notes.push(note);
    cursorYPt += advancePt;
    first = false;
  }
  return Object.freeze(notes);
}

/** Acquire one native reserved separator story at the band cursor through
 * the shared story/paragraph pipeline and project its page occurrence. An
 * empty, guard-only or fully hidden story owns no flow, so none is retained. */
function acquireNativeNoteSeparator(
  context: BodySessionDependencies,
  request: Parameters<NonNullable<BodyLayoutSession['layoutNotes']>>[0],
  definition: NativeNoteSeparatorDefinitionInput,
  yPt: number,
  occurrenceId: string,
): NoteSeparatorLayout | undefined {
  if (!definition.paragraph || definition.paragraph.hidden) return undefined;
  const container = {
    ...request.container,
    id: `${request.container.id}:${sourceKey(definition.root)}`,
    bounds: {
      ...request.container.bounds,
      yPt,
      heightPt: Math.max(0, request.container.bounds.yPt + request.container.bounds.heightPt - yPt),
    },
  };
  let story: StoryLayout;
  try {
    story = acquireBodyStoryLayout(context.storyAcquisitionContext, {
      source: definition.root,
      pageIndex: request.pageIndex,
      section: request.section,
      container,
    });
  } catch (error) {
    if (error instanceof FlowCapacityExceededError && error.containerId === container.id) {
      throw new NoteCapacityExceededError(request.kind, request.pageIndex, request.container.id);
    }
    throw error;
  }
  return placeNoteSeparatorOccurrence(
    noteSeparatorOccurrence(definition, story, container.bounds),
    { occurrenceId, flowDomainId: request.container.id, yPt },
  );
}

/** Acquire one ordinary DOCX selected separator story at the band cursor
 * through the shared story/paragraph pipeline. Returns two retained facts:
 * - `advancePt`, the band height: the paragraph's acquired advance, i.e. its
 *   mark line box under the authored line rule plus authored spacing before
 *   and after (§17.3.1.33 places that spacing outside the paragraph's lines).
 * - `ruleOffsetPt`: the midpoint of the acquired mark line box, measured from
 *   the story's own top in its acquisition frame, so it applies at whatever
 *   band y it is placed. §17.11.23 places the separator mark in its run; the
 *   parser admits only a mark run formatted exactly like the paragraph mark,
 *   so that run's line is this mark-only line box. The midpoint is the
 *   library's existing rule placement, not an Office glyph-baseline
 *   observation.
 * No paragraph layout is retained: the closed run-free shape has neither text
 * nor paragraph ink. */
function acquireSelectedNoteSeparatorBand(
  context: BodySessionDependencies,
  request: Parameters<NonNullable<BodyLayoutSession['layoutNotes']>>[0],
  definition: SelectedNoteSeparatorDefinitionInput,
  yPt: number,
): Readonly<{ advancePt: number; ruleOffsetPt: number }> {
  const container = {
    ...request.container,
    id: `${request.container.id}:${sourceKey(definition.root)}`,
    bounds: {
      ...request.container.bounds,
      yPt,
      heightPt: Math.max(0, request.container.bounds.yPt + request.container.bounds.heightPt - yPt),
    },
  };
  let story: StoryLayout;
  try {
    story = acquireBodyStoryLayout(context.storyAcquisitionContext, {
      source: definition.root,
      pageIndex: request.pageIndex,
      section: request.section,
      container,
    });
  } catch (error) {
    if (error instanceof FlowCapacityExceededError && error.containerId === container.id) {
      throw new NoteCapacityExceededError(request.kind, request.pageIndex, request.container.id);
    }
    throw error;
  }
  const [paragraph, ...rest] = story.blocks;
  // The run-free paragraph acquires no text line; its mark-only line box is
  // the paragraph mark bounds (spacing before excluded, line rule applied).
  const mark = paragraph?.kind === 'paragraph' && rest.length === 0 && paragraph.lines.length === 0
    ? paragraph.paragraphMark : undefined;
  if (!mark || mark.hidden) {
    throw new Error('A note separator story must acquire exactly one visible mark-only paragraph');
  }
  return Object.freeze({
    advancePt: story.advancePt,
    ruleOffsetPt: mark.bounds.yPt + mark.bounds.heightPt / 2 - story.flowBounds.yPt,
  });
}

function measureFollowingBodyBlock(
  context: BodySessionDependencies,
  request: Parameters<NonNullable<BodyLayoutSession['measureFollowingBlock']>>[0],
): ReturnType<NonNullable<BodyLayoutSession['measureFollowingBlock']>> {
  const {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = context;
  const candidate: BodyAcquisitionState = {
    ...state,
    floats: [...state.floats],
    retainedTablesBySourceIndex: new Map(state.retainedTablesBySourceIndex),
  };
  applyBodyAcquisitionLocationTo(services, candidate, request.location);
  if (request.input.kind === 'adjacent-table-group') {
    const records = computeAdjacentTablePtLayouts(candidate, request.input.tables.map((tableInput) => {
      const table = sourceElement(dependencies.source, tableInput.source);
      if (table.type !== 'table') throw new Error('Following table source kind mismatch');
      return { table, sourceIndex: tableInput.source.path[0]! };
    }), request.availableInlineExtentPt);
    const combinedInput = ordinaryAcquisitionInputForAdjacentGroup(
      combineAdjacentTableLayoutInputs(
        request.input.logicalSequenceId,
        records.map((record) => record.input),
      ),
    );
    const layout = layoutRetainedTableInput(
      combinedInput,
      {
        container: {
          id: request.location.flowDomainId,
          kind: 'body',
          bounds: request.location.availableBounds,
        },
        cursor: request.location.cursorPt,
        availableBounds: request.location.availableBounds,
      },
      services,
    ).layout;
    const combined: RetainedTableAcquisition = {
      input: combinedInput,
      layout,
      nestedById: Object.assign({}, ...records.map((record) => record.nestedById)),
      floatingTables: [],
    };
    const ownerExtents = ownerSegmentedFlowExtents(
      combined,
      request.availableInlineExtentPt,
      services,
    );
    return Object.freeze({
      fullExtentPt: ownerExtents?.fullExtentPt ?? layout.advancePt,
      leadContentExtentPt: ownerExtents?.leadContentExtentPt
        ?? layout.rows[0]?.advancePt ?? layout.advancePt,
      fullFootnoteReferenceIds: footnoteIdsInRetainedSlice(layout),
      leadFootnoteReferenceIds: footnoteIdsInRetainedSlice({
        ...layout,
        rows: layout.rows.slice(0, 1),
      }),
    });
  }
  const element = sourceElement(dependencies.source, request.input.source);
  if (request.input.kind === 'paragraph') {
    if (element.type !== 'paragraph') throw new Error('Following paragraph source kind mismatch');
    const { layout } = acquireBodyParagraphAtLocation(
      candidate,
      element,
      request.input.source,
      request.location,
      request.availableInlineExtentPt,
      false,
      undefined,
      sessionState.drawingCollisionRegistry.entries,
    );
    const firstLine = layout.lines[0];
    return Object.freeze({
      fullExtentPt: layout.advancePt,
      // keepNext admits the successor's first content line, including
      // any retained wrap displacement before that line begins.
      leadContentExtentPt: firstLine
        ? firstLine.bounds.yPt + firstLine.advancePt - layout.flowBounds.yPt
        : layout.advancePt,
      fullFootnoteReferenceIds: footnoteIdsInRetainedSlice(layout),
      leadFootnoteReferenceIds: firstLine ? footnoteIdsInRetainedLines([firstLine]) : [],
      pageOwnedAnchorKeysByLine: pageOwnedAnchorKeysByLine(layout),
    });
  }
  if (element.type !== 'table') throw new Error('Following table source kind mismatch');
  const sourceIndex = request.input.source.path[0]!;
  // The same upright physical acquisition owner as actual measurement
  // (bodyTableAcquisitionState), so lookahead sizes the table, and its owner
  // segments, exactly as it will be measured on the page.
  const tableState = bodyTableAcquisitionState(candidate, element, effectiveTablePositioning);
  computeTablePtLayout(tableState, element, request.availableInlineExtentPt, sourceIndex);
  const acquisition = retainedTableRecord(tableState, sourceIndex).acquisition;
  const layout = acquisition.layout;
  const ownerExtents = ownerSegmentedFlowExtents(
    acquisition,
    request.availableInlineExtentPt,
    services,
  );
  return Object.freeze({
    fullExtentPt: ownerExtents?.fullExtentPt ?? layout.advancePt,
    leadContentExtentPt: ownerExtents?.leadContentExtentPt
      ?? layout.rows[0]?.advancePt ?? layout.advancePt,
    fullFootnoteReferenceIds: footnoteIdsInRetainedSlice(layout),
    leadFootnoteReferenceIds: footnoteIdsInRetainedSlice({
      ...layout,
      rows: layout.rows.slice(0, 1),
    }),
  });
}

function prescanBodyPageAnchors(
  context: BodySessionDependencies,
  request: Parameters<NonNullable<BodyLayoutSession['prescanPageAnchors']>>[0],
): ReturnType<NonNullable<BodyLayoutSession['prescanPageAnchors']>> {
  const {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = context;
  const geometry = request.location.section.geometry;
  const marginTopPt = bodyMarginInsetPt(geometry.marginTop);
  const marginBottomPt = bodyMarginInsetPt(geometry.marginBottom);
  const frames = Object.freeze({
    page: Object.freeze({
      xPt: 0,
      yPt: 0,
      widthPt: geometry.pageWidth,
      heightPt: geometry.pageHeight,
    }),
    margin: Object.freeze({
      xPt: geometry.marginLeft,
      yPt: marginTopPt,
      widthPt: Math.max(0, geometry.pageWidth - geometry.marginLeft - geometry.marginRight),
      heightPt: Math.max(0, geometry.pageHeight - marginTopPt - marginBottomPt),
    }),
    column: Object.freeze({
      xPt: request.location.availableBounds.xPt,
      yPt: marginTopPt,
      widthPt: request.availableInlineExtentPt,
      heightPt: Math.max(0, geometry.pageHeight - marginTopPt - marginBottomPt),
    }),
    paragraph: null,
    line: null,
    character: null,
    pageParity: request.location.pageIndex % 2 === 0 ? ('odd' as const) : ('even' as const),
  });
  const publicParagraphs = new Set<ParagraphLayoutSource>();
  const paragraphKey = (source: SourceRef) =>
    `${source.story}:${source.storyInstance}:${source.path.join('.')}`;
  const paragraphIds = new Map<string, number>();
  const paragraphIdFor = (source: SourceRef): number => {
    const key = paragraphKey(source);
    if (!paragraphIds.has(key)) {
      paragraphIds.set(key, sessionState.floatRegistry.nextParagraphId + paragraphIds.size);
    }
    return paragraphIds.get(key)!;
  };
  const entries = request.anchors.flatMap((anchor): readonly FloatRegistryEntryPt[] => {
    if (anchor.kind === 'host-drawing') {
      // WORD_LATER_ANCHOR_EARLIER_LINE_WRAP: register the first-placement
      // geometry at page start; the anchor paragraph keeps the same frame.
      const { carry } = anchor;
      if (!state.frozenAnchorFrames) state.frozenAnchorFrames = new Map();
      state.frozenAnchorFrames.set(anchor.occurrenceId, Object.freeze({ ...carry.bounds }));
      return [
        Object.freeze({
          kind: 'shape' as const,
          occurrenceId: anchor.occurrenceId,
          exclusionId: anchor.occurrenceId,
          paragraphId: paragraphIdFor(anchor.paragraphSource),
          bounds: carry.bounds,
          exclusionBounds: carry.exclusionBounds,
          ...(carry.horizontalOwnership ? { horizontalOwnership: carry.horizontalOwnership } : {}),
          ...(carry.verticalOwnership ? { verticalOwnership: carry.verticalOwnership } : {}),
          wrap: carry.wrap,
          wrapSide: carry.wrapSide ?? null,
          ...(carry.wrapDistances ? { wrapDistances: carry.wrapDistances } : {}),
          ...(carry.wrapPolygon ? { wrapPolygon: carry.wrapPolygon } : {}),
          ...(carry.topEdgeInclusiveFromYPt === undefined
            ? {} : { topEdgeInclusiveFromYPt: carry.topEdgeInclusiveFromYPt }),
          ...(carry.anchorLineExemptTopPt === undefined
            ? {} : { anchorLineExemptTopPt: carry.anchorLineExemptTopPt }),
        }),
      ];
    }
    if (anchor.kind === 'floating-table') {
      // §17.4.57 topFromText is the minimum gap above a positioned
      // table. In controlled Word output, a page-positioned table
      // keeps its authored y while preceding paragraph lines that
      // intersect its exclusion move below it. This holds with and
      // without an intervening empty mark; increasing topFromText
      // can move even the first line. Reuse the first pass's actual
      // page/fragment bounds rather than guessing table height here.
      const table = sourceElement(dependencies.source, anchor.tableSource);
      if (table.type !== 'table') {
        throw new Error('Page-positioned table prescan source kind mismatch');
      }
      const positioning = state.acquisitionInputs.tableFormatInput(table).positioning;
      if (
        !positioning ||
        (positioning.vertAnchor !== 'page' && positioning.vertAnchor !== 'margin')
      ) {
        throw new Error('Page-positioned table prescan requires a page-owned vertical axis');
      }
      const { bounds } = anchor;
      return [
        Object.freeze({
          kind: 'table' as const,
          occurrenceId: anchor.occurrenceId,
          overlap: table.overlap === 'never' ? ('never' as const) : ('overlap' as const),
          paragraphId: paragraphIdFor(anchor.tableSource),
          bounds,
          exclusionBounds: Object.freeze({
            xPt: bounds.xPt - positioning.leftFromTextPt,
            yPt: bounds.yPt - positioning.topFromTextPt,
            widthPt: bounds.widthPt + positioning.leftFromTextPt + positioning.rightFromTextPt,
            heightPt: bounds.heightPt + positioning.topFromTextPt + positioning.bottomFromTextPt,
          }),
        }),
      ];
    }
    const paragraph = sourceElement(dependencies.source, anchor.paragraphSource);
    if (paragraph.type !== 'paragraph') {
      throw new Error('Page-anchor prescan source kind mismatch');
    }
    const acquired = paragraph;
    const hostMatches = acquired.runs.filter(
      (run) => run.type === 'anchorHost' && run.anchorOccurrenceId === anchor.occurrenceId,
    );
    const payloads = acquired.runs
      .map((run, runIndex) => ({ run, runIndex }))
      .filter(
        (
          candidate,
        ): candidate is typeof candidate & {
          run: Extract<
            typeof candidate.run,
            {
              type: 'image' | 'chart' | 'shape' | 'unavailableDrawing';
            }
          > & {
            anchorAcquisitionInput: NonNullable<
              Extract<
                typeof candidate.run,
                {
                  type: 'image' | 'chart' | 'shape' | 'unavailableDrawing';
                }
              >['anchorAcquisitionInput']
            >;
          };
        } =>
          (candidate.run.type === 'image' ||
            candidate.run.type === 'chart' ||
            candidate.run.type === 'shape' ||
            candidate.run.type === 'unavailableDrawing') &&
          candidate.run.anchorAcquisitionInput?.occurrenceId === anchor.occurrenceId,
      )
      .sort(
        (left, right) =>
          (left.run.anchorAcquisitionInput.group?.sourceIndex ?? 0) -
            (right.run.anchorAcquisitionInput.group?.sourceIndex ?? 0) ||
          left.runIndex - right.runIndex,
      );
    if (hostMatches.length !== 1 || payloads.length === 0) {
      const publicRun = paragraph.runs.find(
        (run, runIndex) =>
          publicAnchorBridge(anchor.paragraphSource, runIndex)?.occurrenceId ===
          anchor.occurrenceId,
      );
      if (publicRun) {
        if (
          (publicRun.type === 'image' ||
            publicRun.type === 'chart' ||
            publicRun.type === 'shape') &&
          publicRun.wrapMode === 'none'
        )
          return [];
        const candidate: BodyAcquisitionState = {
          ...state,
          floats: [...state.floats],
          pageAnchorPrescanned: new Set(state.pageAnchorPrescanned),
        };
        applyBodyAcquisitionLocationTo(services, candidate, request.location);
        const publicEntries = acquirePublicParagraphFloats(
          sessionState,
          publicAnchorBridge,
          paragraph,
          anchor.paragraphSource,
          candidate,
          new Set([anchor.occurrenceId]),
          paragraphIdFor(anchor.paragraphSource),
        );
        if (publicEntries.length !== 1) {
          throw new Error(`Public page-anchor prescan occurrence mismatch: ${anchor.occurrenceId}`);
        }
        publicParagraphs.add(paragraph);
        return publicEntries;
      }
      throw new Error(
        `Page-anchor prescan occurrence acquisition mismatch: ${anchor.occurrenceId}`,
      );
    }
    if (payloads[0]!.run.anchorAcquisitionInput.nativeReadingRelocation === 'completeScene') {
      // No page-owned exclusion is claimed for changed placement. Defer only
      // after the same complete scene/resource owner validates all members.
      const scenes = acquireRequestedNativeReadingScenes(acquired, anchor.paragraphSource,
        `prescan:${anchor.occurrenceId}`, pageRegistryFlowDomainId(request.location.pageIndex), paintResourceRegistryOf(services));
      if (!scenes.some(scene => scene.occurrenceId === anchor.occurrenceId)) throw new Error('Reading prescan omitted its complete scene');
      return [];
    }
    const result = resolveAnchorFrame({
      acquisition: payloads[0]!.run.anchorAcquisitionInput,
      frames,
    });
    if (result.status !== 'resolved') {
      throw new Error(`Page-anchor prescan could not resolve occurrence: ${anchor.occurrenceId}`);
    }
    // ECMA-376 §17.6.20 + §§20.4.2.3/.7/.10/.11: positionH/V and
    // extent resolve in the upright physical drawing frame. The page
    // registry is section-logical, so page-start prescan must apply the
    // section writing mode's canonical physical-to-logical affine
    // inverse before the exclusion can affect earlier body content.
    const retainedResult = isVerticalTextDirection(request.location.section.textDirection)
      ? (() => {
          const writingMode = sectionWritingMode(request.location.section);
          const physicalPage = uprightPhysicalExtent(
            {
              widthPt: frames.page.widthPt,
              heightPt: frames.page.heightPt,
            },
            writingMode,
          );
          return projectPhysicalAnchorResult(
            result,
            physicalToLogicalMatrix(writingMode, physicalPage),
          );
        })()
      : result;
    const wrapBounds = retainedResult.geometry.wrapBounds;
    if (wrapBounds === null || retainedResult.geometry.wrap.kind === 'none') {
      return [];
    }
    const polygon =
      retainedResult.geometry.wrap.polygon?.points ??
      Object.freeze([
        Object.freeze({ xPt: wrapBounds.xPt, yPt: wrapBounds.yPt }),
        Object.freeze({ xPt: wrapBounds.xPt + wrapBounds.widthPt, yPt: wrapBounds.yPt }),
        Object.freeze({
          xPt: wrapBounds.xPt + wrapBounds.widthPt,
          yPt: wrapBounds.yPt + wrapBounds.heightPt,
        }),
        Object.freeze({ xPt: wrapBounds.xPt, yPt: wrapBounds.yPt + wrapBounds.heightPt }),
      ]);
    return [
      Object.freeze({
        kind: 'shape' as const,
        occurrenceId: anchor.occurrenceId,
        paragraphId: paragraphIdFor(anchor.paragraphSource),
        bounds: retainedResult.geometry.objectFrame,
        exclusionBounds: wrapBounds,
        wrap: retainedResult.geometry.wrap.kind,
        wrapSide: retainedResult.geometry.wrap.side,
        wrapDistances: retainedResult.geometry.wrap.distances,
        wrapPolygon: Object.freeze([...polygon]),
      }),
    ];
  });
  publicParagraphs.forEach((paragraph) => state.pageAnchorPrescanned?.add(paragraph));
  if (entries.length === 0) return null;
  return Object.freeze({
    floats: Object.freeze({
      coordinateSpace: 'logical-page-points' as const,
      // Page-owned wrap exclusions survive same-page column/section
      // cutovers, so their transaction identity belongs to the physical
      // page rather than the active body flow domain.
      flowDomainId: sessionState.floatRegistry.flowDomainId,
      baseEntries: sessionState.floatRegistry.entries,
      baseNextParagraphId: sessionState.floatRegistry.nextParagraphId,
      nextParagraphId: sessionState.floatRegistry.nextParagraphId + entries.length,
      entries: Object.freeze(entries),
    }),
  });
}

function measureBodyLineNumberGlyph(
  context: BodySessionDependencies,
  text: Parameters<NonNullable<BodyLayoutSession['measureLineNumberGlyph']>>[0],
): ReturnType<NonNullable<BodyLayoutSession['measureLineNumberGlyph']>> {
  const {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = context;
  const previousFont = measureContext.font;
  try {
    const fontSizePt = source.fonts.defaultBodyFontSizePt;
    const font = buildFont(false, false, fontSizePt, null, {});
    measureContext.font = font;
    const metrics = measureContext.measureText(text);
    return Object.freeze({
      widthPt: metrics.width,
      ascentPt:
        metrics.fontBoundingBoxAscent ?? metrics.actualBoundingBoxAscent ?? fontSizePt * 0.8,
      descentPt:
        metrics.fontBoundingBoxDescent ?? metrics.actualBoundingBoxDescent ?? fontSizePt * 0.2,
      font,
    });
  } finally {
    measureContext.font = previousFont;
  }
}

function resetBodyPageAcquisition(
  context: BodySessionDependencies,
  next: Parameters<NonNullable<BodyLayoutSession['resetPageAcquisition']>>[0],
): ReturnType<NonNullable<BodyLayoutSession['resetPageAcquisition']>> {
  const {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = context;
  state.floats = [];
  state.floatParaSeq = 0;
  state.pageAnchorPrescanned = new Set();
  state.frozenAnchorFrames = new Map();
  sessionState.floatRegistry = Object.freeze({
    coordinateSpace: 'logical-page-points' as const,
    flowDomainId: pageRegistryFlowDomainId(next.pageIndex),
    entries: Object.freeze([]),
    nextParagraphId: 0,
  });
  sessionState.drawingCollisionRegistry = createDrawingMLCollisionRegistry(
    pageRegistryFlowDomainId(next.pageIndex),
    'logical-page-points',
  );
  setBodyAcquisitionLocation(services, state, sessionState, next);
}

function bodyFlowRegistrySnapshot(
  context: BodySessionDependencies,
): ReturnType<NonNullable<BodyLayoutSession['flowRegistrySnapshot']>> {
  const {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = context;
  return Object.freeze({
    floats: sessionState.floatRegistry,
    drawingCollisions: sessionState.drawingCollisionRegistry,
  });
}

function commitBodyFlowRegistryDelta(
  context: BodySessionDependencies,
  delta: Parameters<NonNullable<BodyLayoutSession['commitFlowRegistryDelta']>>[0],
): ReturnType<NonNullable<BodyLayoutSession['commitFlowRegistryDelta']>> {
  const {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = context;
  if (!delta.floats && !delta.drawingCollisions) {
    throw new Error('Body flow registry delta must update at least one registry');
  }
  if (delta.floats) {
    validateFloatingTableRegistryDelta(delta.floats, {
      coordinateSpace: sessionState.floatRegistry.coordinateSpace,
      flowDomainId: sessionState.floatRegistry.flowDomainId,
      entries: sessionState.floatRegistry.entries,
      nextParagraphId: sessionState.floatRegistry.nextParagraphId,
    });
  }
  if (delta.drawingCollisions) {
    validateDrawingMLCollisionRegistryDelta(
      sessionState.drawingCollisionRegistry,
      delta.drawingCollisions,
    );
  }
  const nextDrawingCollisionRegistry = delta.drawingCollisions
    ? applyDrawingMLCollisionRegistryDelta(
        sessionState.drawingCollisionRegistry,
        delta.drawingCollisions,
      )
    : sessionState.drawingCollisionRegistry;
  const retainedFloats = (delta.floats?.entries ?? []).map((entry): FloatRect => {
    const left = entry.wrapDistances?.leftPt ?? entry.bounds.xPt - entry.exclusionBounds.xPt;
    const top = entry.wrapDistances?.topPt ?? entry.bounds.yPt - entry.exclusionBounds.yPt;
    const right =
      entry.wrapDistances?.rightPt ??
      entry.exclusionBounds.xPt +
        entry.exclusionBounds.widthPt -
        entry.bounds.xPt -
        entry.bounds.widthPt;
    const bottom =
      entry.wrapDistances?.bottomPt ??
      entry.exclusionBounds.yPt +
        entry.exclusionBounds.heightPt -
        entry.bounds.yPt -
        entry.bounds.heightPt;
    const core = {
      mode: (entry.kind === 'frame' && entry.exclusionMode
        ? entry.exclusionMode
        : entry.wrap === 'topAndBottom' ? 'topAndBottom' : 'square') as FloatRect['mode'],
      ...(entry.kind === 'shape'
        ? {
            anchorOccurrenceId: entry.occurrenceId,
            acquisitionOccurrenceId: entry.occurrenceId,
          }
        : {}),
      ...(entry.wrap
        ? {
            authoredWrap: entry.wrap,
            wrapPolygon: entry.wrapPolygon,
          }
        : {}),
      imageKey:
        entry.exclusionId ?? (entry.kind === 'table' ? `body:float:${entry.paragraphId}` : ''),
      imageX: entry.bounds.xPt,
      imageY: entry.bounds.yPt,
      imageW: entry.bounds.widthPt,
      imageH: entry.bounds.heightPt,
      xLeft: entry.exclusionBounds.xPt,
      xRight: entry.exclusionBounds.xPt + entry.exclusionBounds.widthPt,
      yTop: entry.exclusionBounds.yPt,
      yBottom: entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt,
      side: entry.wrapSide ?? 'bothSides',
      distLeft: left,
      distRight: right,
      distTop: top,
      distBottom: bottom,
      paraId: entry.paragraphId,
      ...(entry.topEdgeInclusiveFromYPt === undefined
        ? {} : { topEdgeInclusiveFromYPt: entry.topEdgeInclusiveFromYPt }),
      ...(entry.anchorLineExemptTopPt === undefined
        ? {} : { exemptLineTopPt: entry.anchorLineExemptTopPt }),
    };
    return entry.kind === 'table'
      ? {
          ...core,
          kind: 'table',
          tableOverlap: entry.overlap,
        }
      : { ...core, kind: entry.kind };
  });
  if (delta.floats) {
    state.floats.push(...retainedFloats);
    sessionState.floatRegistry = Object.freeze({
      ...sessionState.floatRegistry,
      entries: Object.freeze([...sessionState.floatRegistry.entries, ...delta.floats.entries]),
      nextParagraphId: delta.floats.nextParagraphId,
    });
    state.floatParaSeq = delta.floats.nextParagraphId;
  }
  sessionState.drawingCollisionRegistry = nextDrawingCollisionRegistry;
}

function acquirePublicParagraphFloats(
  sessionState: BodySessionDependencies['sessionState'],
  publicAnchorBridge: ConcreteBodyKernelContext['publicAnchorBridge'],
  paragraph: LayoutParagraphBlock,
  source: SourceRef,
  candidate: BodyAcquisitionState,
  onlyOccurrenceIds?: ReadonlySet<string>,
  paragraphId = sessionState.floatRegistry.nextParagraphId,
): readonly FloatRegistryEntryPt[] {
  const committedOccurrenceIds = new Set(
    sessionState.floatRegistry.entries.map((entry) => entry.occurrenceId),
  );
  const publicRuns = paragraph.runs.flatMap((run, runIndex) => {
    if (run.type !== 'shape' && run.type !== 'image' && run.type !== 'chart') return [];
    const bridge = publicAnchorBridge(source, runIndex);
    if (
      !bridge ||
      (onlyOccurrenceIds && !onlyOccurrenceIds.has(bridge.occurrenceId)) ||
      committedOccurrenceIds.has(bridge.occurrenceId) ||
      (bridge.pageOwned && candidate.pageAnchorPrescanned?.has(paragraph))
    )
      return [];
    return [
      {
        run,
        occurrenceId: bridge.occurrenceId,
      },
    ];
  });
  if (publicRuns.length === 0) return Object.freeze([]);
  const baseFloatCount = candidate.floats.length;
  registerAnchorFloats(
    { ...paragraph, runs: publicRuns.map(({ run }) => run) },
    candidate,
    candidate.y,
  );
  const registered = candidate.floats.slice(baseFloatCount);
  if (registered.length !== publicRuns.length) {
    throw new Error('Public paragraph anchor acquisition did not retain every wrap float');
  }
  return Object.freeze(
    registered.map((float, index): FloatRegistryEntryPt => {
      const occurrenceId = publicRuns[index]!.occurrenceId;
      return Object.freeze({
        kind: 'shape',
        occurrenceId,
        exclusionId: occurrenceId,
        paragraphId,
        bounds: Object.freeze({
          xPt: float.imageX,
          yPt: float.imageY,
          widthPt: float.imageW,
          heightPt: float.imageH,
        }),
        exclusionBounds: Object.freeze({
          xPt: float.xLeft,
          yPt: float.yTop,
          widthPt: float.xRight - float.xLeft,
          heightPt: float.yBottom - float.yTop,
        }),
        wrap: publicRuns[index]!.run.wrapMode as NonNullable<FloatRegistryEntryPt['wrap']>,
        wrapSide: float.side,
        wrapDistances: Object.freeze({
          topPt: float.distTop,
          rightPt: float.distRight,
          bottomPt: float.distBottom,
          leftPt: float.distLeft,
        }),
        ...(float.wrapPolygon ? { wrapPolygon: Object.freeze([...float.wrapPolygon]) } : {}),
      });
    }),
  );
}

function retainedBodyParagraphFloatEntries(
  sessionState: BodySessionDependencies['sessionState'],
  layout: ParagraphLayout,
): readonly FloatRegistryEntryPt[] {
  const hostFrames = new Map(
    (layout.anchorFrames ?? []).flatMap((frame) => {
      if (frame.status !== 'resolved') return [];
      const isHostAxis = (axis: typeof frame.axes.horizontal) =>
        axis.status === 'resolved' &&
        (axis.referenceFrame === 'paragraph' ||
          axis.referenceFrame === 'line' ||
          axis.referenceFrame === 'character');
      return isHostAxis(frame.axes.horizontal) || isHostAxis(frame.axes.vertical)
        ? [[frame.occurrenceId, frame] as const]
        : [];
    }),
  );
  if (hostFrames.size === 0) return Object.freeze([]);
  // A drawing carried to this page (WORD_LATER_ANCHOR_EARLIER_LINE_WRAP) is
  // already registered with the same geometry.
  const registered = new Set(sessionState.floatRegistry.entries.map((entry) => entry.occurrenceId));
  const exclusions = new Map(
    layout.exclusions.flatMap((exclusion) =>
      exclusion.anchorOccurrenceId ? [[exclusion.anchorOccurrenceId, exclusion] as const] : [],
    ),
  );
  return Object.freeze(
    (layout.anchorCollisions ?? []).flatMap((collision): FloatRegistryEntryPt[] => {
      const frame = hostFrames.get(collision.occurrenceId);
      if (!frame || registered.has(collision.occurrenceId)) return [];
      if (frame.geometry.wrap.kind === 'none') return [];
      const exclusion = exclusions.get(collision.occurrenceId);
      if (!exclusion) {
        throw new Error(`Wrapped anchor omitted exclusion geometry: ${collision.occurrenceId}`);
      }
      return [
        Object.freeze({
          kind: 'shape' as const,
          occurrenceId: collision.occurrenceId,
          exclusionId: collision.occurrenceId,
          paragraphId: sessionState.floatRegistry.nextParagraphId,
          bounds: collision.bounds,
          exclusionBounds: exclusion.bounds,
          horizontalOwnership: collision.horizontalOwnership,
          verticalOwnership: collision.verticalOwnership,
          wrap: frame.geometry.wrap.kind,
          wrapSide: frame.geometry.wrap.side,
          wrapDistances: frame.geometry.wrap.distances,
          ...(frame.geometry.wrap.polygon
            ? { wrapPolygon: frame.geometry.wrap.polygon.points }
            : {}),
        }),
      ];
    }),
  );
}

function reacquireBodyTableBlock(
  state: BodyAcquisitionState,
  sourceStore: LayoutSourceStore,
  request: PageDependentTableBlockRequest,
): ParagraphLayout | TableLayout {
  if (request.acquired.kind !== 'paragraph') return request.acquired;
  const source = nestedSourceElement(sourceStore, request.acquired.source);
  if (source.type !== 'paragraph') {
    throw new Error('Table paragraph re-acquisition source kind mismatch');
  }
  const referenceDelta = request.paragraphAnchorReferenceDeltaPt !== undefined
    && wordGridFrameCarrierSourceOwnsNoFlow(
      state.acquisitionInputs.paragraphAcquisitionInput(source, request.acquired.source),
    )
    ? request.paragraphAnchorReferenceDeltaPt : undefined;
  const candidate: BodyAcquisitionState = {
    ...withTableCellStory(tableCellOwnerState(state)),
    contentX: 0,
    contentW: request.acquired.flowBounds.widthPt,
    y: request.acquired.flowBounds.yPt,
    floats: (request.floatingTableExclusions ?? []).map(
      (bounds, index): FloatRect => ({
        kind: 'table',
        tableOverlap: 'never',
        mode: 'square',
        imageKey: `${TRANSIENT_TABLE_FINAL_FRAME_EXCLUSION_PREFIX}${index}`,
        imageX: bounds.xPt,
        imageY: bounds.yPt,
        imageW: bounds.widthPt,
        imageH: bounds.heightPt,
        xLeft: bounds.xPt,
        xRight: bounds.xPt + bounds.widthPt,
        yTop: bounds.yPt,
        yBottom: bounds.yPt + bounds.heightPt,
        side: 'bothSides',
        distLeft: 0,
        distRight: 0,
        distTop: 0,
        distBottom: 0,
        paraId: index,
      }),
    ),
    floatParaSeq: request.floatingTableExclusions?.length ?? 0,
    pageAnchorPrescanned: new Set<ParagraphLayoutSource>(),
    cellHostFlowPageTranslationPt: request.hostFlowPageTranslationPt,
    cellParagraphAnchorReferenceDeltaPt: referenceDelta,
  };
  const inheritedAuthority = inheritedParagraphAuthorityForReacquisition(request.acquired);
  const tableAcquisition = state.retainedTableAcquisition;
  return tableAcquisition.acquireParagraph(
    candidate,
    source,
    request.acquired.flowBounds.widthPt,
    request.acquired.source.path,
    request.acquired.flowDomainId,
    undefined,
    inheritedAuthority,
    // A story table's paragraph keeps its own story (headers, notes).
    request.acquired.source,
  );
}

/** Bound acquirers per session callback, owner section context and page. The
 * paragraph acquisition cache keys on the acquirer's identity, so equal
 * authority must yield the identical function. */
const boundCompleteStoryAcquirers = new WeakMap<
  NonNullable<BodyAcquisitionState['acquireCompleteTextBoxStory']>,
  WeakMap<SectionLayoutContext, Map<number, CompleteTextBoxStoryAcquirer>>
>();

/** Bind nested complete-story acquisition to the section and page of the exact
 * state that contains the text box (body location, table cell or story
 * candidate), never to the session's current body location. */
function completeTextBoxStoryAcquirerFor(
  owner: BodyAcquisitionState,
): CompleteTextBoxStoryAcquirer | undefined {
  const acquire = owner.acquireCompleteTextBoxStory;
  if (!acquire) return undefined;
  const { sectionLayout, pageIndex } = owner;
  let bySection = boundCompleteStoryAcquirers.get(acquire);
  if (!bySection) {
    bySection = new WeakMap();
    boundCompleteStoryAcquirers.set(acquire, bySection);
  }
  let byPage = bySection.get(sectionLayout);
  if (!byPage) {
    byPage = new Map();
    bySection.set(sectionLayout, byPage);
  }
  let bound = byPage.get(pageIndex);
  if (!bound) {
    const authority = Object.freeze({ sectionLayout, pageIndex });
    bound = (request) => acquire(authority, request);
    byPage.set(pageIndex, bound);
  }
  return bound;
}

/** Acquire a text box's complete story with its owner's section, page and
 * frame as authority. The section-keyed story cache stays shared because its
 * key includes the full section context. */
function acquireCompleteBodyTextBoxStory(
  owner: CompleteTextBoxStoryOwner,
  storyAcquisitionContext: BodyStoryAcquisitionContext,
  request: Parameters<CompleteTextBoxStoryAcquirer>[0],
): ReturnType<CompleteTextBoxStoryAcquirer> {
  // An upright physical text box is horizontal: it keeps the owner section's
  // physical page box but not that section's native frame.
  const { nativeSectionFlow, ...sectionLayout } = owner.sectionLayout;
  const section =
    request.coordinateSpace === 'upright-physical'
      ? {
          ...sectionLayout,
          geometry: physicalSectionGeometry(owner.sectionLayout.geometry, nativeSectionFlow),
          textDirection: 'lrTb',
        }
      : owner.sectionLayout;
  return acquireBodyStoryLayout(storyAcquisitionContext, {
    source: request.source,
    pageIndex: owner.pageIndex,
    section,
    container: request.container,
    ...(request.pageFrames ? { pageFrames: request.pageFrames } : {}),
  });
}

function openConcreteBodyLayoutSession(
  dependencies: ConcreteBodyKernelContext,
  input: import('./body-layout-kernel.js').BodyLayoutSessionInput,
  services: LayoutServices,
  options: LayoutOptions,
): BodyLayoutSession {
  const {
    source,
    measureContext,
    resolvedLocalFonts,
    documentFontFamilyClasses,
    publicAnchorBridge,
    effectiveTablePositioning,
  } = dependencies;
  if (!measureContext) throw new Error('Body layout acquisition requires a measurement context');
  const physicalSection: SectionProps = {
    ...source.section,
    ...input.section.geometry,
    textDirection: input.section.textDirection,
    vAlign: input.section.verticalAlignment,
  };
  const section = isVerticalTextDirection(physicalSection.textDirection)
    ? verticalLayoutSection(physicalSection, input.section.nativeSectionFlow)
    : physicalSection;
  const state = buildMeasureState(
    measureContext,
    section,
    documentFontFamilyClasses,
    source.documentLayoutSettings,
    resolvedLocalFonts,
    services,
    options,
    input.section.nativeSectionFlow,
  );
  // Markup view only: resolve tracked-change author colours once per
  // session from the main story's document run order (first-appearance
  // policy, layout/track-changes.ts). The default final-view variant
  // never builds or carries this.
  if (options.showTrackedChanges === true) {
    state.revisionAuthorColor = createRevisionAuthorColorResolver(source.blocks.body);
  }
  const sourceFootnotes = source.blocks.footnotes;
  const sourceEndnotes = source.blocks.endnotes;
  const noteSettings = source.bodyLayoutInput.noteLayoutSettings;
  state.noteNumbering = {
    footnote: noteSettings?.footnoteNumbering ?? { format: 'decimal', start: 1 },
    endnote: noteSettings?.endnoteNumbering ?? { format: 'decimal', start: 1 },
  };
  const footnotesById = indexNotes(sourceFootnotes);
  state.noteNumbers = new Map([
    ...[
      ...buildNoteNumberMap(
        sourceFootnotes,
        noteReferenceIdsInDocumentOrder(source.blocks.body, 'footnote'),
        new Set(noteReferenceIdsInDocumentOrder(source.blocks.body, 'footnote', true)),
      ),
    ].map(([id, number]) => [`footnote:${id}`, number] as const),
    ...[
      ...buildNoteNumberMap(
        sourceEndnotes,
        noteReferenceIdsInDocumentOrder(source.blocks.body, 'endnote'),
        new Set(noteReferenceIdsInDocumentOrder(source.blocks.body, 'endnote', true)),
      ),
    ].map(([id, number]) => [`endnote:${id}`, number] as const),
  ]);
  let location = input.initialLocation;
  const pageRegistryFlowDomainId = (pageIndex: number) => `body:page:${pageIndex}:registry`;
  let floatRegistry: FloatRegistrySnapshotPt = Object.freeze({
    coordinateSpace: 'logical-page-points' as const,
    flowDomainId: pageRegistryFlowDomainId(location.pageIndex),
    entries: Object.freeze([]) as readonly FloatRegistryEntryPt[],
    nextParagraphId: 0,
  });
  let drawingCollisionRegistry: DrawingMLCollisionRegistrySnapshotPt =
    createDrawingMLCollisionRegistry(
      pageRegistryFlowDomainId(location.pageIndex),
      'logical-page-points',
    );
  const sessionState = { location, floatRegistry, drawingCollisionRegistry };
  setBodyAcquisitionLocation(services, state, sessionState, sessionState.location);
  const endnotesById = indexNotes(source.blocks.endnotes);
  // Trial misses debit the pagination-scoped miss budget and full
  // continued-note misses the pagination-scoped footnote work budget; both are
  // owned by the execution's service view, so neither resets with this
  // session or between convergence passes.
  const storyLayoutCache = createStoryLayoutCache(() => paragraphAcquisitionCacheOf(services)?.noteMiss());
  const storyAcquisitionContext: BodyStoryAcquisitionContext = {
    source,
    state,
    services,
    storyLayoutCache,
    noteSourceReuse: new WeakMap(),
    publicAnchorBridge,
  };
  state.acquireCompleteTextBoxStory = (owner, request) =>
    acquireCompleteBodyTextBoxStory(owner, storyAcquisitionContext, request);
  const sessionDependencies = {
    source,
    dependencies,
    services,
    state,
    sessionState,
    footnotesById,
    endnotesById,
    storyAcquisitionContext,
    pageRegistryFlowDomainId,
    measureContext,
    publicAnchorBridge,
    effectiveTablePositioning,
  };

  const session: BodyLayoutSession = {
    hasPaginationFields: source.hasPaginationFields,
    measureParagraph(
      request: Parameters<NonNullable<BodyLayoutSession['measureParagraph']>>[0],
    ): ReturnType<NonNullable<BodyLayoutSession['measureParagraph']>> {
      return measureBodyParagraphEntry(sessionDependencies, request);
    },
    measureTable(
      request: Parameters<NonNullable<BodyLayoutSession['measureTable']>>[0],
    ): ReturnType<NonNullable<BodyLayoutSession['measureTable']>> {
      return measureBodyTableEntry(sessionDependencies, request, {
        setBodyAcquisitionLocation, sourceElement, computeTablePtLayout, computeAdjacentTablePtLayouts,
        ordinaryAcquisitionInputForAdjacentGroup, reacquireBodyTableBlock,
      });
    },
    layoutStory: (request) => acquireBodyStoryLayout(storyAcquisitionContext, request),
    layoutNotes(
      request: Parameters<NonNullable<BodyLayoutSession['layoutNotes']>>[0],
    ): ReturnType<NonNullable<BodyLayoutSession['layoutNotes']>> {
      return layoutBodyNotes(sessionDependencies, request);
    },
    measureFollowingBlock(
      request: Parameters<NonNullable<BodyLayoutSession['measureFollowingBlock']>>[0],
    ): ReturnType<NonNullable<BodyLayoutSession['measureFollowingBlock']>> {
      return measureFollowingBodyBlock(sessionDependencies, request);
    },
    prescanPageAnchors(
      request: Parameters<NonNullable<BodyLayoutSession['prescanPageAnchors']>>[0],
    ): ReturnType<NonNullable<BodyLayoutSession['prescanPageAnchors']>> {
      return prescanBodyPageAnchors(sessionDependencies, request);
    },
    measureLineNumberGlyph(
      text: Parameters<NonNullable<BodyLayoutSession['measureLineNumberGlyph']>>[0],
    ): ReturnType<NonNullable<BodyLayoutSession['measureLineNumberGlyph']>> {
      return measureBodyLineNumberGlyph(sessionDependencies, text);
    },
    resetPageAcquisition(
      next: Parameters<NonNullable<BodyLayoutSession['resetPageAcquisition']>>[0],
    ): ReturnType<NonNullable<BodyLayoutSession['resetPageAcquisition']>> {
      return resetBodyPageAcquisition(sessionDependencies, next);
    },
    moveAcquisitionCursor: (next) =>
      setBodyAcquisitionLocation(services, state, sessionState, next),
    flowRegistrySnapshot(): ReturnType<NonNullable<BodyLayoutSession['flowRegistrySnapshot']>> {
      return bodyFlowRegistrySnapshot(sessionDependencies);
    },
    commitFlowRegistryDelta(
      delta: Parameters<NonNullable<BodyLayoutSession['commitFlowRegistryDelta']>>[0],
    ): ReturnType<NonNullable<BodyLayoutSession['commitFlowRegistryDelta']>> {
      return commitBodyFlowRegistryDelta(sessionDependencies, delta);
    },
  };
  return Object.freeze(session);
}

function buildConcreteBodyLayoutKernel(dependencies: ConcreteBodyKernelContext): BodyLayoutKernel {
  return Object.freeze({
    openBodyLayoutSession(
      input: import('./body-layout-kernel.js').BodyLayoutSessionInput,
      services: LayoutServices,
      options: LayoutOptions,
    ) {
      return openConcreteBodyLayoutSession(dependencies, input, services, options);
    },
  });
}

function paraGrid(para: ParagraphLayoutSource, state: BodyMeasurementContext): DocGridCtx {
  return gridForParagraphContext(
    state,
    resolveStateParagraphLayoutContext(state, para),
  );
}

/** Resolve column widths once and retain the acquired table for one top-level
 * body occurrence (read back through retainedTableRecord). A table that
 * continues onto another flow region is measured once per region; while the
 * inline extent is unchanged the retained acquisition is identical for every
 * destination page (page-varying geometry is excluded by
 * retainedTableAcquisitionIsReusableAcrossPages), so reuse the record in
 * constant time instead of re-walking, re-laying out or re-projecting every
 * row — that made pagination cost O(flow-regions × rows), and a table paged
 * one owner segment per request would pay it per segment. */
function computeTablePtLayout(
  state: BodyAcquisitionState,
  table: TableLayoutSource,
  contentWPt: number,
  sourceIndex: number,
  prepared?: Readonly<{ columns: readonly number[]; decision: LogicalTableDecision; member: TableMemberDecision }>,
): void {
  const prior = state.retainedTablesBySourceIndex.get(sourceIndex);
  if (prior?.contentWidthPt === contentWPt && prior.reusableAcrossPages) return;
  const decision = prepared?.decision ?? singleTableDecision(table, contentWPt, state);
  const member = prepared?.member ?? decision.logical;
  const colWidthsPt = prepared ? [...prepared.columns] : resolveColumnWidths(table, contentWPt, state, decision, member);
  const dependencies = state.retainedTableAcquisition;
  const acquired = acquireRetainedTable(
    table,
    colWidthsPt,
    contentWPt,
    state,
    [sourceIndex],
    dependencies,
    member,
    { kind: 'story-root', story: 'body' },
  );
  // Split rows are page-local acquisitions, but an unchanged inline extent
  // retains one authoritative track vector for the table's full occurrence.
  const retained = prior?.contentWidthPt === contentWPt
    ? Object.freeze({
        ...acquired,
        layout: Object.freeze({
          ...acquired.layout,
          columnWidthsPt: prior.acquisition.layout.columnWidthsPt,
        }),
      })
    : acquired;
  state.retainedTablesBySourceIndex.set(sourceIndex, Object.freeze({
    sourceIndex,
    acquisition: retained,
    contentWidthPt: contentWPt,
    reusableAcrossPages: retainedTableAcquisitionIsReusableAcrossPages(retained),
    anchorYPt: state.y,
  }));
}

/** Acquire columns after parser-owned §17.4.37 grouping. Both pagination and
 * keep-next lookahead use this path, so no authored member can independently
 * discard a leading track or restart the logical first-row margin anchor. */
function computeAdjacentTablePtLayouts(
  state: BodyAcquisitionState,
  members: readonly Readonly<{ table: TableLayoutSource; sourceIndex: number }>[],
  contentWPt: number,
): readonly RetainedTableAcquisition[] {
  const prior = members.map(({ sourceIndex }) => state.retainedTablesBySourceIndex.get(sourceIndex));
  if (prior.every((record) => record?.contentWidthPt === contentWPt && record.reusableAcrossPages)) {
    return prior.map((record) => record!.acquisition);
  }
  const first = members[0]!.table;
  const commonFrame = members.every(({ table }) => table.bidiVisual === first.bidiVisual
    && table.jc === first.jc && table.tblInd === first.tblInd);
  const commonGrid = members.every(({ table }) => table.colWidths.length === first.colWidths.length
    && table.colWidths.every((width, i) => width === first.colWidths[i]));
  // §17.4.37: resolve measured table-wide properties only after concatenating
  // logical rows. Keep each row's resolved margins/lexical constraints, but
  // WORD_FIRST_ROW_TABLE_EXCEPTION_SCOPE selects width/layout/indent from
  // the first logical row, which also owns the leading cell-margin anchor.
  // Differing authored frames/grids retain the established union policy.
  const sources = members.map(({ table }) => state.acquisitionInputs.tableSourceAcquisitionInput(table));
  const logicalTable = Object.freeze({
    ...first, rows: Object.freeze(members.flatMap(({ table }) => table.rows)),
  });
  const logicalRows = Object.freeze(sources.flatMap((source) => source.format.rows));
  const logicalSource: TableSourceAcquisitionInput = Object.freeze({
    semantic: Object.freeze({ ...sources[0]!.semantic, rows: Object.freeze(sources.flatMap((source) => source.semantic.rows)) }),
    lexical: Object.freeze({ ...sources[0]!.lexical, rows: Object.freeze(sources.flatMap((source) => source.lexical.rows)) }),
    format: Object.freeze({
      ...sources[0]!.format, rows: logicalRows, firstRowException: logicalRows[0]?.exception ?? null,
    }),
  });
  const decision = decideLogicalTable(logicalTable, logicalSource,
    members.map(({ table }, i) => ({ table, source: sources[i]! })), contentWPt,
    tableDecisionEnvironment(state), commonFrame, commonGrid);
  const logicalColumns = decision.usesLogicalProperties
    ? resolveColumnWidths(logicalTable, contentWPt, state, decision, decision.logical) : null;
  return members.map(({ table, sourceIndex }, i) => {
    const member = decision.members[i]!;
    computeTablePtLayout(state, table, contentWPt, sourceIndex, {
      columns: logicalColumns ?? resolveColumnWidths(table, contentWPt, state, decision, member),
      decision, member,
    });
    return retainedTableRecord(state, sourceIndex).acquisition;
  });
}

/**
 * Acquire the parser/style facts and intrinsic content constraints required by
 * ECMA-376 §17.18.87, then resolve the shared table grid. `tblGrid` is the
 * initial grid, not an oracle containing an application's previous result;
 * authored `tblW`, `tcW`, `wBefore`, and `wAfter` remain active constraints.
 * Exported for the table-layout integration tests.
 */
function tableDecisionEnvironment(state: BodyMeasurementContext) {
  return {
    mode: state.layoutSettings.compat.compatibilityMode,
    story: state.storyContext?.story,
    topLevel: state.storyContext?.containers.length === 0,
    singleColumn: state.sectionLayout?.columns.length === 1,
    textDirection: state.sectionLayout.textDirection,
    inTableCell: state.storyContext?.containers.some((container) => container.kind === 'tableCell') ?? false,
    contentX: state.contentX,
    pageWidth: state.pageWidth,
  };
}

function singleTableDecision(table: TableLayoutSource, contentWPt: number, state: BodyMeasurementContext): LogicalTableDecision {
  const source = state.acquisitionInputs.tableSourceAcquisitionInput(table);
  return decideLogicalTable(table, source, [{ table, source }], contentWPt, tableDecisionEnvironment(state));
}

function resolveColumnWidths(
  table: TableLayoutSource,
  contentWPt: number,
  state: BodyMeasurementContext,
  decision = singleTableDecision(table, contentWPt, state),
  member = decision.logical,
): number[] {
  const input = acquireTableColumnInput(table, contentWPt, state, member);
  const normalized = decision.dropsUnusedLeadingGrid ? {
    ...input,
    gridWidthsPt: input.gridWidthsPt.map((width, i) => i === 0 ? 0 : width),
    gridWidthKeys: input.gridWidthKeys?.map((key, i) => i === 0 ? null : key),
    rows: input.rows.map((row) => ({ ...row, before: null })),
  } : input;
  return [...resolveTableColumnWidths(normalized)];
}

function acquireTableColumnInput(
  table: TableLayoutSource,
  contentWPt: number,
  state: BodyMeasurementContext,
  decision: TableMemberDecision,
): TableColumnLayoutInput {
  const format = decision.source.format;
  const { maximumTableWidthPt, isFixedNestedTable,
    usesLeadingIndentBand, forcedFitOuterMarginsPt } = decision;
  const intrinsicWidthsForTable = (
    owner: TableLayoutSource,
    ownerDecision = decision,
  ): ((cell: DeepReadonly<DocTableCell>, preferredWidth: TablePreferredWidthConstraint | null) => ReturnType<typeof measureTableCellIntrinsicWidths>) => {
    const ownerFormat = ownerDecision.source.format;
    const ownerLayout = ownerDecision.effectiveLayout;
    const ownerMargins = new WeakMap<object, Readonly<{ left: number; right: number }>>();
    owner.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
      const acquired = ownerFormat.rows[rowIndex]?.cells[cellIndex]?.marginsPt;
      ownerMargins.set(cell, acquired ?? effCellMargins(cell, owner));
    }));
    return (cell, preferredWidth) => measureTableCellIntrinsicWidths(
      cell,
      ownerMargins.get(cell as object) ?? effCellMargins(cell, owner),
      {
        paragraph: (paragraph) => {
          const baseContext = resolveParagraphLayoutContext(
            state.layoutSettings,
            state.sectionLayout,
            state.storyContext ?? BODY_STORY_CONTEXT,
            paragraph,
          );
          const markerInput = paragraph.numbering
            ? state.acquisitionInputs.numberingMarkerShapeInput(
                paragraph.numbering,
                getDefaultFontSize(paragraph),
              )
            : undefined;
          const context = applyNumberingBodyOffset(baseContext, {
            numbering: paragraph.numbering,
            ...(markerInput ? { markerInput } : {}),
            authoredFirstIndentPt: paragraph.indentFirst,
            tabStops: paragraph.tabStops,
            defaultTabPt: state.defaultTabPt,
            service: state.layoutServices?.text,
            clusterGeometry: false,
          });
          const numbering = context.numberingMarkerGeometry
            ?? (paragraph.numbering && markerInput && state.layoutServices?.text
              ? resolveNumberingMarkerGeometry(paragraph.numbering, markerInput, {
                  authoredFirstIndentPt: paragraph.indentFirst,
                  physicalIndentLeftPt: context.physicalIndentLeftPt,
                  tabStops: paragraph.tabStops,
                  defaultTabPt: state.defaultTabPt,
                }, state.layoutServices.text, false)
              : undefined);
          const intrinsic = measureParagraphIntrinsicWidths(
            paragraph,
            context,
            // §17.18.87 defines maximum content width without soft wrapping.
            // Capping it at the band loses the relative demand of two growing
            // columns (notably the 5:1 control). Only the measured auto-column
            // growth path needs this uncapped interval; preferred and nested
            // paths retain their existing measurement ceiling.
            usesLeadingIndentBand && owner === table && ownerLayout !== 'fixed'
              && preferredWidth === null
              ? Number.MAX_SAFE_INTEGER
              : contentWPt,
            { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses },
            paragraphMeasurementEnvironment(state),
            numbering,
            { preserveWhitespaceOnlyContent: true },
          );
          if (cell.noWrap !== true || preferredWidth?.kind === 'dxa' || ownerLayout === 'fixed') return intrinsic;
          // The nonbreaking text interval is independent of first-line
          // positioning. Probe without a line-width ceiling: the ordinary
          // maxWidthPt is capped at contentWPt, which is too small for an
          // AutoFit noWrap minimum when tcW is omitted (§17.4.29, §17.4.71).
          // Ordinary min/max still use the authored indent and width limit.
          const unpositioned = measureParagraphIntrinsicWidths(
            paragraph,
            { ...context, firstIndentPt: 0 },
            Number.MAX_SAFE_INTEGER,
            { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses },
            paragraphMeasurementEnvironment(state),
            numbering,
            { preserveWhitespaceOnlyContent: true },
          );
          return { ...intrinsic, noWrapWidthPt: unpositioned.maxWidthPt };
        },
        nestedTable: (nested) => {
          const source = state.acquisitionInputs.tableSourceAcquisitionInput(nested);
          const nestedDecision = decideLogicalTable(nested, source, [{ table: nested, source }], contentWPt,
            { ...tableDecisionEnvironment(state), topLevel: false, inTableCell: true });
          const intrinsic = intrinsicWidthsForTable(nested, nestedDecision.logical);
          return measureTableIntrinsicWidths(projectTableColumnLayoutInput(source, contentWPt,
            (row, cellIndex, preferred) => intrinsic(nested.rows[row]!.cells[cellIndex]!, preferred), contentWPt));
        },
      },
      ownerLayout,
      preferredWidth,
    );
  };
  const maximumWidthPt = isFixedNestedTable
    // ECMA-376 §17.18.87: the containing cell is not an implicit tblW.
    ? null
    : format.ordinaryFlow ? maximumTableWidthPt : Math.max(contentWPt, state.pageWidth);
  const intrinsicWidths = intrinsicWidthsForTable(table);
  const columnInput = projectTableColumnLayoutInput(decision.source, contentWPt,
    (rowIndex, cellIndex, preferred) => intrinsicWidths(table.rows[rowIndex]!.cells[cellIndex]!, preferred),
    maximumWidthPt);
  return {
    ...columnInput,
    outerMarginAllowancePt: forcedFitOuterMarginsPt,
    growUnpreferredColumns: usesLeadingIndentBand,
  };
}

// ===== Text frames & drop caps (ECMA-376 §17.3.1.11) =====

/**
 * One point-space line height of the anchor (following non-frame) paragraph,
 * used to
 * size a drop cap by `lines` (§17.3.1.11). The drop cap height equals
 * `lines` × this. Scans `elements` after the frame element for the first
 * non-frame paragraph; falls back to the frame paragraph's own single-line
 * height when none follows (a degenerate trailing frame).
 */
// Historical name retained while this renderer helper remains on the C3
// migration inventory; the returned value is now canonical points.
function frameAnchorLineHeightPx(
  elements: readonly LayoutStoryBlock[],
  frameEl: LayoutParagraphBlock,
  state: BodyMeasurementContext,
): number {
  const start = elements.indexOf(frameEl);
  for (let j = start + 1; j < elements.length; j++) {
    const e = elements[j];
    if (e.type !== 'paragraph') continue;
    const p = e;
    if (p.framePr) continue; // adjacent frame paragraphs are part of the frame
    return paragraphMarkLineHeight(
      p,
      1,
      paraGrid(p, state),
      resolveBodyParagraphLayoutContext(state, p).hasRuby,
      state.docEastAsian,
      state.ctx,
      state.fontFamilyClasses,
      p.lineSpacing,
      state.resolvedLocalFonts,
      state.layoutServices?.text,
      state.acquisitionInputs.paragraphMarkShapeInput(p),
      state.layoutSettings.compat.useFeLayout,
    );
  }
  const fp = frameEl;
  return paragraphMarkLineHeight(
    fp,
    1,
    paraGrid(fp, state),
    resolveBodyParagraphLayoutContext(state, fp).hasRuby,
    state.docEastAsian,
    state.ctx,
    state.fontFamilyClasses,
    fp.lineSpacing,
    state.resolvedLocalFonts,
    state.layoutServices?.text,
    state.acquisitionInputs.paragraphMarkShapeInput(fp),
    state.layoutSettings.compat.useFeLayout,
  );
}

/** Border adjacency inside one story-local frame group (ECMA-376 §17.3.1.11). */
function storyFrameBorderEdges(
  group: BodyFrameGroup<LayoutParagraphBlock>,
  paragraph: LayoutParagraphBlock,
): ReturnType<typeof resolveParagraphBorderEdges> {
  const index = group.members.indexOf(paragraph);
  return resolveParagraphBorderEdges(
    group.members[index - 1] ?? null,
    paragraph,
    group.members[index + 1] ?? null,
    true,
  );
}

/** Resolve a prepared body frame group and attach its retained member layouts. */
/** How a frame box reports its retained group, and the story context when
 * the frame lives in a header/footer story rather than the body. */
type FrameBoxAcquisitionOptions = Readonly<{
  onAcquired?: (acquired: ReturnType<typeof acquireRetainedFrameGroup>) => void;
  story?: Readonly<{ story: SourceRef['story']; storyInstance: string }>;
  borderEdgesFor?: (
    paragraph: LayoutParagraphBlock,
  ) => ReturnType<typeof bodyParagraphBorderEdgesFor>;
}>;

function resolveFrameBox(
  para: ParagraphLayoutSource,
  group: BodyFrameGroup<LayoutParagraphBlock>,
  state: BodyAcquisitionState,
  anchorLineHPt: number,
  acquisition: FrameBoxAcquisitionOptions,
): FrameBox {
  const { onAcquired, story, borderEdgesFor = bodyParagraphBorderEdgesFor } = acquisition;
  const measurer = { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses };
  const environment = paragraphMeasurementEnvironment(state);
  const borderEdges = group.members.map(borderEdgesFor);
  const horizontalBand = frameXContainer(group.framePr.hAnchor, state);
  const pointPlacement = {
    contentXPt: state.contentX,
    contentWidthPt: state.contentW,
    pageHeightPt: state.pageH,
    yPt: state.y,
    anchorLineHeightPt: anchorLineHPt,
  };
  const acquired = acquireRetainedFrameGroup(group, {
    contexts: group.members.map((paragraph) =>
      resolveBodyParagraphLayoutContext(state, paragraph)),
    inputs: group.members,
    borderEdges,
    borderExtentsPt: group.members.map((paragraph, index) =>
      borderEdges[index]?.bottom === 'none' ? 0 : bottomBorderExtentPt(paragraph.borders)),
    measurer,
    environment,
    containerShading: state.containerShading,
    maximumWidthPt: Math.max(0, horizontalBand.right - horizontalBand.left),
    acquisitionSession: state,
    ...(story ? { story } : {}),
    placementSignature: [
      pointPlacement.contentXPt,
      pointPlacement.contentWidthPt,
      pointPlacement.pageHeightPt,
      pointPlacement.yPt,
      pointPlacement.anchorLineHeightPt,
      state.pageWidth,
      state.marginLeft,
      state.marginRight,
      state.marginTop,
      state.marginBottom,
    ].join('|'),
    place: (contentWidthPt, contentHeightPt) => {
      const box = computeFrameBox(
        group.framePr,
        state,
        pointPlacement.yPt,
        contentWidthPt,
        contentHeightPt,
        pointPlacement.anchorLineHeightPt,
      );
      return Object.freeze({
        bounds: Object.freeze({
          xPt: box.x,
          yPt: box.y,
          widthPt: box.w,
          heightPt: box.h,
        }),
        exclusionBounds: Object.freeze({
          xPt: box.exLeft,
          yPt: box.exTop,
          widthPt: box.exRight - box.exLeft,
          heightPt: box.exBottom - box.exTop,
        }),
      });
    },
    anchorFrames: bodyAnchorReferenceFrames(state),
  });
  onAcquired?.(acquired);
  const box: FrameBox = {
    x: acquired.box.bounds.xPt,
    y: acquired.box.bounds.yPt,
    w: acquired.box.bounds.widthPt,
    h: acquired.box.bounds.heightPt,
    exLeft: acquired.box.exclusionBounds.xPt,
    exTop: acquired.box.exclusionBounds.yPt,
    exRight: acquired.box.exclusionBounds.xPt + acquired.box.exclusionBounds.widthPt,
    exBottom: acquired.box.exclusionBounds.yPt + acquired.box.exclusionBounds.heightPt,
    registerExclusion: true,
    exclusionId: acquired.box.exclusionId,
  };
  return para === group.owner
    ? box
    : { ...box, registerExclusion: false };
}

/**
 * Resolve an anchored shape's point-space bounding box {x,y,w,h}. Retained
 * drawing acquisition and float registration share this geometry so the
 * exclusion band matches the painted box.
 *
 * Mirrors the renderer's sizing: sizeRelH/sizeRelV (ECMA-376 §20.4.2.18)
 * override the static extent, and a wgp child scales by the group ratio with its
 * within-group offset scaled in step; resolveAnchorX/Y then place the box. `w`/`h`
 * may be 0/negative for degenerate line presets; a wrap shape with no area
 * registers no float.
 */
function resolveShapeBox(
  shape: DeepReadonly<ShapeRun>,
  state: AnchorFloatRegistrationState,
  paragraphTopPt: number,
): { x: number; y: number; w: number; h: number } {
  // ECMA-376 §17.6.20 + §20.4.3.x (issue #988 batch-3 adjudication ②): on a
  // vertical (tbRl) page an anchored shape's positionH/V resolve against the
  // PHYSICAL (un-rotated) page — the drawing layer is independent of the
  // section text direction, exactly like the image path (resolveAnchorBox).
  // Resolve in the physical frame, then project into the swapped logical
  // layout frame (w↔h swapped) so the float-exclusion band and the flow all
  // share one geometry. Under `word-vertical-section-physical-drawing-layer`,
  // a `paragraph`/`line`-relative positionV anchors from the PHYSICAL TOP of
  // the anchor paragraph's COLUMN. That physical y is the column
  // band's logical x start (`state.contentX`, since physical y = logical x
  // under the +90° page paint), NOT the paragraph's logical flow
  // `paragraphTopPt`, which lies on the column-progression axis.
  if (state.verticalPhys) {
    const phys = resolveShapeBox(
      shape,
      verticalPhysicalContentState(state),
      physicalColumnTopPt(state, state.verticalPhys),
    );
    return physicalAnchorBoxToLogical(state.verticalPhys, phys.x, phys.y, phys.w, phys.h);
  }
  // ECMA-376 §20.4.2.18: when wp14:sizeRelH/sizeRelV is present it overrides
  // the static wp:extent for that axis. The size is `relativeFrom` container
  // size × pct.
  //
  // For a wgp group with sizeRelH, the parent group resizes and every child
  // shape scales proportionally — so a grouped child's effective width is
  // `original_width × (new_group_w / old_group_w)`, and its within-group
  // offset (carried by anchorXPt) scales by the same ratio. Standalone
  // shapes simply take `container × pct` as their width.
  let w = shape.widthPt;
  let h = shape.heightPt;
  let offsetXPt = shape.anchorXPt;
  let offsetYPt = shape.anchorYPt;
  let alignWidthPt = shape.groupWidthPt ?? null;
  let alignHeightPt = shape.groupHeightPt ?? null;
  if (shape.widthPct != null) {
    const c = xContainer(shape.widthRelativeFrom, false, state);
    const newSizePt = (c.end - c.start) * shape.widthPct;
    if (shape.groupWidthPt != null && shape.groupWidthPt > 0) {
      const ratio = newSizePt / shape.groupWidthPt;
      w = shape.widthPt * ratio;
      offsetXPt = shape.anchorXPt * ratio;
    } else {
      w = newSizePt;
    }
    alignWidthPt = newSizePt;
  }
  if (shape.heightPct != null) {
    const c = yContainer(shape.heightRelativeFrom, false, paragraphTopPt, state);
    const newSizePt = (c.end - c.start) * shape.heightPct;
    if (shape.groupHeightPt != null && shape.groupHeightPt > 0) {
      const ratio = newSizePt / shape.groupHeightPt;
      h = shape.heightPt * ratio;
      offsetYPt = shape.anchorYPt * ratio;
    } else {
      h = newSizePt;
    }
    alignHeightPt = newSizePt;
  }
  const x = resolveAnchorX(
    shape.anchorXAlign, shape.anchorXFromMargin, offsetXPt, w, state,
    shape.anchorXRelativeFrom, shape.pctPosH, alignWidthPt,
  );
  const y = resolveAnchorY(
    shape.anchorYAlign, shape.anchorYFromPara, offsetYPt, h, paragraphTopPt, state,
    shape.anchorYRelativeFrom, shape.pctPosV, alignHeightPt,
  );
  return { x, y, w, h };
}

/**
 * Resolve an anchor image's point-space box origin and dist* padding, shared
 * by legacy float registration and the canonical anchor acquisition bridge.
 *
 * X: margin-relative offsets add section.marginLeft (ECMA-376 §20.4.3.4
 * relativeFrom="margin"); otherwise anchorXPt is already page-absolute.
 * Y: paragraph-relative offsets add `paraBaseY`; otherwise page-absolute. The
 * caller supplies `paraBaseY` = the paragraph's pre-spaceBefore TOP for ALL
 * paragraph-relative floats — wrap and wrapNone alike (ECMA-376 §20.4.3.5: a
 * `positionV relativeFrom="paragraph"` float is positioned relative to the
 * paragraph that contains the anchor, i.e. its top edge before spaceBefore).
 * Page-level floats pass 0 (resolveAnchorY ignores paraBaseY for them). This is
 * the box origin BEFORE the typed float placement policy displaces it.
 *
 * Exported under a `_test` alias for the anchor-image relativeFrom wiring test
 * (the public renderer entry points consume the box internally; pin the
 * positionH/V → xContainer/yContainer plumbing at this seam).
 */
const __test_resolveAnchorBox = (
  img: ImageRun,
  state: AnchorFloatRegistrationState,
  paraBaseY: number,
): { x: number; y: number; w: number; h: number; dl: number; dr: number; dt: number; db: number } =>
  resolveAnchorBox(img, state, paraBaseY);

/** Exported for the vertical shape-anchor test (ECMA-376 §17.6.20 + §20.4.3.x,
 *  issue #988 ②): pins the physical-page resolution (and logical projection) of
 *  an anchored SHAPE's positionH/V on a vertical (tbRl) page. */
const __test_resolveShapeBox = (
  shape: ShapeRun,
  state: AnchorFloatRegistrationState,
  paragraphTopPt: number,
): { x: number; y: number; w: number; h: number } =>
  resolveShapeBox(shape, state, paragraphTopPt);

/** Exported for the vertical header/footer test (ECMA-376 §17.6.20 + §17.10.1,
 *  issue #988): pins the inverse-of-`verticalLayoutSection` page/margin mapping a
 *  vertical section's HORIZONTAL header/footer are laid out in. */
const __test_physicalLayoutSection = (logical: SectionProps): SectionProps =>
  physicalLayoutSection(logical);
const __test_verticalLayoutSection = (phys: SectionProps): SectionProps =>
  verticalLayoutSection(phys);

type AnchorScanBlock =
  | (ParagraphLayoutSource & Readonly<{ type: 'paragraph' }>)
  | DeepReadonly<Exclude<BodyElement, { type: 'paragraph' }>>
  | Extract<LayoutStoryBlock, { type: 'unsupportedTextBoxBlock' }>;

/** Exported for the page-anchor pre-scan test (ECMA-376 §20.4.3.2/§20.4.3.5):
 *  drives {@link preRegisterPageFloats} from a unit test against a stub
 *  AnchorFloatRegistrationState so we can pin which paragraphs get pre-registered and that
 *  duplicate calls are idempotent. */
const __test_preRegisterPageFloats = (
  body: readonly AnchorScanBlock[],
  startIdx: number,
  state: AnchorFloatRegistrationState,
): void => preRegisterPageFloats(body, startIdx, state);

/** ECMA-376 §17.6.20 + §20.4.3.x — an acquisition-state view whose page/margin geometry
 *  is the PHYSICAL (un-rotated) page, used to resolve a DrawingML anchor's
 *  `<wp:positionH/V>` against the physical page for a vertical (tbRl) section
 *  under `word-vertical-section-physical-drawing-layer`. Only
 *  the geometry fields `xContainer`/`yContainer`/`resolveAnchorX`/`resolveAnchorY`
 *  read are overridden (page size, margins, and `pageH`); everything else is
 *  the live logical state. Callers map the resolved physical box back into the
 *  logical layout frame with {@link physicalToLogicalAnchorBox}. */
/** Physical y of the anchor paragraph's column top: the logical inline start
 * in the clockwise frame (physical y = logical x), and the logical inline end
 * in a native counter-clockwise frame (physical y = page height - logical x).
 * The native case is the same generic library policy transformed through its
 * own frame. No comparison with a Word-produced reference has established
 * this placement. */
function physicalColumnTopPt(
  state: AnchorFloatRegistrationState,
  frame: PhysicalAnchorFrame,
): number {
  return frame.nativeSectionFlow == null
    ? state.contentX
    : frame.pageHeight - (state.contentX + state.contentW);
}

/** Project a box resolved on the upright physical page into the section's
 * logical frame. The Transitional vertical frame keeps its established
 * clockwise projection; a native BtoT frame applies the inverse of its
 * counter-clockwise matrix (logical x = page height - physical y, logical
 * y = physical x). */
function physicalAnchorBoxToLogical(
  frame: PhysicalAnchorFrame,
  x: number,
  y: number,
  w: number,
  h: number,
): { x: number; y: number; w: number; h: number } {
  if (frame.nativeSectionFlow == null) {
    return physicalToLogicalAnchorBox(x, y, w, h, frame.physicalPageWidthPt);
  }
  return { x: frame.pageHeight - (y + h), y: x, w: h, h: w };
}

/** Relabel physical dist* padding with the logical edges of the box. Clockwise:
 * physical top/bottom are logical left/right and physical right/left are
 * logical top/bottom. Native BtoT: physical bottom/top are logical left/right
 * and physical left/right are logical top/bottom. */
function physicalDistToLogical(
  frame: PhysicalAnchorFrame,
  dist: Readonly<{ dl: number; dr: number; dt: number; db: number }>,
): { dl: number; dr: number; dt: number; db: number } {
  return frame.nativeSectionFlow == null
    ? { dl: dist.dt, dr: dist.db, dt: dist.dr, db: dist.dl }
    : { dl: dist.db, dr: dist.dt, dt: dist.dl, db: dist.dr };
}

function physicalAnchorState(
  state: AnchorFloatRegistrationState,
): AnchorFloatRegistrationState {
  const p = state.verticalPhys;
  if (!p) return state;
  return {
    ...state,
    pageWidth: p.pageWidth,
    marginLeft: p.marginLeft,
    marginRight: p.marginRight,
    marginTop: p.marginTop,
    marginBottom: p.marginBottom,
    pageH: p.pageHeight,
  };
}

/** ECMA-376 §17.6.20 + §20.4.3.x (issue #988 ②/④) — an acquisition-state view whose
 *  geometry AND text flags are PHYSICAL, for content that stays UPRIGHT inside a
 *  vertical (tbRl) section: anchored shapes and block tables. Under
 *  `word-vertical-section-physical-drawing-layer`, these acquire against the
 *  un-rotated physical page — cell/label text is
 *  horizontal — so on top of {@link physicalAnchorState}'s page/margin un-swap
 *  this view also re-points the content band at the physical margins and clears
 *  the vertical flags (no per-glyph counter-rotation, no +90° text-layer
 *  transform, `resolveShapeBox`/`resolveAnchorBox` take their horizontal path).
 *  `floats` is fresh: the live float set is in LOGICAL flow coordinates and must
 *  not leak into a physical-frame layout (and vice-versa). Clearing
 *  `verticalPhys` also drops any native section frame; the anchor geometry
 *  consumers of this view read only its physical page facts. */
/** Owner state for table cell content. A body table placed upright in the
 * physical page (identity paint root) owns its cells' physical frame: like an
 * upright text box, the cell content is acquired horizontally in the physical
 * page box, with no section counter-turn on its graphics, no vertical glyph
 * flags and no native section frame ({@link verticalPhysicalContentState}).
 * Every other table keeps its section-logical owner, and authored cell text
 * directions are applied by the table itself either way. */
function tableCellOwnerState(state: BodyAcquisitionState): BodyAcquisitionState {
  if (!state.uprightPhysicalTable) return state;
  const { nativeSectionFlow, ...section } = state.sectionLayout;
  return {
    ...(verticalPhysicalContentState(state) as BodyAcquisitionState),
    uprightPhysicalTable: false,
    sectionLayout: {
      ...section,
      geometry: physicalSectionGeometry(state.sectionLayout.geometry, nativeSectionFlow),
      textDirection: 'lrTb',
    },
  };
}

function verticalPhysicalContentState(
  state: AnchorFloatRegistrationState,
): AnchorFloatRegistrationState {
  const p = state.verticalPhys;
  if (!p) return state;
  return {
    ...physicalAnchorState(state),
    contentX: p.marginLeft,
    contentW: p.pageWidth - p.marginLeft - p.marginRight,
    verticalCJK: false,
    verticalAllRotated: false,
    verticalPhys: undefined,
    floats: [],
  };
}

type AnchorBoxSource = Pick<ImageRun,
  | 'widthPt' | 'heightPt'
  | 'anchorXPt' | 'anchorYPt'
  | 'anchorXFromMargin' | 'anchorYFromPara'
  | 'anchorXAlign' | 'anchorYAlign'
  | 'anchorXRelativeFrom' | 'anchorYRelativeFrom'
  | 'distTop' | 'distBottom' | 'distLeft' | 'distRight'
>;

function resolveAnchorBox(
  img: AnchorBoxSource,
  state: AnchorFloatRegistrationState,
  paraBaseY: number,
): { x: number; y: number; w: number; h: number; dl: number; dr: number; dt: number; db: number } {
  const w = img.widthPt;
  const h = img.heightPt;
  const dl = img.distLeft ?? 0;
  const dr = img.distRight ?? 0;
  const dt = img.distTop ?? 0;
  const db = img.distBottom ?? 0;
  // ECMA-376 §20.4.3.1 wp:align — when positionH/V carry <wp:align>, the
  // renderer aligns the image within its relativeFrom container instead of
  // using the (discarded) posOffset. Mirrors resolveShapeBox (the ShapeRun
  // equivalent): we route X/Y through resolveAnchorX/Y with the image's own
  // box size as the align size. The raw §20.4.3.2/§20.4.3.5
  // `<wp:positionH/V>@relativeFrom` string (e.g. "margin", "topMargin") is
  // threaded through so xContainer/yContainer pick the correct container.
  // Without it a `relativeFrom="margin"` + `align="top"` image would degrade
  // to the page-relative top edge (Y=0 → inside the top margin). ImageRun
  // carries no pctPos/sizeRel, so those args remain null and the legacy boolean
  // anchorXFromMargin / anchorYFromPara hints still gate page-vs-margin when
  // no raw relativeFrom is present. When align is absent, resolveAnchorX/Y
  // fall back to the offset path.
  if (state.verticalPhys) {
    // `word-vertical-section-physical-drawing-layer`: resolve positionH/V in
    // the physical page frame independently of rotated text flow, resolve the
    // box there, then project it into the swapped logical layout frame. The
    // float-exclusion band and the retained upright painted image
    // therefore share one geometry. A `paragraph`/`line`-relative positionV
    // anchors from the physical top of the anchor paragraph's column. That y is
    // the column band's logical x start (`state.contentX`; physical y =
    // logical x under the +90° page paint) — NOT the logical flow `paraBaseY`,
    // which lies on the column-progression axis and would rotate the offset.
    const phys = physicalAnchorState(state);
    const px = resolveAnchorX(
      img.anchorXAlign, img.anchorXFromMargin ?? false, img.anchorXPt ?? 0, w, phys,
      img.anchorXRelativeFrom ?? null, null, null,
    );
    const py = resolveAnchorY(
      img.anchorYAlign, img.anchorYFromPara ?? false, img.anchorYPt ?? 0, h,
      physicalColumnTopPt(state, state.verticalPhys), phys,
      img.anchorYRelativeFrom ?? null, null, null,
    );
    const box = physicalAnchorBoxToLogical(state.verticalPhys, px, py, w, h);
    // Rotate the dist* padding one quarter-turn with the box. Symmetric
    // wrapSquare dist is common, but rotate the labels so asymmetric dist
    // stays correct.
    const logicalDist = physicalDistToLogical(state.verticalPhys, { dl, dr, dt, db });
    return { x: box.x, y: box.y, w: box.w, h: box.h, ...logicalDist };
  }
  const x = resolveAnchorX(
    img.anchorXAlign, img.anchorXFromMargin ?? false, img.anchorXPt ?? 0, w, state,
    img.anchorXRelativeFrom ?? null, null, null,
  );
  const y = resolveAnchorY(
    img.anchorYAlign, img.anchorYFromPara ?? false, img.anchorYPt ?? 0, h, paraBaseY, state,
    img.anchorYRelativeFrom ?? null, null, null,
  );
  return { x, y, w, h, dl, dr, dt, db };
}

/** Register float exclusions from a paragraph's anchored images, charts, and
 *  shapes so body text wraps around retained drawings
 *  (ECMA-376 §20.4.2.16/.17).
 *
 *  Page-level floats (positionV relativeFrom ∈ {page, margin, *Margin, column},
 *  ECMA-376 §20.4.3.2/§20.4.3.5) are skipped when this paragraph was already
 *  pre-registered at the current page's start by {@link preRegisterPageFloats}
 *  — re-registering would double-stamp the FloatRect.
 *  Paragraph-local floats (`paragraph`/`line`/`character`) keep the per-
 *  paragraph path so their Y stays anchored at this paragraph's top. */
function registerAnchorFloats(
  para: ParagraphLayoutSource,
  state: AnchorFloatRegistrationState,
  paragraphAnchorY: number,
): void {
  // One id per registerAnchorFloats call ⇒ one id per paragraph. Floats sharing
  // a paraId (e.g. two side-by-side photos in one paragraph) never displace each
  // other; floats from different paragraphs do (de-facto overlap avoidance).
  const paraId = state.floatParaSeq++;
  const prescanned = state.pageAnchorPrescanned?.has(para) ?? false;
  for (const run of para.runs) {
    if (run.type === 'image') {
      const img = run;
      if (prescanned && isPageLevelWrapFloat(img)) continue;
      registerImageFloat(img, state, paragraphAnchorY, paraId);
    } else if (run.type === 'chart') {
      const chart = run;
      if (prescanned && isPageLevelWrapFloat(chart)) continue;
      registerChartFloat(chart, state, paragraphAnchorY, paraId);
    } else if (run.type === 'shape') {
      const shp = run;
      if (prescanned && isPageLevelWrapFloat(shp)) continue;
      registerShapeFloat(shp, state, paragraphAnchorY, paraId);
    }
  }
}

/** Pre-scan upcoming body elements at a page-start moment and register any
 *  page-level (positionV relativeFrom ∈ {page, margin, *Margin, column})
 *  wrap floats they carry. Mirrors Word's layout order: page-level floats are
 *  positioned as soon as the page is opened, so paragraphs that PRECEDE the
 *  anchoring paragraph in source order on the same page wrap around them
 *  (ECMA-376 §20.4.3.2/§20.4.3.5 + §20.4.2.16/.17). Each pre-registered
 *  paragraph is recorded in `state.pageAnchorPrescanned` so the main flow's
 *  {@link registerAnchorFloats} skips its page-level runs (avoiding a
 *  duplicate FloatRect) while still registering its
 *  paragraph-local floats normally.
 *
 *  Bounds: the scan stops at the next forced page boundary that the
 *  paginator/renderer is guaranteed to honor — an explicit `pageBreak`
 *  (§17.18.79 / §17.3.1.20) or a non-continuous `sectionBreak`. Content
 *  overflow may still push paragraphs to later pages mid-scan; the
 *  paginator's `newPage()` resets the float set wholesale, so those
 *  paragraphs get re-pre-scanned on the next page. (Same idempotent flow as
 *  the existing `registerAnchorFloats` post-newPage re-call at the split-
 *  relocation site.) */
function preRegisterPageFloats(
  body: readonly AnchorScanBlock[],
  startIdx: number,
  state: AnchorFloatRegistrationState,
): void {
  if (!state.pageAnchorPrescanned) state.pageAnchorPrescanned = new Set();
  for (let j = startIdx; j < body.length; j++) {
    const el = body[j];
    if (!el) continue;
    if (el.type === 'pageBreak') break;
    if (el.type === 'sectionBreak') {
      const sb = el as unknown as { kind?: string };
      if (sb.kind && sb.kind !== 'continuous') break;
      continue;
    }
    if (el.type !== 'paragraph') continue;
    const para = el;
    // Skip if already pre-registered (renderer may call this once per page;
    // paginator may re-call after newPage(), but newPage clears the set).
    if (state.pageAnchorPrescanned.has(para)) continue;
    let hasPageLevel = false;
    for (const run of para.runs) {
      if (run.type === 'image') {
        if (isPageLevelWrapFloat(run)) { hasPageLevel = true; break; }
      } else if (run.type === 'chart') {
        if (isPageLevelWrapFloat(run)) { hasPageLevel = true; break; }
      } else if (run.type === 'shape') {
        if (isPageLevelWrapFloat(run)) { hasPageLevel = true; break; }
      }
    }
    if (!hasPageLevel) continue;
    // Register only the page-level floats from this paragraph. paraY=0 is safe
    // because resolveAnchorY ignores it for page-level relativeFrom containers
    // (anchor-geometry.ts §20.4.3.x). Allocate a fresh paraId so overlap
    // avoidance treats these like any other anchor-paragraph float.
    const paraId = state.floatParaSeq++;
    for (const run of para.runs) {
      if (run.type === 'image') {
        const img = run;
        if (!isPageLevelWrapFloat(img)) continue;
        registerImageFloat(img, state, 0, paraId);
      } else if (run.type === 'chart') {
        const chart = run;
        if (!isPageLevelWrapFloat(chart)) continue;
        registerChartFloat(chart, state, 0, paraId);
      } else if (run.type === 'shape') {
        const shp = run;
        if (!isPageLevelWrapFloat(shp)) continue;
        registerShapeFloat(shp, state, 0, paraId);
      }
    }
    state.pageAnchorPrescanned.add(para);
  }
}

/** Reserve the float-exclusion rect for one anchored wrap-image. Retained paint
 * owns the bitmap drawing. */
function registerImageFloat(
  img: DeepReadonly<ImageRun>,
  state: AnchorFloatRegistrationState,
  paragraphAnchorY: number,
  paraId: number,
): void {
  if (!img.anchor) return;
  if (!isWrapFloat(img.wrapMode)) return;

  const mode: 'square' | 'topAndBottom' =
    img.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square';

  // Paragraph-relative wrap floats anchor at the pre-spaceBefore paragraph top
  // (paragraphAnchorY), per ECMA-376 §20.4.3.5 — identical to wrapNone images.
  const box = resolveAnchorBox(img, state, paragraphAnchorY);
  const { w, h, dl, dr, dt, db } = box;

  // Overlap avoidance. Spec-mandated part: allowOverlap="false" (ECMA-376
  // §20.4.2.3) REQUIRES repositioning to prevent overlap; "true"/omitted only
  // permits overlap. Default true per §20.4.2.3.
  // Implementation-defined heuristic with no ECMA-376 basis:
  // displacing the later document-order float, the "other paragraphs only"
  // gate under allowOverlap=true, and the right-then-down re-seat using dist
  // padding as the float-to-float gap. See layout/floats.ts.
  const allowOverlap = img.allowOverlap ?? true;
  const key = anchoredImageCollisionKey(img.imagePath, img.colorReplaceFrom, img.duotone);
  pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w, h, dl, dr, dt, db,
    kind: 'shape', // DrawingML anchor (§20.4.2.3); not a floating table.
    mode,
    side: img.wrapSide ?? 'bothSides',
    imageKey: key,
    paraId,
    avoidOverlap: true,
    allowOverlap,
  });
}

/** Reserve the float-exclusion rect for one anchored wrap-chart
 *  (ECMA-376 §20.4.2.3/.16/.17). Retained paint owns chart drawing. */
function registerChartFloat(
  chart: DeepReadonly<Omit<ChartRun, 'chart'>>,
  state: AnchorFloatRegistrationState,
  paragraphAnchorY: number,
  paraId: number,
): void {
  if (!chart.anchor || !isWrapFloat(chart.wrapMode)) return;

  const box = resolveAnchorBox(chart, state, paragraphAnchorY);
  const { w, h, dl, dr, dt, db } = box;
  if (w <= 0 || h <= 0) return;

  pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w, h, dl, dr, dt, db,
    kind: 'shape',
    mode: chart.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square',
    side: chart.wrapSide ?? 'bothSides',
    allowOverlap: chart.allowOverlap ?? true,
    avoidOverlap: true,
    paraId,
    imageKey: '',
  });
}

/** Reserve the float-exclusion rect for one anchored wrap shape. Retained paint
 *  owns the drawing, so this only pushes an already-represented FloatRect. */
function registerShapeFloat(
  shape: DeepReadonly<ShapeRun>,
  state: AnchorFloatRegistrationState,
  paragraphAnchorY: number,
  paraId: number,
): void {
  if (!isWrapFloat(shape.wrapMode)) return;

  // Match resolveShapeBox's paragraphTopPt convention. resolveAnchorY reads
  // paragraphTopPt only for relativeFrom="paragraph"/"line" (anchorYFromPara);
  // wrap floats anchor at the pre-spaceBefore paragraph top (§20.4.3.5),
  // identical to the image path (resolveAnchorBox uses paragraphAnchorY there).
  const { x, y, w, h } = resolveShapeBox(shape, state, paragraphAnchorY);
  // A degenerate (zero/negative-area) box reserves no exclusion band.
  if (w <= 0 || h <= 0) return;

  const mode: 'square' | 'topAndBottom' =
    shape.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square';

  const physicalDist = {
    dl: shape.distLeft ?? 0,
    dr: shape.distRight ?? 0,
    dt: shape.distTop ?? 0,
    db: shape.distBottom ?? 0,
  };
  // §17.6.20 — on a vertical page the box above is the LOGICAL projection of the
  // physically-resolved shape (resolveShapeBox), so rotate the dist* labels one
  // quarter-turn with it, exactly like the image path (resolveAnchorBox).
  const { dl, dr, dt, db } = state.verticalPhys
    ? physicalDistToLogical(state.verticalPhys, physicalDist)
    : physicalDist;

  // Overlap avoidance, kept consistent with the image path. Shapes carry no
  // parsed allowOverlap field; the spec default is true (§20.4.2.3), so
  // same-paragraph floats never displace each other and a lone shape is a no-op
  // here — but running it keeps multi-float behavior identical to images.
  pushFloatRect(state, {
    x, y, w, h, dl, dr, dt, db,
    kind: 'shape', // DrawingML wp:anchor shape (§20.4.2.3); not a floating table.
    mode,
    side: shape.wrapSide ?? 'bothSides',
    imageKey: '',
    paraId,
    avoidOverlap: true,
    allowOverlap: true,
  });
}

/** Effective cell margins (pt). Per-cell `<w:tcMar>` overrides (ECMA-376
 *  §17.4.42) take precedence per edge over the table-level `<w:tblCellMar>`
 *  default (§17.4.41). A résumé template, for example, gives one cell a larger
 *  top margin to add space above its content. */
function effCellMargins(
  cell: TableLayoutSource['rows'][number]['cells'][number],
  table: TableLayoutSource,
): { top: number; bottom: number; left: number; right: number } {
  return {
    top: cell.marginTop ?? table.cellMarginTop,
    bottom: cell.marginBottom ?? table.cellMarginBottom,
    left: cell.marginLeft ?? table.cellMarginLeft,
    right: cell.marginRight ?? table.cellMarginRight,
  };
}

// ───────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.6.5 docGrid CHARACTER grid (字詰め). When the section's docGrid
// `type` is "linesAndChars" AND a `charSpace` is declared, every character gains
// a fixed spacing delta
//   Δpt = charSpace / 4096   in FLAT POINTS (NEGATIVE = tighter)
// that is INDEPENDENT of font size — it is added to the glyph's MEASURED advance
// (≈1em for full-width EA glyphs), NOT scaled by it. (`gridCharDeltaPx` returns
// exactly `charSpacePt * scale` = charSpace/4096 pt in px; it does not multiply
// by the font size.)
//
// ── The single advance model (measure == draw) ──────────────────────────────
// To make line-break MEASUREMENT and the draw ADVANCE provably identical, the
// grid delta enters in exactly ONE way: as a per-code-point spacing selected by
// the retained grid kind. `gridSegDeltaPx` returns the total delta a segment's
// box gains (`len × Δpx` for every `linesAndChars` segment),
// and `segAdvanceWidth` folds it into the run's complete advance together with
// §17.3.2.43 `w:w` and §17.3.2.35 `w:spacing`. BOTH the layout's `measuredWidth`
// and every draw path derive the segment's advance from this SAME quantity:
//   • non-justified draw walks the glyphs via `justifiedPiecePositions(cps,
//     [1..n-1], perGap=0, measure, letterSpacingPx=Δ)`, whose final glyph lands
//     at `measure(whole) + n·Δ` = the box edge;
//   • justified draw reuses the EXISTING `justifiedPiecePositions` path with the
//     same `letterSpacingPx = Δ`, so its box edge is `measure(whole) + n·Δ +
//     nGaps·perGap` = `measuredWidth + internalStretch`.
// Because both come from `measure(prefix) + (cps before)·Δ`, draw never diverges
// from `measuredWidth` by construction — there is no separate per-glyph sum to
// drift against the whole-string measure (約物半角 contextual collapse stays
// honoured). See packages/core/src/text/justify-positions.ts.

  // Keep the canonical layout-runtime transport explicit. Regional glyph
  // selection itself belongs to the text service and is applied only to Han.
  void cjkFallback;
  const kernel = buildConcreteBodyLayoutKernel(concreteBodyKernelContext);
  return Object.freeze({
    kernel,
    internals: Object.freeze({
      resolveColumnWidths,
      resolveAnchorBox: __test_resolveAnchorBox,
      resolveShapeBox: __test_resolveShapeBox,
      physicalLayoutSection: __test_physicalLayoutSection,
      verticalLayoutSection: __test_verticalLayoutSection,
      preRegisterPageFloats: __test_preRegisterPageFloats,
    }),
  });
}
