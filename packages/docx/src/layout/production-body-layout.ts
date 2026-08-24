import type { BodyElement, DocParagraph, DocTable, DocTableCell, DocRun, ImageRun, ChartRun, ShapeRun, SectionProps } from '../types';
import type { ResolvedLocalFontMetric } from '@silurus/ooxml-core';
import { type FloatRect, FLOAT_OVERLAP_EPS, isWrapFloat } from '../float-layout.js';
import { type FrameBox, computeFrameBox, frameXContainer, pushFloatRect } from '../frame-geometry.js';
import { resolveFloatingTableBoxPt } from '../float-table-geometry.js';
import { xContainer, yContainer, resolveAnchorX, resolveAnchorY } from '../anchor-geometry.js';
import { resolveParagraphLayoutContext, resolveSectionLayoutContext, type DocumentLayoutSettings, type SectionLayoutContext } from '../layout-context.js';
import type { BlockLayoutAlgorithms, BodyFlowRegistryDeltaPt, BodyFlowRegistrySnapshotPt, DeepReadonly, DrawingMLCollisionRegistrySnapshotPt, LayoutServices, FloatRegistryEntryPt, FloatRegistrySnapshotPt, FloatingTablePlacementLayout, DrawingMLCollisionEntryPt, NoteLayout, ParagraphLayout, SourceRef, StoryBlockInput, StoryLayout, TableLayout, TableLayoutInput } from './types.js';
import { beginFloatingTablePlacementTransaction, floatingTableRegistryDelta, resolveFloatingTablePlacementInTransaction, validateFloatingTableRegistryDelta } from './floating-table-transaction.js';
import { floatRegistryParticipant, resolveBlockFlowAdmission, resolvePageAnchoredTableDeferral } from './floats.js';
import { ExactConvergenceError, convergeExactState } from './convergence.js';
import { LayoutInvariantError } from './diagnostics.js';
import type { LayoutOptions } from './options.js';
import { createLayoutServicesRuntimeView, fieldAcquisitionContextOf, verticalGlyphMeasurementServiceOf } from './runtime-state.js';
import { attachStoryBlockLayoutAlgorithms, layoutStory as layoutSharedStory } from './stories.js';
import { buildNoteNumberMap, footnoteIdsInRetainedLines, footnoteIdsInRetainedSlice, indexNotes, noteReferenceIdsInDocumentOrder } from './note-reference-ownership.js';
import type { BodyAcquisitionLocation, BodyLayoutKernel, BodyLayoutSession, PageAnchorPrescanInput, BodyParagraphAcquisitionInput, BodyTableAcquisitionInput } from './body-layout-kernel.js';
import { NoteCapacityExceededError } from './body-layout-kernel.js';
import { FlowCapacityExceededError } from './flow.js';
import { projectBodyOccurrence } from './occurrence-projection.js';
import {
  sectionBodyInsetPt as bodyMarginInsetPt,
  physicalSectionGeometry,
} from './context.js';
import { isAllRotatedVerticalTextDirection, isVerticalSection, isVerticalTextDirection, physicalLayoutSection, verticalLayoutSection } from './section-orientation.js';
import { gridForParagraphContext, paragraphMeasurementEnvironment } from './measurement-environment.js';
import { createRevisionAuthorColorResolver } from './track-changes.js';
import { BODY_STORY_CONTEXT, bodyAnchorReferenceFrames, retainedTableRecord, resolveBodyParagraphLayoutContext, resolveStateParagraphLayoutContext, withTableCellStory } from './acquisition-state.js';
import { applyNumberingBodyOffset, resolveNumberingMarkerGeometry } from './numbering-marker.js';
import { measureTableIntrinsicWidths, resolveTableColumnWidths } from './table-columns.js';
import { measureParagraphIntrinsicWidths, measureTableCellIntrinsicWidths } from './intrinsic-width.js';
// ── Line-layout engine (segmentation + line-breaking + measurement) ──────────
// Body acquisition drives the pure root line-layout kernel through this
// one-directional dependency, with mutable acquisition state owned under layout/.
import { buildFont, fontClassesWithPitches, getDefaultFontSize, paragraphMarkLineHeight } from '../line-layout.js';
import type { DocGridCtx } from '../line-layout.js';
import { measureParagraph } from '../paragraph-measure.js';
import { acquireRetainedTable, type RetainedTableAcquisition } from './table-acquisition.js';
import { combineAdjacentTableLayoutInputs } from './adjacent-table-layout-input.js';
import { layoutTable as layoutRetainedTableInput } from './table.js';
import { startTableFragmentCursor, takeTableFragment, type PageDependentTableBlockRequest } from './table-pagination.js';
import { paragraphGapAdjustment } from './paragraph-spacing.js';
import { bottomBorderExtentPt, resolveParagraphBorderEdges } from './paragraph-border-adjacency.js';
import { acquireParagraphResult, acquireRetainedFrameGroup, bodyFrameGroupFor, bodyParagraphBorderEdgesFor, projectPhysicalAnchorResult, retainedFrameMaximumBaselineLoweringPt, type BodyFrameGroup } from './paragraph.js';
import { wordLoweredDropCapAnchorLeadingPt } from './body-pagination-compatibility.js';
import type { CompleteTextBoxStoryAcquirer } from './paragraph.js';
import type { AnchorFloatRegistrationState, BodyAcquisitionState, BodyMeasurementContext, RetainedTableRecord } from './acquisition-context.js';
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
import type { ParagraphChartRun, ParagraphImageRun, ParagraphLayoutSource, ParagraphShapeRun } from './text.js';
import type { TableLayoutSource } from './table-source-acquisition.js';
import { prepareBodyFrameMetadata } from './frame.js';
import {
  physicalToLogicalMatrix,
  uprightPhysicalExtent,
  writingModeFromTextDirection,
} from './coordinate-space.js';

export function createProductionBodyLayoutRuntime(
  source: LayoutSourceStore,
  measureContext: MeasurementTextContext | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>>,
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
  const anchoredImageCollisionKey = (
    imagePath: string,
    colorReplaceFrom?: string,
    duotone?: { readonly clr1: string; readonly clr2: string },
  ): string => `${imagePath}${colorReplaceFrom ? `|clr:${colorReplaceFrom}` : ''}`
    + `${duotone ? `|duo:${duotone.clr1}:${duotone.clr2}` : ''}`;
/** Retained default separator leading used by the shared note story layout. */
const FOOTNOTE_SEPARATOR_GAP_PT = 6;

function buildMeasureState(
  ctx: MeasurementTextContext,
  section: SectionProps,
  fontFamilyClasses: Record<string, string> = {},
  layoutSettings: DocumentLayoutSettings,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>> = {},
  layoutServices: LayoutServices,
  layoutOptions?: LayoutOptions,
): BodyAcquisitionState {
  const sectionLayout = resolveSectionLayoutContext(layoutSettings, section);
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
      resolveColumns: resolveColumnWidths,
      createCellState: (state, contentWidthPt, cell) => ({
        ...withTableCellStory(state),
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
        const layout = acquireRegisteredParagraph(
          cellState,
          cellState.acquisitionInputs.paragraphAcquisitionInput(paragraph, source),
          {
            id: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
            source,
            flowDomainId,
            ordinaryFlow: true,
            context,
            placement: {
              startYPt: cellState.y,
              paragraphXPt: 0,
              availableWidthPt: paragraphWidthPt,
              maximumYPt: cellState.pageH,
              suppressSpaceBefore: true,
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
            acquireCompleteStory: cellState.acquireCompleteTextBoxStory,
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
            beforePt: paragraph.spaceBefore,
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
    // Un-swap via
    // physicalLayoutSection; `physicalPageWidthPt` is the physical page width
    // in canonical points. `verticalCJK` stays unset: acquisition
    // keeps its horizontal glyph metrics (only anchor geometry re-frames).
    // Seeded from the section this measure state is BUILT from (the body-level
    // body-level one); a direction-mixed document then re-seeds it per
    // section via its retained acquisition location (issue #1000), so a
    // mid-body section's anchors resolve against ITS OWN physical frame.
    get verticalCJK() {
      return isVerticalTextDirection(this.sectionLayout.textDirection);
    },
    get verticalAllRotated() {
      return isVerticalTextDirection(this.sectionLayout.textDirection)
        && isAllRotatedVerticalTextDirection(this.sectionLayout.textDirection);
    },
    verticalPhys: isVerticalSection(section)
      ? (() => {
          const phys = physicalLayoutSection(section);
          return {
            pageWidth: phys.pageWidth,
            pageHeight: phys.pageHeight,
            marginLeft: phys.marginLeft,
            marginRight: phys.marginRight,
            marginTop: bodyMarginInsetPt(phys.marginTop),
            marginBottom: bodyMarginInsetPt(phys.marginBottom),
            physicalPageWidthPt: phys.pageWidth,
          };
        })()
      : undefined,
  };
}

function buildConcreteBodyLayoutKernel(
  source: LayoutSourceStore,
  measureContext: MeasurementTextContext | null,
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>>,
): BodyLayoutKernel {
  const ordinaryAcquisitionInputForAdjacentGroup = (
    group: ReturnType<typeof combineAdjacentTableLayoutInputs>,
  ): TableLayoutInput => {
    const noEdges = Object.freeze({
      top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
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
      rows: Object.freeze(group.rows.map((row) => Object.freeze({
        ...row,
        // §17.4.37 gives every authored member table ownership of its own outer
        // border layer; the union grid carries that folded layer per source row.
        exceptionBorders: row.sourceTableEdges,
      }))),
    });
  };
  const sourceElement = (
    ref: SourceRef,
  ): LayoutFlowBlock => {
    if (ref.story !== 'body' || ref.storyInstance !== 'body' || ref.path.length !== 1) {
      throw new Error('Body acquisition requires a top-level body source');
    }
    const element = source.blocks.resolve(ref);
    if (!element || (element.type !== 'paragraph' && element.type !== 'table')) {
      throw new Error(`Body source does not identify a flow block: ${ref.path.join('.')}`);
    }
    return element;
  };
  const nestedSourceElement = (ref: SourceRef): LayoutFlowBlock => source.blocks.resolve(ref);
  /** Body acquisition stays at the kernel adapter because it resolves renderer
   * state into retained paragraph inputs; layout-owned projections below it
   * receive only immutable structural values. */
  const acquireBodyParagraphAtLocation = (
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
  ) => {
    const edges = bodyParagraphBorderEdgesFor(paragraph) ?? {
      top: 'top' as const,
      bottom: 'bottom' as const,
    };
    const context = resolveBodyParagraphLayoutContext(state, paragraph);
    return acquireParagraphResult(
      paragraph,
      {
        id: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
        source,
        flowDomainId: location.flowDomainId,
        ordinaryFlow: true,
        context,
        placement: {
          startYPt: state.y,
          paragraphXPt: location.availableBounds.xPt,
          availableWidthPt: availableInlineExtentPt,
          maximumYPt: state.pageH,
          suppressSpaceBefore,
        },
        measurer: { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses },
        environment: paragraphMeasurementEnvironment(state),
        exclusions: paragraphWrapExclusions(state.floats, location.flowDomainId),
        anchorCollisions: retainedAnchorCollisions
          ?? paragraphAnchorCollisions(state.floats),
        containerShading: state.containerShading,
        paragraphBorderEdges: edges,
        trailingExtentPt: Math.max(
          context.spaceAfterPt,
          edges.bottom === 'none' ? 0 : bottomBorderExtentPt(paragraph.borders),
        ),
        continuesFromPrevious: continuation.boundary !== null,
        ...(continuation.sourceRangeStart === undefined ? {} : {
          sourceRangeStart: continuation.sourceRangeStart,
        }),
        anchorFrames: bodyAnchorReferenceFrames(state),
        acquireCompleteStory: state.acquireCompleteTextBoxStory,
      },
      continuation.boundary === null ? undefined : {
        boundary: continuation.boundary,
        ...(continuation.uniformRubyAdvancePt === undefined ? {} : {
          uniformRubyAdvancePt: continuation.uniformRubyAdvancePt,
        }),
      },
    );
  };
  return Object.freeze({
    openBodyLayoutSession(
      input: import('./body-layout-kernel.js').BodyLayoutSessionInput,
      services: LayoutServices,
      options: LayoutOptions,
    ) {
      if (!measureContext) throw new Error('Body layout acquisition requires a measurement context');
      const physicalSection: SectionProps = {
        ...source.section,
        ...input.section.geometry,
        textDirection: input.section.textDirection,
        vAlign: input.section.verticalAlignment,
      };
      const section = isVerticalTextDirection(physicalSection.textDirection)
        ? verticalLayoutSection(physicalSection)
        : physicalSection;
      const state = buildMeasureState(
        measureContext,
        section,
        documentFontFamilyClasses,
        source.documentLayoutSettings,
        resolvedLocalFonts,
        services,
        options,
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
      const footnotesById = indexNotes(sourceFootnotes);
      state.noteNumbers = new Map([
        ...[...buildNoteNumberMap(
          sourceFootnotes,
          noteReferenceIdsInDocumentOrder(source.blocks.body, 'footnote'),
        )].map(
          ([id, number]) => [`footnote:${id}`, number] as const,
        ),
        ...[...buildNoteNumberMap(
          sourceEndnotes,
          noteReferenceIdsInDocumentOrder(source.blocks.body, 'endnote'),
        )].map(
          ([id, number]) => [`endnote:${id}`, number] as const,
        ),
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
      const applyLocationTo = (target: BodyAcquisitionState, next: BodyAcquisitionLocation) => {
        const geometry = next.section.geometry;
        target.sectionLayout = next.section as SectionLayoutContext;
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
      };
      const applyLocation = (next: BodyAcquisitionLocation) => {
        location = next;
        applyLocationTo(state, next);
      };
      applyLocation(location);
      const publicParagraphFloatAcquisition = (
        paragraph: LayoutParagraphBlock,
        source: SourceRef,
        candidate: BodyAcquisitionState,
        onlyOccurrenceIds?: ReadonlySet<string>,
        paragraphId = floatRegistry.nextParagraphId,
      ): readonly FloatRegistryEntryPt[] => {
        const committedOccurrenceIds = new Set(floatRegistry.entries.map((entry) => entry.occurrenceId));
        const publicRuns = paragraph.runs.flatMap((run, runIndex) => {
          if (run.type !== 'shape' && run.type !== 'image' && run.type !== 'chart') return [];
          const bridge = publicAnchorBridge(source, runIndex);
          if (!bridge
            || (onlyOccurrenceIds && !onlyOccurrenceIds.has(bridge.occurrenceId))
            || committedOccurrenceIds.has(bridge.occurrenceId)
            || (bridge.pageOwned && candidate.pageAnchorPrescanned?.has(paragraph))) return [];
          return [{
            run,
            occurrenceId: bridge.occurrenceId,
          }];
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
        return Object.freeze(registered.map((float, index): FloatRegistryEntryPt => {
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
        }));
      };
      const retainedParagraphFloatEntries = (
        layout: ParagraphLayout,
      ): readonly FloatRegistryEntryPt[] => {
        const hostFrames = new Map((layout.anchorFrames ?? []).flatMap((frame) => {
          if (frame.status !== 'resolved') return [];
          const isHostAxis = (axis: typeof frame.axes.horizontal) => axis.status === 'resolved'
            && (axis.referenceFrame === 'paragraph'
              || axis.referenceFrame === 'line'
              || axis.referenceFrame === 'character');
          return isHostAxis(frame.axes.horizontal) || isHostAxis(frame.axes.vertical)
            ? [[frame.occurrenceId, frame] as const]
            : [];
        }));
        if (hostFrames.size === 0) return Object.freeze([]);
        const exclusions = new Map(layout.exclusions.flatMap((exclusion) =>
          exclusion.anchorOccurrenceId
            ? [[exclusion.anchorOccurrenceId, exclusion] as const]
            : []));
        return Object.freeze((layout.anchorCollisions ?? []).flatMap(
          (collision): FloatRegistryEntryPt[] => {
          const frame = hostFrames.get(collision.occurrenceId);
          if (!frame) return [];
          if (frame.geometry.wrap.kind === 'none') return [];
          const exclusion = exclusions.get(collision.occurrenceId);
          if (!exclusion) {
            throw new Error(`Wrapped anchor omitted exclusion geometry: ${collision.occurrenceId}`);
          }
          return [Object.freeze({
            kind: 'shape' as const,
            occurrenceId: collision.occurrenceId,
            exclusionId: collision.occurrenceId,
            paragraphId: floatRegistry.nextParagraphId,
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
          })];
        }));
      };
      const reacquireTableBlock = (
        request: PageDependentTableBlockRequest,
      ): ParagraphLayout | TableLayout => {
        if (request.acquired.kind !== 'paragraph') return request.acquired;
        const source = nestedSourceElement(request.acquired.source);
        if (source.type !== 'paragraph') {
          throw new Error('Table paragraph re-acquisition source kind mismatch');
        }
        const candidate: BodyAcquisitionState = {
          ...withTableCellStory(state),
          contentX: 0,
          contentW: request.acquired.flowBounds.widthPt,
          y: request.acquired.flowBounds.yPt,
          floats: (request.floatingTableExclusions ?? []).map((bounds, index): FloatRect => ({
            kind: 'table', tableOverlap: 'never', mode: 'square',
            imageKey: `${TRANSIENT_TABLE_FINAL_FRAME_EXCLUSION_PREFIX}${index}`,
            imageX: bounds.xPt, imageY: bounds.yPt,
            imageW: bounds.widthPt, imageH: bounds.heightPt,
            xLeft: bounds.xPt,
            xRight: bounds.xPt + bounds.widthPt,
            yTop: bounds.yPt,
            yBottom: bounds.yPt + bounds.heightPt,
            side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
            paraId: index,
          })),
          floatParaSeq: request.floatingTableExclusions?.length ?? 0,
          pageAnchorPrescanned: new Set<ParagraphLayoutSource>(),
        };
        const inheritedAuthority =
          inheritedParagraphAuthorityForReacquisition(request.acquired);
        const tableAcquisition = state.retainedTableAcquisition;
        return tableAcquisition.acquireParagraph(
          candidate,
          source,
          request.acquired.flowBounds.widthPt,
          request.acquired.source.path,
          request.acquired.flowDomainId,
          undefined,
          inheritedAuthority,
        );
      };
      const endnotesById = indexNotes(source.blocks.endnotes);
      const storyLayoutCache = new Map<string, StoryLayout>();
      const storyRoot = (ref: SourceRef): readonly LayoutStoryBlock[] => {
        if (ref.path.length !== 0) {
          throw new Error('Story acquisition requires a story-root source');
        }
        return source.blocks.storyRoot(ref);
      };
      const storyElement = (sourceRef: SourceRef): LayoutFlowBlock => {
        if (sourceRef.path.length === 0 || (sourceRef.path.length - 1) % 3 !== 0) {
          throw new Error('Story block acquisition requires a canonical source path');
        }
        return source.blocks.resolve(sourceRef);
      };
      const acquireStoryLayout = (
        request: import('./body-layout-kernel.js').StoryLayoutAcquisitionInput,
      ): StoryLayout => {
        const cacheKey = JSON.stringify({
          source: request.source,
          pageIndex: request.pageIndex,
          section: request.section,
          container: request.container,
        });
        const cached = storyLayoutCache.get(cacheKey);
        if (cached) return cached;
        const root = storyRoot(request.source);
        const noteReferenceNumber = request.source.story === 'footnote'
          || request.source.story === 'endnote'
          ? state.noteNumbers?.get(
              `${request.source.story}:${request.source.storyInstance}`,
            )
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
          pageH: request.container.capacity === 'unbounded'
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
          verticalAllRotated: storyVertical
            && isAllRotatedVerticalTextDirection(request.section.textDirection),
          ...(storyVertical ? {} : { verticalPhys: undefined }),
          storyContext: {
            story: request.source.story,
            containers: [],
            lineNumberingEligible: false,
          },
        };
        preRegisterPageFloats(root, 0, candidate);
        const storyServices = createLayoutServicesRuntimeView(services);
        candidate.layoutServices = storyServices;
        const blockInputs: StoryBlockInput[] = root.flatMap((element, index): StoryBlockInput[] => {
          const source: SourceRef = {
            story: request.source.story,
            storyInstance: request.source.storyInstance,
            path: [index],
          };
          if (element.type === 'unsupportedTextBoxBlock') {
            return [{
              type: 'unsupportedTextBoxBlock',
              qName: element.qName,
              sourcePath: element.sourcePath,
            }];
          }
          if (element.type === 'paragraph') return [{ kind: 'paragraph', source }];
          if (element.type !== 'table') {
            throw new Error(`Unsupported ${request.source.story} story block: ${element.type}`);
          }
          const dependencies = candidate.retainedTableAcquisition;
          const table: LayoutTableBlock = element;
          const columns = resolveColumnWidths(
            table,
            request.container.bounds.widthPt,
            candidate,
          );
          return [acquireRetainedTable(
            table,
            columns,
            request.container.bounds.widthPt,
            candidate,
            source,
            dependencies,
          ).input];
        });
        let previousParagraph: LayoutParagraphBlock | null = null;
        const algorithms: BlockLayoutAlgorithms = {
          layoutParagraph(block, placement) {
            const paragraph = storyElement(block.source);
            if (paragraph.type !== 'paragraph') throw new Error('Story paragraph source kind mismatch');
            const sourceIndex = block.source.path[0]!;
            const previousCandidate = sourceIndex > 0 ? root[sourceIndex - 1] : undefined;
            const previous: LayoutParagraphBlock | null = previousCandidate?.type === 'paragraph'
              ? previousCandidate : null;
            const nextCandidate = root[sourceIndex + 1];
            const next: LayoutParagraphBlock | null = nextCandidate?.type === 'paragraph'
              ? nextCandidate : null;
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
            const publicRuns = paragraph.runs.filter((run, runIndex) =>
              publicAnchorBridge(block.source, runIndex) !== null);
            if (publicRuns.length > 0) {
              registerAnchorFloats(
                Object.freeze({ ...paragraph, runs: Object.freeze(publicRuns) }),
                candidate,
                candidate.y,
              );
            }
            const context = resolveStateParagraphLayoutContext(candidate, paragraph);
            const borderEdges = resolveParagraphBorderEdges(previous, paragraph, next);
            const result = acquireRegisteredParagraph(
              candidate,
              paragraph,
              {
                id: `${block.source.story}:${block.source.storyInstance}:${block.source.path.join('.')}`,
                source: block.source,
                flowDomainId: placement.container.id,
                ordinaryFlow: true,
                context,
                placement: {
                  startYPt,
                  paragraphXPt: placement.container.bounds.xPt,
                  availableWidthPt: placement.container.bounds.widthPt,
                  maximumYPt: placement.availableBounds.yPt + placement.availableBounds.heightPt,
                  suppressSpaceBefore: spacing.suppressBefore,
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
                acquireCompleteStory: candidate.acquireCompleteTextBoxStory,
              },
            );
            previousParagraph = paragraph;
            const nextCursor = {
              xPt: placement.cursor.xPt,
              yPt: startYPt + result.layout.advancePt,
            };
            candidate.y = nextCursor.yPt;
            return { layout: result.layout, nextCursor };
          },
          layoutTable(block, placement) {
            previousParagraph = null;
            const normalizedInput: TableLayoutInput = {
              ...block,
              flowDomainId: placement.container.id,
            };
            const result = layoutRetainedTableInput(normalizedInput, placement, storyServices);
            candidate.y = result.nextCursor.yPt;
            return result;
          },
        };
        attachStoryBlockLayoutAlgorithms(storyServices, algorithms);
        const acquired = layoutSharedStory({
          source: request.source,
          container: request.container,
          blocks: Object.freeze(blockInputs),
        }, storyServices);
        const retained = Object.freeze({
          ...acquired,
          blocks: Object.freeze(acquired.blocks.map((block, index) => {
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
          })),
        });
        storyLayoutCache.set(cacheKey, retained);
        return retained;
      };
      const acquireCompleteTextBoxStory: CompleteTextBoxStoryAcquirer = (request) => {
        const section = request.coordinateSpace === 'upright-physical'
          ? {
              ...state.sectionLayout,
              geometry: physicalSectionGeometry(state.sectionLayout.geometry),
              textDirection: 'lrTb',
            }
          : state.sectionLayout;
        return acquireStoryLayout({
          source: request.source,
          pageIndex: state.pageIndex,
          section,
          container: request.container,
        });
      };
      state.acquireCompleteTextBoxStory = acquireCompleteTextBoxStory;
      const session: BodyLayoutSession = {
        hasPaginationFields: source.hasPaginationFields,
        measureParagraph(request: BodyParagraphAcquisitionInput) {
          applyLocation(request.location);
          const paragraph = sourceElement(request.input.source);
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
              (acquired) => { acquiredGroup = acquired; },
            );
            if (!acquiredGroup) throw new Error('Body frame acquisition omitted its retained group');
            const member = acquiredGroup.members.find((candidate) => candidate.paragraph === paragraph);
            if (!member) throw new Error('Body frame acquisition omitted its retained member');
            const dropCapAnchorLeadingPt = paragraph === frameGroup.members.at(-1)
              && frameGroup.framePr.dropCap !== 'none'
              ? wordLoweredDropCapAnchorLeadingPt(
                  retainedFrameMaximumBaselineLoweringPt(acquiredGroup),
                )
              : 0;
            const absoluteVertical = paragraph.framePr.vAnchor === 'page'
              || paragraph.framePr.vAnchor === 'margin';
            const frameOccurrenceId = box.exclusionId
              ?? `frame:${request.input.source.path.join(':')}`;
            const frameEntry: FloatRegistryEntryPt = Object.freeze({
              kind: 'frame',
              occurrenceId: frameOccurrenceId,
              exclusionId: frameOccurrenceId,
              paragraphId: floatRegistry.nextParagraphId,
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
                sectionFlowOwnership: absoluteVertical ? 'page' as const : 'host-flow' as const,
              }),
              ...(paragraph === frameGroup.owner ? {
                // §17.3.1.11 makes identical adjacent framePr paragraphs one frame,
                // so page admission belongs to the owner before any member is painted.
                retainedFootnoteReferenceIds: Object.freeze([...new Set(
                  acquiredGroup.members.flatMap((candidate) =>
                    footnoteIdsInRetainedSlice(candidate.fragment)),
                )]),
              } : {}),
              ...(!absoluteVertical ? {
                relocationBlockExtentPt: Math.max(
                  0,
                  box.y + box.h - request.location.cursorPt.yPt,
                ),
              } : {}),
              ...(box.registerExclusion === false ? {} : {
                flowRegistryDelta: Object.freeze({
                  floats: floatingTableRegistryDelta(
                    floatRegistry,
                    Object.freeze([frameEntry]),
                    floatRegistry.nextParagraphId + 1,
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
          applyLocationTo(candidate, request.location);
          const publicFloats = request.continuation.boundary === null
            ? publicParagraphFloatAcquisition(paragraph, request.input.source, candidate)
            : Object.freeze([]);
          const acquired = acquireBodyParagraphAtLocation(
            candidate,
            paragraph,
            request.input.source,
            request.location,
            request.availableInlineExtentPt,
            request.suppressSpaceBefore,
            request.continuation,
            drawingCollisionRegistry.entries,
          );
          const { measured, layout } = acquired;
          const allBoundaries = measured.lines.map((line) => {
            const boundary = line.layout.consumedEnd;
            if (!boundary) throw new Error('Measured line omitted its source boundary');
            return boundary;
          });
          const retainedFloats = retainedParagraphFloatEntries(layout);
          const floatEntries = Object.freeze([...publicFloats, ...retainedFloats]);
          // This accepted-collision path is intentionally parser-owned. Hand-built
          // public-model anchors still use the compatibility float bridge (and a
          // public wrapNone run therefore has no collision entry) until the Series
          // B/C bridge removal migrates those runs to the retained OOXML contract.
          const collisionEntries = ownedParagraphAnchorCollisions(layout);
          return Object.freeze({
            layout,
            blockExtentPt: layout.advancePt,
            fragmentation: measured.markOnly
              ? Object.freeze({ kind: 'indivisible' as const })
              : Object.freeze({
                  kind: 'splittable' as const,
                  lineEndBoundaries: Object.freeze(allBoundaries),
                }),
            ...(measured.markOnly
              ? { markBelowBaselinePt: measured.lastLineBelowBaselinePt }
              : {}),
            ...(measured.uniformRubyAdvancePt == null
              ? {}
              : { uniformRubyAdvancePt: measured.uniformRubyAdvancePt }),
            ...(floatEntries.length === 0 && collisionEntries.length === 0
              ? {}
              : {
                  flowRegistryDelta: Object.freeze({
                    ...(floatEntries.length === 0 ? {} : {
                      floats: floatingTableRegistryDelta(
                        floatRegistry,
                        floatEntries,
                        floatRegistry.nextParagraphId + floatEntries.length,
                      ),
                    }),
                    ...(collisionEntries.length === 0 ? {} : {
                      drawingCollisions: drawingMLCollisionRegistryDelta(
                        drawingCollisionRegistry,
                        collisionEntries,
                      ),
                    }),
                  }),
                }),
          });
        },
        measureTable(request: BodyTableAcquisitionInput) {
          applyLocation(request.location);
          if (request.input.kind === 'adjacent-table-group') {
            if (request.cursor && request.cursor.kind !== 'adjacent-table-group') {
              throw new Error('Adjacent table group acquisition received an ordinary table cursor');
            }
            const records = request.input.tables.map((tableInput) => {
              const table = sourceElement(tableInput.source);
              if (table.type !== 'table') throw new Error('Table source kind mismatch');
              const sourceIndex = tableInput.source.path[0]!;
              computeTablePtLayout(state, table, request.availableInlineExtentPt, sourceIndex);
              return retainedTableRecord(state, sourceIndex).acquisition;
            });
            const combinedInput = ordinaryAcquisitionInputForAdjacentGroup(
              combineAdjacentTableLayoutInputs(
                request.input.logicalSequenceId,
                records.map((record) => record.input),
              ),
            );
            const placement = {
              container: {
                id: request.location.flowDomainId,
                kind: 'body' as const,
                bounds: {
                  xPt: 0, yPt: 0,
                  widthPt: request.availableInlineExtentPt,
                  heightPt: request.freshPageBlockExtentPt,
                },
              },
              cursor: { xPt: 0, yPt: 0 },
              availableBounds: {
                xPt: 0, yPt: 0,
                widthPt: request.availableInlineExtentPt,
                heightPt: request.freshPageBlockExtentPt,
              },
            };
            const combinedLayout = layoutRetainedTableInput(
              combinedInput,
              placement,
              services,
            ).layout;
            const nestedById: Record<string, RetainedTableAcquisition> = {};
            records.forEach((record) => Object.entries(record.nestedById).forEach(([id, nested]) => {
              if (nestedById[id] && nestedById[id] !== nested) {
                throw new Error(`Adjacent table group has duplicate nested table id: ${id}`);
              }
              nestedById[id] = nested;
            }));
            const combined: RetainedTableAcquisition = Object.freeze({
              input: combinedInput,
              layout: combinedLayout,
              nestedById: Object.freeze(nestedById),
              floatingTables: Object.freeze(records.flatMap((record) => record.floatingTables)),
            });
            const groupCursor: import('./body-layout-kernel.js').AdjacentTableGroupCursor = request.cursor?.cursor ?? Object.freeze({
              tableIndex: 0,
              sourceRowIndex: 0,
            });
            const rowsBefore = request.input.tables
              .slice(0, groupCursor.tableIndex)
              .reduce((sum, tableInput) => sum + (tableInput.rowCount ?? 0), 0);
            const globalRowIndex = rowsBefore + groupCursor.sourceRowIndex;
            const cursor = groupCursor.tableCursor ?? Object.freeze({
              ...startTableFragmentCursor(),
              rowIndex: globalRowIndex,
            });
            if (cursor.rowIndex !== globalRowIndex) {
              throw new Error('Adjacent-table group and table-fragment cursors disagree');
            }
            const result = takeTableFragment(combined, cursor, {
              availableHeightPt: request.availableBlockExtentPt,
              freshPageHeightPt: request.freshPageBlockExtentPt,
              placement,
              services,
              compatibility: 'word',
              page: {
                physicalPageIndex: request.location.pageIndex,
                displayPageNumber: request.location.pageIndex + 1,
                occurrenceId: `${combinedInput.id}:body:${request.location.pageIndex}`,
              },
            });
            if (!result.fragment || result.requiresFreshPage) {
              return Object.freeze({
                layout: combined.layout,
                blockExtentPt: 0,
                nextCursor: Object.freeze({
                  kind: 'adjacent-table-group' as const,
                  cursor: groupCursor,
                }),
                requiresFreshFlowRegion: true,
              });
            }
            const nextGroupCursor = result.nextCursor
              ? (() => {
                  let tableIndex = 0;
                  let firstRow = 0;
                  while (tableIndex < request.input.tables.length) {
                    const rowCount = request.input.tables[tableIndex]!.rowCount ?? 0;
                    if (result.nextCursor!.rowIndex < firstRow + rowCount) break;
                    firstRow += rowCount;
                    tableIndex += 1;
                  }
                  if (tableIndex >= request.input.tables.length) return null;
                  return Object.freeze({
                    tableIndex,
                    sourceRowIndex: result.nextCursor!.rowIndex - firstRow,
                    tableCursor: result.nextCursor!,
                  });
                })()
              : null;
            return Object.freeze({
              layout: result.fragment,
              blockExtentPt: result.fragment.advancePt,
              nextCursor: nextGroupCursor
                ? Object.freeze({ kind: 'adjacent-table-group' as const, cursor: nextGroupCursor })
                : null,
              ...(result.floatingTableRegistryDelta
                ? {
                    flowRegistryDelta: Object.freeze({
                      floats: result.floatingTableRegistryDelta,
                    }),
                  }
                : {}),
            });
          }
          const table = sourceElement(request.input.source);
          if (table.type !== 'table') throw new Error('Table source kind mismatch');
          const sourceIndex = request.input.source.path[0]!;
          computeTablePtLayout(state, table, request.availableInlineExtentPt, sourceIndex);
          const retained = retainedTableRecord(state, sourceIndex).acquisition;
          if (request.cursor && request.cursor.kind !== 'table') {
            throw new Error('Ordinary table acquisition received an adjacent-group cursor');
          }
          const cursor = request.cursor?.cursor ?? startTableFragmentCursor();
          const pageHeightPt = state.pageH;
          const authoredPositioning = state.acquisitionInputs.tableFormatInput(table).positioning;
          if (authoredPositioning) {
            const positioning = request.cursor?.kind === 'table'
              && request.cursor.floatingContinuationFrame === 'fresh-text'
              ? Object.freeze({ ...authoredPositioning, vertAnchor: 'text', yPt: 0, yAlign: undefined })
              : authoredPositioning;
            const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
            const frames = Object.freeze({
              page: Object.freeze({ xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: pageHeightPt }),
              margin: Object.freeze({
                xPt: state.marginLeft, yPt: state.marginTop,
                widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
                heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
              }),
              text: Object.freeze({
                xPt: request.location.cursorPt.xPt,
                yPt: request.location.cursorPt.yPt,
                widthPt: request.availableInlineExtentPt,
                heightPt: retained.layout.advancePt,
              }),
            });
            const raw = resolveFloatingTableBoxPt(
              positioning,
              frames,
              tableWidthPt,
              retained.layout.advancePt,
            );
            const pageAnchoredCollision = request.cursor?.kind !== 'table'
              && (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin')
              && resolvePageAnchoredTableDeferral({
                bounds: {
                  xPt: raw.x,
                  yPt: raw.y,
                  widthPt: raw.w,
                  heightPt: raw.h,
                },
                blockers: floatRegistry.entries.map(floatRegistryParticipant),
                overlapEpsilonPt: FLOAT_OVERLAP_EPS,
              }).defer;
            if (pageAnchoredCollision) {
              // `word-page-anchored-table-collision-deferral`: a fresh page
              // preserves the authored absolute anchor instead of converting
              // the colliding table to a text continuation.
              return Object.freeze({
                layout: retained.layout,
                blockExtentPt: 0,
                nextCursor: Object.freeze({
                  kind: 'table' as const,
                  cursor,
                  floatingContinuationFrame: 'authored' as const,
                }),
                requiresFreshFlowRegion: true,
              });
            }
            const absoluteAnchorMustSplit = (positioning.vertAnchor === 'page'
              || positioning.vertAnchor === 'margin')
              && retained.layout.advancePt > request.freshPageBlockExtentPt;
            const admissionBlockEndPt = absoluteAnchorMustSplit
              ? request.location.availableBounds.yPt
                + request.location.availableBounds.heightPt
              : positioning.vertAnchor === 'page'
                ? frames.page.yPt + frames.page.heightPt
                : positioning.vertAnchor === 'margin'
                  ? frames.margin.yPt + frames.margin.heightPt
                  : request.location.availableBounds.yPt
                    + request.location.availableBounds.heightPt;
            const freshAdmissionHeightPt = absoluteAnchorMustSplit
              ? request.freshPageBlockExtentPt
              : positioning.vertAnchor === 'page'
              ? frames.page.heightPt
              : positioning.vertAnchor === 'margin'
                ? frames.margin.heightPt
                : request.freshPageBlockExtentPt;
            type FloatingParentTransactionPass =
              | Readonly<{
                  kind: 'fresh-flow-region';
                  result: ReturnType<typeof takeTableFragment>;
                }>
              | Readonly<{
                  kind: 'candidate';
                  parentFrame: Readonly<{ xPt: number; yPt: number }>;
                  result: ReturnType<typeof takeTableFragment>;
                  fragment: NonNullable<ReturnType<typeof takeTableFragment>['fragment']>;
                  resolved: ReturnType<typeof resolveFloatingTablePlacementInTransaction>;
                  nestedEntries: readonly FloatRegistryEntryPt[];
                  fingerprint: string;
                }>;
            let transaction: FloatingParentTransactionPass;
            try {
              transaction = convergeExactState<FloatingParentTransactionPass>({
                step: (previous) => {
                  if (previous?.kind === 'fresh-flow-region') return previous;
                  if (previous?.kind === 'candidate'
                    && previous.resolved.placement.xPt === previous.parentFrame.xPt
                    && previous.resolved.placement.yPt === previous.parentFrame.yPt) {
                    return previous;
                  }
                  const parentFrame = previous?.resolved.placement ?? {
                    xPt: raw.x,
                    yPt: raw.y,
                  };
                  const availableHeightPt = Math.max(
                    0,
                    admissionBlockEndPt - parentFrame.yPt,
                  );
                  const result = takeTableFragment(retained, cursor, {
                    availableHeightPt,
                    freshPageHeightPt: freshAdmissionHeightPt,
                    placement: {
                      container: {
                        id: `${request.location.flowDomainId}:floating-table`,
                        kind: 'body',
                        bounds: {
                          xPt: 0,
                          yPt: 0,
                          widthPt: request.availableInlineExtentPt,
                          heightPt: availableHeightPt,
                        },
                      },
                      cursor: { xPt: 0, yPt: 0 },
                      availableBounds: {
                        xPt: 0,
                        yPt: 0,
                        widthPt: request.availableInlineExtentPt,
                        heightPt: availableHeightPt,
                      },
                    },
                    services,
                    compatibility: 'word',
                    oversizedRowPolicy: 'atomic',
                    page: {
                      physicalPageIndex: request.location.pageIndex,
                      displayPageNumber: state.displayPageNumber
                        ?? request.location.pageIndex + 1,
                      occurrenceId: `${retained.input.id}:fitting-outer:${request.location.pageIndex}:${cursor.rowIndex}:${cursor.rowFragmentIndex}`,
                    },
                    floatingTableFrames: {
                      page: frames.page,
                      margin: frames.margin,
                      column: frames.text,
                    },
                    floatingTableRegistry: floatRegistry,
                    finalPlacementTranslationPt: parentFrame,
                    reacquirePageDependentBlock: reacquireTableBlock,
                  });
                  if (!result.fragment || result.requiresFreshPage) {
                    return Object.freeze({
                      kind: 'fresh-flow-region' as const,
                      result,
                    });
                  }
                  const sourcePlacement: FloatingTablePlacementLayout = Object.freeze({
                    kind: 'floating-table-placement',
                    occurrenceId: `${retained.input.id}:root:${request.location.pageIndex}:${cursor.rowIndex}:${cursor.rowFragmentIndex}`,
                    ownership: 'source',
                    physicalPageIndex: request.location.pageIndex,
                    displayPageNumber: state.displayPageNumber
                      ?? request.location.pageIndex + 1,
                    hostCellId: request.location.flowDomainId,
                    sourceBlockIndex: request.input.source.path[0]!,
                    anchorBlockIndex: request.input.source.path[0]!,
                    tableId: result.fragment.id,
                    overlap: table.overlap === 'never' ? 'never' : 'overlap',
                    positioning,
                    anchorBounds: frames.text,
                    child: result.fragment,
                  });
                  const nestedEntries = result.floatingTableRegistryDelta?.entries ?? [];
                  const nestedNextParagraphId =
                    result.floatingTableRegistryDelta?.nextParagraphId
                    ?? floatRegistry.nextParagraphId;
                  const resolved = resolveFloatingTablePlacementInTransaction(
                    sourcePlacement,
                    frames,
                    beginFloatingTablePlacementTransaction(
                      floatRegistry.entries,
                      nestedNextParagraphId,
                      floatRegistry.coordinateSpace,
                      floatRegistry.flowDomainId,
                    ),
                  );
                  const fingerprint = JSON.stringify({
                    parentFrame: {
                      xPt: resolved.placement.xPt,
                      yPt: resolved.placement.yPt,
                    },
                    fragment: result.fragment,
                    nestedEntries,
                    resolvedBounds: resolved.placement.bounds,
                  });
                  return Object.freeze({
                    kind: 'candidate' as const,
                    parentFrame: Object.freeze({
                      xPt: parentFrame.xPt,
                      yPt: parentFrame.yPt,
                    }),
                    result,
                    fragment: result.fragment,
                    resolved,
                    nestedEntries,
                    fingerprint,
                  });
                },
                stateOf: (value) => value.kind === 'fresh-flow-region'
                  ? 'fresh-flow-region'
                  : value.fingerprint,
                limit: 16,
              }).value;
            } catch (error) {
              if (error instanceof ExactConvergenceError) {
                throw new LayoutInvariantError(
                  'NON_CONVERGENCE',
                  error.reason === 'cycle'
                    ? 'Floating table parent/child transaction repeated an exact-state cycle'
                    : 'Floating table parent/child transaction reached the operational pass limit 16',
                );
              }
              throw error;
            }
            if (transaction.kind === 'fresh-flow-region') {
              return Object.freeze({
                layout: retained.layout,
                blockExtentPt: 0,
                nextCursor: Object.freeze({
                  kind: 'table' as const,
                  cursor,
                  floatingContinuationFrame: 'fresh-text' as const,
                }),
                requiresFreshFlowRegion: true,
              });
            }
            const {
              result,
              fragment,
              resolved,
              nestedEntries,
            } = transaction;
            const isFloatingContinuation = request.cursor?.kind === 'table'
              && request.cursor.floatingContinuationFrame !== undefined;
            const admittedBlockEndPt = request.location.availableBounds.yPt
              + request.location.availableBounds.heightPt;
            const hostFlowPlacements = [
              ...fragment.resolvedFloatingTables ?? [],
              resolved.placement,
            ].filter((placement) => placement.source.positioning.vertAnchor === 'text');
            if (!isFloatingContinuation && hostFlowPlacements.some((placement) => (
              placement.exclusionBounds.yPt + placement.exclusionBounds.heightPt
                > admittedBlockEndPt
            ))) {
              return Object.freeze({
                layout: fragment,
                blockExtentPt: 0,
                nextCursor: Object.freeze({
                  kind: 'table' as const,
                  cursor,
                  floatingContinuationFrame: 'fresh-text' as const,
                }),
                requiresFreshFlowRegion: true,
              });
            }
            return Object.freeze({
              layout: fragment,
              blockExtentPt: 0,
              nextCursor: result.nextCursor
                ? Object.freeze({
                    kind: 'table' as const,
                    cursor: result.nextCursor,
                    floatingContinuationFrame: 'fresh-text' as const,
                  })
                : null,
              flowRegistryDelta: Object.freeze({
                floats: floatingTableRegistryDelta(
                  floatRegistry,
                  Object.freeze([...nestedEntries, ...resolved.transaction.delta]),
                  resolved.transaction.nextParagraphId,
                ),
              }),
              placement: Object.freeze({
                coordinateSpace: 'logical-body' as const,
                xPt: resolved.placement.xPt,
                yPt: resolved.placement.yPt,
                sectionFlowOwnership: positioning.vertAnchor === 'page'
                  || positioning.vertAnchor === 'margin'
                  ? 'page' as const
                  : 'host-flow' as const,
              }),
            });
          }
          if (state.verticalPhys && !effectiveTablePositioning(table)) {
            if (request.cursor) {
              throw new Error('An upright physical table must remain atomic');
            }
            const physical = state.verticalPhys;
            const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
            if (tableWidthPt > request.availableBlockExtentPt
              && request.availableBlockExtentPt < request.freshPageBlockExtentPt) {
              return Object.freeze({
                layout: retained.layout,
                blockExtentPt: 0,
                nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
                requiresFreshFlowRegion: true,
              });
            }
            const physicalLeftPt = physical.physicalPageWidthPt
              - request.location.cursorPt.yPt - tableWidthPt;
            const physicalTopPt = request.location.cursorPt.xPt;
            const physicalBandHeightPt = Math.max(
              retained.layout.advancePt,
              physical.pageHeight - physical.marginTop - physical.marginBottom,
            );
            const flowDomainId = `upright-physical-page:${request.location.pageIndex}`;
            const upright = takeTableFragment(retained, startTableFragmentCursor(), {
              availableHeightPt: physicalBandHeightPt,
              freshPageHeightPt: physicalBandHeightPt,
              placement: {
                container: {
                  id: flowDomainId, kind: 'body',
                  bounds: { xPt: 0, yPt: 0, widthPt: tableWidthPt, heightPt: physicalBandHeightPt },
                },
                cursor: { xPt: 0, yPt: 0 },
                availableBounds: {
                  xPt: 0, yPt: 0, widthPt: tableWidthPt, heightPt: physicalBandHeightPt,
                },
              },
              services,
              compatibility: 'word',
              oversizedRowPolicy: 'atomic',
              page: {
                physicalPageIndex: request.location.pageIndex,
                displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
                occurrenceId: `${retained.input.id}:upright-page:${request.location.pageIndex}`,
              },
              floatingTableFrames: {
                page: { xPt: 0, yPt: 0, widthPt: physical.pageWidth, heightPt: physical.pageHeight },
                margin: {
                  xPt: physical.marginLeft, yPt: physical.marginTop,
                  widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
                  heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
                },
                column: {
                  xPt: physical.marginLeft, yPt: physical.marginTop,
                  widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
                  heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
                },
              },
              floatingTableRegistry: Object.freeze({
                coordinateSpace: 'upright-physical-page-points' as const,
                flowDomainId,
                entries: Object.freeze([]),
                nextParagraphId: 0,
              }),
              finalPlacementTranslationPt: { xPt: physicalLeftPt, yPt: physicalTopPt },
              reacquirePageDependentBlock: reacquireTableBlock,
            });
            if (!upright.fragment || upright.nextCursor || upright.requiresFreshPage) {
              throw new Error('Upright table final-frame layout must remain atomic');
            }
            return Object.freeze({
              layout: upright.fragment,
              blockExtentPt: tableWidthPt,
              nextCursor: null,
              placement: Object.freeze({
                coordinateSpace: 'upright-physical' as const,
                xPt: physicalLeftPt + upright.fragment.flowBounds.xPt,
                yPt: physicalTopPt + upright.fragment.flowBounds.yPt,
                sectionFlowOwnership: 'host-flow' as const,
              }),
            });
          }
          const result = takeTableFragment(retained, cursor, {
            availableHeightPt: request.availableBlockExtentPt,
            freshPageHeightPt: request.freshPageBlockExtentPt,
            placement: {
              container: {
                id: request.location.flowDomainId,
                kind: 'body',
                bounds: {
                  xPt: 0, yPt: 0,
                  widthPt: request.availableInlineExtentPt,
                  heightPt: request.availableBlockExtentPt,
                },
              },
              cursor: { xPt: 0, yPt: 0 },
              availableBounds: {
                xPt: 0, yPt: 0,
                widthPt: request.availableInlineExtentPt,
                heightPt: request.availableBlockExtentPt,
              },
            },
            services,
            compatibility: 'word',
            page: {
              physicalPageIndex: request.location.pageIndex,
              displayPageNumber: request.location.pageIndex + 1,
              occurrenceId: `${retained.input.id}:body:${request.location.pageIndex}`,
            },
            floatingTableFrames: {
              page: { xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: pageHeightPt },
              margin: {
                xPt: state.marginLeft,
                yPt: state.marginTop,
                widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
                heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
              },
              column: request.location.availableBounds,
            },
            floatingTableRegistry: floatRegistry,
            finalPlacementTranslationPt: {
              xPt: request.location.availableBounds.xPt,
              yPt: request.location.cursorPt.yPt,
            },
            reacquirePageDependentBlock: reacquireTableBlock,
          });
          const tableInlineStartPt = request.location.availableBounds.xPt
            + retained.layout.flowBounds.xPt;
          const tableInlineEndPt = tableInlineStartPt + retained.layout.flowBounds.widthPt;
          const remainingTableExtentPt = result.fragment?.advancePt ?? 0;
          const retryAtBlockStartPt = resolveBlockFlowAdmission({
            inlineStartPt: tableInlineStartPt,
            inlineEndPt: tableInlineEndPt,
            blockStartPt: request.location.cursorPt.yPt,
            blockExtentPt: remainingTableExtentPt,
            blockers: floatRegistry.entries.map(floatRegistryParticipant),
            overlapEpsilonPt: FLOAT_OVERLAP_EPS,
          }).blockStartPt;
          if (retryAtBlockStartPt > request.location.cursorPt.yPt) {
            return Object.freeze({
              layout: retained.layout,
              blockExtentPt: 0,
              nextCursor: request.cursor ?? null,
              retryAtBlockStartPt,
            });
          }
          if (!result.fragment || result.requiresFreshPage) {
            return Object.freeze({
              layout: retained.layout,
              blockExtentPt: 0,
              nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
              requiresFreshFlowRegion: true,
            });
          }
          return Object.freeze({
            layout: result.fragment,
            blockExtentPt: result.fragment.advancePt,
            nextCursor: result.nextCursor
              ? Object.freeze({ kind: 'table' as const, cursor: result.nextCursor })
              : null,
            ...(result.floatingTableRegistryDelta
              ? {
                  flowRegistryDelta: Object.freeze({
                    floats: result.floatingTableRegistryDelta,
                  }),
                }
              : {}),
          });
        },
        layoutStory: acquireStoryLayout,
        layoutNotes(request) {
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
            const separatorHeightPt = first ? FOOTNOTE_SEPARATOR_GAP_PT : 0;
            const storyContainer = {
              ...request.container,
              id: `${request.container.id}:${request.kind}:${id}`,
              bounds: {
                ...request.container.bounds,
                yPt: cursorYPt + separatorHeightPt,
                heightPt: Math.max(
                  0,
                  request.container.bounds.yPt + request.container.bounds.heightPt
                    - cursorYPt - separatorHeightPt,
                ),
              },
            };
            let story: StoryLayout;
            try {
              story = acquireStoryLayout({
                source,
                pageIndex: request.pageIndex,
                section: request.section,
                container: storyContainer,
              });
            } catch (error) {
              if (error instanceof FlowCapacityExceededError
                && error.containerId === storyContainer.id) {
                throw new NoteCapacityExceededError(
                  request.kind,
                  request.pageIndex,
                  request.container.id,
                );
              }
              throw error;
            }
            const separator = first ? Object.freeze([Object.freeze({
              edge: 'top' as const,
              from: Object.freeze({
                xPt: request.container.bounds.xPt,
                yPt: cursorYPt + separatorHeightPt / 2,
              }),
              to: Object.freeze({
                xPt: request.container.bounds.xPt + request.container.bounds.widthPt / 3,
                yPt: cursorYPt + separatorHeightPt / 2,
              }),
              color: '#000000',
              widthPt: 0.5,
              authoredStyle: 'single',
              style: 'solid' as const,
            })]) : Object.freeze([]);
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
                widthPt: Math.max(
                  flowBounds.xPt + flowBounds.widthPt,
                  story.inkBounds.xPt + story.inkBounds.widthPt,
                ) - Math.min(flowBounds.xPt, story.inkBounds.xPt),
                heightPt: Math.max(
                  flowBounds.yPt + flowBounds.heightPt,
                  story.inkBounds.yPt + story.inkBounds.heightPt,
                ) - Math.min(flowBounds.yPt, story.inkBounds.yPt),
              }),
              clipBounds: request.container.bounds,
              advancePt,
              separator,
              story,
            });
            notes.push(note);
            cursorYPt += advancePt;
            first = false;
          }
          return Object.freeze(notes);
        },
        measureFollowingBlock(request) {
          const candidate: BodyAcquisitionState = {
            ...state,
            floats: [...state.floats],
            retainedTablesBySourceIndex: new Map(state.retainedTablesBySourceIndex),
          };
          applyLocationTo(candidate, request.location);
          if (request.input.kind === 'adjacent-table-group') {
            const records = request.input.tables.map((tableInput) => {
              const table = sourceElement(tableInput.source);
              if (table.type !== 'table') throw new Error('Following table source kind mismatch');
              const sourceIndex = tableInput.source.path[0]!;
              computeTablePtLayout(candidate, table, request.availableInlineExtentPt, sourceIndex);
              return retainedTableRecord(candidate, sourceIndex).acquisition;
            });
            const combinedInput = ordinaryAcquisitionInputForAdjacentGroup(
              combineAdjacentTableLayoutInputs(
                request.input.logicalSequenceId,
                records.map((record) => record.input),
              ),
            );
            const layout = layoutRetainedTableInput(combinedInput, {
              container: {
                id: request.location.flowDomainId,
                kind: 'body',
                bounds: request.location.availableBounds,
              },
              cursor: request.location.cursorPt,
              availableBounds: request.location.availableBounds,
            }, services).layout;
            return Object.freeze({
              fullExtentPt: layout.advancePt,
              leadContentExtentPt: layout.rows[0]?.advancePt ?? layout.advancePt,
              fullFootnoteReferenceIds: footnoteIdsInRetainedSlice(layout),
              leadFootnoteReferenceIds: footnoteIdsInRetainedSlice({
                ...layout,
                rows: layout.rows.slice(0, 1),
              }),
            });
          }
          const element = sourceElement(request.input.source);
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
              drawingCollisionRegistry.entries,
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
              leadFootnoteReferenceIds: firstLine
                ? footnoteIdsInRetainedLines([firstLine])
                : [],
            });
          }
          if (element.type !== 'table') throw new Error('Following table source kind mismatch');
          const sourceIndex = request.input.source.path[0]!;
          computeTablePtLayout(candidate, element, request.availableInlineExtentPt, sourceIndex);
          const layout = retainedTableRecord(candidate, sourceIndex).acquisition.layout;
          return Object.freeze({
            fullExtentPt: layout.advancePt,
            leadContentExtentPt: layout.rows[0]?.advancePt ?? layout.advancePt,
            fullFootnoteReferenceIds: footnoteIdsInRetainedSlice(layout),
            leadFootnoteReferenceIds: footnoteIdsInRetainedSlice({
              ...layout,
              rows: layout.rows.slice(0, 1),
            }),
          });
        },
        prescanPageAnchors(request: PageAnchorPrescanInput) {
          const geometry = request.location.section.geometry;
          const marginTopPt = bodyMarginInsetPt(geometry.marginTop);
          const marginBottomPt = bodyMarginInsetPt(geometry.marginBottom);
          const frames = Object.freeze({
            page: Object.freeze({
              xPt: 0, yPt: 0,
              widthPt: geometry.pageWidth, heightPt: geometry.pageHeight,
            }),
            margin: Object.freeze({
              xPt: geometry.marginLeft, yPt: marginTopPt,
              widthPt: Math.max(0, geometry.pageWidth - geometry.marginLeft - geometry.marginRight),
              heightPt: Math.max(0, geometry.pageHeight - marginTopPt - marginBottomPt),
            }),
            column: Object.freeze({
              xPt: request.location.availableBounds.xPt, yPt: marginTopPt,
              widthPt: request.availableInlineExtentPt,
              heightPt: Math.max(0, geometry.pageHeight - marginTopPt - marginBottomPt),
            }),
            paragraph: null,
            line: null,
            character: null,
            pageParity: request.location.pageIndex % 2 === 0 ? 'odd' as const : 'even' as const,
          });
          const publicParagraphs = new Set<ParagraphLayoutSource>();
          const paragraphKey = (source: SourceRef) =>
            `${source.story}:${source.storyInstance}:${source.path.join('.')}`;
          const paragraphIds = new Map<string, number>();
          const paragraphIdFor = (source: SourceRef): number => {
            const key = paragraphKey(source);
            if (!paragraphIds.has(key)) {
              paragraphIds.set(key, floatRegistry.nextParagraphId + paragraphIds.size);
            }
            return paragraphIds.get(key)!;
          };
          const entries = request.anchors.flatMap((anchor): readonly FloatRegistryEntryPt[] => {
            const paragraph = sourceElement(anchor.paragraphSource);
            if (paragraph.type !== 'paragraph') {
              throw new Error('Page-anchor prescan source kind mismatch');
            }
            const acquired = paragraph;
            const hostMatches = acquired.runs.filter((run) =>
              run.type === 'anchorHost' && run.anchorOccurrenceId === anchor.occurrenceId);
            const payloads = acquired.runs
              .map((run, runIndex) => ({ run, runIndex }))
              .filter((candidate): candidate is typeof candidate & {
                run: Extract<typeof candidate.run, {
                  type: 'image' | 'chart' | 'shape' | 'unavailableDrawing';
                }> & {
                  anchorAcquisitionInput: NonNullable<Extract<typeof candidate.run, {
                    type: 'image' | 'chart' | 'shape' | 'unavailableDrawing';
                  }>['anchorAcquisitionInput']>;
                };
              } => (
                (candidate.run.type === 'image'
                  || candidate.run.type === 'chart'
                  || candidate.run.type === 'shape'
                  || candidate.run.type === 'unavailableDrawing')
                && candidate.run.anchorAcquisitionInput?.occurrenceId === anchor.occurrenceId
              ))
              .sort((left, right) => (
                (left.run.anchorAcquisitionInput.group?.sourceIndex ?? 0)
                  - (right.run.anchorAcquisitionInput.group?.sourceIndex ?? 0)
                  || left.runIndex - right.runIndex
              ));
            if (hostMatches.length !== 1 || payloads.length === 0) {
              const publicRun = paragraph.runs.find((run, runIndex) =>
                publicAnchorBridge(anchor.paragraphSource, runIndex)?.occurrenceId
                  === anchor.occurrenceId);
              if (publicRun) {
                if (
                  (publicRun.type === 'image'
                    || publicRun.type === 'chart'
                    || publicRun.type === 'shape')
                  && publicRun.wrapMode === 'none'
                ) return [];
                const candidate: BodyAcquisitionState = {
                  ...state,
                  floats: [...state.floats],
                  pageAnchorPrescanned: new Set(state.pageAnchorPrescanned),
                };
                applyLocationTo(candidate, request.location);
                const publicEntries = publicParagraphFloatAcquisition(
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
              throw new Error(`Page-anchor prescan occurrence acquisition mismatch: ${anchor.occurrenceId}`);
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
            const retainedResult = isVerticalTextDirection(
              request.location.section.textDirection,
            )
              ? (() => {
                  const writingMode = writingModeFromTextDirection(
                    request.location.section.textDirection as string,
                  );
                  const physicalPage = uprightPhysicalExtent({
                    widthPt: frames.page.widthPt,
                    heightPt: frames.page.heightPt,
                  }, writingMode);
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
            const polygon = retainedResult.geometry.wrap.polygon?.points ?? Object.freeze([
              Object.freeze({ xPt: wrapBounds.xPt, yPt: wrapBounds.yPt }),
              Object.freeze({ xPt: wrapBounds.xPt + wrapBounds.widthPt, yPt: wrapBounds.yPt }),
              Object.freeze({
                xPt: wrapBounds.xPt + wrapBounds.widthPt,
                yPt: wrapBounds.yPt + wrapBounds.heightPt,
              }),
              Object.freeze({ xPt: wrapBounds.xPt, yPt: wrapBounds.yPt + wrapBounds.heightPt }),
            ]);
            return [Object.freeze({
              kind: 'shape' as const,
              occurrenceId: anchor.occurrenceId,
              paragraphId: paragraphIdFor(anchor.paragraphSource),
              bounds: retainedResult.geometry.objectFrame,
              exclusionBounds: wrapBounds,
              wrap: retainedResult.geometry.wrap.kind,
              wrapSide: retainedResult.geometry.wrap.side,
              wrapDistances: retainedResult.geometry.wrap.distances,
              wrapPolygon: Object.freeze([...polygon]),
            })];
          });
          publicParagraphs.forEach((paragraph) => state.pageAnchorPrescanned?.add(paragraph));
          if (entries.length === 0) return null;
          return Object.freeze({
            floats: Object.freeze({
              coordinateSpace: 'logical-page-points' as const,
              // Page-owned wrap exclusions survive same-page column/section
              // cutovers, so their transaction identity belongs to the physical
              // page rather than the active body flow domain.
              flowDomainId: floatRegistry.flowDomainId,
              baseEntries: floatRegistry.entries,
              baseNextParagraphId: floatRegistry.nextParagraphId,
              nextParagraphId: floatRegistry.nextParagraphId + entries.length,
              entries: Object.freeze(entries),
            }),
          });
        },
        measureLineNumberGlyph(text) {
          const previousFont = measureContext.font;
          try {
            const fontSizePt = source.fonts.defaultBodyFontSizePt;
            const font = buildFont(false, false, fontSizePt, null, {});
            measureContext.font = font;
            const metrics = measureContext.measureText(text);
            return Object.freeze({
              widthPt: metrics.width,
              ascentPt: metrics.fontBoundingBoxAscent
                ?? metrics.actualBoundingBoxAscent
                ?? fontSizePt * 0.8,
              descentPt: metrics.fontBoundingBoxDescent
                ?? metrics.actualBoundingBoxDescent
                ?? fontSizePt * 0.2,
              font,
            });
          } finally {
            measureContext.font = previousFont;
          }
        },
        resetPageAcquisition(next) {
          state.floats = [];
          state.floatParaSeq = 0;
          state.pageAnchorPrescanned = new Set();
          floatRegistry = Object.freeze({
            coordinateSpace: 'logical-page-points' as const,
            flowDomainId: pageRegistryFlowDomainId(next.pageIndex),
            entries: Object.freeze([]),
            nextParagraphId: 0,
          });
          drawingCollisionRegistry = createDrawingMLCollisionRegistry(
            pageRegistryFlowDomainId(next.pageIndex),
            'logical-page-points',
          );
          applyLocation(next);
        },
        moveAcquisitionCursor: applyLocation,
        flowRegistrySnapshot(): BodyFlowRegistrySnapshotPt {
          return Object.freeze({
            floats: floatRegistry,
            drawingCollisions: drawingCollisionRegistry,
          });
        },
        commitFlowRegistryDelta(delta: BodyFlowRegistryDeltaPt) {
          if (!delta.floats && !delta.drawingCollisions) {
            throw new Error('Body flow registry delta must update at least one registry');
          }
          if (delta.floats) {
            validateFloatingTableRegistryDelta(delta.floats, {
              coordinateSpace: floatRegistry.coordinateSpace,
              flowDomainId: floatRegistry.flowDomainId,
              entries: floatRegistry.entries,
              nextParagraphId: floatRegistry.nextParagraphId,
            });
          }
          if (delta.drawingCollisions) {
            validateDrawingMLCollisionRegistryDelta(
              drawingCollisionRegistry,
              delta.drawingCollisions,
            );
          }
          const nextDrawingCollisionRegistry = delta.drawingCollisions
            ? applyDrawingMLCollisionRegistryDelta(
                drawingCollisionRegistry,
                delta.drawingCollisions,
              )
            : drawingCollisionRegistry;
          const retainedFloats = (delta.floats?.entries ?? []).map((entry): FloatRect => {
            const left = entry.wrapDistances?.leftPt
              ?? entry.bounds.xPt - entry.exclusionBounds.xPt;
            const top = entry.wrapDistances?.topPt
              ?? entry.bounds.yPt - entry.exclusionBounds.yPt;
            const right = entry.wrapDistances?.rightPt
              ?? entry.exclusionBounds.xPt + entry.exclusionBounds.widthPt
                - entry.bounds.xPt - entry.bounds.widthPt;
            const bottom = entry.wrapDistances?.bottomPt
              ?? entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt
                - entry.bounds.yPt - entry.bounds.heightPt;
            const core = {
              mode: (entry.wrap === 'topAndBottom'
                ? 'topAndBottom'
                : 'square') as FloatRect['mode'],
              ...(entry.kind === 'shape' ? {
                anchorOccurrenceId: entry.occurrenceId,
                acquisitionOccurrenceId: entry.occurrenceId,
              } : {}),
              ...(entry.wrap ? {
                authoredWrap: entry.wrap,
                wrapPolygon: entry.wrapPolygon,
              } : {}),
              imageKey: entry.exclusionId
                ?? (entry.kind === 'table' ? `body:float:${entry.paragraphId}` : ''),
              imageX: entry.bounds.xPt, imageY: entry.bounds.yPt,
              imageW: entry.bounds.widthPt, imageH: entry.bounds.heightPt,
              xLeft: entry.exclusionBounds.xPt,
              xRight: entry.exclusionBounds.xPt + entry.exclusionBounds.widthPt,
              yTop: entry.exclusionBounds.yPt,
              yBottom: entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt,
              side: entry.wrapSide ?? 'bothSides',
              distLeft: left, distRight: right,
              distTop: top, distBottom: bottom,
              paraId: entry.paragraphId,
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
            floatRegistry = Object.freeze({
              ...floatRegistry,
              entries: Object.freeze([...floatRegistry.entries, ...delta.floats.entries]),
              nextParagraphId: delta.floats.nextParagraphId,
            });
            state.floatParaSeq = delta.floats.nextParagraphId;
          }
          drawingCollisionRegistry = nextDrawingCollisionRegistry;
        },
      };
      return Object.freeze(session);
    },
  });
}

function paraGrid(para: ParagraphLayoutSource, state: BodyMeasurementContext): DocGridCtx {
  return gridForParagraphContext(
    state,
    resolveStateParagraphLayoutContext(state, para),
  );
}

/** Resolve column widths once, acquire the retained table, and return its
 * authoritative row advances for one top-level body occurrence. */
function computeTablePtLayout(
  state: BodyAcquisitionState,
  table: TableLayoutSource,
  contentWPt: number,
  sourceIndex: number,
): { colWidthsPt: number[]; rowContentHeightsPt: number[]; rowHeightsPt: number[] } {
  const prior = state.retainedTablesBySourceIndex.get(sourceIndex);
  const colWidthsPt = resolveColumnWidths(table, contentWPt, state);
  const dependencies = state.retainedTableAcquisition;
  const acquired = acquireRetainedTable(
    table,
    colWidthsPt,
    contentWPt,
    state,
    [sourceIndex],
    dependencies,
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
    anchorYPt: state.y,
  }));
  const rowHeightsPt = retained.layout.rows.map((row) => row.advancePt);
  return { colWidthsPt, rowContentHeightsPt: rowHeightsPt, rowHeightsPt };
}

/**
 * Acquire the parser/style facts and intrinsic content constraints required by
 * ECMA-376 §17.18.87, then resolve the shared table grid. `tblGrid` is the
 * initial grid, not an oracle containing an application's previous result;
 * authored `tblW`, `tcW`, `wBefore`, and `wAfter` remain active constraints.
 * Exported for the table-layout integration tests.
 */
function resolveColumnWidths(
  table: TableLayoutSource,
  contentWPt: number,
  state: BodyMeasurementContext,
): number[] {
  const format = state.acquisitionInputs.tableFormatInput(table);
  const baseIndentPt = Number.isFinite(table.tblInd) ? (table.tblInd ?? 0) : 0;
  const rowIndentPts = format.rows.map((row) => {
    const exception = row.exception;
    return exception?.indentAuthored
      ? (exception.indentPt ?? 0)
      : baseIndentPt;
  });
  // §17.18.87 lets AutoFit override its preferred width up to the page width.
  // Keep the ordinary text-band ceiling unless §17.4.50 placement moves a
  // top-level page-owned story table into its semantic leading margin. The
  // promoted ceiling is the physical page distance left after resolving jc +
  // tblInd, rather than the page width in isolation: a partial negative indent
  // does not move the table origin all the way to the page edge. Test the
  // authored indent before bidiVisual reverses its physical translation;
  // table justification reverses under bidiVisual as well. Body tables must
  // also remain in a single-column page band; headers and footers are
  // page-owned stories and do not inherit the body's newspaper columns.
  const story = state.storyContext?.story;
  const isTopLevelPageOwnedStory = state.storyContext?.containers.length === 0
    && (story === 'header'
      || story === 'footer'
      || (story === 'body' && state.sectionLayout?.columns.length === 1));
  const isLeadingMarginPageStoryTable = format.ordinaryFlow
    && isTopLevelPageOwnedStory
    && !isVerticalTextDirection(state.sectionLayout.textDirection)
    && [baseIndentPt, ...rowIndentPts].some((indentPt) => indentPt < 0);
  const rowPlacements = format.rows.length === 0
    ? [{ justification: table.jc, indentPt: baseIndentPt }]
    : format.rows.map((row, rowIndex) => ({
        justification: row.justification ?? table.jc,
        indentPt: rowIndentPts[rowIndex] ?? baseIndentPt,
      }));
  const bidiVisual = table.bidiVisual === true;
  const pageFitCeilingPt = Math.min(
    state.pageWidth,
    ...rowPlacements.map(({ justification, indentPt }) => {
      const trailing = justification === 'right' || justification === 'end';
      const alignment = justification === 'center'
        ? 'center'
        : (bidiVisual ? !trailing : trailing) ? 'right' : 'left';
      const signedIndentPt = bidiVisual ? -indentPt : indentPt;
      if (alignment === 'left') {
        const resolvedOriginPt = state.contentX + signedIndentPt;
        return state.pageWidth - resolvedOriginPt;
      }
      if (alignment === 'right') {
        const resolvedTrailingEdgePt = state.contentX + contentWPt + signedIndentPt;
        return resolvedTrailingEdgePt;
      }
      const resolvedCenterPt = state.contentX + contentWPt / 2 + signedIndentPt;
      return 2 * Math.min(resolvedCenterPt, state.pageWidth - resolvedCenterPt);
    }),
  );
  const maximumTableWidthPt = isLeadingMarginPageStoryTable
    ? Math.max(contentWPt, pageFitCeilingPt)
    : contentWPt;
  const effectiveLayout = format.firstRowException?.layout === 'fixed'
    ? 'fixed'
    : table.layout;
  const isFixedNestedTable = effectiveLayout === 'fixed'
    && state.storyContext?.containers.some((container) => container.kind === 'tableCell');

  const intrinsicWidthsForTable = (
    owner: TableLayoutSource,
    ownerFormat = state.acquisitionInputs.tableFormatInput(owner),
  ): ((cell: DeepReadonly<DocTableCell>) => ReturnType<typeof measureTableCellIntrinsicWidths>) => {
    const ownerMargins = new WeakMap<object, Readonly<{ left: number; right: number }>>();
    owner.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
      const acquired = ownerFormat.rows[rowIndex]?.cells[cellIndex]?.marginsPt;
      ownerMargins.set(cell, acquired ?? effCellMargins(cell, owner));
    }));
    return (cell) => measureTableCellIntrinsicWidths(
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
          return measureParagraphIntrinsicWidths(
            paragraph,
            context,
            contentWPt,
            { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses },
            paragraphMeasurementEnvironment(state),
            numbering,
            { preserveWhitespaceOnlyContent: true },
          );
        },
        nestedTable: (nested) => measureTableIntrinsicWidths(
          state.acquisitionInputs.tableColumnLayoutInput(
            nested,
            contentWPt,
            intrinsicWidthsForTable(nested),
            contentWPt,
          ),
        ),
      },
    );
  };
  return [...resolveTableColumnWidths(state.acquisitionInputs.tableColumnLayoutInput(
    table,
    contentWPt,
    intrinsicWidthsForTable(table, format),
    isFixedNestedTable
      // ECMA-376 §17.18.87 resolves a fixed table from its authored tblGrid,
      // tblW, and tcW constraints. A containing cell is not an implicit tblW:
      // Word permits such a table to overflow its cell instead of scaling it.
      // `null` explicitly removes the unrelated physical cell ceiling; it is
      // not a numeric width and therefore cannot alter authored geometry.
      ? null
      : state.acquisitionInputs.tableParticipatesInOrdinaryFlow(table)
      ? maximumTableWidthPt
      : Math.max(contentWPt, state.pageWidth),
  ))];
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

/** Resolve a prepared body frame group and attach its retained member layouts. */
function resolveFrameBox(
  para: ParagraphLayoutSource,
  group: BodyFrameGroup<LayoutParagraphBlock>,
  state: BodyAcquisitionState,
  anchorLineHPt: number,
  onAcquired?: (acquired: ReturnType<typeof acquireRetainedFrameGroup>) => void,
): FrameBox {
  const measurer = { context: state.ctx, fontFamilyClasses: state.fontFamilyClasses };
  const environment = paragraphMeasurementEnvironment(state);
  const borderEdges = group.members.map(bodyParagraphBorderEdgesFor);
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
      state.contentX,
    );
    return physicalToLogicalAnchorBox(
      phys.x, phys.y, phys.w, phys.h, state.verticalPhys.physicalPageWidthPt,
    );
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
 *  not leak into a physical-frame layout (and vice-versa). */
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
      img.anchorYAlign, img.anchorYFromPara ?? false, img.anchorYPt ?? 0, h, state.contentX, phys,
      img.anchorYRelativeFrom ?? null, null, null,
    );
    const box = physicalToLogicalAnchorBox(
      px,
      py,
      w,
      h,
      state.verticalPhys.physicalPageWidthPt,
    );
    // Rotate the dist* padding one quarter-turn with the box: physical top/bottom
    // become logical left/right; physical right/left become logical top/bottom
    // (logical y runs opposite physical x). Symmetric wrapSquare dist is common,
    // but rotate the labels so asymmetric dist stays correct.
    return { x: box.x, y: box.y, w: box.w, h: box.h, dl: dt, dr: db, dt: dr, db: dl };
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

  const pdl = shape.distLeft ?? 0;
  const pdr = shape.distRight ?? 0;
  const pdt = shape.distTop ?? 0;
  const pdb = shape.distBottom ?? 0;
  // §17.6.20 — on a vertical page the box above is the LOGICAL projection of the
  // physically-resolved shape (resolveShapeBox), so rotate the dist* labels one
  // quarter-turn with it, exactly like the image path (resolveAnchorBox):
  // physical top/bottom ↦ logical left/right, physical right/left ↦ logical
  // top/bottom (logical y runs opposite physical x).
  const vertical = !!state.verticalPhys;
  const dl = vertical ? pdt : pdl;
  const dr = vertical ? pdb : pdr;
  const dt = vertical ? pdr : pdt;
  const db = vertical ? pdl : pdb;

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

  const kernel = buildConcreteBodyLayoutKernel(source, measureContext, resolvedLocalFonts);
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
