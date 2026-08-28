import { LayoutInvariantError } from './diagnostics.js';
import {
  createSectionRegionCoordinateSpace,
  logicalPageExtent,
  transformRect,
  uprightPhysicalExtent,
  writingModeFromTextDirection,
} from './coordinate-space.js';
import { columnSeparatorSegments } from './column-separators.js';
import { orderedPagePaintNodes, pageLayerNodes, PageGraphError } from './page-graph.js';
import {
  deriveRetainedPageBookmarkStarts,
  sectionLayoutContextsEqual,
} from './page-factory.js';
import type {
  DeepReadonly,
  DocumentLayout,
  DrawingPaintCommand,
  DrawingLayout,
  FlowDomain,
  LayoutDiagnosticCode,
  LayoutPage,
  LayoutRect,
  PageSectionRegion,
  PaintNode,
  PointPt,
  SourceRef,
  SectionRegionCoordinateSpace,
  WritingMode,
} from './types.js';
import { unionLayoutRects } from './rect-union.js';

import { documentLayoutValidationEnabled } from './validation-policy.js';
const LAYOUT_DIAGNOSTIC_CODE_MEMBERS = {
  FLOW_OVERLAP: true,
  BOTTOM_MARGIN_INVASION: true,
  FLOW_DOMAIN_INVASION: true,
  INVALID_REFERENCE: true,
  INVALID_GEOMETRY: true,
  INVALID_VALUE: true,
  MISSING_RESOURCE: true,
  NON_CONVERGENCE: true,
  UNSUPPORTED_FEATURE: true,
} as const satisfies Readonly<Record<LayoutDiagnosticCode, true>>;

const SOURCE_STORY_MEMBERS = {
  body: true,
  header: true,
  footer: true,
  footnote: true,
  endnote: true,
  textbox: true,
} as const satisfies Readonly<Record<SourceRef['story'], true>>;

const LAYOUT_DIAGNOSTIC_CODES = new Set<LayoutDiagnosticCode>(
  Object.keys(LAYOUT_DIAGNOSTIC_CODE_MEMBERS) as LayoutDiagnosticCode[],
);
const SOURCE_STORIES = new Set<SourceRef['story']>(
  Object.keys(SOURCE_STORY_MEMBERS) as SourceRef['story'][],
);

function assertPlainData(value: unknown, path: string, ancestors = new WeakSet<object>()): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not finite`);
    }
    return;
  }
  if (typeof value !== 'object') {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} contains ${typeof value}`);
  }
  if (ancestors.has(value)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} contains a cycle`);
  }

  ancestors.add(value);
  try {
    if (Array.isArray(value)) {
      let indexCount = 0;
      for (const key of Reflect.ownKeys(value)) {
        if (key === 'length') continue;
        if (typeof key !== 'string') {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a symbol key`);
        }
        const index = Number(key);
        if (!Number.isInteger(index) || index < 0 || String(index) !== key || index >= value.length) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.${key} is not an array index`);
        }
        const descriptor = Object.getOwnPropertyDescriptor(value, key);
        if (!descriptor?.enumerable || !('value' in descriptor)) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}[${key}] is not plain data`);
        }
        assertPlainData(descriptor.value, `${path}[${key}]`, ancestors);
        indexCount += 1;
      }
      if (indexCount !== value.length) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is sparse`);
      }
      return;
    }

    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a plain record`);
    }
    for (const key of Reflect.ownKeys(value)) {
      if (typeof key !== 'string') {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a symbol key`);
      }
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (!descriptor?.enumerable || !('value' in descriptor)) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.${key} is not plain data`);
      }
      assertPlainData(descriptor.value, `${path}.${key}`, ancestors);
    }
  } finally {
    ancestors.delete(value);
  }
}

function requireFinite(value: number, path: string): void {
  if (!Number.isFinite(value)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not finite`);
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function requirePoint(point: PointPt, path: string): void {
  if (!isRecord(point)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a point`);
  }
  requireFinite(point.xPt, `${path}.xPt`);
  requireFinite(point.yPt, `${path}.yPt`);
}

function requireRect(rect: LayoutRect, path: string): void {
  requirePoint(rect, path);
  requireFinite(rect.widthPt, `${path}.widthPt`);
  requireFinite(rect.heightPt, `${path}.heightPt`);
  if (rect.widthPt < 0 || rect.heightPt < 0) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a negative extent`);
  }
}

function requireMatrix(matrix: unknown, path: string): asserts matrix is SectionRegionCoordinateSpace['logicalToPhysical'] {
  if (!isRecord(matrix)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a matrix`);
  }
  for (const coefficient of ['a', 'b', 'c', 'd', 'e', 'f'] as const) {
    requireFinite(matrix[coefficient] as number, `${path}.${coefficient}`);
  }
}

function requireWritingMode(value: unknown, path: string): asserts value is WritingMode {
  if (value !== 'horizontal-tb' && value !== 'vertical-rl' && value !== 'vertical-lr') {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is unsupported`);
  }
}

function requireCoordinateSpace(
  value: unknown,
  path: string,
): asserts value is SectionRegionCoordinateSpace {
  if (!isRecord(value)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a coordinate space`);
  }
  requireWritingMode(value.writingMode, `${path}.writingMode`);
  requireMatrix(value.logicalToPhysical, `${path}.logicalToPhysical`);
  requireMatrix(value.physicalToLogical, `${path}.physicalToLogical`);
}

function requireDrawingMLShapePlan(
  command: Extract<DrawingPaintCommand, { kind: 'drawingml-shape' | 'drawingml-image-fill' }>,
  path: string,
): void {
  const { plan } = command;
  assertPlainData(plan, `${path}.plan`);
  requireFinite(plan.rect.x, `${path}.plan.rect.x`);
  requireFinite(plan.rect.y, `${path}.plan.rect.y`);
  requireFinite(plan.rect.w, `${path}.plan.rect.w`);
  requireFinite(plan.rect.h, `${path}.plan.rect.h`);
  if (plan.rect.w < 0 || plan.rect.h < 0) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.plan.rect has a negative extent`);
  }
  requireFinite(plan.transform.rotationDeg, `${path}.plan.transform.rotationDeg`);
  if (plan.geometry.kind === 'preset') {
    if (plan.geometry.name.length === 0) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.plan.geometry.name is empty`);
    }
    plan.geometry.adjustments.forEach((adjustment, index) => {
      if (adjustment !== null) {
        requireFinite(adjustment, `${path}.plan.geometry.adjustments[${index}]`);
      }
    });
  } else {
    plan.geometry.subpaths.forEach((subpath, subpathIndex) => {
      subpath.forEach((pathCommand, commandIndex) => {
        if (pathCommand.cmd.length === 0) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path}.plan.geometry.subpaths[${subpathIndex}][${commandIndex}].cmd is empty`,
          );
        }
      });
    });
  }
  if (plan.stroke) {
    requireFinite(plan.stroke.width, `${path}.plan.stroke.width`);
    if (plan.stroke.width < 0) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.plan.stroke.width is negative`);
    }
  }
}

/** Inclusive geometry comparisons must tolerate the single-ULP drift produced
 * when the same authored boundary is reconstructed through different retained
 * subtraction/addition paths. This is floating-point precision, not a layout
 * tolerance: a geometrically distinct point value remains outside. */
function atMostWithinFloatingPrecision(left: number, right: number): boolean {
  if (left <= right) return true;
  const precision = Number.EPSILON * Math.max(1, Math.abs(left), Math.abs(right));
  return left - right <= precision;
}

function strictlyLessBeyondFloatingPrecision(left: number, right: number): boolean {
  return left < right && !atMostWithinFloatingPrecision(right, left);
}

function overlaps(a: LayoutRect, b: LayoutRect): boolean {
  return strictlyLessBeyondFloatingPrecision(a.xPt, b.xPt + b.widthPt)
    && strictlyLessBeyondFloatingPrecision(b.xPt, a.xPt + a.widthPt)
    && strictlyLessBeyondFloatingPrecision(a.yPt, b.yPt + b.heightPt)
    && strictlyLessBeyondFloatingPrecision(b.yPt, a.yPt + a.heightPt);
}

function contains(outer: LayoutRect, inner: LayoutRect): boolean {
  return atMostWithinFloatingPrecision(outer.xPt, inner.xPt)
    && atMostWithinFloatingPrecision(outer.yPt, inner.yPt)
    && atMostWithinFloatingPrecision(
      inner.xPt + inner.widthPt,
      outer.xPt + outer.widthPt,
    )
    && atMostWithinFloatingPrecision(
      inner.yPt + inner.heightPt,
      outer.yPt + outer.heightPt,
    );
}

function containsInlineExtent(outer: LayoutRect, inner: LayoutRect): boolean {
  return atMostWithinFloatingPrecision(outer.xPt, inner.xPt)
    && atMostWithinFloatingPrecision(
      inner.xPt + inner.widthPt,
      outer.xPt + outer.widthPt,
    );
}

function containsBlockInterval(
  blockStartPt: number,
  blockEndPt: number,
  inner: LayoutRect,
): boolean {
  return atMostWithinFloatingPrecision(blockStartPt, inner.yPt)
    && atMostWithinFloatingPrecision(inner.yPt + inner.heightPt, blockEndPt);
}

function containsBlockExtent(outer: LayoutRect, inner: LayoutRect): boolean {
  return atMostWithinFloatingPrecision(outer.yPt, inner.yPt)
    && atMostWithinFloatingPrecision(
      inner.yPt + inner.heightPt,
      outer.yPt + outer.heightPt,
    );
}

function equalRect(left: LayoutRect, right: LayoutRect): boolean {
  return left.xPt === right.xPt
    && left.yPt === right.yPt
    && left.widthPt === right.widthPt
    && left.heightPt === right.heightPt;
}

function equalMatrix(
  left: PageSectionRegion['coordinateSpace']['logicalToPhysical'],
  right: PageSectionRegion['coordinateSpace']['logicalToPhysical'],
): boolean {
  return left.a === right.a && left.b === right.b && left.c === right.c
    && left.d === right.d && left.e === right.e && left.f === right.f;
}

function requirePageBorder(page: LayoutPage, path: string): void {
  const pageBorder = page.pageBorder;
  if (pageBorder === null) return;
  if (pageBorder.zOrder !== 'front' && pageBorder.zOrder !== 'back') {
    throw new LayoutInvariantError('INVALID_REFERENCE', `${path}.zOrder is invalid`);
  }
  requireMatrix(pageBorder.logicalToPhysical, `${path}.logicalToPhysical`);
  const expectedTransform = createSectionRegionCoordinateSpace(
    writingModeFromTextDirection(page.section.textDirection),
    page.geometry,
  ).logicalToPhysical;
  if (!equalMatrix(pageBorder.logicalToPhysical, expectedTransform)) {
    throw new LayoutInvariantError(
      'INVALID_GEOMETRY',
      `${path}.logicalToPhysical contradicts the page-start section`,
    );
  }
  if (!Array.isArray(pageBorder.segments) || pageBorder.segments.length === 0) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.segments is empty`);
  }
  pageBorder.segments.forEach((segment, index) => {
    const segmentPath = `${path}.segments[${index}]`;
    requirePoint(segment.from, `${segmentPath}.from`);
    requirePoint(segment.to, `${segmentPath}.to`);
    requireFinite(segment.widthPt, `${segmentPath}.widthPt`);
    if (
      (
        segment.from.xPt !== segment.to.xPt
        && segment.from.yPt !== segment.to.yPt
      )
    ) {
      throw new LayoutInvariantError(
        'INVALID_GEOMETRY',
        `${segmentPath} is not an axis-aligned page edge`,
      );
    }
    if (!/^#[0-9a-fA-F]{6}$/.test(segment.color)) {
      throw new LayoutInvariantError('INVALID_REFERENCE', `${segmentPath}.color is invalid`);
    }
  });
}

function retainUniqueNodeId(
  id: string,
  pageIds: Set<string>,
  documentIds: Set<string>,
): void {
  if (documentIds.has(id)) {
    throw new LayoutInvariantError('INVALID_REFERENCE', `duplicate retained node id ${id}`);
  }
  documentIds.add(id);
  pageIds.add(id);
}

function collectRetainedNodeIds(
  node: PaintNode,
  pageIds: Set<string>,
  documentIds: Set<string>,
): void {
  retainUniqueNodeId(node.id, pageIds, documentIds);
  if (node.kind === 'paragraph') {
    node.drawings.forEach((drawing) =>
      collectRetainedNodeIds(drawing, pageIds, documentIds));
    node.textBoxes.forEach((textBox) =>
      collectRetainedNodeIds(textBox, pageIds, documentIds));
    return;
  }
  if (node.kind === 'table') {
    node.rows.forEach((row) => {
      retainUniqueNodeId(row.id, pageIds, documentIds);
      row.cells.forEach((cell) => {
        retainUniqueNodeId(cell.id, pageIds, documentIds);
        cell.blocks.forEach((block) =>
          collectRetainedNodeIds(block.layout, pageIds, documentIds));
      });
    });
    return;
  }
  if (node.kind === 'note') {
    node.story.blocks.forEach((block) =>
      collectRetainedNodeIds(block, pageIds, documentIds));
    return;
  }
  if (node.kind === 'textbox') {
    node.story.blocks.forEach((block) =>
      collectRetainedNodeIds(block, pageIds, documentIds));
  }
}

function requireRetainedCollisionGeometry(node: PaintNode, path: string): void {
  if (node.kind === 'paragraph') {
    const expectedCellContainmentBounds = unionLayoutRects(
      node.drawings
        .filter((drawing) => drawing.anchorLayer?.cellContainment === true)
        .map((drawing) => drawing.flowBounds),
    );
    if (node.cellContainmentBounds) {
      requireRect(node.cellContainmentBounds, `${path}.cellContainmentBounds`);
    }
    if (
      (expectedCellContainmentBounds === null) !== (node.cellContainmentBounds === undefined)
      || (
        expectedCellContainmentBounds
        && node.cellContainmentBounds
        && !equalRect(expectedCellContainmentBounds, node.cellContainmentBounds)
      )
    ) {
      throw new LayoutInvariantError(
        'INVALID_GEOMETRY',
        `${path}.cellContainmentBounds does not match its retained layoutInCell drawings`,
      );
    }
    const occurrenceIds = new Set<string>();
    (node.anchorCollisions ?? []).forEach((entry, index) => {
      const entryPath = `${path}.anchorCollisions[${index}]`;
      if (entry.occurrenceId.length === 0 || occurrenceIds.has(entry.occurrenceId)) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `${entryPath}.occurrenceId is empty or duplicated`,
        );
      }
      occurrenceIds.add(entry.occurrenceId);
      requireRect(entry.bounds, `${entryPath}.bounds`);
      if (
        (entry.horizontalOwnership !== 'page' && entry.horizontalOwnership !== 'host')
        || (entry.verticalOwnership !== 'page' && entry.verticalOwnership !== 'host')
      ) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `${entryPath} has invalid axis ownership`,
        );
      }
    });
    node.textBoxes.forEach((textBox, index) =>
      requireRetainedCollisionGeometry(textBox, `${path}.textBoxes[${index}]`));
    return;
  }
  if (node.kind === 'table') {
    node.rows.forEach((row, rowIndex) =>
      row.cells.forEach((cell, cellIndex) =>
        cell.blocks.forEach((block, blockIndex) =>
          requireRetainedCollisionGeometry(
            block.layout,
            `${path}.rows[${rowIndex}].cells[${cellIndex}].blocks[${blockIndex}]`,
          ))));
    return;
  }
  if (node.kind === 'textbox') {
    node.story.blocks.forEach((block, index) =>
      requireRetainedCollisionGeometry(block, `${path}.story.blocks[${index}]`));
  }
}

function requireDrawingGeometry(node: DrawingLayout, path: string): void {
  if (node.orientation === 'upright-physical' && !node.transform) {
    throw new LayoutInvariantError(
      'INVALID_GEOMETRY',
      `${path} upright physical drawing is missing its logical transform`,
    );
  }
  if (node.transform) {
    for (const key of ['a', 'b', 'c', 'd', 'e', 'f'] as const) {
      requireFinite(node.transform[key], `${path}.transform.${key}`);
    }
  }
  if (node.clip?.kind === 'rect') requireRect(node.clip.rect, `${path}.clip.rect`);
  if (node.clip?.kind === 'polygon') {
    node.clip.points.forEach((point, index) => requirePoint(point, `${path}.clip.points[${index}]`));
  }
  node.commands.forEach((command, index) => {
    const commandPath = `${path}.commands[${index}]`;
    if (command.kind === 'noop') return;
    if (command.kind === 'drawingml-shape') {
      requireDrawingMLShapePlan(command, commandPath);
      return;
    }
    if (command.kind === 'drawingml-image-fill') {
      requireDrawingMLShapePlan(command, commandPath);
      if (command.resourceKey.length === 0) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${commandPath}.resourceKey is empty`);
      }
      if (command.fillRect) {
        for (const edge of ['l', 't', 'r', 'b'] as const) {
          requireFinite(command.fillRect[edge], `${commandPath}.fillRect.${edge}`);
        }
      }
      return;
    }
    requireRect(command.rect, `${commandPath}.rect`);
    if (command.kind === 'stroke-rect') {
      requireFinite(command.lineWidthPt, `${commandPath}.lineWidthPt`);
      command.dashPt.forEach((dash, dashIndex) =>
        requireFinite(dash, `${commandPath}.dashPt[${dashIndex}]`));
    }
    if (command.kind === 'text') {
      requireFinite(command.fontSizePt, `${commandPath}.fontSizePt`);
      requireFinite(command.fontWeight, `${commandPath}.fontWeight`);
    }
    if (command.kind === 'watermark-text') {
      requireRect(command.sourceBounds, `${commandPath}.sourceBounds`);
      if (command.sourceBounds.widthPt <= 0 || command.sourceBounds.heightPt <= 0) {
        throw new LayoutInvariantError(
          'INVALID_GEOMETRY',
          `${commandPath}.sourceBounds must have positive extents`,
        );
      }
      requireFinite(command.opacity, `${commandPath}.opacity`);
      requireFinite(command.rotationDeg, `${commandPath}.rotationDeg`);
      requireFinite(command.fontSizePt, `${commandPath}.fontSizePt`);
      if (command.opacity < 0 || command.opacity > 1 || command.fontSizePt <= 0) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${commandPath} has invalid textPath paint metrics`);
      }
      command.spans.forEach((span, spanIndex) => {
        requireFinite(span.advancePt, `${commandPath}.spans[${spanIndex}].advancePt`);
        requireFinite(span.fontWeight, `${commandPath}.spans[${spanIndex}].fontWeight`);
      });
    }
  });
}

function assertDocumentLayoutUnchecked(layout: DocumentLayout): void {
  assertPlainData(layout, 'layout');
  layout.diagnostics.forEach((diagnostic, index) => {
    const path = `diagnostics[${index}]`;
    if (!LAYOUT_DIAGNOSTIC_CODES.has(diagnostic.code)) {
      throw new LayoutInvariantError('INVALID_REFERENCE', `${path}.code is unknown`);
    }
    if (diagnostic.severity !== 'warning' && diagnostic.severity !== 'error') {
      throw new LayoutInvariantError('INVALID_REFERENCE', `${path}.severity is unknown`);
    }
    if (typeof diagnostic.message !== 'string' || diagnostic.message.length === 0) {
      throw new LayoutInvariantError('INVALID_REFERENCE', `${path}.message is empty`);
    }
    if (diagnostic.source !== undefined) {
      if (!SOURCE_STORIES.has(diagnostic.source.story)
        || typeof diagnostic.source.storyInstance !== 'string'
        || diagnostic.source.storyInstance.length === 0
        || !Array.isArray(diagnostic.source.path)
        || diagnostic.source.path.some((entry) =>
          !Number.isSafeInteger(entry) || entry < 0)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `${path}.source is invalid`);
      }
    }
  });
  const documentRetainedNodeIds = new Set<string>();
  layout.pages.forEach((page, pageIndex) => {
    if (!Number.isInteger(page.pageIndex) || page.pageIndex !== pageIndex) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `pages[${pageIndex}] has invalid page index ${page.pageIndex}`,
      );
    }
    requireRect(page.geometry, `pages[${pageIndex}].geometry`);
    requireFinite(page.geometry.contentTopPt, `pages[${pageIndex}].geometry.contentTopPt`);
    requireFinite(page.geometry.contentBottomPt, `pages[${pageIndex}].geometry.contentBottomPt`);
    if (
      page.geometry.widthPt <= 0
      || page.geometry.heightPt <= 0
      || page.geometry.contentTopPt < 0
      || page.geometry.contentTopPt > page.geometry.contentBottomPt
      || page.geometry.contentBottomPt > page.geometry.heightPt
    ) {
      throw new LayoutInvariantError(
        'INVALID_GEOMETRY',
        `pages[${pageIndex}] has invalid effective page edges`,
      );
    }
    requirePageBorder(page, `pages[${pageIndex}].pageBorder`);

    const domains = new Map<string, FlowDomain>();
    page.flowDomains.forEach((domain, domainIndex) => {
      requireRect(
        domain.logicalBounds,
        `pages[${pageIndex}].flowDomains[${domainIndex}].logicalBounds`,
      );
      requireRect(
        domain.physicalBounds,
        `pages[${pageIndex}].flowDomains[${domainIndex}].physicalBounds`,
      );
      if (domains.has(domain.id)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `duplicate flow domain ${domain.id}`);
      }
      domains.set(domain.id, domain);
    });

    if (page.parityBlank && (
      page.flowDomains.length > 0
      || (page.sectionRegions?.length ?? 0) > 0
      || (page.columnSeparators?.length ?? 0) > 0
      || pageLayerNodes(page).length > 0
      || page.layers.roots.length > 0
      || page.readingOrder.length > 0
      || (page.bookmarkStarts?.length ?? 0) > 0
    )) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `pages[${pageIndex}] parity blank retains page content`,
      );
    }

    const sectionOccurrenceIds = new Set<string>();
    if (page.sectionOccurrenceId !== undefined) {
      if (page.sectionOccurrenceId.length === 0) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `pages[${pageIndex}] has an empty section occurrence id`,
        );
      }
      sectionOccurrenceIds.add(page.sectionOccurrenceId);
    }

    const regionByDomain = new Map<string, PageSectionRegion>();
    if (page.sectionRegions) {
      const regionIds = new Set<string>();
      const occurrenceIds = new Set<string>();
      const bodyOwnership = new Map<string, number>();
      const occupiedPhysicalBodyDomains: Array<Readonly<{
        regionId: string;
        bounds: LayoutRect;
      }>> = [];
      let pageWritingMode: WritingMode | undefined;
      page.sectionRegions.forEach((region, regionIndex) => {
        const path = `pages[${pageIndex}].sectionRegions[${regionIndex}]`;
        if (region.id.length === 0 || regionIds.has(region.id)) {
          throw new LayoutInvariantError('INVALID_REFERENCE', `${path} has an invalid region id`);
        }
        regionIds.add(region.id);
        if (region.sectionOccurrenceId.length === 0) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${path} has an empty section occurrence id`,
          );
        }
        if (occurrenceIds.has(region.sectionOccurrenceId)) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${path} has a duplicate section occurrence id`,
          );
        }
        occurrenceIds.add(region.sectionOccurrenceId);
        sectionOccurrenceIds.add(region.sectionOccurrenceId);
        requireCoordinateSpace(region.coordinateSpace, `${path}.coordinateSpace`);
        const writingMode = region.coordinateSpace.writingMode;
        if (pageWritingMode !== undefined && pageWritingMode !== writingMode) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path} mixes coordinate systems on one physical page`,
          );
        }
        pageWritingMode = writingMode;
        let sectionWritingMode: WritingMode;
        try {
          sectionWritingMode = writingModeFromTextDirection(region.section.textDirection);
        } catch (error) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path}.section.textDirection is unsupported: ${(error as Error).message}`,
          );
        }
        if (writingMode !== sectionWritingMode) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path} writing mode contradicts its section text direction`,
          );
        }
        const logicalExtent = logicalPageExtent(page.geometry, writingMode);
        const physicalExtent = uprightPhysicalExtent({
          widthPt: region.section.geometry.pageWidth,
          heightPt: region.section.geometry.pageHeight,
        }, writingMode);
        if (physicalExtent.widthPt !== page.geometry.widthPt
          || physicalExtent.heightPt !== page.geometry.heightPt) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path} section geometry does not match the upright physical page`,
          );
        }
        requireFinite(region.blockStartPt, `${path}.blockStartPt`);
        requireFinite(region.blockEndPt, `${path}.blockEndPt`);
        if (
          region.columnFlowDirection !== 'ltr'
          && region.columnFlowDirection !== 'rtl'
        ) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path} has an invalid column flow direction`,
          );
        }
        const sectionColumnFlowDirection = region.section.sectionBidi === true ? 'rtl' : 'ltr';
        if (region.columnFlowDirection !== sectionColumnFlowDirection) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path} column flow direction contradicts its section bidi`,
          );
        }
        if (region.blockStartPt < 0
          || region.blockEndPt < region.blockStartPt
          || region.blockEndPt > logicalExtent.heightPt) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has an invalid block interval`);
        }
        const expectedCoordinateSpace = createSectionRegionCoordinateSpace(
          region.coordinateSpace.writingMode,
          page.geometry,
        );
        if (!equalMatrix(
          region.coordinateSpace.logicalToPhysical,
          expectedCoordinateSpace.logicalToPhysical,
        ) || !equalMatrix(
          region.coordinateSpace.physicalToLogical,
          expectedCoordinateSpace.physicalToLogical,
        )) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has an invalid coordinate transform`);
        }
        const columnIndexes = region.columnIndexes;
        if (
          region.flowDomainIds.length !== columnIndexes.length
          || columnIndexes.some((columnIndex, index) => (
            !Number.isInteger(columnIndex)
            || columnIndex < 0
            || columnIndex >= region.section.columns.length
            || (index > 0 && columnIndex <= columnIndexes[index - 1]!)
          ))
        ) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} columns contradict its section`);
        }
        let priorInlineEndPt = 0;
        region.flowDomainIds.forEach((domainId, columnPosition) => {
          const domain = domains.get(domainId);
          if (!domain) {
            throw new LayoutInvariantError('INVALID_REFERENCE', `${path} references missing flow domain ${domainId}`);
          }
          if (domain.kind !== 'body') {
            throw new LayoutInvariantError('INVALID_REFERENCE', `${path} owns non-body flow domain ${domainId}`);
          }
          bodyOwnership.set(domainId, (bodyOwnership.get(domainId) ?? 0) + 1);
          regionByDomain.set(domainId, region);
          const bounds = domain.logicalBounds;
          const sectionColumn = region.section.columns[columnIndexes[columnPosition]!];
          if (bounds.widthPt <= 0 || bounds.heightPt < 0
            || bounds.yPt !== region.blockStartPt
            || bounds.heightPt !== region.blockEndPt - region.blockStartPt
            || bounds.xPt < 0
            || bounds.xPt < priorInlineEndPt
            || bounds.xPt + bounds.widthPt > logicalExtent.widthPt
            || sectionColumn === undefined
            || bounds.xPt !== sectionColumn.xPt
            || bounds.widthPt !== sectionColumn.wPt) {
            throw new LayoutInvariantError(
              'INVALID_GEOMETRY',
              `${domainId} is not the section column's non-negative logical region`,
            );
          }
          priorInlineEndPt = bounds.xPt + bounds.widthPt;
          const expectedPhysicalBounds = transformRect(
            region.coordinateSpace.logicalToPhysical,
            domain.logicalBounds,
          );
          if (!equalRect(expectedPhysicalBounds, domain.physicalBounds)) {
            throw new LayoutInvariantError(
              'INVALID_GEOMETRY',
              `${domainId} physical bounds do not match its section region transform`,
            );
          }
          if (!contains(page.geometry, domain.physicalBounds)) {
            throw new LayoutInvariantError(
              'INVALID_GEOMETRY',
              `${domainId} physical bounds leave the upright physical page`,
            );
          }
          if (occupiedPhysicalBodyDomains.some((prior) => (
            prior.regionId !== region.id && overlaps(prior.bounds, domain.physicalBounds)
          ))) {
            throw new LayoutInvariantError(
              'INVALID_GEOMETRY',
              `${domainId} overlaps a body flow domain owned by another section region`,
            );
          }
          occupiedPhysicalBodyDomains.push({
            regionId: region.id,
            bounds: domain.physicalBounds,
          });
        });
      });
      page.flowDomains.filter((domain) => domain.kind === 'body').forEach((domain) => {
        if (bodyOwnership.get(domain.id) !== 1) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${domain.id} has invalid section region ownership`,
          );
        }
      });
      if (!page.parityBlank && page.sectionRegions.length > 0
      ) {
        const firstRegion = page.sectionRegions[0]!;
        if (page.sectionOccurrenceId !== firstRegion.sectionOccurrenceId) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `pages[${pageIndex}] page-start section occurrence does not match its first region`,
          );
        }
        if (!sectionLayoutContextsEqual(page.section, firstRegion.section)) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `pages[${pageIndex}] page-start section facts do not match its first region`,
          );
        }
      }
    }
    const expectedColumnSeparators = columnSeparatorSegments(page.sectionRegions ?? []);
    if (!Array.isArray(page.columnSeparators)
      || page.columnSeparators.length !== expectedColumnSeparators.length
      || page.columnSeparators.some((segment, index) => {
        const expected = expectedColumnSeparators[index];
        return expected === undefined
          || segment.start.xPt !== expected.start.xPt
          || segment.start.yPt !== expected.start.yPt
          || segment.end.xPt !== expected.end.xPt
          || segment.end.yPt !== expected.end.yPt;
      })) {
      throw new LayoutInvariantError(
        'INVALID_GEOMETRY',
        `pages[${pageIndex}].columnSeparators contradict the retained section regions`,
      );
    }
    const regionById = new Map(page.sectionRegions.map((region) => [region.id, region]));
    for (const domain of page.flowDomains) {
      if (domain.kind !== 'footnote' && domain.kind !== 'endnote') continue;
      const storyRegion = domain.sectionRegionId
        ? regionById.get(domain.sectionRegionId)
        : page.sectionRegions[0];
      if (!storyRegion) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `${domain.id} references missing page story region ${domain.sectionRegionId ?? '<default>'}`,
        );
      }
      const expectedPhysicalBounds = transformRect(
        storyRegion.coordinateSpace.logicalToPhysical,
        domain.logicalBounds,
      );
      if (!equalRect(expectedPhysicalBounds, domain.physicalBounds)) {
        throw new LayoutInvariantError(
          'INVALID_GEOMETRY',
          `${domain.id} physical bounds do not match the page story transform`,
        );
      }
      regionByDomain.set(domain.id, storyRegion);
    }
    for (const domain of page.flowDomains) {
      if (!regionByDomain.has(domain.id) && !equalRect(domain.logicalBounds, domain.physicalBounds)) {
        throw new LayoutInvariantError(
          'INVALID_GEOMETRY',
          `${domain.id} has unequal logical and physical bounds without a section region`,
        );
      }
    }

    if (page.pageNumber) {
      requireFinite(page.pageNumber.displayNumber, `pages[${pageIndex}].pageNumber.displayNumber`);
      if (!Number.isInteger(page.pageNumber.displayNumber)) {
        throw new LayoutInvariantError(
          'INVALID_GEOMETRY',
          `pages[${pageIndex}] page number is not an integer`,
        );
      }
      if (
        page.pageNumber.format.length === 0
        || !sectionOccurrenceIds.has(page.pageNumber.sectionOccurrenceId)
      ) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `pages[${pageIndex}] has an invalid page number section owner`,
        );
      }
    }

    const ordinary: PaintNode[] = [];
    try {
      orderedPagePaintNodes(page);
    } catch (error) {
      if (error instanceof PageGraphError) {
        throw new LayoutInvariantError('INVALID_REFERENCE', error.message);
      }
      throw error;
    }
    const nodes = new Map<string, PaintNode>();
    const retainedNodeIds = new Set<string>();
    pageLayerNodes(page).forEach(({ node }, nodeIndex) => {
      const path = `pages[${pageIndex}].nodes[${nodeIndex}]`;
      nodes.set(node.id, node);
      collectRetainedNodeIds(node, retainedNodeIds, documentRetainedNodeIds);
      requireRetainedCollisionGeometry(node, path);
      requireRect(node.flowBounds, `${path}.flowBounds`);
      requireRect(node.inkBounds, `${path}.inkBounds`);
      if (node.clipBounds) requireRect(node.clipBounds, `${path}.clipBounds`);
      requireFinite(node.advancePt, `${path}.advancePt`);
      if (node.kind === 'drawing') requireDrawingGeometry(node, path);
      const domain = domains.get(node.flowDomainId);
      if (!domain) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `${node.id} references missing flow domain ${node.flowDomainId}`);
      }
      if (node.ordinaryFlow && domain.kind === 'body' && domain.logicalBounds.heightPt === 0) {
        throw new LayoutInvariantError(
          'FLOW_DOMAIN_INVASION',
          `${node.id} claims ordinary flow in an empty body domain`,
        );
      }
      if (!node.ordinaryFlow) return;
      const bodyRegion = domain.kind === 'body' ? regionByDomain.get(domain.id) : undefined;
      if (domain.kind === 'body') {
        if (!bodyRegion) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${node.id} references a body flow domain without a section region`,
          );
        }
        if (node.flowBounds.yPt + node.flowBounds.heightPt > bodyRegion.blockEndPt) {
          throw new LayoutInvariantError('BOTTOM_MARGIN_INVASION', `${node.id} crosses logical block end`);
        }
      }
      // Signed w:tblInd and table justification may intentionally put an
      // ordinary table outside the inline text band. It still belongs to this
      // flow domain and must remain contained on the logical block axis.
      const insideFlowDomain = bodyRegion
        ? containsBlockInterval(bodyRegion.blockStartPt, bodyRegion.blockEndPt, node.flowBounds)
          && (node.kind === 'table' || containsInlineExtent(domain.logicalBounds, node.flowBounds))
        : node.kind === 'table'
          ? containsBlockExtent(domain.logicalBounds, node.flowBounds)
          : contains(domain.logicalBounds, node.flowBounds);
      if (!insideFlowDomain) {
        throw new LayoutInvariantError('FLOW_DOMAIN_INVASION', `${node.id} crosses flow domain ${domain.id}`);
      }
      ordinary.push(node);
    });

    const read = new Set<string>();
    page.readingOrder.forEach((nodeId) => {
      if (!nodes.has(nodeId) || read.has(nodeId)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `invalid reading-order reference ${nodeId}`);
      }
      read.add(nodeId);
    });

    if (page.bookmarkStarts !== undefined) {
      const sectionByDomain = new Map(
        [...regionByDomain].map(([domainId, region]) => [
          domainId,
          region.sectionOccurrenceId,
        ]),
      );
      const expectedBookmarkStarts = deriveRetainedPageBookmarkStarts(page, sectionByDomain);
      const bookmarkOwnersAreValid = expectedBookmarkStarts.every((bookmark) =>
        bookmark.sectionOccurrenceId.length > 0
        && sectionOccurrenceIds.has(bookmark.sectionOccurrenceId));
      const bookmarkMetadataMatches = page.bookmarkStarts.length === expectedBookmarkStarts.length
        && page.bookmarkStarts.every((bookmark, index) => {
          const expected = expectedBookmarkStarts[index];
          return expected !== undefined
            && bookmark.name === expected.name
            && bookmark.nodeId === expected.nodeId
            && bookmark.sectionOccurrenceId === expected.sectionOccurrenceId;
        });
      if (!bookmarkOwnersAreValid || !bookmarkMetadataMatches) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `pages[${pageIndex}] bookmark metadata does not match its retained graph (invalid bookmark node or ownership)`,
        );
      }
    }

    for (let index = 0; index < ordinary.length; index += 1) {
      for (let other = index + 1; other < ordinary.length; other += 1) {
        const first = ordinary[index];
        const second = ordinary[other];
        if (!first || !second) continue;
        const firstDomain = domains.get(first.flowDomainId);
        const secondDomain = domains.get(second.flowDomainId);
        const sameDomain = first.flowDomainId === second.flowDomainId;
        const bodyAndNote = (
          firstDomain?.kind === 'body'
          && (secondDomain?.kind === 'footnote' || secondDomain?.kind === 'endnote')
        ) || (
          secondDomain?.kind === 'body'
          && (firstDomain?.kind === 'footnote' || firstDomain?.kind === 'endnote')
        );
        const distinctNoteStories = firstDomain?.id !== secondDomain?.id
          && (firstDomain?.kind === 'footnote' || firstDomain?.kind === 'endnote')
          && (secondDomain?.kind === 'footnote' || secondDomain?.kind === 'endnote');
        if (
          (sameDomain || bodyAndNote || distinctNoteStories)
          && overlaps(first.flowBounds, second.flowBounds)
        ) {
          throw new LayoutInvariantError('FLOW_OVERLAP', `${first.id} overlaps ${second.id}`);
        }
      }
    }
  });
}

export function assertDocumentLayout(layout: DocumentLayout): void {
  try {
    assertDocumentLayoutUnchecked(layout);
  } catch (error) {
    if (error instanceof LayoutInvariantError) throw error;
    if (error instanceof TypeError || error instanceof RangeError) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', error.message);
    }
    throw error;
  }
}

function isLayoutDiagnosticRecord(
  value: Readonly<Record<string, unknown>>,
): boolean {
  return typeof value.code === 'string'
    && LAYOUT_DIAGNOSTIC_CODES.has(value.code as LayoutDiagnosticCode)
    && (value.severity === 'warning' || value.severity === 'error')
    && typeof value.message === 'string';
}

function canonicalize(value: unknown): unknown {
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new LayoutInvariantError('INVALID_GEOMETRY', 'fingerprint input is not finite');
    const normalized = Number(value.toFixed(6));
    return Object.is(normalized, -0) ? 0 : normalized;
  }
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value;
  if (Array.isArray(value)) return value.map((entry) => canonicalize(entry));
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>;
    const diagnostic = isLayoutDiagnosticRecord(record);
    return Object.fromEntries(Object.entries(record)
      .filter(([entryKey]) => !(diagnostic && entryKey === 'message'))
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([entryKey, entry]) => [entryKey, canonicalize(entry)]));
  }
  throw new LayoutInvariantError('INVALID_GEOMETRY', `fingerprint contains ${typeof value}`);
}

export function layoutFingerprint(layout: DocumentLayout): string {
  assertPlainData(layout, 'layout');
  const value = {
    pages: layout.pages,
    diagnostics: layout.diagnostics.map(({ message: _message, ...identity }) => identity),
  };
  return JSON.stringify(canonicalize(value));
}

function deepFreeze<T>(value: T, seen: WeakSet<object>): DeepReadonly<T> {
  if (value === null || typeof value !== 'object') {
    // Non-finite geometry is fatal state (see docs/docx-layout-engine-redesign
    // "Convergence and Errors"): the check rides the freeze walk that always
    // runs, so no retained layout can seal a NaN even on the paths that skip
    // the development-only plain-data pre-pass.
    if (typeof value === 'number' && !Number.isFinite(value)) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', 'retained layout contains a non-finite number');
    }
    return value as DeepReadonly<T>;
  }
  if (seen.has(value)) return value as DeepReadonly<T>;
  seen.add(value);
  for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child, seen);
  return Object.freeze(value) as DeepReadonly<T>;
}

const frozenDocumentLayouts = new WeakSet<object>();
const verifiedFrozenDocumentLayouts = new WeakSet<object>();

function freezeDocumentLayout(layout: DocumentLayout): DeepReadonly<DocumentLayout> {
  if (frozenDocumentLayouts.has(layout)) {
    return layout as DeepReadonly<DocumentLayout>;
  }
  const frozen = deepFreeze(layout, new WeakSet<object>());
  frozenDocumentLayouts.add(frozen);
  return frozen;
}

export function deepFreezeDocumentLayout(layout: DocumentLayout): DeepReadonly<DocumentLayout> {
  if (frozenDocumentLayouts.has(layout)) {
    return layout as DeepReadonly<DocumentLayout>;
  }
  // Freezing (and its fused non-finite check) is unconditional; the
  // path-precise plain-data traversal that precedes it is a development-time
  // diagnostic (see validation-policy.ts).
  if (documentLayoutValidationEnabled()) assertPlainData(layout, 'layout');
  return freezeDocumentLayout(layout);
}

/** Validate the complete retained-layout contract and freeze the same accepted
 * graph without repeating the plain-data traversal.
 *
 * The invariant suite runs UNCONDITIONALLY: non-finite or negative geometry,
 * invalid ownership, and broken layout invariants are fatal state per the
 * layout engine's error contract, with no test/production split. The
 * validation-policy switch gates only redundant path-precise diagnostics
 * elsewhere — never the detection of fatal state. */
export function assertAndDeepFreezeDocumentLayout(
  layout: DocumentLayout,
): DeepReadonly<DocumentLayout> {
  if (verifiedFrozenDocumentLayouts.has(layout)) {
    return layout as DeepReadonly<DocumentLayout>;
  }
  assertDocumentLayout(layout);
  const frozen = freezeDocumentLayout(layout);
  verifiedFrozenDocumentLayouts.add(frozen);
  return frozen;
}
