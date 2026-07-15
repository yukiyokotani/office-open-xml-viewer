import { LayoutInvariantError } from './diagnostics.js';
import { canonicalLogicalToPhysical, mapAffineRect, sameAffine } from './affine.js';
import { translateRect } from './retained-geometry-translation.js';
import { orderedPagePaintEntries, pageLayerNodes, PageGraphError } from './page-graph.js';
import type {
  DeepReadonly,
  DocumentLayout,
  DrawingPaintCommand,
  DrawingLayout,
  FlowDomain,
  LayoutRect,
  LayoutPage,
  PageSectionRegion,
  PaintNode,
  PointPt,
} from './types.js';

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

function requirePoint(point: PointPt, path: string): void {
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

function requireDrawingMLShapePlan(
  command: Extract<DrawingPaintCommand, { kind: 'drawingml-shape' }>,
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

function overlaps(a: LayoutRect, b: LayoutRect): boolean {
  return a.xPt < b.xPt + b.widthPt
    && b.xPt < a.xPt + a.widthPt
    && a.yPt < b.yPt + b.heightPt
    && b.yPt < a.yPt + a.heightPt;
}

function contains(outer: LayoutRect, inner: LayoutRect): boolean {
  return inner.xPt >= outer.xPt
    && inner.yPt >= outer.yPt
    && inner.xPt + inner.widthPt <= outer.xPt + outer.widthPt
    && inner.yPt + inner.heightPt <= outer.yPt + outer.heightPt;
}

function sameRect(left: LayoutRect, right: LayoutRect): boolean {
  return left.xPt === right.xPt && left.yPt === right.yPt
    && left.widthPt === right.widthPt && left.heightPt === right.heightPt;
}

function borderBounds(border: import('./types.js').BorderSegment): LayoutRect {
  const half = border.widthPt / 2;
  const left = Math.min(border.from.xPt, border.to.xPt) - half;
  const top = Math.min(border.from.yPt, border.to.yPt) - half;
  return {
    xPt: left,
    yPt: top,
    widthPt: Math.max(border.from.xPt, border.to.xPt) + half - left,
    heightPt: Math.max(border.from.yPt, border.to.yPt) + half - top,
  };
}

function drawingCommandBounds(command: DrawingPaintCommand): LayoutRect | null {
  if (command.kind === 'noop') return null;
  if (command.kind === 'drawingml-shape') {
    return {
      xPt: command.plan.rect.x,
      yPt: command.plan.rect.y,
      widthPt: command.plan.rect.w,
      heightPt: command.plan.rect.h,
    };
  }
  return command.rect;
}

function retainedPaintBounds(node: PaintNode): readonly LayoutRect[] {
  if (node.kind === 'drawing') {
    return [node.inkBounds, ...node.commands.flatMap((command) => {
      const bounds = drawingCommandBounds(command);
      return bounds ? [bounds] : [];
    })];
  }
  if (node.kind === 'paragraph') {
    return [
      node.inkBounds,
      ...node.borders.map(borderBounds),
      ...node.drawings.flatMap(retainedPaintBounds),
      ...node.textBoxes.flatMap(retainedPaintBounds),
    ];
  }
  if (node.kind === 'textbox') {
    return [node.inkBounds, ...node.paragraphs.flatMap(retainedPaintBounds)];
  }
  if (node.kind === 'table') {
    const bounds: LayoutRect[] = [node.inkBounds, ...node.borders.map(borderBounds)];
    for (const row of node.rows) for (const cell of row.cells) {
      if (cell.background) bounds.push(cell.flowBounds);
      for (const block of cell.blocks) {
        const target = {
          xPt: cell.contentBounds.xPt
            + (block.layout.kind === 'table' ? block.layout.flowBounds.xPt : 0),
          yPt: cell.flowBounds.yPt + block.offsetPt
            + (block.layout.kind === 'table' ? block.layout.flowBounds.yPt : 0),
        };
        const delta = {
          xPt: target.xPt - block.layout.flowBounds.xPt,
          yPt: target.yPt - block.layout.flowBounds.yPt,
        };
        bounds.push(...retainedPaintBounds(block.layout).map((item) => translateRect(item, delta)));
      }
    }
    return bounds;
  }
  return [node.inkBounds];
}

function ownsAnchorStart(cell: unknown, anchorBlockIndex: number): boolean {
  const ranges = (cell as { contentRanges?: ReadonlyArray<{
    kind: string;
    blockIndex: number;
    lineStart?: number;
    childFragmentIndex?: number;
  }> }).contentRanges;
  return ranges?.some((range) => range.blockIndex === anchorBlockIndex && (
    range.kind === 'whole'
    || (range.kind === 'paragraph' && range.lineStart === 0)
    || (range.kind === 'nested-table' && range.childFragmentIndex === 0)
  )) ?? false;
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
  if (node.kind === 'textbox') {
    node.paragraphs.forEach((paragraph) =>
      collectRetainedNodeIds(paragraph, pageIds, documentIds));
  }
}

interface AnchorOccurrenceOwner {
  readonly physicalPageIndex: number;
  readonly flowDomainId: string;
  readonly regionId: string;
}

function requireDrawingGeometry(
  node: DrawingLayout,
  path: string,
  owner: AnchorOccurrenceOwner,
): void {
  if (node.anchorLayer) {
    if (node.anchorLayer.coordinateSpace !== 'physical-page-points') {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `${path}.anchorLayer is not a normalized physical anchor`,
      );
    }
    const normalizedFor = node.anchorLayer.normalizedFor;
    if (!normalizedFor
      || normalizedFor.physicalPageIndex !== owner.physicalPageIndex
      || normalizedFor.flowDomainId !== owner.flowDomainId
      || normalizedFor.regionId !== owner.regionId) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `${path}.anchorLayer has invalid normalized anchor occurrence ownership`,
      );
    }
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

function requirePaintNodeDrawingGeometry(
  node: PaintNode,
  path: string,
  owner: AnchorOccurrenceOwner,
): void {
  if (node.kind === 'drawing') {
    requireDrawingGeometry(node, path, owner);
    return;
  }
  if (node.kind === 'paragraph') {
    node.drawings.forEach((drawing, index) =>
      requireDrawingGeometry(drawing, `${path}.drawings[${index}]`, owner));
    node.textBoxes.forEach((textBox, index) =>
      requirePaintNodeDrawingGeometry(textBox, `${path}.textBoxes[${index}]`, owner));
    return;
  }
  if (node.kind === 'table') {
    node.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => (
      cell.blocks.forEach((block, blockIndex) => requirePaintNodeDrawingGeometry(
        block.layout,
        `${path}.rows[${rowIndex}].cells[${cellIndex}].blocks[${blockIndex}]`,
        owner,
      ))
    )));
    return;
  }
  if (node.kind === 'textbox') {
    node.paragraphs.forEach((paragraph, index) =>
      requirePaintNodeDrawingGeometry(paragraph, `${path}.paragraphs[${index}]`, owner));
  }
}

export function assertDocumentLayout(layout: DocumentLayout): void {
  assertPlainData(layout, 'layout');
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
      page.geometry.contentTopPt < 0
      || page.geometry.contentTopPt > page.geometry.contentBottomPt
      || page.geometry.contentBottomPt > page.geometry.heightPt
    ) {
      throw new LayoutInvariantError(
        'INVALID_GEOMETRY',
        `pages[${pageIndex}] has invalid effective page edges`,
      );
    }

    const domains = new Map<string, FlowDomain>();
    page.flowDomains.forEach((domain, domainIndex) => {
      requireRect(domain.bounds, `pages[${pageIndex}].flowDomains[${domainIndex}].bounds`);
      if (domains.has(domain.id)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `duplicate flow domain ${domain.id}`);
      }
      domains.set(domain.id, domain);
    });

    if (page.parityBlank && (
      page.flowDomains.length > 0
      || page.sectionRegions.length > 0
      || pageLayerNodes(page).length > 0
      || page.layers.paintOrder.length > 0
      || page.readingOrder.length > 0
      || page.bookmarkStarts.length > 0
    )) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `pages[${pageIndex}] parity blank retains page content`,
      );
    }

    const sectionOccurrenceIds = new Set<string>();
    if (page.sectionOccurrenceId.length === 0) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `pages[${pageIndex}] has an empty section occurrence id`,
      );
    }
    sectionOccurrenceIds.add(page.sectionOccurrenceId);

    const regionByDomain = new Map<string, PageSectionRegion>();
    {
      const regionIds = new Set<string>();
      const bodyOwnership = new Map<string, number>();
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
        sectionOccurrenceIds.add(region.sectionOccurrenceId);
        if (!region.coordinateSpace) {
          throw new LayoutInvariantError('INVALID_REFERENCE', `${path} has no coordinate space`);
        }
        const expectedMatrix = canonicalLogicalToPhysical(
          region.coordinateSpace.writingMode,
          page.geometry.widthPt,
        );
        if (!sameAffine(region.coordinateSpace.logicalToPhysical, expectedMatrix)) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a noncanonical matrix`);
        }
        requireFinite(region.blockStartPt, `${path}.blockStartPt`);
        requireFinite(region.blockEndPt, `${path}.blockEndPt`);
        if (region.blockEndPt < region.blockStartPt) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a negative block extent`);
        }
        region.flowDomainIds.forEach((domainId) => {
          if (!domains.has(domainId)) {
            throw new LayoutInvariantError('INVALID_REFERENCE', `${path} references missing flow domain ${domainId}`);
          }
          bodyOwnership.set(domainId, (bodyOwnership.get(domainId) ?? 0) + 1);
          regionByDomain.set(domainId, region);
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
    }

    {
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

    const ordinary: Array<Readonly<{ node: PaintNode; bounds: LayoutRect }>> = [];
    let paintEntries;
    try {
      paintEntries = orderedPagePaintEntries(page);
    } catch (error) {
      if (error instanceof PageGraphError) {
        throw new LayoutInvariantError('INVALID_REFERENCE', error.message);
      }
      throw error;
    }
    const nodes = new Map<string, PaintNode>();
    const retainedNodeIds = new Set<string>();
    paintEntries.forEach((entry, nodeIndex) => {
      const { node } = entry;
      const path = `pages[${pageIndex}].nodes[${nodeIndex}]`;
      nodes.set(node.id, node);
      if (node.kind === 'table' && !node.paintReadyFloatingTables) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `${node.id} has no paint-ready floating-table ownership`,
        );
      }
      collectRetainedNodeIds(node, retainedNodeIds, documentRetainedNodeIds);
      if (node.kind === 'table') {
        switch (node.paintReadyFloatingTables.kind) {
          case 'none':
            break;
          case 'resolved':
            node.paintReadyFloatingTables.unresolved.forEach((placement) =>
              collectRetainedNodeIds(placement.child, retainedNodeIds, documentRetainedNodeIds));
            node.paintReadyFloatingTables.placements.forEach((placement) =>
              collectRetainedNodeIds(placement.child, retainedNodeIds, documentRetainedNodeIds));
            break;
          default:
            throw new LayoutInvariantError(
              'INVALID_REFERENCE',
              `${node.id} has unknown paint-ready floating-table kind`,
            );
        }
      }
      requireRect(node.flowBounds, `${path}.flowBounds`);
      requireRect(node.inkBounds, `${path}.inkBounds`);
      if (node.clipBounds) requireRect(node.clipBounds, `${path}.clipBounds`);
      requireFinite(node.advancePt, `${path}.advancePt`);
      const domain = domains.get(node.flowDomainId);
      if (!domain) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `${node.id} references missing flow domain ${node.flowDomainId}`);
      }
      const region = regionByDomain.get(node.flowDomainId);
      requirePaintNodeDrawingGeometry(node, path, {
        physicalPageIndex: page.pageIndex,
        flowDomainId: node.flowDomainId,
        regionId: region?.id ?? `page-layer:${entry.layer}`,
      });
      let physicalBounds = node.flowBounds;
      if (entry.coordinateSpace === 'logical-body-points') {
        if (!region) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${node.id} has no owning logical section region`,
          );
        }
        physicalBounds = mapAffineRect(
          region.coordinateSpace.logicalToPhysical,
          node.flowBounds,
        );
      }
      if (entry.coordinateSpace === 'upright-physical-page-points') {
        if (node.kind !== 'table') {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${node.id} uses upright physical coordinates without a table root`,
          );
        }
        if (region?.coordinateSpace.writingMode === 'vertical-lr') {
          // Upright root-table placement is an observed vertical-rl Office
          // compatibility contract; no mirrored vertical-lr evidence exists.
          throw new LayoutInvariantError(
            'UNSUPPORTED_FEATURE',
            `${node.id} has unsupported vertical-lr upright table placement`,
          );
        }
      }
      if (node.kind === 'table' && node.paintReadyFloatingTables.kind === 'resolved') {
        const floating = node.paintReadyFloatingTables;
        const expected = entry.coordinateSpace === 'upright-physical-page-points'
          ? 'upright-physical-page-points'
          : 'logical-page-points';
        if (floating.coordinateSpace !== expected) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${node.id} has mismatched floating-table coordinate space`,
          );
        }
        if (!region) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${node.id} floating tables have no owning section region`,
          );
        }
        const cells = new Map(node.rows.flatMap((row) => row.cells.map((cell) => [cell.id, cell] as const)));
        const occurrences = new Set<string>();
        const physicalFloatRect = (bounds: LayoutRect): LayoutRect => (
          floating.coordinateSpace === 'logical-page-points'
            ? mapAffineRect(region.coordinateSpace.logicalToPhysical, bounds)
            : bounds
        );
        const validateSource = (
          source: typeof floating.unresolved[number],
          sourcePath: string,
        ): void => {
          const host = cells.get(source.hostCellId);
          if (source.physicalPageIndex !== page.pageIndex
            || source.displayPageNumber !== page.pageNumber.displayNumber
            || source.tableId !== source.child.id
            || source.child.flowDomainId !== node.flowDomainId
            || !host
            || !Number.isInteger(source.sourceBlockIndex)
            || source.sourceBlockIndex < 0
            || !Number.isInteger(source.anchorBlockIndex)
            || source.anchorBlockIndex < 0
            || source.anchorBlockIndex <= source.sourceBlockIndex) {
            throw new LayoutInvariantError(
              'INVALID_REFERENCE',
              `${sourcePath} has invalid unresolved floating-table ownership`,
            );
          }
          const sourceProof = host.floatingSourceBlocks?.some((proof) => (
            proof.sourceBlockIndex === source.sourceBlockIndex
            && proof.tableId === source.tableId
          )) ?? false;
          if (!sourceProof) {
            throw new LayoutInvariantError(
              'INVALID_REFERENCE',
              `${sourcePath} has no floating source reference`,
            );
          }
          if (!ownsAnchorStart(host, source.anchorBlockIndex)) {
            throw new LayoutInvariantError(
              'INVALID_REFERENCE',
              `${sourcePath} has no floating anchor reference`,
            );
          }
          requireRect(source.anchorBounds, `${sourcePath}.anchorBounds`);
          if (source.columnBounds) requireRect(source.columnBounds, `${sourcePath}.columnBounds`);
          if (!contains(domain.bounds, physicalFloatRect(source.anchorBounds))
            || (source.columnBounds
              && !contains(domain.bounds, physicalFloatRect(source.columnBounds)))) {
            throw new LayoutInvariantError(
              'FLOW_DOMAIN_INVASION',
              `${sourcePath} crosses floating-table destination domain`,
            );
          }
        };
        for (const placement of floating.placements) {
          const source = placement.source;
          validateSource(source, `${placement.occurrenceId}.source`);
          requirePaintNodeDrawingGeometry(
            placement.child,
            `${placement.occurrenceId}.child`,
            {
              physicalPageIndex: page.pageIndex,
              flowDomainId: node.flowDomainId,
              regionId: region.id,
            },
          );
          if (occurrences.has(placement.occurrenceId)
            || placement.occurrenceId !== source.occurrenceId
            || source.tableId !== placement.child.id
            || source.child !== placement.child
            || placement.child.flowDomainId !== node.flowDomainId
            || source.child.flowDomainId !== node.flowDomainId
            || placement.xPt !== placement.bounds.xPt
            || placement.yPt !== placement.bounds.yPt
            || placement.overlap !== source.overlap) {
            throw new LayoutInvariantError(
              'INVALID_REFERENCE',
              `${placement.occurrenceId} has invalid floating-table destination ownership`,
            );
          }
          occurrences.add(placement.occurrenceId);
          requireRect(placement.bounds, `${placement.occurrenceId}.bounds`);
          requireRect(placement.exclusionBounds, `${placement.occurrenceId}.exclusionBounds`);
          const placementDelta = {
            xPt: placement.xPt - placement.child.flowBounds.xPt,
            yPt: placement.yPt - placement.child.flowBounds.yPt,
          };
          const paintedFlow = translateRect(placement.child.flowBounds, placementDelta);
          const paintedInk = translateRect(placement.child.inkBounds, placementDelta);
          if (!sameRect(placement.bounds, paintedFlow)) {
            throw new LayoutInvariantError(
              'INVALID_GEOMETRY',
              `${placement.occurrenceId} declared bounds do not match its paint extent`,
            );
          }
          const physicalPlacement = physicalFloatRect(paintedFlow);
          const physicalInk = physicalFloatRect(paintedInk);
          const physicalExclusion = physicalFloatRect(placement.exclusionBounds);
          const physicalPaint = retainedPaintBounds(placement.child).map((bounds) => (
            physicalFloatRect(translateRect(bounds, placementDelta))
          ));
          if (!contains(domain.bounds, physicalPlacement)
            || !contains(domain.bounds, physicalInk)
            || !contains(domain.bounds, physicalExclusion)
            || physicalPaint.some((bounds) => !contains(domain.bounds, bounds))
          ) {
            throw new LayoutInvariantError(
              'FLOW_DOMAIN_INVASION',
              `${placement.occurrenceId} crosses floating-table destination domain`,
            );
          }
        }
        for (const unresolved of floating.unresolved) {
          validateSource(unresolved, unresolved.occurrenceId);
          if (occurrences.has(unresolved.occurrenceId)) {
            throw new LayoutInvariantError(
              'INVALID_REFERENCE',
              `${unresolved.occurrenceId} has invalid unresolved floating-table ownership`,
            );
          }
          occurrences.add(unresolved.occurrenceId);
        }
      }
      if (!node.ordinaryFlow) return;
      if (entry.layer === 'body') {
        const footprint = entry.logicalBlock;
        if (!footprint
          || !Number.isFinite(footprint.blockStartPt)
          || !Number.isFinite(footprint.blockExtentPt)
          || footprint.blockExtentPt < 0) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${node.id} has an invalid logical block footprint`,
          );
        }
        if (region && (
          footprint.blockStartPt < region.blockStartPt
          || footprint.blockStartPt + footprint.blockExtentPt > region.blockEndPt
        )) {
          throw new LayoutInvariantError(
            'FLOW_DOMAIN_INVASION',
            `${node.id} logical block footprint crosses region ${region.id}`,
          );
        }
      }
      if (!contains(domain.bounds, physicalBounds)) {
        throw new LayoutInvariantError('FLOW_DOMAIN_INVASION', `${node.id} crosses flow domain ${domain.id}`);
      }
      ordinary.push({ node, bounds: physicalBounds });
    });

    const read = new Set<string>();
    page.readingOrder.forEach((nodeId) => {
      if (!nodes.has(nodeId) || read.has(nodeId)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `invalid reading-order reference ${nodeId}`);
      }
      read.add(nodeId);
    });

    const bookmarkNames = new Set<string>();
    page.bookmarkStarts.forEach((bookmark) => {
      if (
        bookmark.name.length === 0
        || bookmarkNames.has(bookmark.name)
        || !retainedNodeIds.has(bookmark.nodeId)
      ) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `invalid bookmark node ${bookmark.nodeId}`,
        );
      }
      if (!sectionOccurrenceIds.has(bookmark.sectionOccurrenceId)) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `bookmark ${bookmark.name} has an invalid section owner`,
        );
      }
      bookmarkNames.add(bookmark.name);
    });

    for (let index = 0; index < ordinary.length; index += 1) {
      for (let other = index + 1; other < ordinary.length; other += 1) {
        const first = ordinary[index];
        const second = ordinary[other];
        if (first && second
          && first.node.flowDomainId === second.node.flowDomainId
          && overlaps(first.bounds, second.bounds)) {
          throw new LayoutInvariantError('FLOW_OVERLAP', `${first.node.id} overlaps ${second.node.id}`);
        }
      }
    }
  });
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
    return Object.fromEntries(Object.entries(value as Record<string, unknown>)
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
  if (value === null || typeof value !== 'object') return value as DeepReadonly<T>;
  if (seen.has(value)) return value as DeepReadonly<T>;
  seen.add(value);
  for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child, seen);
  return Object.freeze(value) as DeepReadonly<T>;
}

export function deepFreezeDocumentLayout(layout: DocumentLayout): DeepReadonly<DocumentLayout> {
  assertPlainData(layout, 'layout');
  return deepFreeze(layout, new WeakSet<object>());
}
