import { LayoutInvariantError } from './diagnostics.js';
import { orderedPagePaintNodes, pageLayerNodes, PageGraphError } from './page-graph.js';
import type {
  DeepReadonly,
  DocumentLayout,
  DrawingPaintCommand,
  DrawingLayout,
  FlowDomain,
  LayoutRect,
  LayoutPage,
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

function requireDrawingGeometry(node: DrawingLayout, path: string): void {
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
      || (page.sectionRegions?.length ?? 0) > 0
      || pageLayerNodes(page).length > 0
      || page.layers.paintOrder.length > 0
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

    if (page.sectionRegions) {
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
      requireRect(node.flowBounds, `${path}.flowBounds`);
      requireRect(node.inkBounds, `${path}.inkBounds`);
      if (node.clipBounds) requireRect(node.clipBounds, `${path}.clipBounds`);
      requireFinite(node.advancePt, `${path}.advancePt`);
      if (node.kind === 'drawing') requireDrawingGeometry(node, path);
      const domain = domains.get(node.flowDomainId);
      if (!domain) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `${node.id} references missing flow domain ${node.flowDomainId}`);
      }
      if (!node.ordinaryFlow) return;
      if (domain.kind === 'body'
        && node.flowBounds.yPt + node.flowBounds.heightPt > page.geometry.contentBottomPt) {
        throw new LayoutInvariantError('BOTTOM_MARGIN_INVASION', `${node.id} crosses contentBottomPt`);
      }
      if (!contains(domain.bounds, node.flowBounds)) {
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

    const bookmarkNames = new Set<string>();
    page.bookmarkStarts?.forEach((bookmark) => {
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
          && first.flowDomainId === second.flowDomainId
          && overlaps(first.flowBounds, second.flowBounds)) {
          throw new LayoutInvariantError('FLOW_OVERLAP', `${first.id} overlaps ${second.id}`);
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
