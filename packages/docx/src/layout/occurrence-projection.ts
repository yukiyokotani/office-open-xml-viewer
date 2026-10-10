import { snapshotPlainData } from './plain-data.js';
import {
  floatingTableAxesFollowHostFlow,
  translateCompleteParagraphLayout,
  translateRect,
  translateTableLayout,
  type LayoutTranslation,
} from './retained-geometry-translation.js';
import type {
  DrawingLayout,
  FloatingTablePlacementLayout,
  LayoutCoordinateSpace,
  ParagraphLayout,
  ParagraphPlacement,
  ResolvedFloatingTablePlacementLayout,
  TableCellLayout,
  TableLayout,
  TableRowLayout,
  TextBoxLayout,
} from './types.js';

export interface BodyOccurrenceDestination {
  readonly coordinateSpace: Extract<LayoutCoordinateSpace, 'logical-page-points'>;
  readonly flowDomainId: string;
  readonly translation: Readonly<{ xPt: number; yPt: number }>;
}

export interface BodyOccurrenceProjectionOptions {
  readonly occurrenceId: string;
  readonly destination: BodyOccurrenceDestination;
}

export function projectedNestedOccurrenceId(
  ownerOccurrenceId: string,
  sourceOccurrenceId: string,
): string {
  const component = encodeURIComponent(sourceOccurrenceId).replaceAll('%3A', ':');
  return `${ownerOccurrenceId}/occurrence/${component}`;
}

function validateTranslation(translation: LayoutTranslation): void {
  if (!Number.isFinite(translation.xPt) || !Number.isFinite(translation.yPt)) {
    throw new RangeError('body occurrence translation must be finite');
  }
}

function validateProjectionOptions(options: BodyOccurrenceProjectionOptions): void {
  if (options.occurrenceId.length === 0) throw new RangeError('occurrenceId must not be empty');
  if (options.destination.flowDomainId.length === 0) throw new RangeError('flowDomainId must not be empty');
  validateTranslation(options.destination.translation);
}

function resolvedFloatingDelta(
  placement: FloatingTablePlacementLayout,
  hostDelta: LayoutTranslation,
): LayoutTranslation {
  const followsHost = floatingTableAxesFollowHostFlow(placement.positioning);
  return {
    xPt: followsHost.x ? hostDelta.xPt : 0,
    yPt: followsHost.y ? hostDelta.yPt : 0,
  };
}

function assertAcyclicLayoutGraph(root: ParagraphLayout | TableLayout): void {
  const visiting = new WeakSet<object>();
  const completed = new WeakSet<object>();
  const visit = (layout: ParagraphLayout | TableLayout): void => {
    if (visiting.has(layout)) throw new TypeError('body occurrence layout graph must be acyclic');
    if (completed.has(layout)) return;
    visiting.add(layout);
    if (layout.kind === 'paragraph') {
      for (const textBox of layout.textBoxes) {
        for (const block of textBox.story.blocks) {
          if (block.kind === 'paragraph' || block.kind === 'table') visit(block);
        }
      }
    } else {
      for (const row of layout.rows) for (const cell of row.cells) {
        for (const block of cell.blocks) visit(block.layout);
      }
      for (const placement of layout.floatingTables ?? []) visit(placement.child);
      for (const placement of layout.resolvedFloatingTables ?? []) {
        visit(placement.source.child);
        visit(placement.child);
      }
    }
    visiting.delete(layout);
    completed.add(layout);
  };
  visit(root);
}

function translateOccurrenceGeometry<T extends ParagraphLayout | TableLayout>(
  retained: T,
  translation: LayoutTranslation,
): T {
  assertAcyclicLayoutGraph(retained);
  const tableMemo = new WeakMap<TableLayout, { key: string; value: TableLayout }>();
  const paragraphMemo = new WeakMap<ParagraphLayout, { key: string; value: ParagraphLayout }>();
  const keyFor = (delta: LayoutTranslation) => `${delta.xPt}\u0000${delta.yPt}`;

  const translateParagraph = (paragraph: ParagraphLayout, delta: LayoutTranslation): ParagraphLayout => {
    const key = keyFor(delta);
    const prior = paragraphMemo.get(paragraph);
    if (prior) {
      if (prior.key !== key) throw new Error('incompatible projection ownership');
      return prior.value;
    }
    const translatedBase = translateCompleteParagraphLayout(paragraph, delta);
    const translated = Object.freeze({
      ...translatedBase,
      ...(paragraph.sectionFlowOwnership === undefined
        ? {} : { sectionFlowOwnership: paragraph.sectionFlowOwnership }),
    });
    paragraphMemo.set(paragraph, { key, value: translated });
    return translated;
  };

  const translateTable = (retainedTable: TableLayout, delta: LayoutTranslation): TableLayout => {
    const key = keyFor(delta);
    const prior = tableMemo.get(retainedTable);
    if (prior) {
      if (prior.key !== key) throw new Error('incompatible projection ownership');
      return prior.value;
    }
    const translated: TableLayout = {
      ...translateTableLayout(retainedTable, delta),
      ...(retainedTable.sectionFlowOwnership === undefined
        ? {} : { sectionFlowOwnership: retainedTable.sectionFlowOwnership }),
    };
    tableMemo.set(retainedTable, { key, value: translated });

    const resolvedDeltaBySource = new Map<FloatingTablePlacementLayout, LayoutTranslation>();
    const finalFrameSources = new Set<FloatingTablePlacementLayout>();
    for (const resolved of retainedTable.resolvedFloatingTables ?? []) {
      const finalFrame = retainedTable.resolvedFloatingTableCoordinateSpace !== undefined;
      resolvedDeltaBySource.set(
        resolved.source,
        finalFrame ? { xPt: 0, yPt: 0 } : resolvedFloatingDelta(resolved.source, delta),
      );
      if (finalFrame) finalFrameSources.add(resolved.source);
    }
    const sourceMemo = new Map<FloatingTablePlacementLayout, FloatingTablePlacementLayout>();
    const translateSource = (source: FloatingTablePlacementLayout): FloatingTablePlacementLayout => {
      const priorSource = sourceMemo.get(source);
      if (priorSource) return priorSource;
      const childDelta = resolvedDeltaBySource.get(source) ?? delta;
      const anchorDelta = finalFrameSources.has(source) ? { xPt: 0, yPt: 0 } : delta;
      const result: FloatingTablePlacementLayout = {
        ...source,
        anchorBounds: translateRect(source.anchorBounds, anchorDelta),
        ...(source.columnBounds ? { columnBounds: translateRect(source.columnBounds, anchorDelta) } : {}),
        child: translateTable(source.child, childDelta),
      };
      sourceMemo.set(source, result);
      return result;
    };
    const floatingTables = (retainedTable.floatingTables ?? []).map(translateSource);
    const resolvedFloatingTables = (retainedTable.resolvedFloatingTables ?? []).map((resolved) => {
      const source = translateSource(resolved.source);
      const ownedDelta = resolvedDeltaBySource.get(resolved.source)
        ?? resolvedFloatingDelta(resolved.source, delta);
      return {
        ...resolved,
        xPt: resolved.xPt + ownedDelta.xPt,
        yPt: resolved.yPt + ownedDelta.yPt,
        bounds: translateRect(resolved.bounds, ownedDelta),
        exclusionBounds: translateRect(resolved.exclusionBounds, ownedDelta),
        child: source.child,
        source,
      } satisfies ResolvedFloatingTablePlacementLayout;
    });
    if (retainedTable.floatingTables || retainedTable.resolvedFloatingTables) {
      Object.assign(translated, { floatingTables, resolvedFloatingTables });
    }
    return translated;
  };

  return (retained.kind === 'paragraph'
    ? translateParagraph(retained, translation)
    : translateTable(retained, translation)) as T;
}

export function translateBodyOccurrence<T extends ParagraphLayout | TableLayout>(
  retained: T,
  translation: Readonly<{ xPt: number; yPt: number }>,
): T {
  validateTranslation(translation);
  return translateOccurrenceGeometry(retained, translation);
}

/** Deterministic ID/domain/freezing policy; these are engine rules, not OOXML rules. */
export function projectBodyOccurrence<T extends ParagraphLayout | TableLayout>(
  retained: T,
  options: BodyOccurrenceProjectionOptions,
): T {
  validateProjectionOptions(options);
  const translated = translateOccurrenceGeometry(retained, options.destination.translation);
  const encodedOccurrence = encodeURIComponent(options.occurrenceId);
  const tableMemo = new WeakMap<TableLayout, { domain: string; value: TableLayout }>();
  const paragraphMemo = new WeakMap<ParagraphLayout, { domain: string; value: ParagraphLayout }>();
  const drawingMemo = new WeakMap<DrawingLayout, { domain: string; value: DrawingLayout }>();
  const anchorOwners = new Map<string, DrawingLayout>();
  const floatingOwners = new Map<string, FloatingTablePlacementLayout>();
  const nodeId = (sourceId: string) => `${options.occurrenceId}/node/${encodeURIComponent(sourceId)}`;
  const anchorId = (sourceId: string) => `${options.occurrenceId}/anchor/${encodeURIComponent(sourceId)}`;
  const occurrenceId = (sourceId: string) =>
    projectedNestedOccurrenceId(options.occurrenceId, sourceId);
  const nestedDomain = (kind: 'cell' | 'textbox', sourceId: string) =>
    `${options.destination.flowDomainId}/occurrence/${encodedOccurrence}/${kind}/${encodeURIComponent(sourceId)}`;

  const projectPlacement = (placement: ParagraphPlacement): ParagraphPlacement => {
    if (placement.kind === 'drawing') return { ...placement, drawingId: nodeId(placement.drawingId) };
    if (placement.kind === 'anchor-host' && placement.anchorOccurrenceId) return {
      ...placement, anchorOccurrenceId: anchorId(placement.anchorOccurrenceId),
    };
    return placement;
  };
  const projectDrawing = (drawing: DrawingLayout, domain: string): DrawingLayout => {
    const memoized = drawingMemo.get(drawing);
    if (memoized) {
      if (memoized.domain !== domain) throw new Error('incompatible projection ownership');
      return memoized.value;
    }
    if (drawing.anchorLayer) {
      const prior = anchorOwners.get(drawing.anchorLayer.occurrenceId);
      if (prior && prior !== drawing) throw new Error('duplicate anchor occurrence owner');
      anchorOwners.set(drawing.anchorLayer.occurrenceId, drawing);
    }
    const projected: DrawingLayout = {
      ...drawing, id: nodeId(drawing.id), flowDomainId: domain,
      ...(drawing.textBoxIds ? { textBoxIds: drawing.textBoxIds.map(nodeId) } : {}),
      ...(drawing.anchorLayer ? { anchorLayer: {
        ...drawing.anchorLayer,
        occurrenceId: anchorId(drawing.anchorLayer.occurrenceId),
        acquisitionOccurrenceId: drawing.anchorLayer.acquisitionOccurrenceId ?? drawing.anchorLayer.occurrenceId,
      } } : {}),
    };
    drawingMemo.set(drawing, { domain, value: projected });
    return projected;
  };
  const projectTextBox = (textBox: TextBoxLayout): TextBoxLayout => {
    const domain = nestedDomain('textbox', textBox.id);
    return {
      ...textBox, id: nodeId(textBox.id), flowDomainId: domain,
      story: {
        ...textBox.story,
        blocks: textBox.story.blocks.map((block) => {
          if (block.kind === 'paragraph') return projectParagraph(block, domain);
          if (block.kind === 'table') return projectTable(block, domain);
          throw new Error(`Text-box story contains unsupported retained node: ${block.kind}`);
        }),
      },
    };
  };
  const projectParagraph = (paragraph: ParagraphLayout, domain: string): ParagraphLayout => {
    const prior = paragraphMemo.get(paragraph);
    if (prior) {
      if (prior.domain !== domain) throw new Error('incompatible projection ownership');
      return prior.value;
    }
    const projected: ParagraphLayout = {
      ...paragraph, id: nodeId(paragraph.id), flowDomainId: domain,
      lines: paragraph.lines.map((line) => ({
        ...line, placements: line.placements.map(projectPlacement),
      })),
      drawings: paragraph.drawings.map((drawing) => projectDrawing(drawing, domain)),
      // Completeness references use the same occurrence-local drawing owners.
      // Keep the acquired source graph intact for other page occurrences.
      ...(paragraph.nativeReadingRelocations ? {
        nativeReadingRelocations: paragraph.nativeReadingRelocations.map(nodeId),
      } : {}),
      textBoxes: paragraph.textBoxes.map(projectTextBox),
      exclusions: paragraph.exclusions.map((exclusion) => ({
        ...exclusion,
        id: exclusion.verticalOwnership === 'page' && !exclusion.anchorOccurrenceId
          ? exclusion.id
          : nodeId(exclusion.id),
        ...(exclusion.anchorOccurrenceId
          ? { anchorOccurrenceId: anchorId(exclusion.anchorOccurrenceId) } : {}),
      })),
      ...(paragraph.anchorCollisions ? {
        anchorCollisions: paragraph.anchorCollisions.map((entry) => ({
          ...entry,
          occurrenceId: anchorId(entry.occurrenceId),
        })),
      } : {}),
      ...(paragraph.anchorFrames ? { anchorFrames: paragraph.anchorFrames.map((frame) => ({
        ...frame, occurrenceId: anchorId(frame.occurrenceId),
      })) } : {}),
    };
    paragraphMemo.set(paragraph, { domain, value: projected });
    return projected;
  };
  const projectCell = (cell: TableCellLayout): TableCellLayout => {
    const domain = nestedDomain('cell', cell.id);
    return {
      ...cell, id: nodeId(cell.id), flowDomainId: domain,
      blocks: cell.blocks.map((block) => ({ ...block, layout: projectBlock(block.layout, domain) })),
    };
  };
  const projectRow = (row: TableRowLayout, domain: string): TableRowLayout => ({
    ...row, id: nodeId(row.id), flowDomainId: domain,
    ...('occurrenceId' in row && typeof row.occurrenceId === 'string'
      ? { occurrenceId: occurrenceId(row.occurrenceId) } : {}),
    cells: row.cells.map(projectCell),
  });
  const projectSource = (source: FloatingTablePlacementLayout): FloatingTablePlacementLayout => {
    const childDomain = nestedDomain('cell', source.hostCellId);
    return {
      ...source,
      occurrenceId: occurrenceId(source.occurrenceId),
      hostCellId: nodeId(source.hostCellId),
      tableId: nodeId(source.tableId),
    child: projectTable(source.child, childDomain),
    };
  };
  const projectTable = (table: TableLayout, domain: string): TableLayout => {
    const prior = tableMemo.get(table);
    if (prior) {
      if (prior.domain !== domain) throw new Error('incompatible projection ownership');
      return prior.value;
    }
    const projected: TableLayout = {
      ...table, id: nodeId(table.id), flowDomainId: domain,
      rows: table.rows.map((row) => projectRow(row, domain)),
    };
    tableMemo.set(table, { domain, value: projected });
    const sourceMemo = new Map<FloatingTablePlacementLayout, FloatingTablePlacementLayout>();
    const sourceFor = (source: FloatingTablePlacementLayout) => {
      const priorSource = sourceMemo.get(source);
      if (priorSource) return priorSource;
      const priorOwner = floatingOwners.get(source.occurrenceId);
      if (priorOwner && priorOwner !== source) {
        throw new Error('duplicate floating placement occurrence owner');
      }
      floatingOwners.set(source.occurrenceId, source);
      const result = projectSource(source);
      sourceMemo.set(source, result);
      return result;
    };
    const floatingTables = (table.floatingTables ?? []).map(sourceFor);
    const resolvedFloatingTables = (table.resolvedFloatingTables ?? []).map((resolved) => {
      const source = sourceFor(resolved.source);
      return {
        ...resolved, occurrenceId: occurrenceId(resolved.occurrenceId), child: source.child, source,
      } satisfies ResolvedFloatingTablePlacementLayout;
    });
    if (table.floatingTables || table.resolvedFloatingTables) {
      Object.assign(projected, { floatingTables, resolvedFloatingTables });
    }
    return projected;
  };
  function projectBlock(layout: ParagraphLayout | TableLayout, domain: string): ParagraphLayout | TableLayout {
    return layout.kind === 'paragraph'
      ? projectParagraph(layout, domain)
      : projectTable(layout, domain);
  }

  const projected = projectBlock(translated, options.destination.flowDomainId);
  // Translation and re-keying create new objects wherever geometry, domains or
  // occurrence IDs differ. An object still aliased from a verified frozen
  // acquisition root is unchanged source data, so the snapshot may share it.
  // Mutable fragments and externally supplied layouts take the cloning path.
  const snapshot = snapshotPlainData(projected, 'DOCX body occurrence projection', retained) as T;
  if (snapshot.kind !== 'table' || retained.kind !== 'table') return snapshot;
  // Pagination fragments rebuild occurrence-local rows, while their acquisition-owned
  // track vector is one canonical geometry value shared by every split occurrence.
  const columnWidthsPt = Object.isFrozen(retained.columnWidthsPt)
    ? retained.columnWidthsPt
    : Object.freeze([...retained.columnWidthsPt]);
  return Object.freeze({ ...snapshot, columnWidthsPt }) as unknown as T;
}
