import { describe, expect, it } from 'vitest';
import type { AnchorFrameResult } from './anchor-frame.js';
import {
  beginFloatingTablePlacementTransaction,
  resolveFloatingTablePlacementInTransaction,
} from './floating-table-transaction.js';
import {
  projectBodyOccurrence,
  type BodyOccurrenceProjectionOptions,
} from './occurrence-projection.js';
import type { TableFragmentLayout } from './table-pagination.js';
import type {
  DrawingLayout,
  LayoutRect,
  PaintNode,
  ParagraphLayout,
  SourceRef,
  TableLayout,
  TextBoxLayout,
} from './types.js';

const source = (path: readonly number[]): SourceRef => ({
  story: 'body',
  storyInstance: 'body',
  path,
});

const rect = (xPt: number, yPt: number, widthPt = 20, heightPt = 10): LayoutRect => ({
  xPt, yPt, widthPt, heightPt,
});

function simpleParagraph(
  id: string,
  path: readonly number[],
  flowDomainId = 'acquired:body',
  writingMode: 'horizontal-tb' | 'vertical-rl' = 'horizontal-tb',
): ParagraphLayout {
  const bounds = rect(1, 2);
  return {
    kind: 'paragraph',
    id,
    source: source(path),
    flowDomainId,
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 10,
    ordinaryFlow: true,
    spacing: { beforePt: 0, afterPt: 0 },
    contextualSpacing: false,
    lines: [{
      range: { start: 0, end: 1 },
      bounds,
      baselinePt: 9,
      advancePt: 10,
      placements: [{
        kind: 'text',
        text: 'x',
        range: { start: 0, end: 1 },
        origin: { xPt: 2, yPt: 9 },
        bounds,
        advancePt: 8,
        clusters: [{
          range: { start: 0, end: 1 },
          offset: { xPt: 0, yPt: 0 },
          advancePt: 8,
        }],
        paintOps: [],
        color: { kind: 'default' },
        fontRoute: {
          familyList: 'sans-serif',
          scope: 'generic',
          fingerprint: 'canvas-font-route-v1:generic:sans-serif',
        },
        fontSizePt: 10,
        fontWeight: 400,
        fontStyle: 'normal',
        direction: 'ltr',
        writingMode,
        decorations: [{
          kind: 'underline',
          from: { xPt: 2, yPt: 10 },
          to: { xPt: 10, yPt: 10 },
          color: '#000000',
          widthPt: 1,
          style: 'solid',
        }],
        hyperlink: 'https://example.test/',
        bookmark: 'authored-bookmark',
      }],
    }],
    borders: [{
      from: { xPt: 1, yPt: 2 },
      to: { xPt: 21, yPt: 2 },
      color: '#000000',
      widthPt: 1,
      authoredStyle: 'single',
      style: 'solid',
    }],
    resources: [{
      kind: 'image',
      resourceKey: 'authored-resource-key',
      intrinsicSize: { widthPt: 20, heightPt: 10 },
    }],
    drawings: [],
    textBoxes: [],
    events: [],
    exclusions: [],
    paragraphMark: { hidden: false, bounds: rect(20, 2, 0, 10) },
    lineNumbers: [{
      lineIndex: 0,
      counterValue: 7,
      bounds: rect(-8, 2, 6, 10),
      paintOps: [{
        kind: 'text',
        text: '7',
        origin: { xPt: -3, yPt: 9 },
        font: '10pt sans-serif',
        color: '#000000',
        textAlign: 'right',
      }],
    }],
  };
}

function anchorFrame(occurrenceId: string): AnchorFrameResult {
  const axis = (axisName: 'horizontal' | 'vertical', origin: number) => ({
    axis: axisName,
    status: 'resolved' as const,
    relativeFrom: 'column',
    referenceFrame: 'column' as const,
    choiceKind: 'offset' as const,
    choiceValue: 0,
    baseStartPt: origin,
    baseEndPt: origin + 100,
    resolvedOriginPt: origin,
    pageParity: null,
  });
  return {
    status: 'resolved',
    occurrenceId,
    axes: { horizontal: axis('horizontal', 4), vertical: axis('vertical', 6) },
    issues: [],
    geometry: {
      objectFrame: rect(4, 6, 20, 10),
      inkBounds: rect(3, 5, 22, 12),
      wrapBounds: rect(2, 4, 24, 14),
      size: {
        horizontal: { source: 'extent', valuePt: 20, relativeFrom: null, referenceFrame: null, fraction: null },
        vertical: { source: 'extent', valuePt: 10, relativeFrom: null, referenceFrame: null, fraction: null },
      },
      parentEffectExtent: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
      wrap: {
        kind: 'tight',
        side: 'bothSides',
        distances: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
        distanceSources: { top: 'implicit-zero', right: 'implicit-zero', bottom: 'implicit-zero', left: 'implicit-zero' },
        effectExtent: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
        effectExtentSource: 'none',
        coordinateSpace: { width: 21600, height: 21600 },
        polygon: { edited: false, points: [{ xPt: 4, yPt: 6 }, { xPt: 24, yPt: 16 }] },
      },
      transform: { coordinateSpace: 'anchor-frame', groupApplication: 'parser-resolved-child-frame', group: null },
    },
  };
}

function complexParagraph(): ParagraphLayout {
  const base = simpleParagraph('paragraph:source', [0]);
  const childBase = simpleParagraph('textbox-paragraph:source', [0, 0]);
  const childDrawing: DrawingLayout = {
    kind: 'drawing', id: 'textbox-drawing:source', source: source([0, 0, 1]),
    flowDomainId: 'acquired:body:textbox', flowBounds: rect(7, 8), inkBounds: rect(7, 8),
    advancePt: 0, ordinaryFlow: false, commands: [],
    anchorLayer: {
      occurrenceId: 'textbox-anchor:source', behindDoc: false, relativeHeight: 1,
      sourceOrder: 0, horizontalOwnership: 'host', verticalOwnership: 'host',
    },
  };
  const child: ParagraphLayout = {
    ...childBase,
    drawings: [childDrawing],
    anchorFrames: [anchorFrame('textbox-anchor:source')],
  };
  const textBox: TextBoxLayout = {
    kind: 'textbox',
    id: 'textbox:source',
    source: source([0, 1]),
    flowDomainId: 'acquired:body:textbox',
    flowBounds: rect(5, 6, 30, 20),
    inkBounds: rect(5, 6, 30, 20),
    contentBounds: rect(6, 7, 28, 18),
    advancePt: 0,
    ordinaryFlow: false,
    paragraphs: [child],
    writingMode: 'horizontal-tb',
    insets: { topPt: 1, rightPt: 1, bottomPt: 1, leftPt: 1 },
  };
  const drawing: DrawingLayout = {
    kind: 'drawing',
    id: 'drawing:source',
    source: source([0, 2]),
    flowDomainId: 'acquired:body',
    flowBounds: rect(4, 6, 20, 10),
    inkBounds: rect(3, 5, 22, 12),
    advancePt: 0,
    ordinaryFlow: false,
    commands: [{ kind: 'fill-rect', rect: rect(4, 6, 20, 10), fill: '#000000' }],
    textBoxIds: ['textbox:source'],
    anchorLayer: {
      occurrenceId: 'anchor:source',
      behindDoc: false,
      relativeHeight: 1,
      sourceOrder: 0,
      horizontalOwnership: 'host',
      verticalOwnership: 'host',
    },
  };
  return {
    ...base,
    lines: [{
      ...base.lines[0]!,
      placements: [
        ...base.lines[0]!.placements,
        { kind: 'drawing', range: { start: 1, end: 2 }, drawingId: drawing.id, bounds: drawing.flowBounds, advancePt: 0 },
        { kind: 'anchor-host', range: { start: 2, end: 2 }, bounds: rect(4, 6, 0, 10), baselinePt: 12, anchorOccurrenceId: 'anchor:source' },
      ],
    }],
    drawings: [drawing],
    textBoxes: [textBox],
    exclusions: [{
      id: 'exclusion:source',
      wrap: 'tight',
      bounds: rect(2, 4, 24, 14),
      polygon: [{ xPt: 2, yPt: 4 }, { xPt: 26, yPt: 18 }],
      anchorOccurrenceId: 'anchor:source',
    }],
    anchorFrames: [anchorFrame('anchor:source')],
  };
}

function table(id: string, path: readonly number[], nested = false): TableLayout {
  const paragraph = simpleParagraph(`${id}:paragraph`, [...path, 0, 0]);
  const nestedTable = nested ? table(`${id}:nested`, [...path, 0, 1], false) : null;
  const cellId = `${id}:cell`;
  const rowId = `${id}:row`;
  const bounds = rect(10, 20, 100, 30);
  return {
    kind: 'table', id, source: source(path), flowDomainId: 'acquired:body',
    flowBounds: bounds, inkBounds: bounds, advancePt: 30, ordinaryFlow: true,
    columnWidthsPt: [100], borders: [{
      from: { xPt: 10, yPt: 20 }, to: { xPt: 110, yPt: 20 }, color: '#000000',
      widthPt: 1, authoredStyle: 'single', style: 'solid',
    }],
    rows: [{
      kind: 'table-row', id: rowId, source: source([...path, 0]), flowDomainId: 'acquired:body',
      flowBounds: bounds, inkBounds: bounds, advancePt: 30, ordinaryFlow: true,
      heightPt: 30, contentHeightPt: 20,
      cells: [{
        kind: 'table-cell', id: cellId, source: source([...path, 0, 0]), flowDomainId: 'acquired:body',
        flowBounds: bounds, inkBounds: bounds, contentBounds: rect(12, 22, 96, 26),
        advancePt: 30, ordinaryFlow: true, verticalMerge: 'none', vAlign: 'top',
        blocks: [
          { layout: paragraph, offsetPt: 2, advancePt: 10 },
          ...(nestedTable ? [{ layout: nestedTable, offsetPt: 12, advancePt: 30 }] : []),
        ],
      }],
    }],
  };
}

function fragment(state: 'unresolved' | 'resolved' = 'unresolved'): TableFragmentLayout {
  const base = table('table:source', [1], false);
  const host = base.rows[0]!.cells[0]!;
  // Production acquisition keeps a floating nested table out of ordinary cell flow.
  const child = table('table:source:floating-child', [1, 0, 1], false);
  const placement = {
    kind: 'floating-table-placement' as const,
    occurrenceId: 'float:source',
    ownership: 'repeated-header' as const,
    physicalPageIndex: 1,
    displayPageNumber: 2,
    hostCellId: host.id,
    sourceBlockIndex: 1,
    anchorBlockIndex: 0,
    tableId: child.id,
    overlap: 'overlap' as const,
    positioning: {
      leftFromTextPt: 0, rightFromTextPt: 0, topFromTextPt: 0, bottomFromTextPt: 0,
      horzAnchor: state === 'resolved' ? 'page' : 'text', horzSpecified: true,
      vertAnchor: state === 'resolved' ? 'margin' : 'text',
      xPt: state === 'resolved' ? 12 : 0, yPt: state === 'resolved' ? 14 : 0,
    },
    // Resolved selections have already accepted final page placement translation.
    anchorBounds: state === 'resolved' ? rect(12, 50, 30, 10) : rect(12, 24, 30, 10),
    child,
  };
  const resolved = resolveFloatingTablePlacementInTransaction(
    placement,
    { page: rect(0, 0, 600, 800), margin: rect(36, 36, 528, 728), text: rect(12, 24, 500, 700) },
    beginFloatingTablePlacementTransaction([], 0),
  ).placement;
  return {
    ...base,
    rows: base.rows.map((row, rowIndex) => ({
      ...row,
      logicalRowIndex: rowIndex,
      fragmentIndex: 0,
      ownership: 'repeated-header',
      occurrenceId: 'row-occurrence:source',
      physicalPageIndex: 1,
      displayPageNumber: 2,
      cells: row.cells.map((cell) => ({
        ...cell,
        contentRanges: [{ kind: 'whole', blockIndex: 0 }],
      })),
    })),
    floatingTables: state === 'unresolved' ? [placement] : [],
    resolvedFloatingTables: state === 'resolved' ? [resolved] : [],
    floatingTableCoordinateSpace: 'logical-page-points',
  };
}

const options = (
  occurrenceId: string,
  xPt = 10,
  yPt = 20,
): BodyOccurrenceProjectionOptions => ({
  occurrenceId,
  destination: {
    coordinateSpace: 'logical-body-points',
    flowDomainId: 'page:2/region:body/column:0',
    translation: { xPt, yPt },
  },
});

function validateGraph(node: PaintNode): Set<string> {
  const ids = new Map<string, object>();
  const references: Array<readonly [string, string]> = [];
  const anchorTargets = new Set<string>();
  const anchorReferences = new Set<string>();
  const visited = new Set<object>();
  const visit = (current: PaintNode): void => {
    if (visited.has(current)) return;
    visited.add(current);
    const prior = ids.get(current.id);
    expect(prior === undefined || prior === current, `layout ID ${current.id} has one object`).toBe(true);
    ids.set(current.id, current);
    if (current.kind === 'paragraph') {
      for (const placement of current.lines.flatMap((line) => line.placements)) {
        if (placement.kind === 'drawing') references.push(['drawing', placement.drawingId]);
      }
      current.drawings.forEach(visit);
      current.textBoxes.forEach(visit);
      for (const drawing of current.drawings) {
        if (drawing.anchorLayer) anchorTargets.add(drawing.anchorLayer.occurrenceId);
      }
      for (const exclusion of current.exclusions) {
        const priorExclusion = ids.get(exclusion.id);
        expect(priorExclusion === undefined || priorExclusion === exclusion).toBe(true);
        ids.set(exclusion.id, exclusion);
        if (exclusion.anchorOccurrenceId) anchorReferences.add(exclusion.anchorOccurrenceId);
      }
      for (const frame of current.anchorFrames ?? []) anchorReferences.add(frame.occurrenceId);
    } else if (current.kind === 'textbox') {
      current.paragraphs.forEach(visit);
    } else if (current.kind === 'table') {
      for (const row of current.rows) {
        const priorRow = ids.get(row.id);
        expect(priorRow === undefined || priorRow === row, `row ID ${row.id} has one object`).toBe(true);
        ids.set(row.id, row);
        for (const cell of row.cells) {
          const priorCell = ids.get(cell.id);
          expect(priorCell === undefined || priorCell === cell, `cell ID ${cell.id} has one object`).toBe(true);
          ids.set(cell.id, cell);
          cell.blocks.forEach((block) => visit(block.layout));
        }
      }
      if ('floatingTables' in current
        && Array.isArray(current.floatingTables)
        && 'resolvedFloatingTables' in current
        && Array.isArray(current.resolvedFloatingTables)) {
        const fragment = current as TableFragmentLayout;
        for (const floating of fragment.floatingTables) {
          references.push(['host', floating.hostCellId], ['table', floating.tableId]);
          visit(floating.child);
        }
        for (const resolved of fragment.resolvedFloatingTables) {
          references.push(['host', resolved.source.hostCellId], ['table', resolved.source.tableId]);
          visit(resolved.child);
          visit(resolved.source.child);
        }
      }
    }
  };
  visit(node);
  for (const [kind, id] of references) {
    expect(ids.has(id), `${kind} reference ${id} has a target`).toBe(true);
  }
  for (const id of anchorReferences) expect(anchorTargets.has(id), `anchor ${id} has a target`).toBe(true);
  return new Set([...ids.keys(), ...anchorTargets]);
}

function sharedIds(left: PaintNode, right: PaintNode): string[] {
  const leftIds = validateGraph(left);
  const rightIds = validateGraph(right);
  return [...leftIds].filter((id) => rightIds.has(id));
}

describe('projectBodyOccurrence', () => {
  it('relocates complete retained paragraph geometry without acquisition', () => {
    const projected = projectBodyOccurrence(complexParagraph(), options('paragraph:first'));
    if (projected.kind !== 'paragraph') throw new Error('expected paragraph');

    expect(projected.flowBounds).toEqual(rect(11, 22));
    expect(projected.lines[0]?.bounds).toEqual(rect(11, 22));
    expect(projected.lines[0]?.placements[0]).toMatchObject({ origin: { xPt: 12, yPt: 29 } });
    expect(projected.borders[0]).toMatchObject({ from: { xPt: 11, yPt: 22 } });
    expect(projected.paragraphMark?.bounds).toEqual(rect(30, 22, 0, 10));
    expect(projected.lineNumbers?.[0]).toMatchObject({
      bounds: rect(2, 22, 6, 10),
      paintOps: [{ origin: { xPt: 7, yPt: 29 } }],
    });
    expect(projected.anchorFrames?.[0]).toMatchObject({
      geometry: {
        objectFrame: rect(14, 26, 20, 10),
        wrap: { polygon: { points: [{ xPt: 14, yPt: 26 }, { xPt: 34, yPt: 36 }] } },
      },
    });
    const child = projected.textBoxes[0]!.paragraphs[0]!;
    expect(child.lineNumbers?.[0]).toMatchObject({
      bounds: rect(2, 22, 6, 10),
      paintOps: [{ origin: { xPt: 7, yPt: 29 } }],
    });
    expect(child.anchorFrames?.[0]).toMatchObject({
      geometry: { objectFrame: rect(14, 26, 20, 10) },
      axes: {
        horizontal: { baseStartPt: 14, baseEndPt: 114, resolvedOriginPt: 14 },
        vertical: { baseStartPt: 26, baseEndPt: 126, resolvedOriginPt: 26 },
      },
    });
  });

  it('rekeys drawing, text-box, anchor, exclusion, and placement references', () => {
    const projected = projectBodyOccurrence(complexParagraph(), options('paragraph:first'));
    if (projected.kind !== 'paragraph') throw new Error('expected paragraph');
    const drawing = projected.drawings[0]!;
    const textBox = projected.textBoxes[0]!;
    const drawingPlacement = projected.lines[0]!.placements.find((item) => item.kind === 'drawing');
    const anchorHost = projected.lines[0]!.placements.find((item) => item.kind === 'anchor-host');

    expect(drawing.id).not.toBe('drawing:source');
    expect(textBox.id).not.toBe('textbox:source');
    expect(drawingPlacement).toMatchObject({ drawingId: drawing.id });
    expect(drawing.textBoxIds).toEqual([textBox.id]);
    expect(anchorHost).toMatchObject({ anchorOccurrenceId: drawing.anchorLayer?.occurrenceId });
    expect(projected.exclusions[0]?.anchorOccurrenceId).toBe(drawing.anchorLayer?.occurrenceId);
    expect(projected.anchorFrames?.[0]?.occurrenceId).toBe(drawing.anchorLayer?.occurrenceId);
    expect(projected.exclusions[0]?.id).not.toBe('exclusion:source');
    expect(projected.lines[0]!.placements[0]).toMatchObject({
      bookmark: 'authored-bookmark', hyperlink: 'https://example.test/',
    });
    expect(projected.resources[0]?.resourceKey).toBe('authored-resource-key');
  });

  it('gives split/repeated paragraph occurrences disjoint IDs and stable sources', () => {
    const acquired = complexParagraph();
    const first = projectBodyOccurrence(acquired, options('paragraph:first'));
    const second = projectBodyOccurrence(acquired, options('paragraph:second'));

    expect(sharedIds(first, second)).toEqual([]);
    expect(first.source).toEqual(acquired.source);
    expect(second.source).toEqual(acquired.source);
    expect(first.kind === 'paragraph' && first.textBoxes[0]?.paragraphs[0]?.source)
      .toEqual(source([0, 0]));
  });

  it('rekeys and relocates a table root, rows, cells, and nested blocks', () => {
    const projected = projectBodyOccurrence(fragment(), options('table:first'));
    if (projected.kind !== 'table') throw new Error('expected table');
    const row = projected.rows[0]!;
    const cell = row.cells[0]!;
    const paragraph = cell.blocks[0]!.layout;

    expect(projected.flowBounds).toEqual(rect(20, 40, 100, 30));
    expect(row.flowBounds).toEqual(rect(20, 40, 100, 30));
    expect(cell.contentBounds).toEqual(rect(22, 42, 96, 26));
    // Cell block layouts retain their own coordinate space. The table painter
    // places them from contentBounds/offsetPt; adding the outer delta here would
    // apply it twice for nested tables whose local alignment is in flowBounds.
    expect(paragraph.flowBounds).toEqual(rect(1, 2));
    expect(cell.blocks).toHaveLength(1);
    expect(new Set([projected.id, row.id, cell.id, paragraph.id]).size).toBe(4);
    expect([projected.id, row.id, cell.id, paragraph.id])
      .not.toContain('table:source');
  });

  it('gives repeated-header and split table occurrences disjoint descendant IDs', () => {
    const acquired = fragment();
    const header = projectBodyOccurrence(acquired, options('table:header'));
    const split = projectBodyOccurrence(acquired, options('table:split'));

    expect(sharedIds(header, split)).toEqual([]);
    expect(header.source).toEqual(split.source);
  });

  it('rewrites an unresolved-only floating-table graph without putting the child in cell flow', () => {
    const projected = projectBodyOccurrence(fragment('unresolved'), options('table:first')) as TableFragmentLayout;
    const cell = projected.rows[0]!.cells[0]!;
    const floating = projected.floatingTables[0]!;

    expect(projected.resolvedFloatingTables).toEqual([]);
    expect(cell.blocks).toHaveLength(1);
    expect(floating.hostCellId).toBe(cell.id);
    expect(floating.tableId).toBe(floating.child.id);
    expect(floating.anchorBounds).toEqual(rect(22, 44, 30, 10));
    validateGraph(projected);
  });

  it('preserves production resolved child aliasing and final page-local geometry', () => {
    const projected = projectBodyOccurrence(fragment('resolved'), options('table:first')) as TableFragmentLayout;
    const resolved = projected.resolvedFloatingTables[0]!;

    expect(projected.floatingTables).toEqual([]);
    expect(resolved.child).toBe(resolved.source.child);
    expect(resolved.source.tableId).toBe(resolved.child.id);
    expect(resolved.bounds).toEqual(rect(12, 50, 100, 30));
    expect(resolved.source.anchorBounds).toEqual(rect(12, 50, 30, 10));
    validateGraph(projected);
    const cloned = structuredClone(projected);
    expect(cloned.resolvedFloatingTables[0]!.child)
      .toBe(cloned.resolvedFloatingTables[0]!.source.child);
    expect(Object.isFrozen(resolved.child)).toBe(true);
  });

  it('rejects one retained table object reused under incompatible geometry ownership', () => {
    const invalid = fragment('unresolved');
    const shared = invalid.floatingTables[0]!.child;
    const cell = invalid.rows[0]!.cells[0]!;
    const hybrid: TableFragmentLayout = {
      ...invalid,
      floatingTables: invalid.floatingTables.map((floating) => ({
        ...floating,
        hostCellId: `${cell.id}:incompatible-owner`,
      })),
      rows: [{
        ...invalid.rows[0]!,
        cells: [{
          ...cell,
          blocks: [...cell.blocks, { layout: shared, offsetPt: 12, advancePt: shared.advancePt }],
        }],
      }],
    };

    expect(() => projectBodyOccurrence(hybrid, options('table:invalid')))
      .toThrow(/incompatible projection ownership/);
  });

  it('assigns deterministic occurrence-local domains to nested cell and text-box content', () => {
    const paragraph = projectBodyOccurrence(complexParagraph(), options('paragraph:first'));
    const tableNode = projectBodyOccurrence(fragment(), options('table:first'));
    if (paragraph.kind !== 'paragraph' || tableNode.kind !== 'table') throw new Error('unexpected node');
    const textBox = paragraph.textBoxes[0]!;
    const cell = tableNode.rows[0]!.cells[0]!;

    expect(paragraph.flowDomainId).toBe('page:2/region:body/column:0');
    expect(textBox.flowDomainId).not.toBe(paragraph.flowDomainId);
    expect(textBox.paragraphs[0]?.flowDomainId).toBe(textBox.flowDomainId);
    expect(cell.flowDomainId).not.toBe(tableNode.flowDomainId);
    expect(cell.blocks[0]?.layout.flowDomainId).toBe(cell.flowDomainId);
    expect(projectBodyOccurrence(fragment(), options('table:first')).kind === 'table'
      && projectBodyOccurrence(fragment(), options('table:first')).rows[0]?.cells[0]?.flowDomainId)
      .toBe(cell.flowDomainId);
  });

  it('is structured-clone safe, deeply immutable, and leaves the source graph unchanged', () => {
    const acquired = fragment();
    const before = structuredClone(acquired);
    const projected = projectBodyOccurrence(acquired, options('table:first'));

    expect(() => structuredClone(projected)).not.toThrow();
    expect(acquired).toEqual(before);
    expect(projected).not.toBe(acquired);
    expect(Object.isFrozen(projected)).toBe(true);
    expect(Object.isFrozen(projected.rows)).toBe(true);
  });

  it('applies one logical translation to horizontal and vertical-logical geometry', () => {
    const horizontal = projectBodyOccurrence(
      simpleParagraph('horizontal', [0], 'acquired:body', 'horizontal-tb'),
      options('horizontal', 13, 17),
    );
    const vertical = projectBodyOccurrence(
      simpleParagraph('vertical', [1], 'acquired:body', 'vertical-rl'),
      options('vertical', 13, 17),
    );
    if (horizontal.kind !== 'paragraph' || vertical.kind !== 'paragraph') throw new Error('unexpected node');

    expect(horizontal.flowBounds).toEqual(rect(14, 19));
    expect(vertical.flowBounds).toEqual(rect(14, 19));
    expect(vertical.lines[0]?.placements[0]).toMatchObject({
      writingMode: 'vertical-rl', origin: { xPt: 15, yPt: 26 },
    });
  });

  it('preserves page-owned drawing axes while relocating host-owned paragraph geometry', () => {
    const acquired = complexParagraph();
    const pageOwned: ParagraphLayout = {
      ...acquired,
      drawings: acquired.drawings.map((drawing) => ({
        ...drawing,
        anchorLayer: {
          ...drawing.anchorLayer!,
          horizontalOwnership: 'page',
          verticalOwnership: 'page',
        },
      })),
      exclusions: acquired.exclusions.map((exclusion) => ({
        ...exclusion,
        verticalOwnership: 'page',
      })),
    };
    const projected = projectBodyOccurrence(pageOwned, options('paragraph:first'));
    if (projected.kind !== 'paragraph') throw new Error('expected paragraph');

    expect(projected.flowBounds).toEqual(rect(11, 22));
    expect(projected.drawings[0]?.flowBounds).toEqual(rect(4, 6, 20, 10));
    expect(projected.exclusions[0]?.bounds).toEqual(rect(2, 4, 24, 14));
    expect(projected.anchorFrames?.[0]).toMatchObject({
      geometry: { objectFrame: rect(4, 6, 20, 10) },
    });
  });

  it('preserves page-owned nested anchors and vertical text-box child coordinates', () => {
    const acquired = complexParagraph();
    const textBox = acquired.textBoxes[0]!;
    const nested = textBox.paragraphs[0]!;
    const pageOwnedNested: ParagraphLayout = {
      ...nested,
      drawings: nested.drawings.map((drawing) => ({
        ...drawing,
        anchorLayer: {
          ...drawing.anchorLayer!,
          horizontalOwnership: 'page',
          verticalOwnership: 'page',
        },
      })),
    };
    const pageRelative = projectBodyOccurrence({
      ...acquired,
      textBoxes: [{ ...textBox, paragraphs: [pageOwnedNested] }],
    }, options('paragraph:page-owned')) as ParagraphLayout;
    const projectedNested = pageRelative.textBoxes[0]!.paragraphs[0]!;
    expect(projectedNested.flowBounds).toEqual(rect(11, 22));
    expect(projectedNested.lineNumbers?.[0]?.bounds).toEqual(rect(2, 22, 6, 10));
    const pageOwnedFrame = projectedNested.anchorFrames?.[0];
    expect(pageOwnedFrame?.status).toBe('resolved');
    if (pageOwnedFrame?.status !== 'resolved') throw new Error('expected resolved anchor frame');
    expect(pageOwnedFrame.geometry.objectFrame).toEqual(rect(4, 6, 20, 10));

    const vertical = projectBodyOccurrence({
      ...acquired,
      textBoxes: [{ ...textBox, verticalMode: 'vert', paragraphs: [nested] }],
    }, options('paragraph:vertical-textbox')) as ParagraphLayout;
    const local = vertical.textBoxes[0]!.paragraphs[0]!;
    expect(local.flowBounds).toEqual(nested.flowBounds);
    expect(local.lineNumbers).toEqual(nested.lineNumbers);
    const localFrame = local.anchorFrames?.[0];
    const nestedFrame = nested.anchorFrames?.[0];
    expect(localFrame?.status).toBe('resolved');
    expect(nestedFrame?.status).toBe('resolved');
    if (localFrame?.status !== 'resolved' || nestedFrame?.status !== 'resolved') {
      throw new Error('expected resolved anchor frames');
    }
    expect(localFrame.geometry).toEqual(nestedFrame.geometry);
    expect(localFrame.axes).toEqual(nestedFrame.axes);
  });
});
