import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const checker = resolve(import.meta.dirname, 'check-docx-layout-boundaries.mjs');

function write(root, path, contents) {
  const absolute = join(root, path);
  mkdirSync(resolve(absolute, '..'), { recursive: true });
  writeFileSync(absolute, contents);
}

function command(root, executable, args) {
  const result = spawnSync(executable, args, { cwd: root, encoding: 'utf8' });
  return {
    status: result.status,
    output: `${result.stdout ?? ''}${result.stderr ?? ''}`,
  };
}

function git(root, ...args) {
  const result = command(root, 'git', args);
  assert.equal(result.status, 0, result.output);
}

function runChecker(root, ...args) {
  return command(root, process.execPath, [checker, '--root', root, ...args]);
}

function initializeRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-'));
  write(root, 'packages/docx/src/renderer.ts', 'function buildMeasureState(ctx: unknown, fonts: unknown) { return [ctx, fonts]; }\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}) { const measure = buildMeasureState(ctx, resolvedLocalFonts); return [measure]; }\nexport function computeTableLayout() { return []; }\n');
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  return root;
}

const computePagesTableStampBaseline = `
export function computePages(sp: any, measureState: any) {
  const tableW = (sp.tableColWidthsPt ?? []).reduce((s, w) => s + w, 0) * measureState.scale;
  const sliceH = (sp.tableRowHeightsPt ?? []).reduce((s, h) => s + h, 0) * measureState.scale;
  return [tableW, sliceH];
}
export function computeTableLayout() { return []; }
`;

const computePagesRetainedSliceMigration = `
export function computePages(sp: any, measureState: any) {
  const { widthPx: tableW, heightPx: sliceH } = retainedTableSliceSize(
    sp, measureState.scale,
  );
  return [tableW, sliceH];
}
export function computeTableLayout() { return []; }
`;

const computePagesUprightBaseline = `
export function computePages() {
  const y = 10, h = 20, tblReservePt = 0, colTopY = 0, i = 0;
  const tableEl = {}, colWidthsPt = [], rowHeightsPt = [], bandPt = 100;
  stampTableLayout(tableEl, colWidthsPt, rowHeightsPt, bandPt);
  if (y + h > effContentH() - tblReservePt && y > colTopY) nextColumnOrPage(i);
  return [];
}
export function computeTableLayout() { return []; }
`;

const computePagesUprightMigration = `
export function computePages() {
  const y = 10, h = 20, tblReservePt = 0, colTopY = 0, i = 0;
  const tableEl = {}, colWidthsPt = [], rowHeightsPt = [], bandPt = 100;
  if (y + h > effContentH() - tblReservePt && y > colTopY) nextColumnOrPage(i);
  withColumnBand(() => stampTableLayout(
    tableEl,
    colWidthsPt,
    rowHeightsPt,
    bandPt,
    {
      ...measureState,
      pageIndex: pages.length - 1,
      displayPageNumber: pages.length,
    },
  ));
  return [];
}
export function computeTableLayout() { return []; }
`;

const computePagesUprightIntermediate = `
export function computePages() {
  const y = 10, h = 20, tblReservePt = 0, colTopY = 0, i = 0;
  const tableEl = {}, colWidthsPt = [], rowHeightsPt = [], bandPt = 100;
  stampTableLayout(tableEl, colWidthsPt, rowHeightsPt, bandPt);
  if (y + h > effContentH() - tblReservePt && y > colTopY) nextColumnOrPage(i);
  const sourceIndex = bodySourceIndexFor(tbl);
  const retained = sourceIndex === undefined
    ? undefined
    : measureState.retainedTablesBySourceIndex?.get(sourceIndex);
  if (sourceIndex === undefined || retained === undefined) {
    throw new Error('Upright vertical table requires retained physical geometry');
  }
  attachRetainedTablePlacement(tableEl, retained.layout, sourceIndex, {
    xPt: colX(),
    yPt: measureState.y,
    widthPt: colW(),
    flowAdvancePt: h,
    columnIndex: colIndex,
  });
  return [];
}
export function computeTableLayout() { return []; }
`;

const computePagesEnvelopeBaseline = `
export function computePages() {
  attachRetainedTableEnvelope(first, firstTable, widths, heights, band, state, firstPlacement);
  attachRetainedTableEnvelope(second, secondTable, widths, heights, band, state, secondPlacement);
  return [];
}
export function computeTableLayout() { return []; }
`;

const computePagesEnvelopeMigration = computePagesEnvelopeBaseline
  .replaceAll('attachRetainedTableEnvelope', 'attachTableFragment');

function initializeComputePagesUprightRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-upright-finalize-'));
  write(root, 'packages/docx/src/renderer.ts', computePagesUprightBaseline);
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

function initializeComputePagesEnvelopeRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-envelope-rename-'));
  write(root, 'packages/docx/src/renderer.ts', computePagesEnvelopeBaseline);
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

const computePagesFittingOuterBaseline = `
export function computePages() {
  withColumnBand(() => {
    const side = floatTableWrapSide(first.box, measureState);
    registerTableFloat(first.box, tp, measureState, side, tbl.overlap !== 'never');
  });
  stampTableLayout(
    el as PaginatedBodyElement,
    first.layout.colWidths,
    first.layout.rowHeights,
    first.contentWPt,
  );
  pushTagged(el as PaginatedBodyElement);
  return [];
}
export function computeTableLayout() { return []; }
`;

const computePagesFittingOuterMigration = `
export function computePages() {
  withColumnBand(() => {
    stampTableLayout(
      el as PaginatedBodyElement,
      first.layout.colWidths,
      first.layout.rowHeights,
      first.contentWPt,
    );
    const side = floatTableWrapSide(first.box, measureState);
    registerTableFloat(first.box, tp, measureState, side, tbl.overlap !== 'never');
  });
  pushTagged(el as PaginatedBodyElement);
  return [];
}
export function computeTableLayout() { return []; }
`;

function initializeComputePagesFittingOuterRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-fitting-outer-finalize-'));
  write(root, 'packages/docx/src/renderer.ts', computePagesFittingOuterBaseline);
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

function priorFittingOuterProbeSource(current) {
  const replaceRange = (source, startMarker, endMarker, replacement) => {
    const start = source.indexOf(startMarker);
    const end = source.indexOf(endMarker, start);
    assert.ok(start >= 0 && end > start, `${startMarker}..${endMarker}`);
    return source.slice(0, start) + replacement + source.slice(end);
  };
  let prior = replaceRange(
    current,
    '        const measureFloat = () =>\n',
    '        let first = measureFloat();',
    `        const measureFloat = () =>
          withColumnBand(() => {
            const cW = colW() * measureState.scale;
            const layout = computeTableLayout(tbl, cW, measureState);
            const tableH = layout.rowHeights.reduce((s, x) => s + x, 0);
            const box = computeFloatTableBox(tp, measureState, measureState.y, layout.tableW, tableH);
            const rawBox = computeFloatTableBox(tp, measureState, measureState.y, layout.tableW, tableH, true);
            return { box, rawBox, layout, contentWPt: cW / measureState.scale };
          });
`,
  );
  prior = prior.replace(
    `        if (first.requiresCanonicalSplit
          || (isTextAnchored && tableOverflowsHere)
          || pageAnchoredOverflows) {`,
    '        if ((isTextAnchored && tableOverflowsHere) || pageAnchoredOverflows) {',
  );
  prior = replaceRange(
    prior,
    "        if (!first.prepared) {\n          throw new Error('Fitting outer table acceptance requires a whole prepared fragment');\n        }\n",
    '        continue;\n      }\n\n      // ECMA-376 §17.11.10',
    `        withColumnBand(() => {
          stampTableLayout(
            el as PaginatedBodyElement,
            first.layout.colWidths,
            first.layout.rowHeights,
            first.contentWPt,
          );
          const side = floatTableWrapSide(first.box, measureState);
          registerTableFloat(first.box, tp, measureState, side, tbl.overlap !== 'never');
        });
        pushTagged(el as PaginatedBodyElement);
`,
  );
  return prior;
}

function initializeComposedFittingOuterProbeRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-composed-fitting-probe-'));
  const productionRoot = resolve(import.meta.dirname, '..');
  const current = readFileSync(join(productionRoot, 'packages/docx/src/renderer.ts'), 'utf8');
  write(root, 'packages/docx/src/renderer.ts', priorFittingOuterProbeSource(current));
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base with earlier A5 transformations');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return { root, current };
}

function priorOccurrenceOwnerSource(current) {
  const replacements = [
    ['computeTableRowHeights(ms, t, colWPt, j)', 'computeTableRowHeights(ms, t, colWPt)'],
    [
      'estimateTableHeight(measureState, nxt as unknown as DocTable, colW(), startIdx)',
      'estimateTableHeight(measureState, nxt as unknown as DocTable, colW())',
    ],
    [
      'measureState, para, i, colW, suppressBefore, colX,',
      'measureState, para, colW, suppressBefore, colX,',
    ],
    [
      `        const occurrenceEl = { ...el } as PaginatedElementWithLines;
        attachBodyParagraphFragment(occurrenceEl, para, measureState, i, {
          paragraphXPt: colX(),
          availableWidthPt: colW(),
          suppressSpaceBefore: suppressBefore,
          columnIndex: colIndex,
        }, fitMeasured);
        pushTagged(occurrenceEl);`,
      `        attachBodyParagraphFragment(el as PaginatedElementWithLines, para, measureState, i, {
          paragraphXPt: colX(),
          availableWidthPt: colW(),
          suppressSpaceBefore: suppressBefore,
          columnIndex: colIndex,
        }, fitMeasured);
        pushTagged(el as PaginatedBodyElement);`,
    ],
    [
      'attachBodyParagraphFragment(el as PaginatedElementWithLines, para, measureState, i, {',
      'attachBodyParagraphFragment(el as PaginatedElementWithLines, para, measureState, {',
    ],
    ['computeTableLayout(tbl, cW, measureState, i)', 'computeTableLayout(tbl, cW, measureState)'],
    ['            const retainedRecord = retainedTableRecord(measureState, i);\n', ''],
    [
      `              const prepared = prepareFittingOuterFragment(
                tbl, i, retainedRecord, finalState, box,
              );`,
      `              const prepared = bodyFlowFragments.sourceIndices.retainedTableMeasureBySource
                .prepareFittingOuterFragment(tbl, finalState, box);`,
    ],
    [
      `        const occurrenceEl = { ...el } as PaginatedBodyElement;
        withColumnBand(() => {
          stampTableLayout(
            occurrenceEl,
            first.layout.colWidths,
            first.layout.rowHeights,
            first.contentWPt,
            i,
            retainedTableRecord(measureState, i),
            measureState,
            undefined,
            acceptedPrepared,
          );
          const side = floatTableWrapSide(first.box, measureState);
          registerTableFloat(
            first.box, tp, measureState, side, tbl.overlap !== 'never', true,
          );
        });
        pushTagged(occurrenceEl);`,
      `        withColumnBand(() => {
          stampTableLayout(
            el as PaginatedBodyElement,
            first.layout.colWidths,
            first.layout.rowHeights,
            first.contentWPt,
            i,
            retainedTableRecord(measureState, i),
            measureState,
            undefined,
            acceptedPrepared,
          );
          const side = floatTableWrapSide(first.box, measureState);
          registerTableFloat(
            first.box, tp, measureState, side, tbl.overlap !== 'never', true,
          );
        });
        pushTagged(el as PaginatedBodyElement);`,
    ],
    [
      `            first.contentWPt,
            i,
            retainedTableRecord(measureState, i),
            measureState,
            undefined,
            acceptedPrepared,`,
      `            first.contentWPt,
            undefined,
            acceptedPrepared,`,
    ],
    [
      '            { sourceIndex: i, record: retainedTableRecord(measureState, i), state: measureState },\n',
      '',
    ],
    ['computeTablePtLayout(measureState, tbl, bandPt, i)', 'computeTablePtLayout(measureState, tbl, bandPt)'],
    [
      `          bandPt,
          i,
          retainedTableRecord(measureState, i),
          measureState,
          {`,
      `          bandPt,
          {`,
    ],
    [
      'computeTablePtLayout(measureState, tbl, tblContentWPt, i)',
      'computeTablePtLayout(measureState, tbl, tblContentWPt)',
    ],
    [
      `        tableContentH,
        measureState,
        i,
        true,`,
      `        tableContentH,
        measureState,
        true,`,
    ],
    [
      `              tblContentWPt,
              measureState,
              i,
              {`,
      `              tblContentWPt,
              measureState,
              {`,
    ],
    ['                fragment: meta.fragment,\n', ''],
    [
      `          () => currentSectionFrame.textDirection,
          i,
        );`,
      `          () => currentSectionFrame.textDirection,
        );`,
    ],
    [
      `          tblContentWPt,
          measureState,
          i,
          {`,
      `          tblContentWPt,
          measureState,
          {`,
    ],
  ];
  let prior = current;
  for (const [expected, replacement] of replacements) {
    assert.equal(prior.split(expected).length, 2, expected);
    prior = prior.replace(expected, replacement);
  }
  return prior;
}

function initializeOccurrenceOwnerRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-occurrence-owner-'));
  const productionRoot = resolve(import.meta.dirname, '..');
  const current = readFileSync(join(productionRoot, 'packages/docx/src/renderer.ts'), 'utf8');
  write(root, 'packages/docx/src/renderer.ts', priorOccurrenceOwnerSource(current));
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base before occurrence-owned table state');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return { root, current };
}

function initializeSplitFloatLivePageRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-split-float-live-page-'));
  const productionRoot = resolve(import.meta.dirname, '..');
  const current = readFileSync(join(productionRoot, 'packages/docx/src/renderer.ts'), 'utf8');
  const livePageCallback = '            () => pages.length - 1,\n';
  assert.ok(current.includes(livePageCallback));
  write(root, 'packages/docx/src/renderer.ts', current.replace(livePageCallback, ''));
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base before split float live page callback');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return { root, current };
}

function priorSplitParentCommitSource(current) {
  const startMarker = '            (sliceTp, tableWidthPt, tableHeightPt, externalRegistry) =>';
  const endMarker = `            (sliceEl) => pushTagged(sliceEl),
            () => pages.length - 1,
            { sourceIndex: i, record: retainedTableRecord(measureState, i), state: measureState },
          );`;
  const start = current.indexOf(startMarker);
  const end = current.indexOf(endMarker, start);
  assert.ok(start >= 0 && end > start);
  const legacy = `            (sliceEl) => {
              pushTagged(sliceEl);
              return withColumnBand(() => {
                const sp = sliceEl as PaginatedBodyElement;
                const sliceTp = (sliceEl as unknown as DocTable).tblpPr as TblpPr;
                const { widthPx: tableW, heightPx: sliceH } = retainedTableSliceSize(
                  sp, measureState.scale,
                );
                const skipVClamp = sliceTp.vertAnchor === 'page' || sliceTp.vertAnchor === 'margin';
                return computeFloatTableBox(
                  sliceTp, measureState, measureState.y, tableW, sliceH, skipVClamp,
                  { allowOverlap: tbl.overlap !== 'never' },
                );
              });
            },
            { sourceIndex: i, record: retainedTableRecord(measureState, i), state: measureState },
          );
`;
  return current.slice(0, start) + legacy + current.slice(end + endMarker.length);
}

function initializeSplitParentCommitRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-split-parent-commit-'));
  const productionRoot = resolve(import.meta.dirname, '..');
  const current = readFileSync(join(productionRoot, 'packages/docx/src/renderer.ts'), 'utf8');
  write(root, 'packages/docx/src/renderer.ts', priorSplitParentCommitSource(current));
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base with ordinary split parent registration');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return { root, current };
}

function initializeComputePagesTableStampRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-compute-pages-table-stamp-'));
  write(root, 'packages/docx/src/renderer.ts', computePagesTableStampBaseline);
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

const computeTableLayoutCommonTail = `
  const colWidths = resolveColumnWidths(table, contentWPt1, state).map((w) => w * scale);
  const tableW = colWidths.reduce((s, w) => s + w, 0);
  const rowContentHeights = resolveTableRowContentHeights(table, colWidths, scale, (cell, cellW) =>
    measureCellContentHeightPx(cell, table, cellW, scale, state),
  );
  const rowHeights = applyTableRowBoundaryFootprints(table, rowContentHeights, scale);
  return { colWidths, tableW, rowContentHeights, rowHeights };
}
`;

const computeTableLayoutStampBaseline = `
export function computePages() { return []; }
export function computeTableLayout(table: any, contentWPx: number, state: any) {
  const { scale } = state;
  const contentHeightsFromResolved = (rowHeights: number[]): number[] => {
    const footprints = applyTableRowBoundaryFootprints(table, new Array<number>(table.rows.length).fill(0), scale);
    return rowHeights.map((height, index) => height - (footprints[index] ?? 0));
  };
  const stamped = table as PaginatedBodyElement;
  const contentWPt1 = contentWPx / scale;
  const placedFragment = bodyFlowFragments.get(table as object);
  const fragmentBandPt = tableFragmentBandPt.get(table as object);
  if (tableReuseEnabled && placedFragment !== undefined && placedFragment.fragment.kind === 'table' && fragmentBandPt !== undefined && placedFragment.fragment.rows.length === table.rows.length && Math.abs(fragmentBandPt - contentWPt1) <= 1e-6 * Math.max(1, Math.abs(contentWPt1))) {
    const fragment = placedFragment.fragment;
    const colWidths = fragment.columnWidthsPt.map((w) => w * scale);
    const rowHeights = fragment.rows.map((r) => r.heightPt * scale);
    return { colWidths, tableW: colWidths.reduce((s, w) => s + w, 0), rowContentHeights: contentHeightsFromResolved(rowHeights), rowHeights };
  }
  const reuseInputs = stamped.tableLayoutInputs;
  const reuse = tableReuseEnabled && reuseInputs !== undefined && stamped.tableColWidthsPt !== undefined && stamped.tableRowHeightsPt !== undefined && reuseInputs.scale === 1 && stamped.tableRowHeightsPt.length === table.rows.length && Math.abs(reuseInputs.contentWPt - contentWPt1) <= 1e-6 * Math.max(1, Math.abs(contentWPt1));
  if (reuse) {
    const colWidths = (stamped.tableColWidthsPt as number[]).map((w) => w * scale);
    const rowHeights = (stamped.tableRowHeightsPt as number[]).map((h) => h * scale);
    return { colWidths, tableW: colWidths.reduce((s, w) => s + w, 0), rowContentHeights: contentHeightsFromResolved(rowHeights), rowHeights };
  }
${computeTableLayoutCommonTail.slice(1)}`;

const computeTableLayoutRetainedMigration = `
export function computePages() { return []; }
export function computeTableLayout(table: any, contentWPx: number, state: any) {
  const { scale } = state;
  const contentWPt1 = contentWPx / scale;
  if (state.retainedTableAcquisition && bodySourceIndexFor(table) !== undefined) {
    const retained = computeTablePtLayout(state, table, contentWPt1);
    const colWidths = retained.colWidthsPt.map((width) => width * scale);
    const rowHeights = retained.rowHeightsPt.map((height) => height * scale);
    return {
      colWidths,
      tableW: colWidths.reduce((sum, width) => sum + width, 0),
      rowContentHeights: retained.rowContentHeightsPt.map((height) => height * scale),
      rowHeights,
    };
  }
${computeTableLayoutCommonTail.slice(1)}`;

function initializeComputeTableLayoutRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-compute-table-layout-'));
  write(root, 'packages/docx/src/renderer.ts', computeTableLayoutStampBaseline);
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

function removeComputeTableLayoutStampCounts(root) {
  const baselinePath = join(root, 'scripts/docx-layout-boundary-baseline.json');
  const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
  for (const symbol of [
    'tableReuseEnabled',
    'tableColWidthsPt',
    'tableRowHeightsPt',
    'tableLayoutInputs',
  ]) {
    delete baseline.legacySymbolCounts[`packages/docx/src/renderer.ts#${symbol}`];
  }
  baseline.migrationIdentifierCounts = {};
  writeFileSync(baselinePath, `${JSON.stringify(baseline, null, 2)}\n`);
}

function removeComputePagesTableStampCounts(root) {
  const baselinePath = join(root, 'scripts/docx-layout-boundary-baseline.json');
  const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
  for (const symbol of ['tableColWidthsPt', 'tableRowHeightsPt']) {
    delete baseline.legacySymbolCounts[`packages/docx/src/renderer.ts#${symbol}`];
  }
  writeFileSync(baselinePath, `${JSON.stringify(baseline, null, 2)}\n`);
}

const computePagesAcquisitionBaseline = `
const EMPTY_HEADERS_FOOTERS = {};
function buildMeasureState(ctx: unknown, section: unknown, fontFamilyClasses: unknown, documentSettings: unknown, resolvedLocalFonts: unknown) { return { ctx, section, fontFamilyClasses, documentSettings, resolvedLocalFonts }; }
function createLayoutServices(input: unknown, options: unknown) { return { input, options }; }
function splitParagraphAcrossPages(...args: unknown[]) { return args; }
function attachBodyParagraphFragment(...args: unknown[]) { return args; }
function attachTableFragment(...args: unknown[]) { return args; }
export function computePages(body: unknown[], section: unknown, ctx: unknown, fontFamilyClasses: unknown = {}, footnotes: unknown[] = [], settings?: unknown, resolvedLocalFonts: unknown = {}) {
  const documentSettings = settings;
  const measureState = buildMeasureState(ctx, section, fontFamilyClasses, documentSettings, resolvedLocalFonts);
  const pages: unknown[][] = [[]];
  const footnoteReservePt: number[] = [];
  let pageNoteIds = new Set<string>();
  const stampPageMeta = () => {};
  const startPageBookkeeping = () => {
    footnoteReservePt[pages.length - 1] = 0;
    pageNoteIds = new Set<string>();
    stampPageMeta();
  };
  const para = {};
  const el = {};
  const i = 0;
  splitParagraphAcrossPages(measureState, para, 'tail');
  attachBodyParagraphFragment(el as PaginatedElementWithLines, para, measureState, {});
  attachTableFragment(el, el, [], [], 0, measureState, { columnIndex: 0 });
  attachTableFragment(el, el, [], [], 0, measureState, { columnIndex: 1 });
  startPageBookkeeping();
  return pages;
}
export function computeTableLayout() { return []; }
`;

const computePagesAcquisitionMigration = `
const EMPTY_HEADERS_FOOTERS = {};
function buildMeasureState(ctx: unknown, section: unknown, fontFamilyClasses: unknown, documentSettings: unknown, resolvedLocalFonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return { ctx, section, fontFamilyClasses, documentSettings, resolvedLocalFonts, services, options }; }
function createLayoutServices(input: unknown, options: unknown) { return { input, options }; }
function splitParagraphAcrossPages(...args: unknown[]) { return args; }
function attachBodyParagraphFragment(...args: unknown[]) { return args; }
function attachTableFragment(...args: unknown[]) { return args; }
export function computePages(body: unknown[], section: unknown, ctx: unknown, fontFamilyClasses: unknown = {}, footnotes: unknown[] = [], settings?: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions) {
  const documentSettings = settings;
  const effectiveLayoutServices = layoutServices ?? createLayoutServices({
    section,
    body,
    headers: EMPTY_HEADERS_FOOTERS,
    footers: EMPTY_HEADERS_FOOTERS,
    ...(settings === undefined ? {} : { settings }),
    ...(footnotes.length === 0 ? {} : { footnotes }),
    fontFamilyClasses,
  }, {
    measureContext: ctx,
    localMetrics: resolvedLocalFonts,
  });
  const measureState = buildMeasureState(ctx, section, fontFamilyClasses, documentSettings, resolvedLocalFonts, effectiveLayoutServices, layoutOptions);
  const pages: unknown[][] = [[]];
  const footnoteReservePt: number[] = [];
  let pageNoteIds = new Set<string>();
  const stampPageMeta = () => {};
  const startPageBookkeeping = () => {
    footnoteReservePt[pages.length - 1] = 0;
    pageNoteIds = new Set<string>();
    measureState.pageIndex = pages.length - 1;
    measureState.displayPageNumber = undefined;
    stampPageMeta();
  };
  const para = {};
  const el = {};
  const i = 0;
  splitParagraphAcrossPages(measureState, para, i, 'tail');
  attachBodyParagraphFragment(el, para, i, measureState, {});
  attachTableFragment(el, el, [], [], 0, measureState, { columnIndex: 0, sourcePath: [i] });
  attachTableFragment(el, el, [], [], 0, measureState, { columnIndex: 1, sourcePath: [i] });
  startPageBookkeeping();
  return pages;
}
export function computeTableLayout() { return []; }
`;

function initializeComputePagesAcquisitionRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-compute-pages-'));
  write(root, 'packages/docx/src/renderer.ts', computePagesAcquisitionBaseline);
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

const tablePaintGateBaseline = `
const fragmentPaintEnabled = true;
function tableRequiresLegacyPaint(table: any) { return table.legacy; }
export function isFragmentPaintableTable(table: any, placed: any, state: any): boolean {
  if (
    !fragmentPaintEnabled ||
    placed === undefined ||
    placed.fragment.kind !== 'table' ||
    table.tblpPr != null ||
    state.verticalCJK ||
    tableRequiresLegacyPaint(table)
  ) return false;
  const bandPt = table.fragmentBand;
  if (bandPt === undefined) return false;
  const paintBandPt = state.contentW / state.scale;
  return Math.abs(bandPt - paintBandPt) <= 1e-6 * Math.max(1, Math.abs(paintBandPt));
}
`;

const tablePaintGateMigration = tablePaintGateBaseline
  .replace('const fragmentPaintEnabled = true;\n', '')
  .replace('    !fragmentPaintEnabled ||\n', '');

function initializeTablePaintGateRepository(pruneDeletedFlag = true) {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-table-gate-'));
  write(root, 'packages/docx/src/renderer.ts', tablePaintGateBaseline);
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  if (pruneDeletedFlag) {
    const baselinePath = join(root, 'scripts/docx-layout-boundary-baseline.json');
    const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
    for (const group of [
      'legacySymbolCounts',
      'migrationIdentifierCounts',
      'legacyDeclarationHashes',
    ]) {
      for (const key of Object.keys(baseline[group])) {
        if (key.includes('fragmentPaintEnabled')) delete baseline[group][key];
      }
    }
    baseline.nonLayoutDeclarationKeys = baseline.nonLayoutDeclarationKeys.filter(
      (key) => !key.includes('fragmentPaintEnabled'),
    );
    writeFileSync(baselinePath, `${JSON.stringify(baseline, null, 2)}\n`);
  }
  return root;
}

function initializeLayoutParserBoundaryRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-parser-boundary-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/parser-model.ts', 'export function normalizeInternalDocumentModel(document: unknown) { return { document, mathOccurrences: [] }; }\nexport const parserFacts = true;\nexport interface ParserFacts { value: string; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  return root;
}

const exactParserGatewaySource =
  "import { normalizeInternalDocumentModel } from '../parser-model.js';\n"
  + 'export function documentMathOccurrences(doc: unknown): unknown[] {\n'
  + '  return [...normalizeInternalDocumentModel(doc).mathOccurrences];\n'
  + '}\n';

function initializeShapeRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-shape-'));
  write(root, 'packages/docx/src/renderer.ts', 'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>) { return String([bold, italic, size, family, classes]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown }, classes: Record<string, string>) { return buildFont(s.bold, s.italic, s.size, s.family, classes); }\n');
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

function initializeMetricShapeRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-shape-metric-'));
  write(root, 'packages/docx/src/renderer.ts', 'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown; eaFloorRoute?: unknown }, classes: Record<string, string>) { const shapeLineMetrics = (family: string | null) => { return buildFont(s.bold, s.italic, s.size, family, classes); }; return shapeLineMetrics(s.family); }\n');
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

function initializeNumberingShapeRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-shape-numbering-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function renderShapeText(block: any, ctx: any, scale: number, effState: any, eaVertUpright: boolean, markerX: number, baseline: number) { const markerText = markerDisplayText(block.numbering); const markerW = ctx.measureText(markerText).width; if (eaVertUpright) { drawVerticalRun(ctx, markerText, markerX, baseline, block.fontSizePt * scale, 0); } else { ctx.fillText(markerText, markerX, baseline); } return markerW; }\n');
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  establishA1Baseline(root);
  return root;
}

const exactNumberingShapeSource = 'export function renderShapeText(block: any, ctx: any, scale: number, effState: any, eaVertUpright: boolean, markerX: number, baseline: number) { const markerText = markerDisplayText(block.numbering); const markerShapeInput = numberingMarkerShapeInput(block.numbering, block.fontSizePt); const markerTextLayout = shapeNumberingMarkerText(markerShapeInput, markerText, scale, effState.layoutServices?.text,); const markerW = markerTextLayout?.shape.advancePt ?? ctx.measureText(markerText).width; if (markerTextLayout) { paintNumberingMarkerText(ctx, markerTextLayout, markerX, baseline, eaVertUpright ? (paintCtx, text, drawX, drawBaseline, fontSizePx) => { drawVerticalRun(paintCtx, text, drawX, drawBaseline, fontSizePx, 0); } : undefined,); } else if (eaVertUpright) { drawVerticalRun(ctx, markerText, markerX, baseline, block.fontSizePt * scale, 0); } else { ctx.fillText(markerText, markerX, baseline); } return markerW; }\n';

function establishA1Baseline(root) {
  const writeResult = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');
  assert.equal(writeResult.status, 0, writeResult.output);
  const checkResult = runChecker(root, '--base-ref', 'main');
  assert.equal(checkResult.status, 0, checkResult.output);
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'establish boundary');
  git(root, 'switch', 'main');
  git(root, 'merge', '--ff-only', 'a1');
  git(root, 'switch', '-c', 'a2');
}

test('allows a migrated legacy declaration to be deleted from source and baseline', () => {
  const root = initializeShapeRepository();
  write(root, 'packages/docx/src/renderer.ts', 'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>) { return String([bold, italic, size, family, classes]); }\n');
  const baselinePath = join(root, 'scripts/docx-layout-boundary-baseline.json');
  const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
  for (const group of ['legacySymbolCounts', 'migrationIdentifierCounts', 'legacyDeclarationHashes']) {
    for (const key of Object.keys(baseline[group])) {
      if (key.includes('renderShapeText')) delete baseline[group][key];
    }
  }
  baseline.nonLayoutDeclarationKeys = baseline.nonLayoutDeclarationKeys.filter(
    (key) => !key.includes('renderShapeText'),
  );
  writeFileSync(baselinePath, `${JSON.stringify(baseline, null, 2)}\n`);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('rejects a transitive paint edge to a measurement module', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-edge-'));
  write(root, 'packages/docx/src/renderer.ts', 'export const adapter = true;\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { helper } from './helper.js';\nexport { helper };\n");
  write(root, 'packages/docx/src/paint/helper.ts', "import { layoutLines } from '../line-layout.js';\nexport const helper = layoutLines;\n");
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
  assert.match(result.output, /canvas-page\.ts.*helper\.ts.*line-layout\.ts/s);
});

test('rejects new parser-model dependencies inside retained layout modules', () => {
  const root = initializeRepository();
  establishA1Baseline(root);
  write(root, 'packages/docx/src/parser-model.ts', 'export const parserFacts = true;\n');
  write(
    root,
    'packages/docx/src/layout/numbering-marker.ts',
    "import { parserFacts } from '../parser-model.js';\nexport const marker = parserFacts;\n",
  );

  const result = runChecker(root, '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /LAYOUT_PARSER_MODEL_DEPENDENCY/);
  assert.match(result.output, /layout\/numbering-marker\.ts.*parser-model\.ts/s);
});

test('rejects transitive layout-to-parser-model bridge modules', () => {
  const root = initializeLayoutParserBoundaryRepository();
  write(
    root,
    'packages/docx/src/layout/numbering-marker.ts',
    "import { bridgeFacts as markerFacts } from '../parser-bridge.js';\nexport const marker = markerFacts;\n",
  );
  write(
    root,
    'packages/docx/src/parser-bridge.ts',
    "import { parserFacts } from './parser-model.js';\nexport const bridgeFacts = parserFacts;\n",
  );

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /LAYOUT_PARSER_MODEL_DEPENDENCY/);
  assert.match(
    result.output,
    /layout\/numbering-marker\.ts.*parser-bridge\.ts.*parser-model\.ts/s,
  );
});

test('rejects runtime impurity reachable from retained occurrence projection', () => {
  const root = initializeRepository();
  write(
    root,
    'packages/docx/src/paragraph-measure.ts',
    'export const measureParagraph = () => 1;\n',
  );
  establishA1Baseline(root);
  write(
    root,
    'packages/docx/src/layout/retained-geometry-translation.ts',
    "import { measureParagraph } from '../paragraph-measure.js';\nexport const translate = measureParagraph;\n",
  );
  write(
    root,
    'packages/docx/src/layout/occurrence-projection.ts',
    "import { translate } from './retained-geometry-translation.js';\nexport const project = translate;\n",
  );

  const result = runChecker(root, '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY/);
  assert.match(
    result.output,
    /occurrence-projection\.ts.*retained-geometry-translation\.ts.*paragraph-measure\.ts/s,
  );
});

test('allows the reviewed occurrence projection runtime helper graph', () => {
  const root = initializeRepository();
  establishA1Baseline(root);
  write(
    root,
    'packages/docx/src/layout/plain-data.ts',
    'export const snapshotPlainData = (value) => value;\n',
  );
  write(
    root,
    'packages/docx/src/layout/retained-geometry-translation.ts',
    "import { snapshotPlainData } from './plain-data.js';\nexport const translate = snapshotPlainData;\n",
  );
  write(
    root,
    'packages/docx/src/layout/occurrence-projection.ts',
    "import { translate } from './retained-geometry-translation.js';\nexport const project = translate;\n",
  );

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

for (const [path, specifier] of [
  ['packages/docx/src/layout/text-measure.ts', './text-measure.js'],
  ['packages/docx/src/layout/plain-data.ts', './plain-data.js?worker&inline'],
  [null, 'node:worker_threads'],
  [null, 'jsdom'],
]) {
  test(`rejects occurrence projection runtime dependency outside the allowlist: ${specifier}`, () => {
    const root = initializeRepository();
    establishA1Baseline(root);
    if (path) write(root, path, 'export const forbiddenRuntime = 1;\n');
    write(
      root,
      'packages/docx/src/layout/occurrence-projection.ts',
      `import * as forbiddenRuntime from '${specifier}';\nexport const project = forbiddenRuntime;\n`,
    );

    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, specifier);
    assert.match(result.output, /OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY/, specifier);
  });
}

test('allows retained and legacy layout to share parser-independent primitives', () => {
  const root = initializeLayoutParserBoundaryRepository();
  write(
    root,
    'packages/docx/src/line-layout-primitives.ts',
    'export const nextTabStop = () => 36;\n',
  );
  write(
    root,
    'packages/docx/src/line-layout.ts',
    "import { nextTabStop } from './line-layout-primitives.js';\n"
      + 'export const measuredLine = nextTabStop();\n',
  );
  write(
    root,
    'packages/docx/src/layout/numbering-marker.ts',
    "import { nextTabStop } from '../line-layout-primitives.js';\n"
      + 'export const markerStop = nextTabStop();\n',
  );

  const result = runChecker(root, '--final');

  assert.equal(result.status, 0, result.output);
});

test('rejects literal dynamic and CommonJS layout-to-parser-model bridges', () => {
  for (const source of [
    "export const load = () => import('../parser-bridge.js');\n",
    "export const load = () => require('../parser-bridge.js');\n",
  ]) {
    const root = initializeLayoutParserBoundaryRepository();
    write(root, 'packages/docx/src/layout/numbering-marker.ts', source);
    write(
      root,
      'packages/docx/src/parser-bridge.ts',
      "export { parserFacts } from './parser-model.js';\n",
    );

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0);
    assert.match(result.output, /LAYOUT_PARSER_MODEL_DEPENDENCY/);
    assert.match(result.output, /numbering-marker\.ts.*parser-bridge\.ts.*parser-model\.ts/s);
  }
});

test('rejects non-literal dynamic and CommonJS imports reachable from layout', () => {
  for (const source of [
    "export const load = (name: string) => import(`../${name}.js`);\n",
    "export const load = (name: string) => require('../' + name + '.js');\n",
  ]) {
    const root = initializeLayoutParserBoundaryRepository();
    write(root, 'packages/docx/src/layout/numbering-marker.ts', source);

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0);
    assert.match(result.output, /NON_LITERAL_LAYOUT_MODULE_EDGE/);
    assert.match(result.output, /layout\/numbering-marker\.ts/);
  }
});

test('allows only the exact parser normalization gateway and erased type-only contracts', () => {
  const gateway = initializeLayoutParserBoundaryRepository();
  write(
    gateway,
    'packages/docx/src/layout/numbering-marker.ts',
    "import { documentMathOccurrences } from './resources.js';\nexport const marker = documentMathOccurrences;\n",
  );
  write(
    gateway,
    'packages/docx/src/layout/resources.ts',
    exactParserGatewaySource,
  );
  assert.equal(runChecker(gateway, '--final').status, 0);

  const typeOnly = initializeLayoutParserBoundaryRepository();
  write(
    typeOnly,
    'packages/docx/src/layout/numbering-marker.ts',
    "import type { BridgeFacts } from '../parser-contract.js';\nexport type MarkerFacts = BridgeFacts;\n",
  );
  write(
    typeOnly,
    'packages/docx/src/parser-contract.ts',
    "import { parserFacts } from './parser-model.js';\nexport interface BridgeFacts { value: typeof parserFacts; }\n",
  );
  assert.equal(runChecker(typeOnly, '--final').status, 0);
});

test('rejects non-exact parser-model syntax in the parser normalization gateway', () => {
  for (const source of [
    "import { normalizeInternalDocumentModel, parserFacts } from '../parser-model.js';\nexport const value = [normalizeInternalDocumentModel, parserFacts];\n",
    "import { normalizeInternalDocumentModel as normalize } from '../parser-model.js';\nexport const value = normalize;\n",
    "import * as parserModel from '../parser-model.js';\nexport const value = parserModel;\n",
    "export { normalizeInternalDocumentModel } from '../parser-model.js';\n",
    "export * from '../parser-model.js';\n",
  ]) {
    const root = initializeLayoutParserBoundaryRepository();
    write(root, 'packages/docx/src/layout/resources.ts', source);

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LAYOUT_PARSER_MODEL_DEPENDENCY/, source);
    assert.match(result.output, /layout\/resources\.ts.*parser-model\.ts/s, source);
  }
});

test('rejects parser normalizer re-exports, leaks, and local aliases in the gateway', () => {
  for (const source of [
    `${exactParserGatewaySource}export { normalizeInternalDocumentModel };\n`,
    `${exactParserGatewaySource}export const leak = normalizeInternalDocumentModel;\n`,
    "import { normalizeInternalDocumentModel } from '../parser-model.js';\nconst normalize = normalizeInternalDocumentModel;\nexport function documentMathOccurrences(doc: unknown): unknown[] { return [...normalize(doc).mathOccurrences]; }\n",
  ]) {
    const root = initializeLayoutParserBoundaryRepository();
    write(root, 'packages/docx/src/layout/resources.ts', source);

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LAYOUT_PARSER_MODEL_DEPENDENCY/, source);
    assert.match(result.output, /layout\/resources\.ts.*normalizeInternalDocumentModel/s, source);
  }
});

test('traverses gateway bridges and rejects literal dynamic and CommonJS parser edges', () => {
  const bridge = initializeLayoutParserBoundaryRepository();
  write(
    bridge,
    'packages/docx/src/layout/resources.ts',
    "import { normalizeInternalDocumentModel } from '../parser-model.js';\nimport { bridgeFacts } from '../parser-bridge.js';\nexport const value = [normalizeInternalDocumentModel, bridgeFacts];\n",
  );
  write(
    bridge,
    'packages/docx/src/parser-bridge.ts',
    "import { parserFacts } from './parser-model.js';\nexport const bridgeFacts = parserFacts;\n",
  );
  const bridged = runChecker(bridge, '--final');
  assert.notEqual(bridged.status, 0);
  assert.match(bridged.output, /LAYOUT_PARSER_MODEL_DEPENDENCY/);
  assert.match(bridged.output, /layout\/resources\.ts.*parser-bridge\.ts.*parser-model\.ts/s);

  for (const source of [
    "export const load = () => import('../parser-model.js');\n",
    "export const load = () => require('../parser-model.js');\n",
  ]) {
    const root = initializeLayoutParserBoundaryRepository();
    write(root, 'packages/docx/src/layout/resources.ts', source);

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LAYOUT_PARSER_MODEL_DEPENDENCY/, source);
  }
});

test('rejects non-literal dynamic and CommonJS edges in the parser gateway', () => {
  for (const source of [
    "export const load = (name: string) => import(`../${name}.js`);\n",
    "export const load = (name: string) => require('../' + name + '.js');\n",
  ]) {
    const root = initializeLayoutParserBoundaryRepository();
    write(root, 'packages/docx/src/layout/resources.ts', source);

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /NON_LITERAL_LAYOUT_MODULE_EDGE/, source);
    assert.match(result.output, /layout\/resources\.ts/, source);
  }
});

test('rejects any paint runtime dependency outside the paint owner directory', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-arbitrary-edge-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { helper } from '../text-wrap.js';\nexport { helper };\n");
  write(root, 'packages/docx/src/text-wrap.ts', 'export const helper = true;\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
  assert.match(result.output, /canvas-page\.ts.*text-wrap\.ts/s);
});

test('allows only the layout contract as a type-only paint dependency', () => {
  const allowed = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-type-edge-'));
  write(allowed, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(allowed, 'packages/docx/src/layout/types.ts', 'export interface Layout { pages: number; }\n');
  write(allowed, 'packages/docx/src/paint/canvas-page.ts', "import type { Layout } from '../layout/types.js';\nexport type Page = Layout;\n");
  assert.equal(runChecker(allowed, '--final').status, 0);

  write(allowed, 'packages/docx/src/layout/flow.ts', 'export interface Flow { y: number; }\n');
  write(allowed, 'packages/docx/src/paint/canvas-page.ts', "import type { Flow } from '../layout/flow.js';\nexport type Page = Flow;\n");
  const forbidden = runChecker(allowed, '--final');
  assert.notEqual(forbidden.status, 0);
  assert.match(forbidden.output, /FORBIDDEN_PAINT_EDGE/);
});

test('allows only named shared atomic painters from core', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-shared-paint-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { renderChart } from '@silurus/ooxml-core';\nexport { renderChart };\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { canvasFontString } from '@silurus/ooxml-core';\nexport { canvasFontString };\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { drawImageCropped } from '@silurus/ooxml-core';\nexport { drawImageCropped };\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { paintDrawingMLShape } from '@silurus/ooxml-core';\nexport { paintDrawingMLShape };\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { autoContrastColor } from '@silurus/ooxml-core';\nexport { autoContrastColor };\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { resolveFill } from '@silurus/ooxml-core';\nexport { resolveFill };\n");
  assert.equal(runChecker(root, '--final').status, 0);

  for (const name of ['crispOffset', 'doubleRailGeometry', 'fillDoubleBorder']) {
    write(
      root,
      'packages/docx/src/paint/canvas-page.ts',
      `import { ${name} } from '@silurus/ooxml-core';\nvoid ${name};\n`,
    );
    assert.equal(runChecker(root, '--final').status, 0, name);
  }

  write(root, 'packages/docx/src/paint/canvas-page.ts', "import type { HyperlinkTarget } from '@silurus/ooxml-core';\nexport type Target = HyperlinkTarget;\n");
  assert.equal(runChecker(root, '--final').status, 0);

  for (const source of [
    "import { createCanvasFontRoute } from '@silurus/ooxml-core';\nexport { createCanvasFontRoute };\n",
    "import { resolveFill as localFill } from '@silurus/ooxml-core';\nexport { localFill };\n",
    "import type { CanvasFontRoute } from '@silurus/ooxml-core';\nexport type Route = CanvasFontRoute;\n",
    "import { HyperlinkTarget } from '@silurus/ooxml-core';\nexport { HyperlinkTarget };\n",
    "import type { renderChart } from '@silurus/ooxml-core';\nexport type Renderer = typeof renderChart;\n",
    "import { canvasFontString as fontString } from '@silurus/ooxml-core';\nexport { fontString };\n",
    "import { crispOffset as snap } from '@silurus/ooxml-core';\nvoid snap;\n",
    "import * as core from '@silurus/ooxml-core';\nexport { core };\n",
    "import { measureTextWidth } from '@silurus/ooxml-core';\nvoid measureTextWidth;\n",
    "export { fillDoubleBorder } from '@silurus/ooxml-core';\n",
    "export const load = () => import('@silurus/ooxml-core');\n",
  ]) {
    write(root, 'packages/docx/src/paint/canvas-page.ts', source);
    const rejected = runChecker(root, '--final');
    assert.notEqual(rejected.status, 0);
    assert.match(rejected.output, /FORBIDDEN_PAINT_EDGE/);
  }

  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { measureTextWidth as renderChart } from '@silurus/ooxml-core';\nexport { renderChart };\n");
  const result = runChecker(root, '--final');
  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
});

test('rejects forbidden core APIs reached through a paint helper', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-shared-paint-transitive-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { paint } from './helper.js';\nvoid paint;\n");
  write(root, 'packages/docx/src/paint/helper.ts', "import { measureTextWidth } from '@silurus/ooxml-core';\nexport const paint = measureTextWidth;\n");

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
  assert.match(result.output, /canvas-page\.ts.*helper\.ts.*@silurus\/ooxml-core/s);
});

test('keeps layout-only border treatment outside the paint dependency graph', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-border-treatment-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/layout/types.ts', 'export interface Layout { pages: number; }\n');
  write(root, 'packages/docx/src/layout/border-treatment.ts', "import { docxBorderDashArray } from '@silurus/ooxml-core';\nexport const treatment = docxBorderDashArray('dotDash', 1);\n");
  write(root, 'packages/docx/src/layout/paragraph.ts', "import { treatment } from './border-treatment.js';\nexport const paragraph = treatment;\n");
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import type { Layout } from '../layout/types.js';\nexport type Page = Layout;\n");

  const result = runChecker(root, '--final');

  assert.equal(result.status, 0, result.output);
});

test('rejects paint dependencies on layout resource implementations even when type-only', () => {
  for (const source of [
    "import { createPaintResourceRegistry } from '../layout/paint-resources.js';\nexport const registry = createPaintResourceRegistry([]);\n",
    "import type { PaintResourceRegistry } from '../layout/paint-resources.js';\nexport type Registry = PaintResourceRegistry;\n",
  ]) {
    const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-resource-owner-'));
    write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
    write(root, 'packages/docx/src/layout/paint-resources.ts', 'export interface PaintResourceRegistry {}\nexport function createPaintResourceRegistry() {}\n');
    write(root, 'packages/docx/src/paint/canvas-page.ts', source);

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /FORBIDDEN_PAINT_EDGE/, source);
  }
});

test('audits dependencies of the retained page graph allowed in paint', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-page-graph-edge-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { orderedPagePaintNodes } from '../layout/page-graph.js';\nexport { orderedPagePaintNodes };\n");
  write(root, 'packages/docx/src/layout/page-graph.ts', "import { measureTextWidth } from '../measurement.js';\nexport const orderedPagePaintNodes = measureTextWidth;\n");
  write(root, 'packages/docx/src/measurement.ts', 'export function measureTextWidth() { return 1; }\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
  assert.match(result.output, /canvas-page\.ts.*page-graph\.ts.*measurement\.ts/s);
});

test('rejects non-literal dynamic paint imports', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-dynamic-edge-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export const load = (name: string) => import(`../${name}.js`);\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /NON_LITERAL_MODULE_EDGE/);
});

test('rejects computed measurement access in paint and display inputs in layout TSX', () => {
  const paintRoot = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-computed-'));
  write(paintRoot, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(paintRoot, 'packages/docx/src/paint/canvas-page.ts', "export const width = (ctx: CanvasRenderingContext2D) => ctx['measureText']('x').width;\n");
  const paintResult = runChecker(paintRoot, '--final');
  assert.notEqual(paintResult.status, 0);
  assert.match(paintResult.output, /PAINT_CAPABILITY/);

  const layoutRoot = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-layout-display-'));
  write(layoutRoot, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(layoutRoot, 'packages/docx/src/layout/page.tsx', 'export const Page = ({ dpr }: { dpr: number }) => dpr;\n');
  const layoutResult = runChecker(layoutRoot, '--final');
  assert.notEqual(layoutResult.status, 0);
  assert.match(layoutResult.output, /LAYOUT_DISPLAY_CAPABILITY/);
});

test('rejects layout acquisition from the retained body paint adapter', () => {
  for (const call of [
    'layoutLines()',
    'measureParagraph()',
    'ctx.measureText()',
    "ctx['measureText']()",
    'buildSegments()',
    'contextualSpacingAdjust()',
    'estimateParagraphHeight()',
    'paragraphLayoutFromMeasurement()',
    'paragraphGapAdjustment()',
    'parasShareBorderBox()',
    'resolveParagraphBorderEdges()',
    'acquireParagraphLayout()',
    'acquireRetainedFrameGroup()',
    'renderFrameParagraph()',
    'renderParagraph()',
    'services.resolveFrameBox()',
  ]) {
    const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-body-paint-'));
    write(
      root,
      'packages/docx/src/renderer.ts',
      `function renderBodyElements(elements: any[]) { for (const el of elements) { if (el.type === 'paragraph') { ${call}; } } }\n`,
    );
    write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0, call);
    assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/, call);
  }
});

test('rejects aliased, computed, and unresolved calls from retained body paint', () => {
  const renderers = [
    `import { measureParagraph as paintRetainedParagraph } from './paragraph-measure.js';
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') paintRetainedParagraph();
     }`,
    `const paintRetainedParagraph = measureParagraph;
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') paintRetainedParagraph();
     }`,
    `const paintRetainedParagraph = layout.measureParagraph;
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') paintRetainedParagraph();
     }`,
    `const { measureParagraph: paintRetainedParagraph } = layout;
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') paintRetainedParagraph();
     }`,
    `const operations = { paintPlacedParagraphLayout: measureParagraph };
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') operations.paintPlacedParagraphLayout();
     }`,
    `const operations = { paintPlacedParagraphLayout: measureParagraph };
     const { paintPlacedParagraphLayout: paintRetainedParagraph } = operations;
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') paintRetainedParagraph();
     }`,
    `function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') layout['measure' + 'Paragraph']();
     }`,
    `const operation = 'measureParagraph';
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') layout[operation]();
     }`,
    `function paintPlacedParagraphLayout() { measureParagraph(); }
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') paintPlacedParagraphLayout();
     }`,
    `function bodyFragmentFor() { measureParagraph(); }
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') bodyFragmentFor();
     }`,
    `import { paintPlacedParagraphLayout } from './paragraph-measure.js';
     function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') paintPlacedParagraphLayout();
     }`,
    `function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') layout.save();
     }`,
    `function renderBodyElements(elements: any[]) {
       for (const el of elements) if (el.type === 'paragraph') opaquePaint();
     }`,
  ];

  for (const renderer of renderers) {
    const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-body-paint-alias-'));
    write(root, 'packages/docx/src/renderer.ts', renderer);
    write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');

    const result = runChecker(root, '--final');

    assert.notEqual(result.status, 0, renderer);
    assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/, renderer);
  }
});

test('rejects body fragment lookup receivers that can override the canonical WeakMap get', () => {
  const spoofedReceivers = [
    `const bodyFlowFragments = {
      get(_element: object) { measureParagraph(); }
    };`,
    `const bodyFlowFragments = Object.assign(new WeakMap<object, unknown>(), {
      ...{ get(_element: object) { measureParagraph(); } }
    });`,
    `const bodyFlowFragments = Object.assign(new WeakMap<object, unknown>(), {
      [['g'][0] + 'et'](_element: object) { measureParagraph(); }
    });`,
    `const bodyFlowFragments = Object.assign(new WeakMap<object, unknown>(), {
      get sourceIndices() { measureParagraph(); return new WeakMap<object, number>(); }
    });`,
    `const bodyFlowFragments = Object.assign(new WeakMap<object, unknown>(), {
      sourceIndices: { get(_element: object) { measureParagraph(); } }
    });`,
    `const bodyFlowFragments = Object.freeze(Object.assign(new WeakMap<object, unknown>(), {
      sourceIndices: (() => new WeakMap<object, number>())(),
      framePlacement: new WeakMap<object, unknown>(),
    }));`,
    `const bodyFlowFragments = Object.freeze(Object.assign(new WeakMap<object, unknown>(), {
      sourceIndices: new WeakMap<object, number>(),
      framePlacement: new WeakMap<object, unknown>(),
      extraSidecar: new WeakMap<object, unknown>(),
    }));`,
    `const bodyFlowFragments = Object.freeze(Object.assign(new WeakMap<object, unknown>(), {
      sourceIndices: new WeakMap<object, number>(),
    }));`,
    `const bodyFlowFragments = Object.freeze(Object.assign(new WeakMap<object, unknown>(), {
      sourceIndices: new WeakMap<object, number>(),
      sourceIndices: new WeakMap<object, number>(),
    }));`,
  ];
  for (const receiver of spoofedReceivers) {
    const root = initializeRepository();
    const retainedBody = (mapDeclaration) => `${mapDeclaration}
      function bodyFragmentFor(element: object) { return bodyFlowFragments.get(element); }
      function renderBodyElements(elements: any[]) {
        for (const el of elements) {
          if (el.type === 'paragraph') bodyFragmentFor(el);
        }
      }
    `;
    write(
      root,
      'packages/docx/src/renderer.ts',
      retainedBody(`const bodyFlowFragments = Object.freeze(Object.assign(
        new WeakMap<object, unknown>(),
        {
          sourceIndices: new WeakMap<object, number>(),
          framePlacement: new WeakMap<object, unknown>(),
        },
      ));`),
    );
    establishA1Baseline(root);
    write(root, 'packages/docx/src/renderer.ts', retainedBody(receiver));

    const result = runChecker(root, '--base-ref', 'main');

    assert.notEqual(result.status, 0, receiver);
    assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/, receiver);
  }
});

test('rejects mutable body fragment receivers whose authority changes after initialization', () => {
  const mutations = [
    { frozen: false, declaration: 'const', source: `bodyFlowFragments.get = (_element: object) => {
      measureParagraph();
      return undefined;
    };` },
    { frozen: false, declaration: 'const', source: `Object.defineProperty(bodyFlowFragments, 'get', {
      value: (_element: object) => {
        measureParagraph();
        return undefined;
      },
    });` },
    { frozen: true, declaration: 'let', source: `bodyFlowFragments = {
      get(_element: object) { measureParagraph(); }
    } as unknown as typeof bodyFlowFragments;` },
  ];
  for (const mutation of mutations) {
    const root = initializeRepository();
    const retainedBody = (declaration, frozen, extra = '') => `
      ${declaration} bodyFlowFragments = ${frozen ? 'Object.freeze(' : ''}Object.assign(new WeakMap<object, unknown>(), {
        sourceIndices: new WeakMap<object, number>(),
        framePlacement: new WeakMap<object, unknown>(),
      })${frozen ? ')' : ''};
      ${extra}
      function bodyFragmentFor(element: object) { return bodyFlowFragments.get(element); }
      function renderBodyElements(elements: any[]) {
        for (const el of elements) {
          if (el.type === 'paragraph') bodyFragmentFor(el);
        }
      }
    `;
    write(root, 'packages/docx/src/renderer.ts', retainedBody('const', true));
    establishA1Baseline(root);
    write(
      root,
      'packages/docx/src/renderer.ts',
      retainedBody(mutation.declaration, mutation.frozen, mutation.source),
    );

    const result = runChecker(root, '--base-ref', 'main');

    assert.notEqual(result.status, 0, mutation.source);
    assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/, mutation.source);
  }
});

test('rejects late mutations of canonical body fragment sidecars', () => {
  const mutations = [
    `Object.assign(bodyFlowFragments.sourceIndices, {
      retainedTableMeasureBySource: new WeakMap<object, unknown>(),
    });`,
    `Object.assign(bodyFlowFragments['sourceIndices'], {
      retainedTableMeasureBySource: new WeakMap<object, unknown>(),
    });`,
    `Object.defineProperty(bodyFlowFragments.sourceIndices, 'retainedTableMeasureBySource', {
      value: new WeakMap<object, unknown>(),
    });`,
    `bodyFlowFragments.sourceIndices.retainedTableMeasureBySource =
      new WeakMap<object, unknown>();`,
  ];
  for (const mutation of mutations) {
    const root = initializeRepository();
    const retainedBody = (extra = '') => `
      const bodyFlowFragments = Object.freeze(Object.assign(new WeakMap<object, unknown>(), {
        sourceIndices: new WeakMap<object, number>(),
        framePlacement: new WeakMap<object, unknown>(),
      }));
      ${extra}
      function bodyFragmentFor(element: object) { return bodyFlowFragments.get(element); }
      function renderBodyElements(elements: any[]) {
        for (const el of elements) {
          if (el.type === 'paragraph') bodyFragmentFor(el);
        }
      }
    `;
    write(root, 'packages/docx/src/renderer.ts', retainedBody());
    establishA1Baseline(root);
    write(root, 'packages/docx/src/renderer.ts', retainedBody(mutation));

    const result = runChecker(root, '--base-ref', 'main');

    assert.notEqual(result.status, 0, mutation);
    assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/, mutation);
  }
});

test('rejects transitive layout acquisition from a local body paragraph paint helper', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-body-paint-transitive-'));
  write(
    root,
    'packages/docx/src/renderer.ts',
    `function renderBodyElements(elements: any[]) {
      const paintRetainedParagraph = () => paintBodyParagraph();
      for (const el of elements) {
        if (el.type === 'paragraph') paintRetainedParagraph();
        else if (el.type === 'table') renderTable();
      }
    }
    function paintBodyParagraph() { acquireParagraphLayout(); }
    function renderTable() { measureParagraph(); }
    `,
  );
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/);
  assert.match(result.output, /acquireParagraphLayout/);
});

test('rejects layout acquisition hidden inside an inline retained-paint callback', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-body-paint-callback-'));
  write(
    root,
    'packages/docx/src/renderer.ts',
    `function renderBodyElements(elements: any[]) {
      const paintRetainedParagraph = () => paintPlacedParagraphLayout({
        deferFrontDrawing: () => measureParagraph(),
      });
      for (const el of elements) {
        if (el.type === 'paragraph') paintRetainedParagraph();
      }
    }
    `,
  );
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/);
  assert.match(result.output, /measureParagraph/);
});

test('follows the else branch of a negated paragraph dispatch', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-body-paint-negated-'));
  write(
    root,
    'packages/docx/src/renderer.ts',
    `function renderBodyElements(elements: any[]) {
      for (const el of elements) {
        if (el.type !== 'paragraph') renderTable();
        else measureParagraph();
      }
    }
    function renderTable() {}
    `,
  );
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/);
  assert.match(result.output, /measureParagraph/);
});

test('fails closed when the body paragraph dispatch cannot be audited', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-body-paint-opaque-'));
  write(root, 'packages/docx/src/renderer.ts', 'function renderBodyElements() { dispatchBody(); }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BODY_PAINT_LAYOUT_CAPABILITY/);
  assert.match(result.output, /no statically auditable paragraph branch/);
});

test('allows the exact frozen production fragment map and retained paragraph paint boundary', () => {
  const root = initializeRepository();
  write(
    root,
    'packages/docx/src/renderer.ts',
    `import { paintPlacedParagraphLayout as retainedPaint } from './paint/canvas-text.js';
    const paintOperations = { paintPlacedParagraphLayout: retainedPaint };
    const { paintPlacedParagraphLayout } = paintOperations;
    const bodyFlowFragments = Object.freeze(Object.assign(new WeakMap(), {
      sourceIndices: new WeakMap(),
      framePlacement: new WeakMap(),
    }));
    function bodyFragmentFor(element: object) { return bodyFlowFragments.get(element); }
    function paintRetainedParagraph(element: object) {
      bodyFragmentFor(element);
      paintPlacedParagraphLayout();
    }
    function renderBodyElements(elements: any[]) {
      for (const el of elements) {
        if (el.type === 'paragraph') paintRetainedParagraph(el);
        else if (el.type === 'table') renderTable();
      }
    }
    function renderTable() { measureParagraph(); }
    function renderFrameParagraph() { resolveFrameBox(); renderParagraph(); }
    const headerFooterOperations = { paintFrameParagraph: renderFrameParagraph };
    `,
  );

  establishA1Baseline(root);
  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('rejects a CommonJS require edge that bypasses static ESM imports', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-require-'));
  write(root, 'packages/docx/src/renderer.ts', 'export const adapter = true;\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "const helper = require('./helper.js');\nexport { helper };\n");
  write(root, 'packages/docx/src/paint/helper.ts', "const measured = require('../line-layout.js');\nexport { measured };\n");
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
});

test('writes the A1 baseline only when the merge base has none', () => {
  const root = initializeRepository();

  const result = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
  assert.match(readFileSync(join(root, 'scripts/docx-layout-boundary-baseline.json'), 'utf8'), /computePages/);
  assert.equal(runChecker(root, '--base-ref', 'main').status, 0);
});

test('rejects a head baseline that expands the merge-base allowances', () => {
  const root = initializeRepository();
  establishA1Baseline(root);
  const baselinePath = join(root, 'scripts/docx-layout-boundary-baseline.json');
  const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
  baseline.legacySymbolCounts.tableReuseEnabled = 1;
  writeFileSync(baselinePath, `${JSON.stringify(baseline, null, 2)}\n`);

  const result = runChecker(root, '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BASELINE_EXPANSION/);
  assert.match(result.output, /tableReuseEnabled/);
});

test('rejects moving a legacy symbol to another file without increasing its global count', () => {
  const root = initializeRepository();
  establishA1Baseline(root);
  write(root, 'packages/docx/src/renderer.ts', 'export const adapter = true;\n');
  write(root, 'packages/docx/src/legacy-copy.ts', 'export function computePages() { return []; }\n');

  const result = runChecker(root, '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BASELINE_EXPANSION/);
});

test('rejects renaming a migration flag and changing a hash-frozen leaf declaration', () => {
  const renamed = initializeRepository();
  establishA1Baseline(renamed);
  write(renamed, 'packages/docx/src/new-switch.ts', 'export const useLegacyLayout = true;\n');
  const renamedResult = runChecker(renamed, '--base-ref', 'main');
  assert.notEqual(renamedResult.status, 0);
  assert.match(renamedResult.output, /BASELINE_EXPANSION/);

  const changed = initializeRepository();
  establishA1Baseline(changed);
  write(changed, 'packages/docx/src/renderer.ts', 'export function computePages() { return []; }\nexport function computeTableLayout() { return [1]; }\n');
  const changedResult = runChecker(changed, '--base-ref', 'main');
  assert.notEqual(changedResult.status, 0);
  assert.match(changedResult.output, /LEGACY_DECLARATION_CHANGED/);
});

test('allows only exact A2 service and option dependency threading through computePages', () => {
  const root = initializeRepository();
  establishA1Baseline(root);
  write(root, 'packages/docx/src/renderer.ts', 'function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function createLayoutServices() {}\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions) { const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return [measure]; }\nexport function computeTableLayout() { return []; }\n');

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('allows only the exact A5 retained slice size replacement inside computePages', () => {
  const root = initializeComputePagesTableStampRepository();
  write(root, 'packages/docx/src/renderer.ts', computePagesRetainedSliceMigration);
  removeComputePagesTableStampCounts(root);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('allows only destination-first upright finalization inside computePages', () => {
  const root = initializeComputePagesUprightRepository();
  write(root, 'packages/docx/src/renderer.ts', computePagesUprightMigration);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('normalizes both intermediate and destination-first upright forms to one canonical shape', () => {
  for (const source of [computePagesUprightIntermediate, computePagesUprightMigration]) {
    const root = initializeComputePagesUprightRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.equal(result.status, 0, result.output);
  }
});

test('rejects partial, duplicated, or adjacent intermediate upright transactions', () => {
  const variants = [
    computePagesUprightIntermediate.replace(
      '  attachRetainedTablePlacement(tableEl, retained.layout, sourceIndex, {',
      '  attachRetainedTablePlacement(tableEl, retained.layout, sourceIndex + 1, {',
    ),
    computePagesUprightIntermediate.replace(
      '  const sourceIndex = bodySourceIndexFor(tbl);',
      '  const sourceIndex = bodySourceIndexFor(tbl);\n  const duplicateSourceIndex = bodySourceIndexFor(tbl);',
    ),
    computePagesUprightIntermediate.replace(
      '  const sourceIndex = bodySourceIndexFor(tbl);',
      '  sideEffect();\n  const sourceIndex = bodySourceIndexFor(tbl);',
    ),
  ];
  for (const source of variants) {
    const root = initializeComputePagesUprightRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('allows only the exact two-site retained-envelope callee rename', () => {
  const root = initializeComputePagesEnvelopeRepository();
  write(root, 'packages/docx/src/renderer.ts', computePagesEnvelopeMigration);
  const result = runChecker(root, '--base-ref', 'main');
  assert.equal(result.status, 0, result.output);
});

test('rejects one-site, three-site, or argument-changing retained-envelope renames', () => {
  const variants = [
    computePagesEnvelopeBaseline.replace('attachRetainedTableEnvelope', 'attachTableFragment'),
    computePagesEnvelopeMigration.replace(
      '  return [];',
      '  attachTableFragment(third, thirdTable, widths, heights, band, state, thirdPlacement);\n  return [];',
    ),
    computePagesEnvelopeMigration.replace('secondPlacement', 'firstPlacement'),
  ];
  for (const source of variants) {
    const root = initializeComputePagesEnvelopeRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('rejects altered or adjacent upright finalization edits inside computePages', () => {
  const variants = [
    computePagesUprightMigration.replace('pages.length - 1', 'pages.length'),
    computePagesUprightMigration.replace('displayPageNumber: pages.length', 'displayPageNumber: 1'),
    computePagesUprightMigration.replace('withColumnBand(() =>', 'withColumnBand(() => sideEffect(),'),
    computePagesUprightMigration.replace('  return [];', '  sideEffect();\n  return [];'),
  ];
  for (const source of variants) {
    const root = initializeComputePagesUprightRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('allows only child-before-parent finalization for fitting outer floats', () => {
  const root = initializeComputePagesFittingOuterRepository();
  write(root, 'packages/docx/src/renderer.ts', computePagesFittingOuterMigration);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('rejects altered or adjacent fitting outer-float finalization edits', () => {
  const variants = [
    computePagesFittingOuterMigration.replace('first.contentWPt', 'measureState.contentW'),
    computePagesFittingOuterMigration.replace(
      "tbl.overlap !== 'never'",
      'true',
    ),
    computePagesFittingOuterMigration.replace(
      '    const side =',
      '    sideEffect();\n    const side =',
    ),
    computePagesFittingOuterMigration.replace('  return [];', '  sideEffect();\n  return [];'),
  ];
  for (const source of variants) {
    const root = initializeComputePagesFittingOuterRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('composes fitting outer probe normalization over an earlier A5 base', () => {
  const { root, current } = initializeComposedFittingOuterProbeRepository();
  write(root, 'packages/docx/src/renderer.ts', current);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('composes all exact A5 computePages normalizers over production A5 bases', () => {
  const productionRoot = resolve(import.meta.dirname, '..');
  for (const ref of ['ec4e046', 'aa02bbc']) {
    const result = runChecker(productionRoot, '--base-ref', ref);
    assert.equal(result.status, 0, `${ref}: ${result.output}`);
  }
});

test('allows only the exact occurrence-owned table state threading', () => {
  const { root, current } = initializeOccurrenceOwnerRepository();
  write(root, 'packages/docx/src/renderer.ts', current);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('rejects altered, duplicated, or adjacent occurrence-owner edits', () => {
  const variants = [
    [
      '{ sourceIndex: i, record: retainedTableRecord(measureState, i), state: measureState }',
      '{ sourceIndex: i + 1, record: retainedTableRecord(measureState, i), state: measureState }',
    ],
    [
      '            const retainedRecord = retainedTableRecord(measureState, i);\n',
      `            const retainedRecord = retainedTableRecord(measureState, i);
            retainedTableRecord(measureState, i);
`,
    ],
    [
      `            () => pages.length - 1,
            { sourceIndex: i, record: retainedTableRecord(measureState, i), state: measureState },
`,
      `            () => pages.length - 1,
            sideEffect(),
            { sourceIndex: i, record: retainedTableRecord(measureState, i), state: measureState },
`,
    ],
    [
      '        const occurrenceEl = { ...el } as PaginatedElementWithLines;\n',
      '        const occurrenceEl = { ...para } as PaginatedElementWithLines;\n',
    ],
    [
      '        attachBodyParagraphFragment(occurrenceEl, para, measureState, i, {\n',
      `        const duplicateOccurrenceEl = { ...el } as PaginatedElementWithLines;
        attachBodyParagraphFragment(occurrenceEl, para, measureState, i, {
`,
    ],
    [
      '        attachBodyParagraphFragment(occurrenceEl, para, measureState, i, {\n',
      `        sideEffect();
        attachBodyParagraphFragment(occurrenceEl, para, measureState, i, {
`,
    ],
  ];
  for (const [expected, replacement] of variants) {
    const { root, current } = initializeOccurrenceOwnerRepository();
    assert.ok(current.includes(expected), expected);
    write(root, 'packages/docx/src/renderer.ts', current.replace(expected, replacement));
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, `${expected} -> ${replacement}`);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('allows only the live physical-page callback for split floating tables', () => {
  const { root, current } = initializeSplitFloatLivePageRepository();
  write(root, 'packages/docx/src/renderer.ts', current);
  const result = runChecker(root, '--base-ref', 'main');
  assert.equal(result.status, 0, result.output);
});

test('rejects altered or effectful split floating-table page callbacks', () => {
  const variants = [
    ['() => pages.length - 1', '() => pages.length'],
    ['() => pages.length - 1', '() => { sideEffect(); return pages.length - 1; }'],
  ];
  for (const [expected, replacement] of variants) {
    const { root, current } = initializeSplitFloatLivePageRepository();
    assert.ok(current.includes(expected));
    write(root, 'packages/docx/src/renderer.ts', current.replace(expected, replacement));
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, `${expected} -> ${replacement}`);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('allows only converged split parent resolution before the child commit', () => {
  const { root, current } = initializeSplitParentCommitRepository();
  write(root, 'packages/docx/src/renderer.ts', current);
  const result = runChecker(root, '--base-ref', 'main');
  assert.equal(result.status, 0, result.output);
});

test('rejects altered or adjacent converged split parent resolution', () => {
  const variants = [
    [
      "                  { allowOverlap: tbl.overlap !== 'never' },",
      '                  { allowOverlap: true },',
    ],
    ['                  floats: [...externalRegistry.floats],', '                  floats: measureState.floats,'],
    ['            (sliceEl) => pushTagged(sliceEl),', '            (sliceEl) => { sideEffect(); pushTagged(sliceEl); },'],
  ];
  for (const [expected, replacement] of variants) {
    const { root, current } = initializeSplitParentCommitRepository();
    assert.ok(current.includes(expected));
    write(root, 'packages/docx/src/renderer.ts', current.replace(expected, replacement));
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, `${expected} -> ${replacement}`);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('rejects altered fitting outer probe transactions over an earlier A5 base', () => {
  const variants = [
    ['pass < 4', 'pass < 5'],
    ['pageIndex: physicalPageIndex', 'pageIndex: physicalPageIndex + 1'],
    ['box: first.box,', 'box: first.rawBox,'],
    ['first.requiresCanonicalSplit', 'false'],
    ["{ allowOverlap: tbl.overlap !== 'never' },", '{ allowOverlap: true },'],
    ["tbl.overlap !== 'never', true,", "tbl.overlap !== 'never', false,"],
  ];
  for (const [expected, replacement] of variants) {
    const { root, current } = initializeComposedFittingOuterProbeRepository();
    assert.match(current, new RegExp(expected.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
    write(root, 'packages/docx/src/renderer.ts', current.replace(expected, replacement));
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, `${expected} -> ${replacement}`);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('keeps upright paint resolver-free and the table transaction point-only', () => {
  const productionRoot = resolve(import.meta.dirname, '..');
  const renderer = readFileSync(join(productionRoot, 'packages/docx/src/renderer.ts'), 'utf8');
  const retainedPainter = renderer.indexOf('retainedTablePainter:');
  const uprightStart = renderer.indexOf('if (state.verticalPhys) {', retainedPainter);
  const uprightEnd = renderer.indexOf('      const placement = {', uprightStart);
  assert.ok(retainedPainter >= 0 && uprightStart > retainedPainter && uprightEnd > uprightStart);
  const uprightPaint = renderer.slice(uprightStart, uprightEnd);
  assert.doesNotMatch(uprightPaint, /resolveFloatingTablePlacement/);

  const transaction = readFileSync(
    join(productionRoot, 'packages/docx/src/layout/floating-table-transaction.ts'),
    'utf8',
  );
  assert.doesNotMatch(transaction, /\bFloatRect\b|\bscale\b|\bCanvas\w*\b|\bDPR\b|\bdpr\b/);
});

test('allows only the exact A5 retained body prefix inside computeTableLayout', () => {
  const root = initializeComputeTableLayoutRepository();
  write(root, 'packages/docx/src/renderer.ts', computeTableLayoutRetainedMigration);
  removeComputeTableLayoutStampCounts(root);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('rejects partial, duplicated, altered, or adjacent A5 computeTableLayout edits', () => {
  const variants = [
    computeTableLayoutRetainedMigration.replace(
      'state.retainedTableAcquisition',
      'state.retainedTablePainter',
    ),
    computeTableLayoutRetainedMigration.replace(
      'bodySourceIndexFor(table)',
      'bodySourceIndexFor(state)',
    ),
    computeTableLayoutRetainedMigration.replace(
      '  const colWidths = resolveColumnWidths',
      '  const duplicate = computeTablePtLayout(state, table, contentWPt1);\n  const colWidths = resolveColumnWidths',
    ),
    computeTableLayoutRetainedMigration.replace(
      'tableW: colWidths.reduce((sum, width) => sum + width, 0),',
      'tableW: colWidths.reduce((sum, width) => sum + width, 1),',
    ),
    computeTableLayoutRetainedMigration.replace(
      '  const colWidths = resolveColumnWidths',
      '  state.unrelated = true;\n  const colWidths = resolveColumnWidths',
    ),
  ];

  for (const source of variants) {
    const root = initializeComputeTableLayoutRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    removeComputeTableLayoutStampCounts(root);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('rejects partial, duplicated, altered, or adjacent A5 computePages edits', () => {
  const variants = [
    computePagesRetainedSliceMigration.replace(
      'sp, measureState.scale,',
      'sp, measureState.dpr,',
    ),
    computePagesRetainedSliceMigration.replace(
      'heightPx: sliceH',
      'heightPx: tableHeight',
    ),
    computePagesRetainedSliceMigration.replace(
      '  return [tableW, sliceH];',
      '  const { widthPx: duplicateW, heightPx: duplicateH } = retainedTableSliceSize(sp, measureState.scale);\n  return [tableW, sliceH];',
    ),
    computePagesRetainedSliceMigration.replace(
      'return [tableW, sliceH];',
      'return [tableW, sliceH, 1];',
    ),
    computePagesTableStampBaseline.replace(
      'const tableW = (sp.tableColWidthsPt ?? []).reduce((s, w) => s + w, 0) * measureState.scale;',
      'const { widthPx: tableW } = retainedTableSliceSize(sp, measureState.scale);',
    ),
  ];

  for (const source of variants) {
    const root = initializeComputePagesTableStampRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    removeComputePagesTableStampCounts(root);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, source);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('rejects retained-acquisition edits inside computePages', () => {
  const root = initializeComputePagesAcquisitionRepository();
  write(root, 'packages/docx/src/renderer.ts', computePagesAcquisitionMigration);

  const result = runChecker(root, '--base-ref', 'main');

  assert.notEqual(result.status, 0, result.output);
  assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
});

test('rejects partial, duplicated, moved, or altered computePages acquisition seams', () => {
  const variants = [
    computePagesAcquisitionMigration.replace(
      'settings === undefined ? {} : { settings }',
      'settings == null ? {} : { settings }',
    ),
    computePagesAcquisitionMigration.replace(
      'measureState.displayPageNumber = undefined;',
      'measureState.pageIndex = pages.length - 1;\n    measureState.displayPageNumber = undefined;',
    ),
    computePagesAcquisitionMigration.replace(
      'measureState.displayPageNumber = undefined;\n    stampPageMeta();',
      'stampPageMeta();\n    measureState.displayPageNumber = undefined;',
    ),
    computePagesAcquisitionMigration.replace(
      'splitParagraphAcrossPages(measureState, para, i,',
      'splitParagraphAcrossPages(measureState, para, i + 0,',
    ),
    computePagesAcquisitionMigration.replace(
      'attachBodyParagraphFragment(el, para, i, measureState,',
      'attachBodyParagraphFragment(el, para, measureState,',
    ),
    computePagesAcquisitionMigration.replace(
      'attachBodyParagraphFragment(el, para, i, measureState,',
      'attachBodyParagraphFragment(other, para, i, measureState,',
    ),
    computePagesAcquisitionMigration.replace(', sourcePath: [i] });', ' });'),
    computePagesAcquisitionMigration.replace('sourcePath: [i]', 'sourcePath: [i, 0]'),
    computePagesAcquisitionMigration.replace(
      'return pages;',
      'sideEffect();\n  return pages;',
    ),
  ];
  for (const [index, source] of variants.entries()) {
    const root = initializeComputePagesAcquisitionRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, `variant ${index}: ${source}`);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('allows only deleting the exact leading fragmentPaintEnabled table conjunct', () => {
  const root = initializeTablePaintGateRepository();
  write(root, 'packages/docx/src/renderer.ts', tablePaintGateMigration);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('requires targeted baseline removal for a completely deleted migration symbol', () => {
  const root = initializeTablePaintGateRepository(false);
  write(root, 'packages/docx/src/renderer.ts', tablePaintGateMigration);

  const result = runChecker(root, '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BASELINE_MISMATCH/);
});

test('rejects other table paint gate predicate changes and non-exact flag removal', () => {
  const variants = [
    tablePaintGateMigration.replace('placed === undefined', 'placed == null'),
    tablePaintGateMigration.replace("placed.fragment.kind !== 'table' ||\n", ''),
    tablePaintGateMigration.replace(
      'table.tblpPr != null ||',
      'state.extraGate ||\n    table.tblpPr != null ||',
    ),
    tablePaintGateBaseline.replace(
      '!fragmentPaintEnabled ||',
      'fragmentPaintEnabled === false ||',
    ),
    tablePaintGateBaseline.replace(
      '!fragmentPaintEnabled ||',
      '!fragmentPaintEnabled ||\n    !fragmentPaintEnabled ||',
    ),
  ];
  for (const [index, source] of variants.entries()) {
    const root = initializeTablePaintGateRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, `variant ${index}: ${source}`);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('allows only an exact A2 Canvas route argument on renderShapeText font calls', () => {
  const root = initializeShapeRepository();
  write(root, 'packages/docx/src/renderer.ts', 'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown }, classes: Record<string, string>) { return buildFont(s.bold, s.italic, s.size, s.family, classes, s.fontRoute); }\n');

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('allows only exact A2 routes on renderShapeText line-metric probes', () => {
  const root = initializeMetricShapeRepository();
  write(root, 'packages/docx/src/renderer.ts', 'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown; eaFloorRoute?: unknown }, classes: Record<string, string>) { const shapeLineMetrics = (family: string | null, familyRoute?: CanvasFontRoute, familyEaRoute?: CanvasFontRoute) => { const measureRoute = eaIntended > asciiIntended ? familyEaRoute : familyRoute; return buildFont(s.bold, s.italic, s.size, family, classes, measureRoute); }; return shapeLineMetrics(s.family, s.fontRoute, s.eaFloorRoute); }\n');

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('allows only exact A2 numbering snapshot, shape, and retained paint threading in renderShapeText', () => {
  const root = initializeNumberingShapeRepository();
  write(root, 'packages/docx/src/renderer.ts', exactNumberingShapeSource);

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('rejects non-exact renderShapeText numbering shape and paint migrations', () => {
  const cases = [
    exactNumberingShapeSource.replace('block.fontSizePt);', 'block.fontSizePt + 1);'),
    exactNumberingShapeSource.replace('markerTextLayout?.shape.advancePt', 'markerTextLayout?.shape.advancePt + 1'),
    exactNumberingShapeSource.replace('paintNumberingMarkerText(ctx, markerTextLayout', 'paintNumberingMarkerText(ctx, alteredLayout'),
    exactNumberingShapeSource.replace('const markerShapeInput = numberingMarkerShapeInput(block.numbering, block.fontSizePt); ', ''),
    exactNumberingShapeSource.replace('return markerW;', 'sideEffect(); return markerW;'),
  ];
  for (const source of cases) {
    const root = initializeNumberingShapeRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED/);
  }
});

test('rejects non-exact renderShapeText line-metric route threading', () => {
  for (const source of [
    'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown; eaFloorRoute?: unknown }, classes: Record<string, string>) { const shapeLineMetrics = (family: string | null, familyRoute?: CanvasFontRoute, familyEaRoute?: CanvasFontRoute) => { const measureRoute = s.size > 0 ? familyEaRoute : undefined; return buildFont(s.bold, s.italic, s.size, family, classes, measureRoute); }; return shapeLineMetrics(s.family, s.fontRoute, s.eaFloorRoute); }\n',
    'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown; eaFloorRoute?: unknown }, classes: Record<string, string>) { const shapeLineMetrics = (family: string | null, familyRoute?: unknown, familyEaRoute?: CanvasFontRoute) => { const measureRoute = s.size > 0 ? familyEaRoute : familyRoute; return buildFont(s.bold, s.italic, s.size, family, classes, measureRoute); }; return shapeLineMetrics(s.family, s.fontRoute, s.eaFloorRoute); }\n',
    'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown; eaFloorRoute?: unknown }, classes: Record<string, string>) { const shapeLineMetrics = (family: string | null, familyRoute?: CanvasFontRoute, familyEaRoute?: CanvasFontRoute) => { const measureRoute = s.size > 0 ? familyEaRoute : familyRoute; buildFont(s.bold, s.italic, s.size, family, classes, measureRoute); return "changed"; }; return shapeLineMetrics(s.family, s.fontRoute, s.eaFloorRoute); }\n',
    'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown; eaFloorRoute?: unknown }, classes: Record<string, string>) { const shapeLineMetrics = (family: string | null, familyRoute?: CanvasFontRoute, familyEaRoute?: CanvasFontRoute) => { const measureRoute = sideEffect() ? familyEaRoute : familyRoute; return buildFont(s.bold, s.italic, s.size, family, classes, measureRoute); }; return shapeLineMetrics(s.family, s.fontRoute, s.eaFloorRoute); }\n',
    'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown; eaFloorRoute?: unknown }, classes: Record<string, string>) { const shapeLineMetrics = (family: string | null, familyRoute?: CanvasFontRoute, familyEaRoute?: CanvasFontRoute) => { let measureRoute = eaIntended > asciiIntended ? familyEaRoute : familyRoute; return buildFont(s.bold, s.italic, s.size, family, classes, measureRoute); }; return shapeLineMetrics(s.family, s.fontRoute, s.eaFloorRoute); }\n',
    'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown; eaFloorRoute?: unknown }, classes: Record<string, string>) { const shapeLineMetrics = (familyRoute?: CanvasFontRoute, familyEaRoute?: CanvasFontRoute, family: string | null) => { const measureRoute = eaIntended > asciiIntended ? familyEaRoute : familyRoute; return buildFont(s.bold, s.italic, s.size, family, classes, measureRoute); }; return shapeLineMetrics(s.fontRoute, s.eaFloorRoute, s.family); }\n',
  ]) {
    const root = initializeMetricShapeRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED/);
  }
});

test('rejects other renderShapeText changes beside exact Canvas route threading', () => {
  for (const source of [
    'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown }, classes: Record<string, string>) { buildFont(s.bold, s.italic, s.size, s.family, classes, s.fontRoute); return "changed"; }\n',
    'function buildFont(bold: boolean, italic: boolean, size: number, family: string | null, classes: Record<string, string>, route?: unknown) { return String([bold, italic, size, family, classes, route]); }\nexport function renderShapeText(s: { bold: boolean; italic: boolean; size: number; family: string | null; fontRoute?: unknown }, classes: Record<string, string>) { return buildFont(s.bold, s.italic, s.size, s.family, classes, undefined); }\n',
  ]) {
    const root = initializeShapeRepository();
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED/);
  }
});

test('rejects unrelated computePages control-flow, calls, and parameters during A2 threading', () => {
  const cases = [
    'function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions) { const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return []; }\nexport function computeTableLayout() { return []; }\n',
    'function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions) { buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return [measure]; }\nexport function computeTableLayout() { return []; }\n',
    'function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions, unrelated?: boolean) { const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return [measure]; }\nexport function computeTableLayout() { return []; }\n',
  ];
  for (const source of cases) {
    const root = initializeRepository();
    establishA1Baseline(root);
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('rejects any non-exact A2 parameter syntax and pass-through expression', () => {
  const variants = [
    'layoutServices: LayoutServices, layoutOptions?: LayoutOptions',
    'layoutServices?: LayoutServices = undefined, layoutOptions?: LayoutOptions',
    '...layoutServices: LayoutServices[], layoutOptions?: LayoutOptions',
    'readonly layoutServices?: LayoutServices, layoutOptions?: LayoutOptions',
    'layoutOptions?: LayoutOptions, layoutServices?: LayoutServices',
  ];
  for (const tail of variants) {
    const root = initializeRepository();
    establishA1Baseline(root);
    write(root, 'packages/docx/src/renderer.ts', `function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function createLayoutServices() {}\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, ${tail}) { const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return [measure]; }\nexport function computeTableLayout() { return []; }\n`);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, tail);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }

  const expressions = [
    'layoutServices as LayoutServices, layoutOptions',
    'layoutServices, layoutOptions ?? undefined',
    '{ ...layoutServices }, layoutOptions',
  ];
  for (const args of expressions) {
    const root = initializeRepository();
    establishA1Baseline(root);
    write(root, 'packages/docx/src/renderer.ts', `function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function createLayoutServices() {}\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions) { const measure = buildMeasureState(ctx, resolvedLocalFonts, ${args}); return [measure]; }\nexport function computeTableLayout() { return []; }\n`);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0, args);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('final mode enforces the renderer adapter export and import allowlists', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-final-adapter-'));
  write(root, 'packages/docx/src/layout/document.ts', 'export function layoutDocument() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paintLayoutPage() {}\n');
  write(root, 'packages/docx/src/renderer.ts', "import { layoutDocument } from './layout/document.js';\nimport { paintLayoutPage } from './paint/canvas-page.js';\nexport function paginateDocument() { return layoutDocument(); }\nexport function renderDocumentToCanvas() { return paintLayoutPage(); }\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/renderer.ts', "import { hidden } from './hidden-layout.js';\nexport function paginateDocument() { return hidden(); }\nexport function renderDocumentToCanvas() {}\nexport function accidentalAlgorithm() {}\n");
  write(root, 'packages/docx/src/hidden-layout.ts', 'export function hidden() {}\n');
  const result = runChecker(root, '--final');
  assert.notEqual(result.status, 0);
  assert.match(result.output, /FINAL_ADAPTER_/);
});

test('final mode rejects layout logic hidden inside an allowed renderer adapter', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-inline-adapter-'));
  write(root, 'packages/docx/src/renderer.ts', `
export function paginateDocument(items: unknown[]) {
  const pages = [];
  for (const item of items) pages.push([item]);
  return pages;
}
export function renderDocumentToCanvas() {}
`);

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FINAL_ADAPTER_BODY/);
});

test('final mode rejects renamed fallback and style-cascade capabilities', () => {
  const diagnostic = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-diagnostic-fallback-'));
  write(diagnostic, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(diagnostic, 'packages/docx/src/layout/page.ts', "export const diagnosticFallback = { code: 'UNSUPPORTED_FEATURE' };\n");
  assert.equal(runChecker(diagnostic, '--final').status, 0);

  const fallback = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-old-engine-'));
  write(fallback, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(fallback, 'packages/docx/src/layout/page.ts', 'export const useOldEngine = true;\n');
  const fallbackResult = runChecker(fallback, '--final');
  assert.notEqual(fallbackResult.status, 0);
  assert.match(fallbackResult.output, /FINAL_LEGACY_BOUNDARY/);

  const style = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-fold-style-'));
  write(style, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(style, 'packages/docx/src/layout/page.ts', 'export function foldRunFormatting(base: object, direct: object) { return { ...base, ...direct }; }\n');
  const styleResult = runChecker(style, '--final');
  assert.notEqual(styleResult.status, 0);
  assert.match(styleResult.output, /LAYOUT_STYLE_CAPABILITY/);
});

test('final mode rejects star exports from renderer', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-star-export-'));
  write(root, 'packages/docx/src/layout/page.ts', 'export function accidentalAlgorithm() {}\n');
  write(root, 'packages/docx/src/renderer.ts', "export * from './layout/page.js';\nexport function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n");

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FINAL_ADAPTER_EXPORT/);
});

test('excludes test-support modules from production inventory but rejects production imports', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-test-support-'));
  write(root, 'packages/docx/src/retained.test-support.ts', 'export function acquireForTest() { return 1; }\n');
  write(root, 'packages/docx/src/retained.test.ts', "import { acquireForTest } from './retained.test-support.js';\nvoid acquireForTest();\n");
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');

  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/renderer.ts', "import { acquireForTest } from './retained.test-support.js';\nexport function paginateDocument() { return acquireForTest(); }\nexport function renderDocumentToCanvas() {}\n");
  const result = runChecker(root, '--final');
  assert.notEqual(result.status, 0);
  assert.match(result.output, /PRODUCTION_TEST_SUPPORT_IMPORT/);
});

test('rejects rewriting a transitional baseline after A1', () => {
  const root = initializeRepository();
  establishA1Baseline(root);

  const result = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /TRANSITIONAL_BASELINE_EXISTS/);
});

test('rejects final mode while a transitional baseline remains', () => {
  const root = initializeRepository();
  establishA1Baseline(root);

  const result = runChecker(root, '--final', '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FINAL_BASELINE_PRESENT/);
});
