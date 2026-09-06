import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
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

function runChecker(root, ...args) {
  return command(root, process.execPath, [checker, '--root', root, ...args]);
}

function expectDiagnostic(root, diagnostic, message, ...args) {
  const result = runChecker(root, ...args);
  assert.notEqual(result.status, 0, message ?? result.output);
  assert.match(result.output, new RegExp(`^${diagnostic}:`, 'm'), message ?? result.output);
  return result;
}

const canonicalBodyPaginator =
  "import { assertAndDeepFreezeDocumentLayout } from './invariants.js';\n"
  + 'export type PaginationSteps<T> = Generator<number, T, void>;\n'
  + 'export function drainPagination(steps) {\n'
  + '  let step = steps.next();\n'
  + '  while (!step.done) step = steps.next();\n'
  + '  return step.value;\n'
  + '}\n'
  + 'export function* paginateBodySteps(input, services, options) {\n'
  + '  const layout = { pages: [], diagnostics: [], input, services, options };\n'
  + '  yield 0;\n'
  + '  return assertAndDeepFreezeDocumentLayout(layout);\n'
  + '}\n'
  + 'export function paginateBody(input, services, options) {\n'
  + '  return drainPagination(paginateBodySteps(input, services, options));\n'
  + '}\n';

const canonicalRenderer = `
import { layoutDocument } from './layout/document.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import { renderSelectedDocumentPage } from './paint/canvas-document.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
function createConcreteBodyLayoutKernel(doc, ctx, localMetrics) {
  return { doc, ctx, localMetrics };
}
function createLayoutServices(doc, ctx, localMetrics) {
  const services = {};
  attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));
  return services;
}
function normalizeRenderOptions(services, input, pageIndex) {
  const selection = selectDocumentLayoutPage(services, input, pageIndex);
  layoutDocument();
  paintLayoutPage(selection.page);
  return { selection, paintOptions: {} };
}
export function renderLayoutSourceToCanvas(source, canvas, pageIndex, opts) {
  const normalized = normalizeRenderOptions(source, canvas, pageIndex, opts);
  return renderSelectedDocumentPage(
    normalized.selection.layout,
    normalized.selection.page,
    canvas,
    normalized.paintOptions,
  );
}
export function renderDocumentToCanvas(doc, canvas, pageIndex, opts) {
  return renderLayoutSourceToCanvas(layoutSourceStore(doc), canvas, pageIndex, opts);
}
`;

const canonicalLayoutRuntime = `
function createConcreteBodyLayoutKernel(source, context, fontMetrics) {
  return { source, context, fontMetrics };
}
export function createLayoutServices(input, context, fontMetrics) {
  const source = layoutSourceStore(input);
  const services = {};
  attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(
    source,
    context,
    fontMetrics,
  ));
  return services;
}
`;

function initializeCanonicalFixture(prefix = 'docx-layout-boundary-canonical-') {
  const root = mkdtempSync(join(tmpdir(), prefix));
  write(root, 'packages/docx/src/layout/validation-policy.ts',
    'let enabled = false;\n'
      + 'export function documentLayoutValidationEnabled() { return enabled; }\n'
      + 'export function setDocumentLayoutValidation(next) { enabled = next; }\n');
  write(root, 'packages/docx/src/layout/plain-data.ts',
    "import { documentLayoutValidationEnabled } from './validation-policy.js';\n"
      + 'export function snapshotPlainData(value) {\n'
      + '  if (documentLayoutValidationEnabled()) assertPlainData(value);\n'
      + '  return Object.freeze(value);\n'
      + '}\n');
  write(root, 'packages/docx/src/layout/retained-geometry-translation.ts',
    "import type { PointPt } from './types.js';\nexport const translate = (point: PointPt) => point;\n");
  write(root, 'packages/docx/src/layout/occurrence-projection.ts',
    "import { snapshotPlainData } from './plain-data.js';\nexport const project = snapshotPlainData;\n");
  write(root, 'packages/docx/src/layout/types.ts',
    'export interface PointPt { xPt: number; yPt: number }\n');
  write(root, 'packages/docx/src/layout/acquisition-context.ts',
    'export interface AnchorFloatRegistrationState {}\n'
      + 'export interface AnchorGeometryContext {}\n'
      + 'export interface BodyAcquisitionState {\n'
      + '  retainedTableAcquisition: unknown;\n'
      + '  retainedTablesBySourceIndex: Map<number, unknown>;\n'
      + '}\n'
      + 'export type BodyMeasurementContext = Readonly<BodyAcquisitionState>;\n'
      + 'export interface FloatRegistrationState {}\n'
      + 'export interface PhysicalAnchorFrame {}\n'
      + 'export interface RetainedTableRecord {}\n');
  write(root, 'packages/docx/src/layout/production-body-layout.ts', `
function acquireRetainedTable() { return { layout: { rows: [] } }; }
function acquireRetainedFrameGroup() { return {}; }
function resolveParagraphLayoutContext() { return {}; }
function computeTablePtLayout(state, table, contentWPt, sourceIndex: number) {
  return acquireRetainedTable(table, [], contentWPt, state, [sourceIndex], state.retainedTableAcquisition);
}
function resolveColumnWidths(state, paragraph) {
  const baseContext = resolveParagraphLayoutContext(
    state.layoutSettings,
    state.sectionLayout,
    state.storyContext,
    paragraph,
  );
  return baseContext;
}
function resolveFrameBox(para, group: BodyFrameGroup, state, anchorLineHPt, onAcquired?) {
  return acquireRetainedFrameGroup(group, { state, anchorLineHPt, onAcquired });
}
function production(state, table, para, group) {
  computeTablePtLayout(state, table, 100, 0);
  resolveColumnWidths(state, para);
  resolveFrameBox(para, group, state, 12, undefined);
}
`);
  write(root, 'packages/docx/src/layout/acquisition-input-projections.ts',
    'export interface BodyAcquisitionInputProjections {\n'
      + '  numberingMarkerShapeInput(): unknown;\n'
      + '  paragraphMarkShapeInput(): unknown;\n'
      + '  tableFormatInput(): unknown;\n'
      + '  tableColumnLayoutInput(): unknown;\n'
      + '  tableParticipatesInOrdinaryFlow(): unknown;\n'
      + '  paragraphAcquisitionInput(): unknown;\n'
      + '}\n');
  write(root, 'packages/docx/src/layout/acquisition-state.ts',
    'export const BODY_STORY_CONTEXT = {};\n'
      + 'export function bodyAnchorReferenceFrames() {}\n'
      + 'export function resolveBodyParagraphLayoutContext() {}\n'
      + 'export function resolveStateParagraphLayoutContext() {}\n'
      + 'export function withTableCellStory() {}\n'
      + 'export function retainedTableRecord() {}\n');
  write(root, 'packages/docx/src/layout/anchor-classification.ts',
    'export function isPageLevelAnchorY() {}\n'
      + 'export function isPageLevelWrapFloat() {}\n');
  write(root, 'packages/docx/src/layout/measurement-environment.ts',
    'export function canonicalParagraphTextScaleEligible() {}\n'
      + 'export function docDefaultFontSizePt() {}\n'
      + 'export function paragraphMeasurementEnvironment() {}\n'
      + 'export function segmentEnvironmentOf() {}\n'
      + 'export function snapParaLineToGrid() {}\n'
      + 'export function gridForParagraphContext() {}\n');
  write(root, 'packages/docx/src/layout/measurement-capabilities.ts',
    'export interface MeasurementTextContext {\n'
      + '  font: string;\n'
      + '  letterSpacing: string;\n'
      + '  fontKerning: unknown;\n'
      + '  measureText(text: string): unknown;\n'
      + '}\n'
      + 'export interface VerticalGlyphMeasurementService {\n'
      + '  fingerprint: string;\n'
      + '  measureRunInkExtra(text: string): number;\n'
      + '  planRun(input: unknown): unknown;\n'
      + '}\n');
  write(root, 'packages/docx/src/layout/section-orientation.ts',
    'export function isVerticalSection() {}\n'
      + 'export function isVerticalTextDirection() {}\n'
      + 'export function isAllRotatedVerticalTextDirection() {}\n'
      + 'export function verticalLayoutSection() {}\n'
      + 'export function verticalLayoutDoc() {}\n'
      + 'export function physicalLayoutSection() {}\n');
  write(root, 'packages/docx/src/layout/affine.ts',
    "import type { PointPt } from './types.js';\n"
      + 'export const composeAffine = (point: PointPt) => point;\n');
  write(root, 'packages/docx/src/layout/coordinate-space.ts',
    "import type { PointPt } from './types.js';\nexport const coordinate = (point: PointPt) => point;\n");
  write(root, 'packages/docx/src/layout/page-graph.ts', 'export const PAGE_LAYER_IDS = [];\n');
  write(root, 'packages/docx/src/layout/page-factory.ts',
    "import { coordinate } from './coordinate-space.js';\n"
      + "import { PAGE_LAYER_IDS } from './page-graph.js';\n"
      + "import type { BodyOccurrenceDestination } from './occurrence-projection.js';\n"
      + 'export const pageFactory = [coordinate, PAGE_LAYER_IDS] satisfies unknown;\n'
      + 'export type Destination = BodyOccurrenceDestination;\n');
  write(root, 'packages/docx/src/layout/invariants.ts',
    'export function assertAndDeepFreezeDocumentLayout(value) { return Object.freeze(value); }\n');
  write(root, 'packages/docx/src/layout/body-paginator.ts', canonicalBodyPaginator);
  write(root, 'packages/docx/src/layout/document.ts',
    "import { paginateBody } from './body-paginator.js';\n"
      + 'export function layoutDocument(input, services, options) {\n'
      + '  return paginateBody(input, services, options);\n'
      + '}\n');
  write(root, 'packages/docx/src/layout/document-layout-variants.ts',
    'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
      + '  const store = layoutVariantStoreOf(services);\n'
      + "  if (!store) throw new Error('Document layout variant store is not attached');\n"
      + '  return store.selectPage(layoutOptionsForRender(input), pageIndex);\n'
      + '}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts',
    'export function paintLayoutPage(page) { return page; }\n');
  write(root, 'packages/docx/src/paint/canvas-document.ts',
    'export function renderSelectedDocumentPage() {}\n');
  write(root, 'packages/docx/src/layout-source-model-adapter.ts',
    'export function layoutSourceStore(value) { return value; }\n');
  write(root, 'packages/docx/src/render-worker-layout.ts',
    "import { layoutDocument } from './renderer.js';\n"
      + "import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';\n"
      + 'export interface RetainedRenderWorkerDocumentLayout {\n'
      + '  layoutServices; layoutVariants; defaultCurrentDateMs;\n'
      + '}\n'
      + 'export function retainRenderWorkerDocumentLayout(source, layoutServices, defaultCurrentDateMs) {\n'
      + '  const variants = attachDocumentLayoutVariants({ source, services: layoutServices,\n'
      + '    defaultCurrentDateMs,\n'
      + '    buildLayout: (options) => layoutDocument(source, layoutServices, options) });\n'
      + '  return { layoutServices, layoutVariants: variants.store, defaultCurrentDateMs };\n'
      + '}\n');
  write(root, 'packages/docx/src/render-worker-metadata.ts',
    'export function projectRenderWorkerLayoutMeta(layout, source, review, options = {}) {\n'
      + '  const hasComments = review.comments.length > 0;\n'
      + '  const hasRevisions = review.revisions.length > 0;\n'
      + '  const reviewIndex = hasComments || hasRevisions\n'
      + '    ? reviewProjectionIndexForDocument(layout)\n'
      + '    : undefined;\n'
      + '  const reviewProjection = options.provisional && reviewIndex\n'
      + '    ? { completedSourceKeys: reviewIndex.completedSourceKeys }\n'
      + '    : undefined;\n'
      + '  return {\n'
      + '    pageCount: layout.pages.length,\n'
      + '    pageSizes: layout.pages.map((page) => page.geometry),\n'
      + '    bookmarkPages: [...buildBookmarkPageMap(layout)],\n'
      + '    commentAnchorRanges: hasComments\n'
      + '      ? collectLayoutSourceCommentRangesIfPresent(review.comments, source, reviewIndex!.renderedRunIndex, reviewProjection)\n'
      + '      : [],\n'
      + '    revisionAnchorRanges: hasRevisions\n'
      + '      ? collectLayoutSourceRevisionRangesIfPresent(review.revisions, source, reviewIndex!.renderedRunIndex, reviewProjection)\n'
      + '      : [],\n'
      + '  };\n'
      + '}\n'
      + 'export function renderWorkerLayoutMeta(doc, review, currentDateMs, showTrackedChanges) {\n'
      + '  const source = layoutSourceStoreOf(doc.layoutServices);\n'
      + '  const layout = doc.layoutVariants.layoutFor(normalizeLayoutOptions(\n'
      + '    currentDateMs, doc.defaultCurrentDateMs, showTrackedChanges));\n'
      + '  return projectRenderWorkerLayoutMeta(layout, source, review);\n'
      + '}\n');
  write(root, 'packages/docx/src/render-worker.ts',
    "import { renderLayoutSourceToCanvas } from './renderer.js';\n"
      + "import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';\n"
      + 'export function initializeWorker(source, layoutServices, req) {\n'
      + '  const doc = retainRenderWorkerDocumentLayout(source, layoutServices, req.defaultCurrentDateMs);\n'
      + '  const layoutOptions = normalizeLayoutOptions(req.currentDateMs,\n'
      + '    req.defaultCurrentDateMs, req.showTrackedChanges);\n'
      + '  const layout = doc.layoutVariants.layoutFor(layoutOptions);\n'
      + '  const reviewIndexInput = { comments: [], revisions: [] };\n'
      + '  const meta = {\n'
      + '    ...projectRenderWorkerLayoutMeta(layout, source, reviewIndexInput) };\n'
      + '  return { doc, meta };\n'
      + '}\n'
      + 'export function selectWorkerLayoutView(doc, reviewIndexInput, req) {\n'
      + '  return renderWorkerLayoutMeta(\n'
      + '    doc, reviewIndexInput, req.currentDateMs, req.showTrackedChanges);\n'
      + '}\n'
      + 'export function renderWorkerPage(doc, canvas, req, options) {\n'
      + '  const source = layoutSourceStoreOf(doc.layoutServices);\n'
      + "  if (!source) throw new Error('missing source');\n"
      + '  return renderLayoutSourceToCanvas(source, canvas, req.pageIndex, { ...options,\n'
      + '    layoutServices: doc.layoutServices,\n'
      + '    defaultCurrentDateMs: doc.defaultCurrentDateMs });\n'
      + '}\n'
      + 'export function collectWorkerRuns() { return []; }\n');
  write(root, 'packages/docx/src/renderer.ts', canonicalRenderer);
  write(root, 'packages/docx/src/layout-runtime.ts', canonicalLayoutRuntime);
  write(root, 'packages/docx/tests/visual/conformance-fixture.html', `
<script type="module">
  import { layoutDocument } from '/src/document-layout.ts';
  import { createLayoutServices } from '/src/layout-runtime.ts';
</script>
`);
  return root;
}

const canonicalBodyLayoutAdapter =
  "import { bodyLayoutAcquisitionInput } from './parser-model.js';\n"
  + "import type { DocxDocumentModel } from './types.js';\n"
  + "import { projectBodyLayoutInput, type BodyLayoutInput } from './layout/body-layout-input.js';\n"
  + 'export function createBodyLayoutInput(document: DocxDocumentModel): BodyLayoutInput {\n'
  + '  return projectBodyLayoutInput(bodyLayoutAcquisitionInput(document));\n'
  + '}\n';

function initializeBodyLayoutAdapterFixture(prefix = 'docx-layout-boundary-input-adapter-') {
  const root = initializeCanonicalFixture(prefix);
  write(root, 'packages/docx/src/parser-model.ts',
    'export function bodyLayoutAcquisitionInput(document) { return document; }\n');
  write(root, 'packages/docx/src/types.ts', 'export interface DocxDocumentModel {}\n');
  write(root, 'packages/docx/src/layout/body-layout-input.ts',
    'export interface BodyLayoutInput {}\n'
      + 'export function projectBodyLayoutInput(value) { return value; }\n');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter);
  return root;
}

const canonicalParagraphAnchorFrameAdapter =
  "import type { AnchorReferenceFramesInput } from './layout/anchor-frame.js';\n"
  + 'export interface ParagraphAnchorReferenceFrameSnapshot { readonly scale: number }\n'
  + 'export function paragraphAnchorReferenceFrames(snapshot: ParagraphAnchorReferenceFrameSnapshot): '
  + "Readonly<Pick<AnchorReferenceFramesInput, 'page'>> {\n"
  + '  return { page: { xPt: 0, yPt: 0, widthPt: snapshot.scale, heightPt: 1 } };\n'
  + '}\n';

function installParagraphAnchorFrameAdapter(root) {
  write(root, 'packages/docx/src/layout/anchor-frame.ts',
    'export interface AnchorReferenceFramesInput { page: unknown }\n');
  write(root, 'packages/docx/src/paragraph-anchor-frame-adapter.ts',
    canonicalParagraphAnchorFrameAdapter);
  const rendererPath = join(root, 'packages/docx/src/renderer.ts');
  write(root, 'packages/docx/src/renderer.ts',
    "import { paragraphAnchorReferenceFrames } from './paragraph-anchor-frame-adapter.js';\n"
      + readFileSync(rendererPath, 'utf8')
        .replace('return { doc, ctx, localMetrics };',
          'paragraphAnchorReferenceFrames({ scale: 1 });\n  return { doc, ctx, localMetrics };'));
}

test('accepts the final canonical producer, retained model, selected variant, and worker route', () => {
  const root = initializeCanonicalFixture();
  const result = runChecker(root, '--final');
  assert.equal(result.status, 0, result.output);
});

test('requires one private concrete body-kernel owner with exact loud attachment', () => {
  for (const [name, source] of [
    ['missing implementation', canonicalLayoutRuntime.replace(
      /function createConcreteBodyLayoutKernel[\s\S]*?\n}\n/,
      '',
    )],
    ['missing attachment', canonicalLayoutRuntime.replace(
      /attachBodyLayoutKernel\(services, createConcreteBodyLayoutKernel\([\s\S]*?\n  \)\);/,
      'createConcreteBodyLayoutKernel(source, context, fontMetrics);',
    )],
    ['wrong owner arguments', canonicalLayoutRuntime.replace(
      '    context,\n    fontMetrics,',
      '    fontMetrics,',
    )],
    ['wrong source', canonicalLayoutRuntime.replace(
      '    source,',
      '    input,',
    )],
    ['duplicate owner', canonicalLayoutRuntime.replace(
      'return services;',
      'attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(\n'
        + '    source, context, fontMetrics,\n'
        + '  ));\n  return services;',
    )],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-owner-${name}-`);
    write(root, 'packages/docx/src/layout-runtime.ts', source);
    expectDiagnostic(root, 'BODY_KERNEL_SERVICE_OWNER', name, '--final');
  }
});

test('rejects every deleted body producer, raw page route, and test adapter symbol', () => {
  const deleted = [
    'computePages',
    'paginateDocument',
    'paginateWithHeaderFooterReserve',
    'physicalPageSizeForPage',
    'PaginatedBodyElement',
    'prebuiltPages',
    'retainedLayout',
    'bodyFragmentFor',
    'sectionBreakSpacer',
    'collapsedSpacer',
    'leadsCollapsedRun',
    'hiddenCollapsed',
  ];
  for (const name of deleted) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-deleted-${name}-`);
    write(root, 'packages/docx/src/layout/deleted-path.ts', `export const ${name} = true;\n`);
    expectDiagnostic(root, 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER', name, '--final');
  }
});

test('rejects migration flags and silent alternate layout fallbacks', () => {
  for (const [name, diagnostic] of [
    ['useOldLayoutEngine', 'FINAL_LEGACY_BOUNDARY'],
    ['preferAlternateLayoutPath', 'FINAL_LEGACY_BOUNDARY'],
    ['bodyLayoutFallback', 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-migration-${name}-`);
    write(root, 'packages/docx/src/layout/path.ts', `export const ${name} = true;\n`);
    expectDiagnostic(root, diagnostic, name, '--final');
  }
});

test('rejects renderer-owned acquisition state from the final architecture', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-render-state-');
  write(
    root,
    'packages/docx/src/layout/obsolete-state.ts',
    'export interface RenderState { readonly dryRun: boolean }\n',
  );
  expectDiagnostic(root, 'FINAL_LEGACY_BOUNDARY', 'RenderState', '--final');
});

test('layout acquisition contexts reject paint capabilities and renderer back-edges', () => {
  const scaleRoot = initializeCanonicalFixture('docx-layout-boundary-acquisition-scale-');
  write(
    scaleRoot,
    'packages/docx/src/layout/acquisition-context.ts',
    'export interface AnchorFloatRegistrationState {}\n'
      + 'export interface AnchorGeometryContext { scale: number }\n'
      + 'export interface BodyAcquisitionState { retainedTableAcquisition: unknown; retainedTablesBySourceIndex: Map<number, unknown> }\n'
      + 'export type BodyMeasurementContext = Readonly<BodyAcquisitionState>;\n'
      + 'export interface FloatRegistrationState {}\n'
      + 'export interface PhysicalAnchorFrame {}\n'
      + 'export interface RetainedTableRecord {}\n',
  );
  expectDiagnostic(
    scaleRoot,
    'ACQUISITION_COORDINATE_SCALE',
    'direct scale member',
    '--final',
  );

  const scalePickRoot = initializeCanonicalFixture('docx-layout-boundary-acquisition-scale-pick-');
  write(
    scalePickRoot,
    'packages/docx/src/layout/acquisition-context.ts',
    'export interface AnchorFloatRegistrationState {}\n'
      + 'export interface AnchorGeometryContext {}\n'
      + 'export interface BodyAcquisitionState { retainedTableAcquisition: unknown; retainedTablesBySourceIndex: Map<number, unknown> }\n'
      + "export type BodyMeasurementContext = Readonly<Pick<BodyAcquisitionState, 'scale'>>;\n"
      + 'export interface FloatRegistrationState {}\n'
      + 'export interface PhysicalAnchorFrame {}\n'
      + 'export interface RetainedTableRecord {}\n',
  );
  expectDiagnostic(
    scalePickRoot,
    'ACQUISITION_COORDINATE_SCALE',
    'scale pick',
    '--final',
  );

  const paintRoot = initializeCanonicalFixture('docx-layout-boundary-acquisition-paint-');
  write(
    paintRoot,
    'packages/docx/src/layout/acquisition-context.ts',
    'export interface AnchorFloatRegistrationState {}\n'
      + 'export interface AnchorGeometryContext {}\n'
      + 'export interface BodyAcquisitionState { images: Map<string, unknown>; retainedTableAcquisition: unknown; retainedTablesBySourceIndex: Map<number, unknown> }\n'
      + 'export type BodyMeasurementContext = Readonly<BodyAcquisitionState>;\n'
      + 'export interface FloatRegistrationState {}\n'
      + 'export interface PhysicalAnchorFrame {}\n'
      + 'export interface RetainedTableRecord {}\n',
  );
  expectDiagnostic(
    paintRoot,
    'ACQUISITION_PAINT_CAPABILITY',
    'images',
    '--final',
  );

  const measurementRoot = initializeCanonicalFixture('docx-layout-boundary-measurement-paint-');
  write(
    measurementRoot,
    'packages/docx/src/layout/measurement-capabilities.ts',
    'export interface MeasurementTextContext { canvas: HTMLCanvasElement }\n'
      + 'export interface VerticalGlyphMeasurementService { measureRunInkExtra(text: string): number; planRun(input: unknown): unknown }\n',
  );
  expectDiagnostic(
    measurementRoot,
    'ACQUISITION_PAINT_CAPABILITY',
    'canvas',
    '--final',
  );

  for (const [label, source, detail] of [
    [
      'unlisted member',
      'export interface MeasurementTextContext { font: string; letterSpacing: string; fontKerning: unknown; measureText(text: string): unknown; fillRect(): void }\n'
        + 'export interface VerticalGlyphMeasurementService { fingerprint: string; measureRunInkExtra(text: string): number; planRun(input: unknown): unknown }\n',
      'fillRect',
    ],
    [
      'heritage clause',
      'export interface MeasurementTextContext extends CanvasRenderingContext2D { font: string; letterSpacing: string; fontKerning: unknown; measureText(text: string): unknown }\n'
        + 'export interface VerticalGlyphMeasurementService { fingerprint: string; measureRunInkExtra(text: string): number; planRun(input: unknown): unknown }\n',
      'heritage',
    ],
    [
      'extra declaration',
      'export interface MeasurementTextContext { font: string; letterSpacing: string; fontKerning: unknown; measureText(text: string): unknown }\n'
        + 'export interface VerticalGlyphMeasurementService { fingerprint: string; measureRunInkExtra(text: string): number; planRun(input: unknown): unknown }\n'
        + 'export interface PaintEscape { fillRect(): void }\n',
      'PaintEscape',
    ],
  ]) {
    const exactRoot = initializeCanonicalFixture(`docx-layout-boundary-measurement-${label.replace(' ', '-')}-`);
    write(exactRoot, 'packages/docx/src/layout/measurement-capabilities.ts', source);
    expectDiagnostic(
      exactRoot,
      'ACQUISITION_CONTEXT_SURFACE',
      detail,
      '--final',
    );
  }

  const edgeRoot = initializeCanonicalFixture('docx-layout-boundary-acquisition-edge-');
  write(
    edgeRoot,
    'packages/docx/src/layout/acquisition-context.ts',
    "import type { Hidden } from '../renderer.js';\n"
      + 'export interface AnchorFloatRegistrationState {}\n'
      + 'export interface AnchorGeometryContext {}\n'
      + 'export interface BodyAcquisitionState { hidden?: Hidden; retainedTableAcquisition: unknown; retainedTablesBySourceIndex: Map<number, unknown> }\n'
      + 'export type BodyMeasurementContext = Readonly<BodyAcquisitionState>;\n'
      + 'export interface FloatRegistrationState {}\n'
      + 'export interface PhysicalAnchorFrame {}\n'
      + 'export interface RetainedTableRecord {}\n',
  );
  expectDiagnostic(
    edgeRoot,
    'ACQUISITION_RENDERER_DEPENDENCY',
    'renderer.ts',
    '--final',
  );
});

test('layout acquisition ownership cannot be bypassed by removing a required module', () => {
  for (const module of [
    'acquisition-context.ts',
    'acquisition-input-projections.ts',
    'acquisition-state.ts',
    'anchor-classification.ts',
    'measurement-capabilities.ts',
    'measurement-environment.ts',
    'section-orientation.ts',
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-${module}-missing-`);
    rmSync(join(root, 'packages/docx/src/layout', module));
    expectDiagnostic(
      root,
      'ACQUISITION_CONTEXT_SURFACE',
      `missing ${module} must fail the ownership ratchet`,
      '--final',
    );
  }
});

test('table acquisition capabilities are required final-state inputs', () => {
  for (const property of ['retainedTableAcquisition', 'retainedTablesBySourceIndex']) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-required-${property}-`);
    const path = join(root, 'packages/docx/src/layout/acquisition-context.ts');
    write(root, 'packages/docx/src/layout/acquisition-context.ts', readFileSync(path, 'utf8').replace(
      `${property}:`,
      `${property}?:`,
    ));
    expectDiagnostic(root, 'RETAINED_TABLE_AUTHORITY', property, '--final');
  }
});

test('production table and frame acquisition cannot regain local fallback measurement', () => {
  for (const [name, mutate] of [
    ['optional source index', (source) => source.replace('sourceIndex: number', 'sourceIndex?: number')],
    ['omitted source index', (source) => source.replace(
      'computeTablePtLayout(state, table, 100, 0)',
      'computeTablePtLayout(state, table, 100)',
    )],
    ['table row fallback', (source) => source.replace(
      'return acquireRetainedTable(table, [], contentWPt, state, [sourceIndex], state.retainedTableAcquisition);',
      'return resolveTableRowContentHeights(table);',
    )],
    ['aliased table row fallback', (source) => (
      "import { resolveTableRowContentHeights as remeasureRows } from '../table-geometry.js';\n"
        + source.replace(
          'return acquireRetainedTable(table, [], contentWPt, state, [sourceIndex], state.retainedTableAcquisition);',
          'remeasureRows(table);\n  return acquireRetainedTable(table, [], contentWPt, state, [sourceIndex], state.retainedTableAcquisition);',
        )
    )],
    ['optional frame group', (source) => source.replace('group: BodyFrameGroup', 'group?: BodyFrameGroup')],
    ['local frame measurement', (source) => source.replace(
      'return acquireRetainedFrameGroup(group, { state, anchorLineHPt, onAcquired });',
      'return layoutLines(para);',
    )],
    ['aliased local frame measurement', (source) => (
      "import { layoutLines as localMeasure } from '../line-layout.js';\n"
        + source.replace(
          'return acquireRetainedFrameGroup(group, { state, anchorLineHPt, onAcquired });',
          'localMeasure(para);\n  return acquireRetainedFrameGroup(group, { state, anchorLineHPt, onAcquired });',
        )
    )],
    ['reduced column context fallback', (source) => source.replace(
      'const baseContext = resolveParagraphLayoutContext(',
      'const baseContext = state.layoutSettings ? resolveParagraphLayoutContext(',
    ).replace(
      '    paragraph,\n  );\n  return baseContext;',
      '    paragraph,\n  ) : {};\n  return baseContext;',
    )],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-acquisition-fallback-${name.replaceAll(' ', '-')}-`);
    const path = join(root, 'packages/docx/src/layout/production-body-layout.ts');
    write(root, 'packages/docx/src/layout/production-body-layout.ts', mutate(readFileSync(path, 'utf8')));
    expectDiagnostic(root, 'PRODUCTION_ACQUISITION_AUTHORITY', name, '--final');
  }
});

test('renderer body acquisition cannot bypass its injected parser projections', () => {
  for (const [label, importSource] of [
    [
      'direct paragraph projection',
      "import { paragraphAcquisitionInput } from './parser-model.js';\n",
    ],
    [
      'aliased table projection',
      "import { tableFormatInput as formatTable } from './parser-model.js';\n",
    ],
    [
      'namespace projection access',
      "import * as parserModel from './parser-model.js';\n",
    ],
    [
      'CommonJS projection access',
      "const parserModel = require('./parser-model.js');\n",
    ],
    [
      'dynamic projection access',
      "const parserModelPromise = import('./parser-model.js');\n",
    ],
  ]) {
    const root = initializeCanonicalFixture(
      `docx-layout-boundary-renderer-projection-${label.replaceAll(' ', '-')}-`,
    );
    write(
      root,
      'packages/docx/src/renderer.ts',
      importSource + canonicalRenderer,
    );
    expectDiagnostic(
      root,
      'RENDERER_ACQUISITION_PROJECTION_BYPASS',
      label,
      '--final',
    );
  }

  for (const [label, rendererSource] of [
    [
      'global projection record use',
      "import { bodyAcquisitionInputProjections } from './parser-model.js';\n"
        + canonicalRenderer.replace(
          'return { doc, ctx, localMetrics };',
          'bodyAcquisitionInputProjections.paragraphAcquisitionInput();\n'
            + '  return { doc, ctx, localMetrics };',
        ),
    ],
    [
      'aliased projection record',
      "import { bodyAcquisitionInputProjections as projections } from './parser-model.js';\n"
        + canonicalRenderer,
    ],
    [
      'nested owner impersonation',
      "import { bodyAcquisitionInputProjections } from './parser-model.js';\n"
        + canonicalRenderer.replace(
          'return { doc, ctx, localMetrics };',
          'function buildMeasureState() {\n'
            + '    bodyAcquisitionInputProjections.tableFormatInput();\n'
            + '  }\n'
            + '  buildMeasureState();\n'
            + '  return { doc, ctx, localMetrics };',
        ),
    ],
    [
      'computed non-injected projection access',
      canonicalRenderer.replace(
        'return { doc, ctx, localMetrics };',
        "doc['parserFacts']['paragraph' + 'AcquisitionInput']();\n"
          + '  return { doc, ctx, localMetrics };',
      ),
    ],
  ]) {
    const root = initializeCanonicalFixture(
      `docx-layout-boundary-renderer-projection-${label.replaceAll(' ', '-')}-`,
    );
    write(root, 'packages/docx/src/renderer.ts', rendererSource);
    expectDiagnostic(
      root,
      'RENDERER_ACQUISITION_PROJECTION_BYPASS',
      label,
      '--final',
    );
  }
});

test('coordinate-space and page-factory accept only their explicit dependencies', () => {
  assert.equal(runChecker(initializeCanonicalFixture(), '--final').status, 0);
  for (const [name, path, source] of [
    ['renderer', 'coordinate-space.ts', "import { value } from '../renderer.js';\nexport const coordinate = value;\n"],
    ['parser', 'coordinate-space.ts', "import { value } from '../parser-model.js';\nexport const coordinate = value;\n"],
    ['paint', 'coordinate-space.ts', "import { value } from '../paint/canvas-page.js';\nexport const coordinate = value;\n"],
    ['dynamic', 'coordinate-space.ts', "export const coordinate = import('./types.js');\n"],
    ['dynamic nonliteral', 'coordinate-space.ts', "const moduleName = './types.js';\nexport const coordinate = import(moduleName);\n"],
    ['decorated', 'coordinate-space.ts', "import type { PointPt } from './types.js?raw';\nexport const coordinate = (point: PointPt) => point;\n"],
    ['worker', 'coordinate-space.ts', "import { value } from '../render-worker.js';\nexport const coordinate = value;\n"],
    ['shaping', 'coordinate-space.ts', "import { value } from './text.js';\nexport const coordinate = value;\n"],
    ['package', 'coordinate-space.ts', "import value from 'canvas';\nexport const coordinate = value;\n"],
    ['projection runtime', 'page-factory.ts', "import { project } from './occurrence-projection.js';\nexport const pageFactory = project;\n"],
    ['page factory decorated type', 'page-factory.ts', "import type { PointPt } from './types.js?raw';\nexport type Point = PointPt;\n"],
    ['page factory package runtime', 'page-factory.ts', "import value from 'canvas';\nexport const pageFactory = value;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-coordinate-${name}-`);
    write(root, `packages/docx/src/layout/${path}`, source);
    expectDiagnostic(root, 'COORDINATE_SPACE_RUNTIME_DEPENDENCY', name, '--final');
  }
});

test('layout affine algebra is a pinned type-only dependency seam', () => {
  assert.equal(runChecker(initializeCanonicalFixture(), '--final').status, 0);
  for (const [name, source] of [
    ['runtime layout', "import { coordinate } from './coordinate-space.js';\nexport const composeAffine = coordinate;\n"],
    ['paint', "import { cssTransformFor } from '../paint/affine.js';\nexport const composeAffine = cssTransformFor;\n"],
    ['package', "import value from 'canvas';\nexport const composeAffine = value;\n"],
    ['dynamic', "export const composeAffine = import('./types.js');\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-affine-${name}-`);
    write(root, 'packages/docx/src/layout/affine.ts', source);
    expectDiagnostic(root, 'AFFINE_RUNTIME_DEPENDENCY', name, '--final');
  }

  const missing = initializeCanonicalFixture('docx-layout-boundary-affine-missing-');
  rmSync(join(missing, 'packages/docx/src/layout/affine.ts'));
  expectDiagnostic(missing, 'AFFINE_RUNTIME_DEPENDENCY', 'missing affine seam', '--final');
});

test('page-factory may retain section decoration geometry through its dedicated layout helper', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-page-decoration-');
  write(root, 'packages/docx/src/layout/column-separators.ts',
    'export const columnSeparatorSegments = () => [];\n');
  write(root, 'packages/docx/src/layout/page-factory.ts',
    "import { coordinate } from './coordinate-space.js';\n"
      + "import { columnSeparatorSegments } from './column-separators.js';\n"
      + "import { PAGE_LAYER_IDS } from './page-graph.js';\n"
      + 'export const pageFactory = [coordinate, columnSeparatorSegments, PAGE_LAYER_IDS] satisfies unknown;\n');

  assert.equal(runChecker(root, '--final').status, 0);
});

test('occurrence projection accepts only translation and plain-data runtime dependencies', () => {
  for (const [name, path, source] of [
    ['external layout', 'occurrence-projection.ts', "import { value } from '../paragraph-measure.js';\nexport const project = value;\n"],
    ['reverse edge', 'retained-geometry-translation.ts', "import { project } from './occurrence-projection.js';\nexport const translate = project;\n"],
    ['parser edge', 'plain-data.ts', "import { value } from '../parser-model.js';\nexport const snapshotPlainData = value;\n"],
    ['dynamic edge', 'occurrence-projection.ts', "export const project = import('./plain-data.js');\n"],
    ['dynamic nonliteral edge', 'occurrence-projection.ts', "const moduleName = './plain-data.js';\nexport const project = import(moduleName);\n"],
    ['decorated edge', 'occurrence-projection.ts', "import { snapshotPlainData } from './plain-data.js?raw';\nexport const project = snapshotPlainData;\n"],
    ['package edge', 'occurrence-projection.ts', "import value from 'measurement-service';\nexport const project = value;\n"],
    ['policy layout edge', 'validation-policy.ts', "import { value } from '../paragraph-measure.js';\nexport const documentLayoutValidationEnabled = value;\n"],
    ['policy projection edge', 'validation-policy.ts', "import { snapshotPlainData } from './plain-data.js';\nexport const documentLayoutValidationEnabled = snapshotPlainData;\n"],
    ['policy consumer edge', 'occurrence-projection.ts', "import { documentLayoutValidationEnabled } from './validation-policy.js';\nexport const project = documentLayoutValidationEnabled;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-occurrence-${name}-`);
    write(root, `packages/docx/src/layout/${path}`, source);
    expectDiagnostic(root, 'OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', name, '--final');
  }
});

test('coordinate-space and page-factory seams are mandatory', () => {
  for (const name of ['coordinate-space.ts', 'page-factory.ts']) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-coordinate-missing-${name}-`);
    rmSync(join(root, 'packages/docx/src/layout', name));
    expectDiagnostic(root, 'COORDINATE_SPACE_RUNTIME_DEPENDENCY', name, '--final');
  }
});

test('every occurrence-projection seam is mandatory', () => {
  for (const name of [
    'occurrence-projection.ts',
    'retained-geometry-translation.ts',
    'plain-data.ts',
    'validation-policy.ts',
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-occurrence-missing-${name}-`);
    rmSync(join(root, 'packages/docx/src/layout', name));
    expectDiagnostic(root, 'OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', name, '--final');
  }
});

test('keeps layout display-independent and paint measurement-free', () => {
  for (const [name, path, source, diagnostic] of [
    ['paint measurement', 'paint/canvas-page.ts', 'export function paint(ctx) { return ctx.measureText("x"); }\n', 'PAINT_CAPABILITY'],
    ['paint style cascade', 'paint/canvas-page.ts', 'export function resolveRunStyle() {}\n', 'PAINT_CAPABILITY'],
    ['layout display scale', 'layout/document.ts', 'export const displayScale = 2;\n', 'LAYOUT_DISPLAY_CAPABILITY'],
    ['layout canvas context', 'layout/document.ts', 'export let CanvasRenderingContext2D;\n', 'LAYOUT_DISPLAY_CAPABILITY'],
    ['layout style cascade', 'layout/document.ts', 'export function mergeRunProperties() {}\n', 'LAYOUT_STYLE_CAPABILITY'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-capability-${name}-`);
    write(root, `packages/docx/src/${path}`, source);
    expectDiagnostic(root, diagnostic, name, '--final');
  }
});

test('paint imports only retained layout contracts and reviewed atomic core painters', () => {
  const allowed = initializeCanonicalFixture('docx-layout-boundary-paint-allowed-');
  write(allowed, 'packages/docx/src/paint/helper.ts',
    "import type { PointPt } from '../layout/types.js';\n"
      + "import { crispOffset } from '@silurus/ooxml-core';\n"
      + 'export const helper = [crispOffset] satisfies unknown as PointPt[];\n');
  assert.equal(runChecker(allowed, '--final').status, 0);

  for (const [name, source] of [
    ['layout implementation', "import { paginateBody } from '../layout/body-paginator.js';\nexport const helper = paginateBody;\n"],
    ['measurement module', "import { measure } from '../paragraph-measure.js';\nexport const helper = measure;\n"],
    ['forbidden core API', "import { measureText } from '@silurus/ooxml-core';\nexport const helper = measureText;\n"],
    ['aliased core API', "import { crispOffset as crisp } from '@silurus/ooxml-core';\nexport const helper = crisp;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-paint-${name}-`);
    write(root, 'packages/docx/src/paint/helper.ts', source);
    expectDiagnostic(root, 'FORBIDDEN_PAINT_EDGE', name, '--final');
  }
});

test('paint reaches layout affine algebra only through the exact reviewed facade', () => {
  const names = [
    'composeAffine',
    'inverseMapAffinePoint',
    'inverseMapAffineVector',
    'mapAffinePoint',
    'quarterTurnAffine',
    'scaleAffine',
    'translationAffine',
  ];
  const allowed = initializeCanonicalFixture('docx-layout-boundary-paint-affine-allowed-');
  write(allowed, 'packages/docx/src/paint/affine.ts',
    `export { ${names.join(', ')} } from '../layout/affine.js';\n`);
  assert.equal(runChecker(allowed, '--final').status, 0);

  const incomplete = initializeCanonicalFixture('docx-layout-boundary-paint-affine-incomplete-');
  write(incomplete, 'packages/docx/src/paint/affine.ts',
    `export { ${names.slice(0, -1).join(', ')} } from '../layout/affine.js';\n`);
  expectDiagnostic(incomplete, 'FORBIDDEN_PAINT_EDGE', 'incomplete affine facade', '--final');

  const aliased = initializeCanonicalFixture('docx-layout-boundary-paint-affine-aliased-');
  write(aliased, 'packages/docx/src/paint/affine.ts',
    "export { composeAffine,\n"
      + '  inverseMapAffinePoint as inverseMapAffineVector,\n'
      + '  inverseMapAffineVector as inverseMapAffinePoint,\n'
      + "  mapAffinePoint, quarterTurnAffine, scaleAffine, translationAffine } from '../layout/affine.js';\n");
  expectDiagnostic(aliased, 'FORBIDDEN_PAINT_EDGE', 'aliased affine facade', '--final');

  const direct = initializeCanonicalFixture('docx-layout-boundary-paint-affine-direct-');
  write(direct, 'packages/docx/src/paint/helper.ts',
    "import { composeAffine } from '../layout/affine.js';\nvoid composeAffine;\n");
  expectDiagnostic(direct, 'FORBIDDEN_PAINT_EDGE', 'direct affine edge', '--final');
});

test('every reviewed atomic core paint binding remains explicitly allowed', () => {
  const valueBindings = [
    'autoContrastColor',
    'canvasFontString',
    'captureDecodedBitmapCacheEpoch',
    'crispOffset',
    'deferBitmapCloseWhileLeased',
    'drawImageCropped',
    'doubleRailGeometry',
    'fillDoubleBorder',
    'isTiffDecodeError',
    'paintDrawingMLShape',
    'resolveFill',
    'renderChart',
  ];
  for (const name of valueBindings) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-core-${name}-`);
    write(root, 'packages/docx/src/paint/core-binding.ts',
      `import { ${name} } from '@silurus/ooxml-core';\nvoid ${name};\n`);
    const result = runChecker(root, '--final');
    assert.equal(result.status, 0, `${name}: ${result.output}`);
  }
  const typeRoot = initializeCanonicalFixture('docx-layout-boundary-core-hyperlink-');
  write(typeRoot, 'packages/docx/src/paint/core-binding.ts',
    "import type { HyperlinkTarget } from '@silurus/ooxml-core';\nexport type Target = HyperlinkTarget;\n");
  const result = runChecker(typeRoot, '--final');
  assert.equal(result.status, 0, result.output);
});

test('paint boundary follows transitive edges and rejects sibling or public package imports', () => {
  const cases = [
    ['transitive local', (root) => {
      write(root, 'packages/docx/src/paint/helper.ts',
        "import { measure } from '../paragraph-measure.js';\nexport const helper = measure;\n");
      write(root, 'packages/docx/src/paragraph-measure.ts', 'export const measure = 1;\n');
      return "import { helper } from './helper.js';\nexport const paint = helper;\n";
    }],
    ['sibling package', () => "import { render } from '@silurus/ooxml-pptx';\nexport const paint = render;\n"],
    ['public package', () => "import { render } from '@silurus/ooxml';\nexport const paint = render;\n"],
    ['bare side effect', () => "import '@silurus/ooxml-core';\nexport const paint = true;\n"],
  ];
  for (const [name, source] of cases) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-paint-graph-${name}-`);
    write(root, 'packages/docx/src/paint/graph.ts', source(root));
    expectDiagnostic(root, 'FORBIDDEN_PAINT_EDGE', name, '--final');
  }
});

test('paint cannot import the retained page graph implementation', () => {
  const rejected = initializeCanonicalFixture('docx-layout-boundary-page-graph-');
  write(rejected, 'packages/docx/src/paint/graph.ts',
    "import { PAGE_LAYER_IDS } from '../layout/page-graph.js';\nexport const paint = PAGE_LAYER_IDS;\n");
  expectDiagnostic(rejected, 'FORBIDDEN_PAINT_EDGE', 'page graph edge', '--final');
});

test('layout-only treatment stays outside paint while layout resource implementations stay forbidden', () => {
  const isolated = initializeCanonicalFixture('docx-layout-boundary-layout-only-treatment-');
  write(isolated, 'packages/docx/src/layout/border-treatment.ts',
    "import { docxBorderDashArray } from '@silurus/ooxml-core';\n"
      + "export const treatment = docxBorderDashArray('dotDash', 1);\n");
  const isolatedResult = runChecker(isolated, '--final');
  assert.equal(isolatedResult.status, 0, isolatedResult.output);

  for (const [name, source] of [
    ['value', "import { createRegistry } from '../layout/paint-resources.js';\nexport const paint = createRegistry();\n"],
    ['type', "import type { Registry } from '../layout/paint-resources.js';\nexport type PaintRegistry = Registry;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-layout-resource-${name}-`);
    write(root, 'packages/docx/src/layout/paint-resources.ts',
      'export interface Registry {}\nexport function createRegistry() { return {}; }\n');
    write(root, 'packages/docx/src/paint/resource.ts', source);
    expectDiagnostic(root, 'FORBIDDEN_PAINT_EDGE', name, '--final');
  }
});

test('computed paint measurement and TSX layout display inputs are audited', () => {
  const paint = initializeCanonicalFixture('docx-layout-boundary-computed-measurement-');
  write(paint, 'packages/docx/src/paint/computed.ts',
    "export const width = (ctx) => ctx['measureText']('x').width;\n");
  expectDiagnostic(paint, 'PAINT_CAPABILITY', 'computed measureText', '--final');

  const layout = initializeCanonicalFixture('docx-layout-boundary-layout-tsx-');
  write(layout, 'packages/docx/src/layout/display.tsx',
    'export const Layout = ({ dpr }: { dpr: number }) => dpr;\n');
  expectDiagnostic(layout, 'LAYOUT_DISPLAY_CAPABILITY', 'TSX dpr', '--final');
});

test('paint CommonJS edges cannot bypass the dependency graph', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-paint-require-');
  write(root, 'packages/docx/src/paint/require-edge.ts',
    "const helper = require('./require-helper.js');\nexport { helper };\n");
  write(root, 'packages/docx/src/paint/require-helper.ts',
    "const measured = require('../paragraph-measure.js');\nexport { measured };\n");
  write(root, 'packages/docx/src/paragraph-measure.ts', 'export const measured = true;\n');
  expectDiagnostic(root, 'FORBIDDEN_PAINT_EDGE', 'CommonJS paint edge', '--final');
});

test('allows only the exact parser normalization gateway and erased contracts', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-parser-gateway-');
  write(root, 'packages/docx/src/parser-model.ts',
    'export function normalizeInternalDocumentModel(document) { return { document, mathOccurrences: [] }; }\n');
  write(root, 'packages/docx/src/layout/resources.ts',
    "import { normalizeInternalDocumentModel } from '../parser-model.js';\n"
      + 'export function documentMathOccurrences(doc) {\n'
      + '  return [...normalizeInternalDocumentModel(doc).mathOccurrences];\n'
      + '}\n'
      + 'export function productionDocumentInput(doc) {\n'
      + '  return normalizeInternalDocumentModel(doc);\n'
      + '}\n');
  const allowed = runChecker(root, '--final');
  assert.equal(allowed.status, 0, allowed.output);

  write(root, 'packages/docx/src/layout/resources.ts',
    "import { normalizeInternalDocumentModel as normalize } from '../parser-model.js';\n"
      + 'export function documentMathOccurrences(doc) { return [...normalize(doc).mathOccurrences]; }\n');
  expectDiagnostic(root, 'LAYOUT_PARSER_MODEL_DEPENDENCY', 'aliased gateway', '--final');
});

test('layout rejects direct and transitive parser-model runtime dependencies', () => {
  for (const [name, setup] of [
    ['direct', (root) => write(root, 'packages/docx/src/layout/parser-edge.ts',
      "import { parserFact } from '../parser-model.js';\nexport const fact = parserFact;\n")],
    ['transitive', (root) => {
      write(root, 'packages/docx/src/layout/parser-edge.ts',
        "import { parserFact } from '../parser-bridge.js';\nexport const fact = parserFact;\n");
      write(root, 'packages/docx/src/parser-bridge.ts',
        "import { parserFact } from './parser-model.js';\nexport { parserFact };\n");
    }],
    ['literal dynamic', (root) => write(root, 'packages/docx/src/layout/parser-edge.ts',
      "export const load = () => import('../parser-model.js');\n")],
    ['CommonJS', (root) => write(root, 'packages/docx/src/layout/parser-edge.ts',
      "export const load = () => require('../parser-model.js');\n")],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-parser-${name}-`);
    write(root, 'packages/docx/src/parser-model.ts', 'export const parserFact = true;\n');
    setup(root);
    expectDiagnostic(root, 'LAYOUT_PARSER_MODEL_DEPENDENCY', name, '--final');
  }
});

test('parser normalization gateway rejects aliases, re-exports, leaks, and extra bindings', () => {
  const exact =
    "import { normalizeInternalDocumentModel } from '../parser-model.js';\n"
    + 'export function documentMathOccurrences(doc) {\n'
    + '  return [...normalizeInternalDocumentModel(doc).mathOccurrences];\n'
    + '}\n';
  const sources = [
    "import { normalizeInternalDocumentModel, parserFact } from '../parser-model.js';\n"
      + 'export const leaked = [normalizeInternalDocumentModel, parserFact];\n',
    "import * as parserModel from '../parser-model.js';\nexport const leaked = parserModel;\n",
    "export { normalizeInternalDocumentModel } from '../parser-model.js';\n",
    "export * from '../parser-model.js';\n",
    `${exact}export { normalizeInternalDocumentModel };\n`,
    `${exact}export const leaked = normalizeInternalDocumentModel;\n`,
    "import { normalizeInternalDocumentModel } from '../parser-model.js';\n"
      + 'const normalize = normalizeInternalDocumentModel;\n'
      + 'export function documentMathOccurrences(doc) { return [...normalize(doc).mathOccurrences]; }\n',
  ];
  for (const [index, source] of sources.entries()) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-parser-gateway-${index}-`);
    write(root, 'packages/docx/src/parser-model.ts',
      'export const parserFact = true;\n'
        + 'export function normalizeInternalDocumentModel(document) { return { document, mathOccurrences: [] }; }\n');
    write(root, 'packages/docx/src/layout/resources.ts', source);
    expectDiagnostic(root, 'LAYOUT_PARSER_MODEL_DEPENDENCY', `gateway case ${index}`, '--final');
  }
});

test('layout may share parser-independent values and erased parser-adjacent contracts', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-parser-erased-');
  write(root, 'packages/docx/src/parser-model.ts', 'export const parserFact = true;\n');
  write(root, 'packages/docx/src/parser-contract.ts',
    "import { parserFact } from './parser-model.js';\nexport interface ParserContract { value: typeof parserFact }\n");
  write(root, 'packages/docx/src/shared-primitive.ts', 'export const nextTabStop = 36;\n');
  write(root, 'packages/docx/src/layout/parser-contract-user.ts',
    "import type { ParserContract } from '../parser-contract.js';\n"
      + "import { nextTabStop } from '../shared-primitive.js';\n"
      + 'export type Contract = ParserContract;\nexport const stop = nextTabStop;\n');
  const result = runChecker(root, '--final');
  assert.equal(result.status, 0, result.output);
});

test('body-layout input adapter has one exact parser acquisition projection', () => {
  const root = initializeBodyLayoutAdapterFixture();
  assert.equal(runChecker(root, '--final').status, 0);
});

test('body-layout input adapter rejects non-exact import sets', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-import-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    "import type { DocxDocumentModel } from './types.js';\n",
    '',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_IMPORT', 'incomplete import set', '--final');
});

test('body-layout input adapter rejects incomplete bindings within a reviewed module', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-incomplete-binding-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    'import { projectBodyLayoutInput, type BodyLayoutInput }',
    'import { projectBodyLayoutInput }',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_IMPORT', 'incomplete reviewed binding', '--final');
});

test('body-layout input adapter rejects duplicate, default, and unreviewed imports', () => {
  for (const [name, source] of [
    ['duplicate', canonicalBodyLayoutAdapter.replace(
      "import type { DocxDocumentModel } from './types.js';\n",
      "import type { DocxDocumentModel } from './types.js';\n"
        + "import type { DocxDocumentModel as DuplicateModel } from './types.js';\n",
    )],
    ['default', canonicalBodyLayoutAdapter.replace(
      "import type { DocxDocumentModel } from './types.js';",
      "import DocxDocumentModel from './types.js';",
    )],
    ['unreviewed', `${canonicalBodyLayoutAdapter}import { extra } from './extra.js';\nvoid extra;\n`],
  ]) {
    const root = initializeBodyLayoutAdapterFixture(`docx-layout-boundary-input-import-${name}-`);
    write(root, 'packages/docx/src/extra.ts', 'export const extra = true;\n');
    write(root, 'packages/docx/src/body-layout-input.ts', source);
    expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_IMPORT', name, '--final');
  }
});

test('body-layout input adapter rejects a non-literal import specifier', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-dynamic-specifier-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    "import { bodyLayoutAcquisitionInput } from './parser-model.js';",
    'import { bodyLayoutAcquisitionInput } from parserModule;',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_IMPORT', 'non-literal import', '--final');
});

test('body-layout input adapter rejects aliased or incorrectly typed bindings', () => {
  for (const [name, source] of [
    ['alias', canonicalBodyLayoutAdapter.replace(
      '{ bodyLayoutAcquisitionInput }',
      '{ bodyLayoutAcquisitionInput as acquireBodyLayoutInput }',
    )],
    ['value-as-type', canonicalBodyLayoutAdapter.replace(
      'import type { DocxDocumentModel }',
      'import { DocxDocumentModel }',
    )],
  ]) {
    const root = initializeBodyLayoutAdapterFixture(`docx-layout-boundary-input-binding-${name}-`);
    write(root, 'packages/docx/src/body-layout-input.ts', source);
    expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_BINDING', name, '--final');
  }
});

test('body-layout input adapter rejects additional declarations', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-declaration-');
  write(root, 'packages/docx/src/body-layout-input.ts',
    `${canonicalBodyLayoutAdapter}\nfunction hiddenProjection() {}\n`);
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_DECLARATION', 'additional declaration', '--final');
});

test('body-layout input adapter requires exactly one declaration', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-missing-declaration-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    /export function createBodyLayoutInput[\s\S]*?\n}\n/,
    '',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_DECLARATION', 'missing declaration', '--final');
});

test('body-layout input adapter rejects alternate export forms', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-export-');
  write(root, 'packages/docx/src/body-layout-input.ts',
    `${canonicalBodyLayoutAdapter}\nexport { createBodyLayoutInput as default };\n`);
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_EXPORT', 'export declaration', '--final');
});

test('body-layout input adapter rejects a default declaration export', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-default-export-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    'export function createBodyLayoutInput',
    'export default function createBodyLayoutInput',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_EXPORT', 'default declaration', '--final');
});

test('body-layout input adapter rejects any body other than the exact projection', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-body-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    'return projectBodyLayoutInput(bodyLayoutAcquisitionInput(document));',
    'return projectBodyLayoutInput(document);',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_BODY', 'non-acquisition body', '--final');
});

test('paragraph anchor frame adapter is rejected from the final architecture', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-final-paragraph-anchor-adapter-');
  installParagraphAnchorFrameAdapter(root);

  expectDiagnostic(root, 'FINAL_PARAGRAPH_ANCHOR_ADAPTER', 'final adapter', '--final');
});

test('retained body paint cannot reacquire layout', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-body-reacquisition-');
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() {}\n');
  write(root, 'packages/docx/src/renderer.ts', `
import { layoutLines } from './line-layout.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
function createConcreteBodyLayoutKernel(doc, ctx, localMetrics) { return { doc, ctx, localMetrics }; }
function createLayoutServices(doc, ctx, localMetrics) {
  const services = {};
  attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));
  return services;
}
function renderBodyElements(elements) {
  for (const element of elements) {
    if (element.type === 'paragraph') layoutLines(element);
  }
}
function renderDocumentToCanvas(services, input, pageIndex) {
  return selectDocumentLayoutPage(services, input, pageIndex);
}
`);
  expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', 'body reacquisition', '--final');
});

test('retained body paint rejects aliased, transitive, callback, and unresolved layout calls', () => {
  const renderers = [
    ['alias', `const paintRetained = measureParagraph;
function renderBodyElements(elements) {
  for (const element of elements) if (element.type === 'paragraph') paintRetained(element);
}`],
    ['transitive', `function paintRetained() { acquireParagraphLayout(); }
function renderBodyElements(elements) {
  for (const element of elements) if (element.type === 'paragraph') paintRetained(element);
}`],
    ['callback', `function renderBodyElements(elements) {
  for (const element of elements) if (element.type === 'paragraph') paintPlacedParagraphLayout({
    deferFrontDrawing: () => measureParagraph(element),
  });
}`],
    ['unresolved', `function renderBodyElements(elements) {
  for (const element of elements) if (element.type === 'paragraph') opaquePaint(element);
}`],
  ];
  for (const [name, renderer] of renderers) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-body-${name}-`);
    write(root, 'packages/docx/src/renderer.ts', renderer);
    expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', name, '--final');
  }
});

test('every retained-body layout acquisition capability is rejected directly', () => {
  const calls = [
    'acquireParagraphLayout()',
    'acquireRetainedFrameGroup()',
    'buildSegments()',
    'contextualSpacingAdjust()',
    'estimateParagraphHeight()',
    'layoutLines()',
    'measureParagraph()',
    'measureText()',
    'paragraphGapAdjustment()',
    'paragraphLayoutFromMeasurement()',
    'parasShareBorderBox()',
    'renderFrameParagraph()',
    'renderParagraph()',
    'resolveParagraphBorderEdges()',
    'resolveFrameBox()',
  ];
  for (const call of calls) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-body-capability-${call}-`);
    write(root, 'packages/docx/src/renderer.ts',
      `function renderBodyElements(elements) {
        for (const element of elements) if (element.type === 'paragraph') ${call};
      }
      `);
    expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', call, '--final');
  }
});

test('late mutation of retained body sidecars is rejected before paint dispatch', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-body-sidecar-mutation-');
  write(root, 'packages/docx/src/renderer.ts',
    `const bodyFlowFragments = Object.freeze(Object.assign(new WeakMap(), {
      sourceIndices: new WeakMap(),
      framePlacement: new WeakMap(),
    }));
    bodyFlowFragments.sourceIndices.retainedTableMeasureBySource = new WeakMap();
    function renderBodyElements(elements) {
      for (const element of elements) {
        if (element.type === 'paragraph') state.onTextRun(element);
      }
    }
    `);
  expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', 'late sidecar mutation', '--final');
});

test('only the float placement authority may import the displacement kernel', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-float-authority-');
  write(root, 'packages/docx/src/layout/axis-aligned-overlap.ts',
    'export interface AxisAlignedRect { left; right; top; bottom }\n'
      + 'export function axisAlignedRectsOverlap() { return false; }\n'
      + 'export function resolveAxisAlignedOverlap(value) { return value; }\n');
  write(root, 'packages/docx/src/layout/floats.ts',
    "import { axisAlignedRectsOverlap, resolveAxisAlignedOverlap, type AxisAlignedRect } from './axis-aligned-overlap.js';\n"
      + 'export function place(value: AxisAlignedRect) {\n'
      + '  axisAlignedRectsOverlap();\n'
      + '  return resolveAxisAlignedOverlap(value);\n'
      + '}\n');
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/layout/paragraph.ts',
    "import { resolveAxisAlignedOverlap as displace } from './axis-aligned-overlap.js';\n"
      + 'export const placeParagraph = displace;\n');
  expectDiagnostic(
    root,
    'FLOAT_PLACEMENT_AUTHORITY',
    'an aliased direct kernel import must not bypass layout/floats.ts',
    '--final',
  );
});

test('float placement authority rejects every module-edge and property-access bypass', () => {
  const bypasses = new Map([
    ['namespace',
      "import * as overlap from './axis-aligned-overlap.js';\nexport const place = overlap.resolveAxisAlignedOverlap;\n"],
    ['default',
      "import overlap from './axis-aligned-overlap.js';\nexport const place = overlap;\n"],
    ['export-star',
      "export * from './axis-aligned-overlap.js';\n"],
    ['named-re-export',
      "export { resolveAxisAlignedOverlap as place } from './axis-aligned-overlap.js';\n"],
    ['dynamic-import',
      "export const place = () => import('./axis-aligned-overlap.js');\n"],
    ['commonjs-require',
      "export const place = () => require('./axis-aligned-overlap.js');\n"],
    ['string-key',
      "const overlap = { resolveAxisAlignedOverlap() {} };\n"
        + "export const place = overlap['resolveAxisAlignedOverlap'];\n"],
  ]);
  for (const [name, source] of bypasses) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-float-${name}-`);
    write(root, 'packages/docx/src/layout/axis-aligned-overlap.ts',
      'export default function overlap(value) { return value; }\n'
        + 'export function resolveAxisAlignedOverlap(value) { return value; }\n');
    write(root, 'packages/docx/src/layout/paragraph.ts', source);
    expectDiagnostic(root, 'FLOAT_PLACEMENT_AUTHORITY', name, '--final');
  }
});

test('float compatibility and numeric policies have exact declaration owners', () => {
  const compatibilityRoot = initializeCanonicalFixture(
    'docx-layout-boundary-float-compatibility-owner-',
  );
  write(compatibilityRoot, 'packages/docx/src/layout/paragraph.ts',
    "export const WORD_FLOAT_UNTRACKED_HEURISTIC = 'hidden';\n");
  expectDiagnostic(
    compatibilityRoot,
    'FLOAT_COMPATIBILITY_AUTHORITY',
    'float compatibility declaration outside compatibility.ts',
    '--final',
  );

  const numericRoot = initializeCanonicalFixture(
    'docx-layout-boundary-float-numeric-owner-',
  );
  write(numericRoot, 'packages/docx/src/layout/floats.ts',
    'export const FLOAT_OVERLAP_EPS = 0.02;\n'
      + 'export const FLOAT_PAGE_RIGHT_SLACK = 0.5;\n');
  expectDiagnostic(
    numericRoot,
    'FLOAT_NUMERIC_POLICY',
    'float numerical policy value change',
    '--final',
  );
});

test('retained body paint fails closed without an auditable paragraph branch', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-body-fail-closed-');
  write(root, 'packages/docx/src/renderer.ts',
    'function renderBodyElements(elements) { return dispatchBody(elements); }\n');
  const result = expectDiagnostic(
    root,
    'BODY_PAINT_LAYOUT_CAPABILITY',
    'opaque body dispatch',
    '--final',
  );
  assert.match(result.output, /no statically auditable paragraph branch/);
});

test('retained body paint follows the else branch of a negated paragraph dispatch', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-body-negated-');
  write(root, 'packages/docx/src/renderer.ts',
    `function renderBodyElements(elements) {
      for (const element of elements) {
        if (element.type !== 'paragraph') paintTable(element);
        else measureParagraph(element);
      }
    }
    function paintTable() {}
    `);
  expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', 'negated paragraph dispatch', '--final');
});

test('final renderer rejects transitional exports, hidden algorithms, and non-layout imports', () => {
  for (const [name, source, diagnostic] of [
    ['deleted pagination export', canonicalRenderer + '\nexport function paginateDocument() {}\n', 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER'],
    ['star export', canonicalRenderer + "\nexport * from './layout/document.js';\n", 'FINAL_ADAPTER_EXPORT'],
    ['hidden algorithm', canonicalRenderer + '\nexport function accidentalAlgorithm() {}\n', 'FINAL_ADAPTER_DECLARATION'],
    ['source layout loop', canonicalRenderer.replace(
      'const normalized = normalizeRenderOptions(source, canvas, pageIndex, opts);',
      'const normalized = normalizeRenderOptions(source, canvas, pageIndex, opts);\n'
        + '  for (const item of source) renderSelectedDocumentPage(item);',
    ), 'FINAL_ADAPTER_BODY'],
    ['extra source layout', canonicalRenderer.replace(
      'const normalized = normalizeRenderOptions(source, canvas, pageIndex, opts);',
      'const normalized = normalizeRenderOptions(source, canvas, pageIndex, opts);\n'
        + '  layoutDocument(source);',
    ), 'FINAL_ADAPTER_BODY'],
    ['extra source paint', canonicalRenderer.replace(
      'return renderSelectedDocumentPage(',
      'renderSelectedDocumentPage(normalized.selection.layout);\n  return renderSelectedDocumentPage(',
    ), 'FINAL_ADAPTER_BODY'],
    ['foreign source normalization', canonicalRenderer.replace(
      'normalizeRenderOptions(source, canvas, pageIndex, opts)',
      'normalizeRenderOptions(foreignSource, canvas, pageIndex, opts)',
    ), 'FINAL_ADAPTER_BODY'],
    ['foreign source paint canvas', canonicalRenderer.replace(
      'normalized.selection.page,\n    canvas,',
      'normalized.selection.page,\n    foreignCanvas,',
    ), 'FINAL_ADAPTER_BODY'],
    ['foreign model adapter', canonicalRenderer.replace(
      'layoutSourceStore(doc), canvas, pageIndex, opts',
      'layoutSourceStore(foreignDoc), canvas, pageIndex, opts',
    ), 'FINAL_ADAPTER_BODY'],
    ['non-layout import', "import { hidden } from './hidden.js';\n" + canonicalRenderer, 'FINAL_ADAPTER_IMPORT'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-adapter-${name}-`);
    write(root, 'packages/docx/src/hidden.ts', 'export function hidden() {}\n');
    write(root, 'packages/docx/src/renderer.ts', source);
    expectDiagnostic(root, diagnostic, name, '--final');
  }
});

test('text-run projection has an exact root-adapter declaration surface', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-text-run-adapter-');
  write(root, 'packages/docx/src/layout/text-index.ts',
    'export interface TextRunGeometry {}\nexport function textRunGeometryForPage() { return []; }\n');
  write(root, 'packages/docx/src/paint/affine.ts',
    'export function cssTransformFor() {}\n');
  const adapter = `
import { canvasFontString, PT_TO_PX } from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './renderer.js';
import { composeAffine, mapAffinePoint, scaleAffine } from './layout/affine.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import type { TextRunGeometry } from './layout/text-index.js';
import type { DocumentLayout, LayoutServices, Matrix2DData } from './layout/types.js';
import { cssTransformFor } from './paint/affine.js';
export interface SelectedTextRunsForPageOptions {}
export interface TextRunsForPageOptions {}
function projectTextRun() { return [canvasFontString, composeAffine, mapAffinePoint, scaleAffine,
  textRunGeometryForPage, cssTransformFor] satisfies unknown; }
export function textRunsForPage(): DocxTextRunInfo[] { return projectTextRun() as DocxTextRunInfo[]; }
export function textRunsForSelectedPage(
  _services?: LayoutServices,
  _layout?: DocumentLayout,
  _matrix?: Matrix2DData,
) { return [selectDocumentLayoutPage, PT_TO_PX, {} as TextRunGeometry]; }
`;
  write(root, 'packages/docx/src/text-run-projection.ts', adapter);
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/text-run-projection.ts',
    `${adapter}\nexport const accidentalProjectionPolicy = true;\n`);
  expectDiagnostic(
    root,
    'TEXT_RUN_PROJECTION_DECLARATION',
    'extra text-run adapter declaration',
    '--final',
  );

  write(root, 'packages/docx/src/hidden.ts', 'export const hidden = true;\n');
  write(root, 'packages/docx/src/text-run-projection.ts',
    `import { hidden } from './hidden.js';\n${adapter}\nvoid hidden;\n`);
  expectDiagnostic(root, 'TEXT_RUN_PROJECTION_IMPORT', 'extra text-run adapter import', '--final');
});

test('final renderer import boundary rejects dynamic, bare, and unresolved imports exactly', () => {
  for (const [name, prefix] of [
    ['dynamic', 'import(globalThis.moduleName);\n'],
    ['bare', "import '@silurus/ooxml-core';\n"],
    ['unresolved', "import { missing } from './layout/missing.js';\nvoid missing;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-final-import-${name}-`);
    write(root, 'packages/docx/src/renderer.ts', `${prefix}${canonicalRenderer}`);
    expectDiagnostic(root, 'FINAL_ADAPTER_IMPORT', name, '--final');
  }
});

test('browser conformance assets cannot import removed renderer architecture bindings', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-browser-asset-import-');
  write(root, 'packages/docx/tests/visual/conformance-fixture.html', `
<script type="module">
  import { createLayoutServices, layoutDocument } from '/src/renderer.ts';
</script>
`);
  const result = expectDiagnostic(
    root,
    'FINAL_RENDERER_ASSET_IMPORT',
    'removed renderer import in browser asset',
    '--final',
  );
  assert.match(result.output, /conformance-fixture\.html -> createLayoutServices/);
});

test('the final boundary requires the browser conformance asset', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-browser-asset-missing-');
  rmSync(join(root, 'packages/docx/tests/visual/conformance-fixture.html'));
  expectDiagnostic(root, 'FINAL_RENDERER_ASSET_MISSING', 'missing browser asset', '--final');
});

test('test-only adapters stay excluded from production imports', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-test-support-');
  write(root, 'packages/docx/src/canonical-layout-test-adapter.test.ts',
    'export function acquireForTest() {}\n');
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/document.ts',
    "import { acquireForTest } from './canonical-layout-test-adapter.test.js';\n"
      + 'export const value = acquireForTest();\n');
  expectDiagnostic(root, 'PRODUCTION_TEST_SUPPORT_IMPORT', 'production test adapter import', '--final');
});

test('canonical producer must validate and deeply freeze its retained document layout', () => {
  for (const [name, source, diagnostic] of [
    ['second producer',
      `${canonicalBodyPaginator}export function alternateBodyProducer() { return {}; }\n`,
      'CANONICAL_LAYOUT_PRODUCER'],
    ['eager route bypasses the generator',
      canonicalBodyPaginator.replace(
        'return drainPagination(paginateBodySteps(input, services, options));',
        'return assertAndDeepFreezeDocumentLayout({ pages: [], diagnostics: [] });',
      ),
      'CANONICAL_LAYOUT_PRODUCER'],
    ['eager route drains an impostor generator',
      canonicalBodyPaginator.replace(
        'drainPagination(paginateBodySteps(input, services, options))',
        'drainPagination(alternatePaginationSteps(input, services, options))',
      ),
      'CANONICAL_LAYOUT_PRODUCER'],
    ['stepless producer',
      canonicalBodyPaginator
        .replace('export function* paginateBodySteps', 'export function paginateBodySteps')
        .replace('  yield 0;\n', ''),
      'CANONICAL_LAYOUT_PRODUCER'],
    ['missing drain seam',
      canonicalBodyPaginator.replace('export function drainPagination', 'function drainPagination'),
      'CANONICAL_LAYOUT_PRODUCER'],
    ['missing validation',
      canonicalBodyPaginator.replace(
        'return assertAndDeepFreezeDocumentLayout(layout);',
        'return deepFreezeDocumentLayout(layout);',
      ),
      'RETAINED_LAYOUT_IMMUTABILITY'],
    ['mutable return',
      canonicalBodyPaginator.replace(
        'return assertAndDeepFreezeDocumentLayout(layout);',
        'assertAndDeepFreezeDocumentLayout(layout);\n  return layout;',
      ),
      'RETAINED_LAYOUT_IMMUTABILITY'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-producer-${name}-`);
    write(root, 'packages/docx/src/layout/body-paginator.ts', source);
    expectDiagnostic(root, diagnostic, name, '--final');
  }
});

test('selected variant owns normalization and page validation', () => {
  for (const [name, source] of [
    ['silent missing store',
      'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
        + '  const store = layoutVariantStoreOf(services);\n'
        + '  return store?.selectPage(layoutOptionsForRender(input), pageIndex);\n}\n'],
    ['raw layout selection',
      'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
        + '  const store = layoutVariantStoreOf(services);\n'
        + "  if (!store) throw new Error('missing');\n"
        + '  return store.layoutFor(layoutOptionsForRender(input)).pages[pageIndex];\n}\n'],
    ['unnormalized selection',
      'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
        + '  const store = layoutVariantStoreOf(services);\n'
        + "  if (!store) throw new Error('missing');\n"
        + '  return store.selectPage(input, pageIndex);\n}\n'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-selection-${name}-`);
    write(root, 'packages/docx/src/layout/document-layout-variants.ts', source);
    expectDiagnostic(root, 'SELECTED_LAYOUT_VARIANT', name, '--final');
  }
});

test('worker retains keyed parity and never duplicates selected-page validation', () => {
  for (const [name, source] of [
    ['duplicate worker selection',
      "import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';\n"
        + 'export function render() { return selectDocumentLayoutPage(services, input, pageIndex); }\n'],
    ['raw worker pages', 'export const pages = [];\n'],
    ['unkeyed builder',
      "import { renderDocumentToCanvas } from './renderer.js';\n"
        + 'export function render() { return renderDocumentToCanvas(); }\n'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-worker-${name}-`);
    write(root, 'packages/docx/src/render-worker.ts', source);
    expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', name, '--final');
  }
});

test('a transitional baseline is forbidden by the default CI invocation', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-default-final-');
  write(root, 'scripts/docx-layout-boundary-baseline.json', '{}\n');
  expectDiagnostic(root, 'FINAL_BASELINE_PRESENT', 'final baseline');
});

test('reports an unknown CLI argument exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-cli-');
  expectDiagnostic(root, 'UNKNOWN_ARGUMENT', 'unknown argument', '--definitely-unknown');
});

test('reports a missing final renderer exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-missing-renderer-');
  rmSync(join(root, 'packages/docx/src/renderer.ts'));
  expectDiagnostic(root, 'FINAL_ADAPTER_MISSING', 'missing renderer', '--final');
});

test('reports final legacy inventory without accepting an alternate diagnostic', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-final-legacy-');
  write(root, 'packages/docx/src/layout/legacy-route.ts',
    'export const useOldLayoutEngine = true;\n');
  expectDiagnostic(root, 'FINAL_LEGACY_BOUNDARY', 'legacy migration flag', '--final');
});

test('reports non-literal module edges from paint exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-nonliteral-paint-');
  write(root, 'packages/docx/src/paint/dynamic.ts',
    "const moduleName = './canvas-page.js';\nexport const load = () => import(moduleName);\n");
  expectDiagnostic(root, 'NON_LITERAL_MODULE_EDGE', 'paint dynamic import', '--final');
});

test('reports non-literal module edges from layout exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-nonliteral-layout-');
  write(root, 'packages/docx/src/layout/dynamic.ts',
    "const moduleName = './types.js';\nexport const load = () => import(moduleName);\n");
  expectDiagnostic(root, 'NON_LITERAL_LAYOUT_MODULE_EDGE', 'layout dynamic import', '--final');
});

test('benchmarks are development tooling, not audited production layout code', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-bench-');
  write(root, 'packages/docx/src/parser-model.ts', 'export const model = 1;\n');
  write(root, 'packages/docx/src/layout/layout.bench.ts',
    "import { model } from '../parser-model.js';\nexport const benchModel = model;\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/document.ts',
    "import { benchModel } from './layout/layout.bench.js';\nexport const value = benchModel;\n");
  expectDiagnostic(root, 'PRODUCTION_TEST_SUPPORT_IMPORT', 'production bench import', '--final');
});

test('production rejects both test and test-support module imports', () => {
  for (const suffix of ['test', 'test-support']) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-production-${suffix}-`);
    write(root, `packages/docx/src/fixture.${suffix}.ts`, 'export const fixtureValue = 1;\n');
    write(root, 'packages/docx/src/document.ts',
      `import { fixtureValue } from './fixture.${suffix}.js';\nexport const value = fixtureValue;\n`);
    expectDiagnostic(root, 'PRODUCTION_TEST_SUPPORT_IMPORT', suffix, '--final');
  }
});

test('every deleted final producer and legacy stamp identifier is rejected exactly', () => {
  const deletedProducers = [
    'bodyFragmentFor',
    'bodyLayoutFallback',
    'computePages',
    'paginateDocument',
    'paginateWithHeaderFooterReserve',
    'PaginatedBodyElement',
    'physicalPageSizeForPage',
    'prebuiltPages',
    'retainedLayout',
    'sectionBreakSpacer',
    'collapsedSpacer',
    'leadsCollapsedRun',
    'hiddenCollapsed',
    'computeTableRowHeights',
    'estimateTableHeight',
    'measureRetainedCellContentHeightPt',
    'measureCellParagraphWindow',
    'measureCellElementHeight',
    'trimTrailingStructuralMarker',
    'rescaleLayoutLines',
    'paragraphSegsStateSensitive',
  ];
  const deletedStampProperties = [
    'colIndex',
    'colGeom',
    'colTopPt',
    'sectionHF',
    'sectionGeom',
    'sectionPageNumType',
    'sectionTextDirection',
  ];
  for (const name of deletedProducers) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-deleted-exact-${name}-`);
    write(root, 'packages/docx/src/deleted-path.ts', `export const ${name} = true;\n`);
    expectDiagnostic(root, 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER', name, '--final');
  }
  for (const name of deletedStampProperties) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-deleted-property-${name}-`);
    write(root, 'packages/docx/src/deleted-path.ts',
      `export const legacyStamp = { ${name}: true };\n`);
    expectDiagnostic(root, 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER', name, '--final');
  }
});

test('FloatRect transport cannot regain the deleted drawn state', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-float-drawn-state-');
  write(root, 'packages/docx/src/layout/float-wrap.ts', `
interface FloatRectCore {
  drawn: boolean;
}
export type FloatRect = FloatRectCore & { kind: 'shape' };
`);
  expectDiagnostic(root, 'FLOAT_RECT_TRANSITIONAL_STATE', 'FloatRect.drawn', '--final');
});

test('legacy stamp property guards do not ban unrelated local variable names', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-local-column-index-');
  write(root, 'packages/docx/src/layout/local-index.ts',
    'export function localIndex() { const colIndex = 0; return colIndex + 1; }\n');
  const result = runChecker(root, '--final');
  assert.equal(result.status, 0, result.output);
});

test('the concrete body kernel is private even though its declaration is allowed', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-exported-kernel-');
  write(root, 'packages/docx/src/renderer.ts', canonicalRenderer.replace(
    'function createConcreteBodyLayoutKernel',
    'export function createConcreteBodyLayoutKernel',
  ));
  expectDiagnostic(root, 'FINAL_ADAPTER_EXPORT', 'exported concrete kernel', '--final');
});

test('writing a transitional baseline is rejected after final cutover', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-first-baseline-');
  expectDiagnostic(
    root,
    'UNKNOWN_ARGUMENT',
    'removed baseline writer',
    '--write-transitional-baseline',
  );
});

test('the canonical body producer file is mandatory', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-producer-missing-');
  rmSync(join(root, 'packages/docx/src/layout/body-paginator.ts'));
  expectDiagnostic(root, 'CANONICAL_LAYOUT_PRODUCER', 'missing producer', '--final');
});

test('the canonical body producer rejects alternate exported value producers', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-producer-value-');
  const path = join(root, 'packages/docx/src/layout/body-paginator.ts');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    `${readFileSync(path, 'utf8')}\nexport const alternateBodyProducer = () => ({ pages: [] });\n`);
  expectDiagnostic(root, 'CANONICAL_LAYOUT_PRODUCER', 'alternate value producer', '--final');
});

test('the canonical body producer rejects runtime re-exports', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-producer-reexport-');
  const path = join(root, 'packages/docx/src/layout/body-paginator.ts');
  write(root, 'packages/docx/src/layout/alternate-producer.ts',
    'export const alternateBodyProducer = () => ({ pages: [] });\n');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    `${readFileSync(path, 'utf8')}\nexport { alternateBodyProducer } from './alternate-producer.js';\n`);
  expectDiagnostic(root, 'CANONICAL_LAYOUT_PRODUCER', 'runtime re-export', '--final');
});

test('retained layout validation and freezing must wrap the returned value', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-immutability-identity-');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    canonicalBodyPaginator.replace(
      '  return assertAndDeepFreezeDocumentLayout(layout);\n',
      '  const returned = { pages: [], diagnostics: [] };\n'
        + '  assertAndDeepFreezeDocumentLayout(layout);\n'
        + '  return returned;\n',
    ));
  expectDiagnostic(root, 'RETAINED_LAYOUT_IMMUTABILITY', 'unvalidated returned value', '--final');
});

test('retained layout validation uses the reviewed invariants import', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-immutability-impostor-');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    canonicalBodyPaginator.replace(
      "import { assertAndDeepFreezeDocumentLayout } from './invariants.js';\n",
      'function assertAndDeepFreezeDocumentLayout(value) { return Object.freeze(value); }\n',
    ));
  expectDiagnostic(root, 'RETAINED_LAYOUT_IMMUTABILITY', 'local impostor invariants', '--final');
});

test('the selected-layout variant module is mandatory', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-selection-missing-');
  rmSync(join(root, 'packages/docx/src/layout/document-layout-variants.ts'));
  expectDiagnostic(root, 'SELECTED_LAYOUT_VARIANT', 'missing selection module', '--final');
});

test('selected-layout validation rejects extra raw page selection even with the canonical statements present', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-selection-extra-');
  write(root, 'packages/docx/src/layout/document-layout-variants.ts',
    'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
      + '  if (input.rawLayout) return input.rawLayout.pages[pageIndex];\n'
      + '  const store = layoutVariantStoreOf(services);\n'
      + "  if (!store) throw new Error('missing');\n"
      + '  return store.selectPage(layoutOptionsForRender(input), pageIndex);\n'
      + '}\n');
  expectDiagnostic(root, 'SELECTED_LAYOUT_VARIANT', 'extra raw path', '--final');
});

test('renderer must delegate page validation to the selected-layout boundary exactly once', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-renderer-selection-');
  write(root, 'packages/docx/src/renderer.ts', canonicalRenderer.replace(
    'const selection = selectDocumentLayoutPage(services, input, pageIndex);',
    'const selection = { page: input.pages[pageIndex] };',
  ));
  expectDiagnostic(root, 'SELECTED_LAYOUT_VARIANT', 'renderer raw page selection', '--final');
});

test('the worker canonical selection route is mandatory', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-worker-missing-');
  rmSync(join(root, 'packages/docx/src/render-worker.ts'));
  expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', 'missing worker', '--final');
});

test('worker selection validates call ownership rather than only call counts', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-worker-call-shape-');
  write(root, 'packages/docx/src/render-worker.ts',
    'export function initializeWorker(model, layoutServices) {\n'
      + '  layoutDocument();\n'
      + '  return attachDocumentLayoutVariants({ buildLayout: () => ({ pages: [] }) });\n'
      + '}\n'
      + 'export function renderWorkerPage() { return renderDocumentToCanvas(); }\n'
      + 'export function collectWorkerRuns() { return renderDocumentToCanvas(); }\n');
  expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', 'wrong worker call ownership', '--final');
});

test('worker run collection cannot add a second dry-render call', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-worker-dry-collect-');
  const workerPath = join(root, 'packages/docx/src/render-worker.ts');
  const source = readFileSync(workerPath, 'utf8');
  write(root, 'packages/docx/src/render-worker.ts', source.replace(
    'export function collectWorkerRuns() { return []; }',
    'export function collectWorkerRuns(doc, source, canvas, req, options) {\n'
      + '  return renderLayoutSourceToCanvas(source, canvas, req.pageIndex, { ...options,\n'
      + '    layoutServices: doc.layoutServices,\n'
      + '    defaultCurrentDateMs: doc.defaultCurrentDateMs });\n'
      + '}',
  ));
  expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', 'dry run collection', '--final');
});

test('worker render calls require the exact retained worker layout authority', () => {
  for (const [name, from, to] of [
    ['foreign source', 'renderLayoutSourceToCanvas(source, canvas, req.pageIndex', 'renderLayoutSourceToCanvas(foreignSource, canvas, req.pageIndex'],
    ['foreign canvas', 'renderLayoutSourceToCanvas(source, canvas, req.pageIndex', 'renderLayoutSourceToCanvas(source, foreignCanvas, req.pageIndex'],
    ['foreign page', 'renderLayoutSourceToCanvas(source, canvas, req.pageIndex', 'renderLayoutSourceToCanvas(source, canvas, foreignPageIndex'],
    ['foreign services', 'layoutServices: doc.layoutServices', 'layoutServices: foreignServices'],
    ['derived services', 'layoutServices: doc.layoutServices', 'layoutServices: servicesFor(doc)'],
    ['foreign default date', 'defaultCurrentDateMs: doc.defaultCurrentDateMs', 'defaultCurrentDateMs: foreignDefaultCurrentDateMs'],
    ['derived default date', 'defaultCurrentDateMs: doc.defaultCurrentDateMs', 'defaultCurrentDateMs: defaultDateFor(doc)'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-worker-retained-${name}-`);
    const workerPath = join(root, 'packages/docx/src/render-worker.ts');
    const source = readFileSync(workerPath, 'utf8');
    write(root, 'packages/docx/src/render-worker.ts', source.replace(from, to));
    expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', name, '--final');
  }
});

test('worker construction requires the canonical retained-layout wiring seam', () => {
  for (const [name, from, to] of [
    ['foreign source', 'retainRenderWorkerDocumentLayout(source, layoutServices, req.defaultCurrentDateMs)', 'retainRenderWorkerDocumentLayout(foreignSource, layoutServices, req.defaultCurrentDateMs)'],
    ['foreign retained services', 'retainRenderWorkerDocumentLayout(source, layoutServices, req.defaultCurrentDateMs)', 'retainRenderWorkerDocumentLayout(source, foreignServices, req.defaultCurrentDateMs)'],
    ['foreign retained default date', 'retainRenderWorkerDocumentLayout(source, layoutServices, req.defaultCurrentDateMs)', 'retainRenderWorkerDocumentLayout(source, layoutServices, foreignDefaultCurrentDateMs)'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-worker-construction-${name}-`);
    const workerPath = join(root, 'packages/docx/src/render-worker.ts');
    const source = readFileSync(workerPath, 'utf8');
    write(root, 'packages/docx/src/render-worker.ts', source.replace(from, to));
    expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', name, '--final');
  }
});

test('worker retention attaches the exact parameter identities', () => {
  for (const [name, mutate] of [
    ['foreign attachment source', (source) => source
      .replace('{ source, services: layoutServices,', '{ source: foreignSource, services: layoutServices,')
      .replace('layoutDocument(source, layoutServices, options)', 'layoutDocument(foreignSource, layoutServices, options)')],
    ['foreign attachment services', (source) => source
      .replace('{ source, services: layoutServices,', '{ source, services: foreignServices,')
      .replace('layoutDocument(source, layoutServices, options)', 'layoutDocument(source, foreignServices, options)')],
    ['foreign attachment default date', (source) => source.replace(
      '    defaultCurrentDateMs,',
      '    defaultCurrentDateMs: foreignDefaultCurrentDateMs,',
    )],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-worker-attachment-${name}-`);
    const path = join(root, 'packages/docx/src/render-worker-layout.ts');
    write(root, 'packages/docx/src/render-worker-layout.ts',
      mutate(readFileSync(path, 'utf8')));
    expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', name, '--final');
  }
});

test('worker retention returns exactly the retained parameter identities and attached store', () => {
  for (const [name, from, to] of [
    ['foreign returned services', 'return { layoutServices,', 'return { layoutServices: foreignServices,'],
    ['derived returned services', 'return { layoutServices,', 'return { layoutServices: servicesFor(layoutServices),'],
    ['foreign returned store', 'layoutVariants: variants.store', 'layoutVariants: foreignStore'],
    ['foreign returned default date', 'defaultCurrentDateMs };', 'defaultCurrentDateMs: foreignDefaultCurrentDateMs };'],
    ['missing retained field', 'layoutVariants: variants.store, ', ''],
    ['extra retained field', 'defaultCurrentDateMs };', 'defaultCurrentDateMs, duplicateStore: variants.store };'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-worker-return-${name}-`);
    const path = join(root, 'packages/docx/src/render-worker-layout.ts');
    const source = readFileSync(path, 'utf8');
    const mutated = source.replace(from, to);
    assert.notEqual(mutated, source, `${name} mutation must change the fixture`);
    write(root, 'packages/docx/src/render-worker-layout.ts', mutated);
    expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', name, '--final');
  }
});

test('worker retention rejects declarations outside its exact ownership seam', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-worker-extra-declaration-');
  const path = join(root, 'packages/docx/src/render-worker-layout.ts');
  write(root, 'packages/docx/src/render-worker-layout.ts',
    `${readFileSync(path, 'utf8')}\nfunction alternateWorkerLayout() {}\n`);
  expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', 'extra worker declaration', '--final');
});

test('worker parse metadata comes from the retained selected-variant route', () => {
  for (const [name, file, from, to] of [
    ['foreign metadata layout', 'render-worker.ts', 'doc.layoutVariants.layoutFor(layoutOptions)', 'foreignLayout'],
    ['derived metadata layout', 'render-worker.ts', 'doc.layoutVariants.layoutFor(layoutOptions)', 'layoutForMetadata(doc)'],
    // Reporting the DEFAULT variant is the specific regression this pin exists
    // for: a tracked-changes or explicit-date load paginates differently, so
    // default-variant metadata describes pages the worker will never paint.
    ['default metadata layout', 'render-worker.ts', 'doc.layoutVariants.layoutFor(layoutOptions)', 'doc.layoutVariants.defaultLayout'],
    ['foreign metadata variant', 'render-worker.ts', 'layoutFor(layoutOptions)', 'layoutFor(foreignOptions)'],
    // The selected variant must come from the request's own view fields, not a
    // fabricated one, or the worker paginates a view the host never asked for.
    ['derived metadata variant', 'render-worker.ts', 'normalizeLayoutOptions(req.currentDateMs,', 'normalizeLayoutOptions(Date.now(),'],
    ['dropped tracked-changes axis', 'render-worker.ts', 'req.defaultCurrentDateMs, req.showTrackedChanges)', 'req.defaultCurrentDateMs, false)'],
    ['foreign initial projection', 'render-worker.ts', 'projectRenderWorkerLayoutMeta(layout, source, reviewIndexInput)', 'projectRenderWorkerLayoutMeta(foreignLayout, source, reviewIndexInput)'],
    ['dropped runtime tracked axis', 'render-worker.ts', 'req.currentDateMs, req.showTrackedChanges)', 'req.currentDateMs, false)'],
    ['foreign metadata page count', 'render-worker-metadata.ts', 'pageCount: layout.pages.length', 'pageCount: foreignLayout.pages.length'],
    ['foreign metadata page sizes', 'render-worker-metadata.ts', 'pageSizes: layout.pages.map', 'pageSizes: foreignLayout.pages.map'],
    ['foreign metadata bookmarks', 'render-worker-metadata.ts', 'buildBookmarkPageMap(layout)', 'buildBookmarkPageMap(foreignLayout)'],
    ['foreign review projection index', 'render-worker-metadata.ts', 'reviewProjectionIndexForDocument(layout)', 'reviewProjectionIndexForDocument(foreignLayout)'],
    ['unconditional comment projection', 'render-worker-metadata.ts', 'commentAnchorRanges: hasComments', 'commentAnchorRanges: true'],
    ['foreign comment run index', 'render-worker-metadata.ts', 'reviewIndex!.renderedRunIndex', 'foreignReviewIndex.renderedRunIndex'],
    ['foreign runtime metadata layout', 'render-worker-metadata.ts', 'doc.layoutVariants.layoutFor(normalizeLayoutOptions(', 'foreignLayoutFor(normalizeLayoutOptions('],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-worker-metadata-${name}-`);
    const path = join(root, `packages/docx/src/${file}`);
    write(root, `packages/docx/src/${file}`,
      readFileSync(path, 'utf8').replace(from, to));
    expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', name, '--final');
  }
});
