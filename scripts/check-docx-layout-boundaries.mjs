#!/usr/bin/env node

import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { createRequire } from 'node:module';
import { basename, dirname, join, relative, resolve, sep } from 'node:path';
import { pathToFileURL } from 'node:url';

// TypeScript 7 intentionally exposes no stable Compiler API yet. This
// architecture guard is root-only tooling, so it uses the explicitly named
// TypeScript 6 compatibility dependency while library builds stay on TS 7.
const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript-compiler-api');

const BASELINE_PATH = 'scripts/docx-layout-boundary-baseline.json';
const DOCX_SOURCE = 'packages/docx/src';
const PAINT_SOURCE = `${DOCX_SOURCE}/paint`;
const LAYOUT_SOURCE = `${DOCX_SOURCE}/layout`;
const PARSER_MODEL = `${DOCX_SOURCE}/parser-model.ts`;
const BODY_LAYOUT_ADAPTER = `${DOCX_SOURCE}/body-layout-input.ts`;
const PARAGRAPH_ANCHOR_FRAME_ADAPTER = `${DOCX_SOURCE}/paragraph-anchor-frame-adapter.ts`;
const WORKER_LAYOUT_RETENTION = `${DOCX_SOURCE}/render-worker-layout.ts`;
const TEXT_RUN_PROJECTION_ADAPTER = `${DOCX_SOURCE}/text-run-projection.ts`;
const LAYOUT_RUNTIME_ADAPTER = `${DOCX_SOURCE}/layout-runtime.ts`;
const ACQUISITION_CONTEXT = `${LAYOUT_SOURCE}/acquisition-context.ts`;
const PRODUCTION_BODY_LAYOUT = `${LAYOUT_SOURCE}/production-body-layout.ts`;
const ACQUISITION_STATE = `${LAYOUT_SOURCE}/acquisition-state.ts`;
const ACQUISITION_INPUT_PROJECTIONS = `${LAYOUT_SOURCE}/acquisition-input-projections.ts`;
const ANCHOR_CLASSIFICATION = `${LAYOUT_SOURCE}/anchor-classification.ts`;
const FLOAT_WRAP = `${LAYOUT_SOURCE}/float-wrap.ts`;
const MEASUREMENT_ENVIRONMENT = `${LAYOUT_SOURCE}/measurement-environment.ts`;
const MEASUREMENT_CAPABILITIES = `${LAYOUT_SOURCE}/measurement-capabilities.ts`;
const SECTION_ORIENTATION = `${LAYOUT_SOURCE}/section-orientation.ts`;
const LAYOUT_PARSER_MODEL_GATEWAY = `${LAYOUT_SOURCE}/resources.ts`;
const LAYOUT_AFFINE = `${LAYOUT_SOURCE}/affine.ts`;
const CONFORMANCE_FIXTURE = 'packages/docx/tests/visual/conformance-fixture.html';
const LAYOUT_PARSER_MODEL_GATEWAY_IMPORT = '../parser-model.js';
const LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL = 'normalizeInternalDocumentModel';

const CANONICAL_PAGINATION_PRODUCER = 'paginateBodySteps';
const CANONICAL_PAGINATION_DRAIN = 'drainPagination';
const CANONICAL_PAGINATION_EAGER_ROUTE = 'paginateBody';
const CANONICAL_PAGINATOR_EXPORTS = [
  CANONICAL_PAGINATION_DRAIN,
  CANONICAL_PAGINATION_PRODUCER,
  CANONICAL_PAGINATION_EAGER_ROUTE,
];

const FINAL_RENDERER_EXPORTS = new Set([
  'DocxTextRunInfo',
  'RenderDocumentOptions',
  'clearResolvedLocalFonts',
  'documentHasMath',
  'dropColorReplacedCache',
  'prepareMathRuns',
  'renderDocumentToCanvas',
  'renderLayoutSourceToCanvas',
  'setResolvedLocalFonts',
]);

const FINAL_RENDERER_DECLARATIONS = new Set([
  ...FINAL_RENDERER_EXPORTS,
  'createConcreteBodyLayoutKernel',
  'createLayoutServices',
  'normalizeRenderOptions',
]);

const BODY_LAYOUT_ADAPTER_DECLARATIONS = new Set(['createBodyLayoutInput']);
const WORKER_LAYOUT_RETENTION_DECLARATIONS = new Set([
  'RetainedRenderWorkerDocumentLayout',
  'retainRenderWorkerDocumentLayout',
]);
const TEXT_RUN_PROJECTION_ADAPTER_DECLARATIONS = new Set([
  'SelectedTextRunsForPageOptions',
  'TextRunsForPageOptions',
  'projectTextRun',
  'textRunsForPage',
  'textRunsForSelectedPage',
]);
const LAYOUT_AFFINE_EXPORTS = new Set([
  'composeAffine',
  'inverseMapAffinePoint',
  'inverseMapAffineVector',
  'mapAffinePoint',
  'quarterTurnAffine',
  'scaleAffine',
  'translationAffine',
]);
const BODY_LAYOUT_ADAPTER_IMPORT_BINDINGS = new Map([
  [PARSER_MODEL, new Map([['bodyLayoutAcquisitionInput', 'value']])],
  [`${DOCX_SOURCE}/types.ts`, new Map([['DocxDocumentModel', 'type']])],
  [`${LAYOUT_SOURCE}/body-layout-input.ts`, new Map([
    ['projectBodyLayoutInput', 'value'],
    ['BodyLayoutInput', 'type'],
  ])],
]);

const SHARED_PAINT_IMPORTS = new Map([
  ['@silurus/ooxml-core', new Map([
    ['acquireBitmapCacheLease', 'value'],
    ['applyDuotone', 'value'],
    ['autoContrastColor', 'value'],
    ['captureDecodedBitmapCacheEpoch', 'value'],
    ['canvasFontString', 'value'],
    ['clampCanvasSize', 'value'],
    // Chart picture-marker discovery and lookup warm the document-owned image
    // cache before measurement-free chart paint. They expose no layout API.
    ['chartImageFillKey', 'value'],
    ['chartImageFillUsageSize', 'value'],
    ['collectChartImageFillUsages', 'value'],
    ['collectChartImageFillUsagesForCharts', 'value'],
    ['collectChartMarkerImageFills', 'value'],
    ['collectChartMarkerImageFillsForCharts', 'value'],
    // DrawingML silhouette clipping is shared paint behavior used by
    // resource-backed shape fills; it does not expose layout acquisition.
    ['clipDrawingMLShape', 'value'],
    ['crispOffset', 'value'],
    ['defaultDpr', 'value'],
    ['deferBitmapCloseWhileLeased', 'value'],
    ['docxBorderDashArray', 'value'],
    ['drawImageCropped', 'value'],
    ['doubleRailGeometry', 'value'],
    ['fillDoubleBorder', 'value'],
    // Shared decoded-surface ownership is a paint resource concern. DOCX may
    // create/drop derived clrChange/duotone surfaces without importing layout.
    ['dropCachedDerivedBitmapNamespace', 'value'],
    ['decodedBitmapTargetResizeOptions', 'value'],
    ['duotoneImageData', 'value'],
    ['getCachedBitmapByPath', 'value'],
    ['getCachedDerivedBitmap', 'value'],
    ['getCachedSvgImageByPath', 'value'],
    // Header inspection reuses the document-owned byte profile to decide
    // whether browser target-resize is available. Both predicates remain
    // paint-resource admission and cannot affect acquired layout geometry.
    ['inspectCachedRasterSource', 'value'],
    ['isBrowserResizableRasterMimeType', 'value'],
    ['isDecodeTargetResizableRasterFormat', 'value'],
    ['isOoxmlDecodedImageLimitError', 'value'],
    ['isTiffDecodeError', 'value'],
    // Missing optional image codecs are contained at paint-resource acquisition;
    // the placeholder consumes only the already-retained authored bounds.
    ['isOptionalImageCodecUnavailableError', 'value'],
    ['paintOptionalImagePlaceholder', 'value'],
    ['OptionalImageCodecUnavailableError', 'type'],
    ['HyperlinkTarget', 'type'],
    ['imageNaturalSize', 'value'],
    ['isHTMLCanvas', 'value'],
    ['mathToMathML', 'value'],
    // Math engine output is decoded before layout; paint consumes only the
    // resulting CanvasImageSource keyed by the immutable resource record.
    ['rasterizeMathSvg', 'value'],
    ['metafileRasterSize', 'value'],
    // Paint-time image admission uses the completed layout box and effective
    // DPR to choose a retained decode. These resource-policy primitives neither
    // measure nor feed geometry back into DocumentLayout.
    ['MAX_RASTER_PIXELS', 'value'],
    // Decoded-image policy is shared across all format painters. DOCX supplies
    // immutable paint demands and consumes the returned target plan; neither
    // operation may feed measurement or pagination.
    ['normalizeImageResourceOptions', 'value'],
    ['planDecodedImageTargets', 'value'],
    ['paintDrawingMLShape', 'value'],
    ['withVertFeature', 'value'],
    ['preferVectorBlip', 'value'],
    ['releaseOwnedBitmap', 'value'],
    ['resolvedCachedBitmapVariantKey', 'value'],
    ['sourceRasterTargetSize', 'value'],
    ['PT_TO_PX', 'value'],
    // Shared fill resolution keeps gradient/no-fill semantics identical across
    // DOCX, PPTX, and XLSX painters; paint may consume it but not layout APIs.
    ['resolveFill', 'value'],
    ['recolorSvg', 'value'],
    ['renderChart', 'value'],
    ['withDrawingMLShapeTransform', 'value'],
    // Serial admission and bitmap pinning are document-owned paint-resource
    // lifecycle concerns, not layout acquisition.
    ['withBitmapCacheLease', 'value'],
    // Optional renderers and image codecs are paint-only capabilities threaded
    // into existing shared painters. Their erased contracts do not expose
    // layout acquisition or permit a renderer back-edge.
    ['ChartThreeDRenderer', 'type'],
    ['ChartRegionMapRenderer', 'type'],
    ['ChartExRenderer', 'type'],
    ['ChartImageLookup', 'type'],
    ['Duotone', 'type'],
    ['ImageResourceOptions', 'type'],
    ['MathRenderer', 'type'],
    ['TiffRenderer', 'type'],
  ])],
]);

const LEGACY_SYMBOLS = [
  'fitMeasureReuseEnabled',
  'fragmentPaintEnabled',
  'lineReuseEnabled',
  'isFragmentPaintableParagraph',
  'layoutLinesInputs',
  'stampParagraphLines',
  'renderBodyParagraphLines',
  'renderShapeText',
  'tableRequiresLegacyPaint',
  'isFragmentPaintableTable',
  'tableReuseEnabled',
  'renderTableFragment',
  'computePages',
  'computeTableLayout',
  'calculateRowHeight',
  'measureCellContentHeightPx',
  'buildTableCellBlocks',
  'renderHeaderFooter',
  'measureFootnoteHeight',
  'RenderState',
  'deferFront',
  'deferFrontDrawing',
  'deferBehindDrawing',
  'deferredFrontPaint',
  'deferredPaintWrapper',
  'bodyDrawingPass',
  'sectionBreakSpacer',
  'collapsedSpacer',
  'leadsCollapsedRun',
  'hiddenCollapsed',
  'tableColWidthsPt',
  'tableRowHeightsPt',
  'tableLayoutInputs',
];

const DELETED_PAGE_PRODUCER_IDENTIFIERS = new Set([
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
  'paragraphSegsStateSensitive',
  'rescaleLayoutLines',
]);

const DELETED_LEGACY_STAMP_PROPERTIES = new Set([
  'colIndex',
  'colGeom',
  'colTopPt',
  'sectionHF',
  'sectionGeom',
  'sectionPageNumType',
  'sectionTextDirection',
]);

const LEGACY_RENDERER_IMPORTS = new Set([
  'layout-context.ts',
  'layout-fragments.ts',
  'line-layout.ts',
  'paragraph-measure.ts',
  'table-fragments.ts',
  'table-geometry.ts',
]);

const ACQUISITION_CONTEXT_DECLARATIONS = new Set([
  'AnchorFloatRegistrationState',
  'AnchorGeometryContext',
  'BodyAcquisitionState',
  'BodyMeasurementContext',
  'FloatRegistrationState',
  'PhysicalAnchorFrame',
  'RetainedTableRecord',
]);

const MEASUREMENT_CAPABILITY_DECLARATIONS = new Set([
  'MeasurementTextContext',
  'VerticalGlyphMeasurementService',
]);

const ACQUISITION_INPUT_PROJECTION_DECLARATIONS = new Set([
  'BodyAcquisitionInputProjections',
]);

const ACQUISITION_STATE_DECLARATIONS = new Set([
  'BODY_STORY_CONTEXT',
  'bodyAnchorReferenceFrames',
  'resolveBodyParagraphLayoutContext',
  'resolveStateParagraphLayoutContext',
  'withTableCellStory',
  'retainedTableRecord',
]);

const ANCHOR_CLASSIFICATION_DECLARATIONS = new Set([
  'isPageLevelAnchorY',
  'isPageLevelWrapFloat',
]);

const MEASUREMENT_ENVIRONMENT_DECLARATIONS = new Set([
  'canonicalParagraphTextScaleEligible',
  'docDefaultFontSizePt',
  'paragraphMeasurementEnvironment',
  'segmentEnvironmentOf',
  'snapParaLineToGrid',
  'gridForParagraphContext',
]);

const SECTION_ORIENTATION_DECLARATIONS = new Set([
  'isVerticalSection',
  'isVerticalTextDirection',
  'isAllRotatedVerticalTextDirection',
  'verticalLayoutSection',
  'verticalLayoutDoc',
  'physicalLayoutSection',
]);

const EXACT_ACQUISITION_SURFACE_MEMBERS = new Map([
  [ACQUISITION_INPUT_PROJECTIONS, new Map([
    ['BodyAcquisitionInputProjections', new Set([
      'numberingMarkerShapeInput',
      'paragraphMarkShapeInput',
      'tableFormatInput',
      'tableColumnLayoutInput',
      'tableParticipatesInOrdinaryFlow',
      'paragraphAcquisitionInput',
    ])],
  ])],
  [MEASUREMENT_CAPABILITIES, new Map([
    ['MeasurementTextContext', new Set([
      'font',
      'letterSpacing',
      'fontKerning',
      'measureText',
    ])],
    ['VerticalGlyphMeasurementService', new Set([
      'fingerprint',
      'measureRunInkExtra',
      'planRun',
    ])],
  ])],
]);

const ACQUISITION_CONTEXT_CONSUMERS = [
  `${DOCX_SOURCE}/anchor-geometry.ts`,
  `${DOCX_SOURCE}/float-table-geometry.ts`,
  `${DOCX_SOURCE}/frame-geometry.ts`,
  `${DOCX_SOURCE}/line-layout.ts`,
  `${DOCX_SOURCE}/paragraph-measure.ts`,
];

const ACQUISITION_PAINT_PROPERTIES = new Set([
  'canvas',
  'defaultColor',
  'dpr',
  'drawImage',
  'dryRun',
  'fillText',
  'images',
  'restore',
  'save',
]);

function fail(code, detail) {
  throw new Error(`${code}: ${detail}`);
}

function posixPath(path) {
  return path.split(sep).join('/');
}

function isProductionTypeScript(path) {
  return /\.tsx?$/.test(path)
    && !path.endsWith('.d.ts')
    // `.bench.ts` files are `vitest bench` tooling, not shipped code: they are
    // free to reach across layout boundaries the way tests are, and production
    // may not import them (see `assertNoProductionTestSupportImports`).
    && !/\.(test|spec|stories|test-support|bench)\.tsx?$/.test(path)
    && !path.includes('/wasm/');
}

function assertNoProductionTestSupportImports(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    for (const edge of moduleEdges(path)) {
      if (!edge.literal || !edge.specifier.startsWith('.')) continue;
      const dependency = resolveLocalImport(path, edge.specifier);
      if (dependency && /\.(?:test|test-support|bench)\.tsx?$/.test(dependency)) {
        fail(
          'PRODUCTION_TEST_SUPPORT_IMPORT',
          `${posixPath(relative(root, path))} -> ${posixPath(relative(root, dependency))}`,
        );
      }
    }
  }
}

function assertAcquisitionContextBoundary(root) {
  const surfaces = new Map([
    [ACQUISITION_CONTEXT, ACQUISITION_CONTEXT_DECLARATIONS],
    [ACQUISITION_STATE, ACQUISITION_STATE_DECLARATIONS],
    [ACQUISITION_INPUT_PROJECTIONS, ACQUISITION_INPUT_PROJECTION_DECLARATIONS],
    [ANCHOR_CLASSIFICATION, ANCHOR_CLASSIFICATION_DECLARATIONS],
    [MEASUREMENT_ENVIRONMENT, MEASUREMENT_ENVIRONMENT_DECLARATIONS],
    [MEASUREMENT_CAPABILITIES, MEASUREMENT_CAPABILITY_DECLARATIONS],
    [SECTION_ORIENTATION, SECTION_ORIENTATION_DECLARATIONS],
  ]);
  for (const [file, requiredDeclarations] of surfaces) {
    const path = resolve(root, file);
    if (!existsSync(path)) {
      fail('ACQUISITION_CONTEXT_SURFACE', `${file} missing`);
    }
    const source = sourceFile(path);
    const declarations = new Set(source.statements.flatMap(declarationNames));
    for (const name of requiredDeclarations) {
      if (!declarations.has(name)) {
        fail('ACQUISITION_CONTEXT_SURFACE', `${file} missing ${name}`);
      }
    }
    if (file === ACQUISITION_CONTEXT) {
      const state = source.statements.find((statement) => (
        ts.isInterfaceDeclaration(statement) && statement.name.text === 'BodyAcquisitionState'
      ));
      for (const required of ['retainedTableAcquisition', 'retainedTablesBySourceIndex']) {
        const member = state?.members.find((candidate) => (
          ts.isPropertySignature(candidate)
          && candidate.name
          && (ts.isIdentifier(candidate.name) || ts.isStringLiteralLike(candidate.name))
          && candidate.name.text === required
        ));
        if (!member || member.questionToken) {
          fail('RETAINED_TABLE_AUTHORITY', `${file}#${required}`);
        }
      }
    }
    for (const statement of source.statements) {
      if (file === ACQUISITION_CONTEXT) {
        const rejectScaleMember = (node) => {
          if ((ts.isPropertySignature(node) || ts.isMethodSignature(node))
            && node.name
            && (ts.isIdentifier(node.name) || ts.isStringLiteralLike(node.name))
            && node.name.text === 'scale') {
            fail('ACQUISITION_COORDINATE_SCALE', `${file}#scale`);
          }
          if (ts.isLiteralTypeNode(node)
            && ts.isStringLiteralLike(node.literal)
            && node.literal.text === 'scale') {
            fail('ACQUISITION_COORDINATE_SCALE', `${file}#'scale'`);
          }
          ts.forEachChild(node, rejectScaleMember);
        };
        rejectScaleMember(statement);
      }
      if (!ts.isInterfaceDeclaration(statement)) continue;
      for (const member of statement.members) {
        if ((!ts.isPropertySignature(member) && !ts.isMethodSignature(member)) || !member.name) {
          continue;
        }
        const name = ts.isIdentifier(member.name) || ts.isStringLiteralLike(member.name)
          ? member.name.text
          : null;
        if (name && ACQUISITION_PAINT_PROPERTIES.has(name)) {
          fail('ACQUISITION_PAINT_CAPABILITY', `${file}#${name}`);
        }
      }
    }
    const exactInterfaces = EXACT_ACQUISITION_SURFACE_MEMBERS.get(file);
    if (!exactInterfaces) continue;
    const unexpectedDeclarations = [...declarations]
      .filter((name) => !requiredDeclarations.has(name));
    if (unexpectedDeclarations.length > 0) {
      fail(
        'ACQUISITION_CONTEXT_SURFACE',
        `${file} extra declarations ${unexpectedDeclarations.sort().join(',')}`,
      );
    }
    for (const [interfaceName, expectedMembers] of exactInterfaces) {
      const declaration = source.statements.find((statement) => (
        ts.isInterfaceDeclaration(statement) && statement.name.text === interfaceName
      ));
      if (!declaration || declaration.heritageClauses?.length) {
        fail('ACQUISITION_CONTEXT_SURFACE', `${file}#${interfaceName} heritage`);
      }
      const memberNames = declaration.members.map((member) => (
        member.name && (ts.isIdentifier(member.name) || ts.isStringLiteralLike(member.name))
          ? member.name.text
          : null
      ));
      const actualMembers = new Set(memberNames.filter((name) => name !== null));
      const exact = memberNames.length === expectedMembers.size
        && actualMembers.size === expectedMembers.size
        && [...expectedMembers].every((name) => actualMembers.has(name));
      if (!exact) {
        fail(
          'ACQUISITION_CONTEXT_SURFACE',
          `${file}#${interfaceName} members ${[...actualMembers].sort().join(',')}`,
        );
      }
    }
  }
  const inspected = [
    ...surfaces.keys(),
    ...ACQUISITION_CONTEXT_CONSUMERS,
  ];
  for (const file of inspected) {
    const path = resolve(root, file);
    if (!existsSync(path)) continue;
    for (const edge of moduleEdges(path)) {
      if (!edge.literal || !edge.specifier.startsWith('.')) continue;
      const dependency = resolveLocalImport(path, edge.specifier);
      if (dependency && posixPath(relative(root, dependency)) === `${DOCX_SOURCE}/renderer.ts`) {
        fail(
          'ACQUISITION_RENDERER_DEPENDENCY',
          `${file} -> ${DOCX_SOURCE}/renderer.ts`,
        );
      }
    }
  }
}

function assertProductionBodyAcquisitionAuthority(root) {
  const path = resolve(root, PRODUCTION_BODY_LAYOUT);
  if (!existsSync(path)) {
    fail('PRODUCTION_ACQUISITION_AUTHORITY', `${PRODUCTION_BODY_LAYOUT} missing`);
  }
  const source = sourceFile(path);
  const forbiddenImportsByTarget = new Map([
    [`${DOCX_SOURCE}/line-layout.ts`, new Set([
      'buildSegments',
      'gridCharDeltaPx',
      'layoutLines',
      'lineBoxHeight',
    ])],
    [`${DOCX_SOURCE}/table-geometry.ts`, new Set([
      'applyTableRowBoundaryFootprints',
      'resolveTableRowContentHeights',
    ])],
  ]);
  for (const edge of moduleEdges(path)) {
    if (!edge.literal || !edge.specifier.startsWith('.')) continue;
    const resolvedTarget = resolveLocalImport(path, edge.specifier);
    const fallbackTarget = resolve(dirname(path), edge.specifier)
      .replace(/\.(?:[cm]?js)$/u, '.ts');
    const target = posixPath(relative(root, resolvedTarget ?? fallbackTarget));
    const forbidden = forbiddenImportsByTarget.get(target);
    if (!forbidden) continue;
    const importedNames = edge.importedNames ?? [];
    if (edge.kind !== 'import'
      || importedNames.includes('*')
      || importedNames.includes('default')
      || importedNames.some((name) => forbidden.has(name))) {
      fail(
        'PRODUCTION_ACQUISITION_AUTHORITY',
        `${PRODUCTION_BODY_LAYOUT} imports fallback measurement from ${target}`,
      );
    }
  }
  const functions = [];
  const visitFunctions = (node) => {
    if (ts.isFunctionDeclaration(node) && node.name) functions.push(node);
    ts.forEachChild(node, visitFunctions);
  };
  visitFunctions(source);
  const uniqueFunction = (name) => {
    const matches = functions.filter((declaration) => declaration.name?.text === name);
    return matches.length === 1 ? matches[0] : undefined;
  };
  const columns = uniqueFunction('resolveColumnWidths');
  const baseContexts = [];
  const visitBaseContexts = (node) => {
    if (ts.isVariableDeclaration(node)
      && ts.isIdentifier(node.name)
      && node.name.text === 'baseContext') baseContexts.push(node);
    ts.forEachChild(node, visitBaseContexts);
  };
  if (columns?.body) visitBaseContexts(columns.body);
  const paragraphContextCalls = columns?.body
    ? callsNamed(columns.body, 'resolveParagraphLayoutContext')
    : [];
  if (!columns?.body
    || baseContexts.length !== 1
    || paragraphContextCalls.length !== 1
    || callOf(baseContexts[0].initializer, 'resolveParagraphLayoutContext')
      !== paragraphContextCalls[0]) {
    fail('PRODUCTION_ACQUISITION_AUTHORITY', `${PRODUCTION_BODY_LAYOUT}#resolveColumnWidths`);
  }

  const table = uniqueFunction('computeTablePtLayout');
  const tableSourceIndex = table?.parameters[3];
  const tableBodyCalls = table?.body ? callsNamed(table.body, 'acquireRetainedTable') : [];
  const tableCalls = callsNamed(source, 'computeTablePtLayout');
  const forbiddenTableCalls = [
    'applyTableRowBoundaryFootprints',
    'resolveTableRowContentHeights',
    'measureRetainedCellContentHeightPt',
  ].flatMap((name) => table?.body ? callsNamed(table.body, name) : []);
  if (!table?.body
    || table.parameters.length !== 4
    || !tableSourceIndex
    || tableSourceIndex.questionToken
    || tableSourceIndex.type?.kind !== ts.SyntaxKind.NumberKeyword
    || tableBodyCalls.length !== 1
    || tableCalls.length === 0
    || tableCalls.some((call) => call.arguments.length !== 4)
    || forbiddenTableCalls.length !== 0) {
    fail('PRODUCTION_ACQUISITION_AUTHORITY', `${PRODUCTION_BODY_LAYOUT}#computeTablePtLayout`);
  }

  const frame = uniqueFunction('resolveFrameBox');
  const frameGroup = frame?.parameters[1];
  const frameBodyCalls = frame?.body ? callsNamed(frame.body, 'acquireRetainedFrameGroup') : [];
  const frameCalls = callsNamed(source, 'resolveFrameBox');
  const frameType = frameGroup?.type;
  const forbiddenFrameCalls = [
    'buildSegments',
    'gridCharDeltaPx',
    'layoutLines',
    'lineBoxHeight',
  ].flatMap((name) => frame?.body ? callsNamed(frame.body, name) : []);
  if (!frame?.body
    || !frameGroup
    || frameGroup.questionToken
    || !frameType
    || !ts.isTypeReferenceNode(frameType)
    || !ts.isIdentifier(frameType.typeName)
    || frameType.typeName.text !== 'BodyFrameGroup'
    || frameBodyCalls.length !== 1
    || frameCalls.length === 0
    || frameCalls.some((call) => call.arguments.length !== 5)
    || forbiddenFrameCalls.length !== 0) {
    fail('PRODUCTION_ACQUISITION_AUTHORITY', `${PRODUCTION_BODY_LAYOUT}#resolveFrameBox`);
  }
}

function assertRendererAcquisitionProjectionBoundary(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) return;
  const parserModel = resolve(root, PARSER_MODEL);
  const injectedProjectionMembers = new Set([
    'paragraphAcquisitionInput',
    'tableFormatInput',
  ]);

  for (const edge of moduleEdges(renderer)) {
    if (!edge.literal) continue;
    const targetsParserModel = resolveLocalImport(renderer, edge.specifier) === parserModel
      || edge.specifier === './parser-model.js';
    if (!targetsParserModel) continue;
    const importsProjectionDirectly = edge.importedNames?.some((name) => (
      name === '*' || name === 'default' || injectedProjectionMembers.has(name)
    ));
    if (edge.kind !== 'import' || importsProjectionDirectly) {
      fail(
        'RENDERER_ACQUISITION_PROJECTION_BYPASS',
        `${DOCX_SOURCE}/renderer.ts ${edge.kind} ${edge.specifier}`,
      );
    }
  }

  const source = sourceFile(renderer);
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement)
      || !ts.isStringLiteral(statement.moduleSpecifier)) continue;
    const targetsParserModel = resolveLocalImport(renderer, statement.moduleSpecifier.text)
      === parserModel || statement.moduleSpecifier.text === './parser-model.js';
    if (!targetsParserModel) continue;
    const bindings = statement.importClause?.namedBindings;
    if (!bindings || !ts.isNamedImports(bindings)) continue;
    for (const element of bindings.elements) {
      const imported = element.propertyName?.text ?? element.name.text;
      if (imported === 'bodyAcquisitionInputProjections'
        && element.name.text !== 'bodyAcquisitionInputProjections') {
        fail(
          'RENDERER_ACQUISITION_PROJECTION_BYPASS',
          `${DOCX_SOURCE}/renderer.ts aliases bodyAcquisitionInputProjections`,
        );
      }
    }
  }
  const staticString = (expression) => {
    if (ts.isParenthesizedExpression(expression)) return staticString(expression.expression);
    if (ts.isStringLiteralLike(expression)) return expression.text;
    if (ts.isBinaryExpression(expression)
      && expression.operatorToken.kind === ts.SyntaxKind.PlusToken) {
      const left = staticString(expression.left);
      const right = staticString(expression.right);
      return left === null || right === null ? null : left + right;
    }
    return null;
  };
  const projectionMember = (expression) => {
    if (ts.isPropertyAccessExpression(expression)) return expression.name.text;
    if (ts.isElementAccessExpression(expression)
      && expression.argumentExpression) {
      return staticString(expression.argumentExpression);
    }
    return null;
  };
  const projectionReceiver = (expression) => (
    ts.isPropertyAccessExpression(expression) || ts.isElementAccessExpression(expression)
      ? expression.expression
      : null
  );
  const isInjectedReceiver = (expression) => (
    expression !== null
    && (
      (ts.isPropertyAccessExpression(expression)
        && expression.name.text === 'acquisitionInputs')
      || (ts.isElementAccessExpression(expression)
        && expression.argumentExpression
        && staticString(expression.argumentExpression) === 'acquisitionInputs')
    )
  );
  const buildMeasureStateOwner = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement)
    && statement.name?.text === 'buildMeasureState'
  ));
  const enclosingFunction = (node) => {
    let current = node.parent;
    while (current) {
      if (ts.isFunctionDeclaration(current)
        || ts.isFunctionExpression(current)
        || ts.isArrowFunction(current)
        || ts.isMethodDeclaration(current)
        || ts.isGetAccessorDeclaration(current)
        || ts.isSetAccessorDeclaration(current)
        || ts.isConstructorDeclaration(current)) return current;
      current = current.parent;
    }
    return null;
  };
  const visit = (node) => {
    if (ts.isIdentifier(node)
      && node.text === 'bodyAcquisitionInputProjections'
      && !ts.isImportSpecifier(node.parent)
      && enclosingFunction(node) !== buildMeasureStateOwner) {
      fail(
        'RENDERER_ACQUISITION_PROJECTION_BYPASS',
        `${DOCX_SOURCE}/renderer.ts uses bodyAcquisitionInputProjections outside buildMeasureState`,
      );
    }
    if (ts.isCallExpression(node)) {
      if (ts.isIdentifier(node.expression)
        && injectedProjectionMembers.has(node.expression.text)) {
        fail(
          'RENDERER_ACQUISITION_PROJECTION_BYPASS',
          `${DOCX_SOURCE}/renderer.ts calls ${node.expression.text} directly`,
        );
      }
      const member = projectionMember(node.expression);
      if (member && injectedProjectionMembers.has(member)
        && !isInjectedReceiver(projectionReceiver(node.expression))) {
        fail(
          'RENDERER_ACQUISITION_PROJECTION_BYPASS',
          `${DOCX_SOURCE}/renderer.ts calls ${member} outside acquisitionInputs`,
        );
      }
    }
    ts.forEachChild(node, visit);
  };
  ts.forEachChild(source, visit);
}

function assertFloatRectTransportBoundary(root) {
  const path = resolve(root, FLOAT_WRAP);
  if (!existsSync(path)) return;
  const source = sourceFile(path);
  const visit = (node) => {
    if ((ts.isIdentifier(node) || ts.isStringLiteralLike(node))
      && node.text === 'drawn') {
      const parent = node.parent;
      const isStaticPropertyName = (
        (ts.isPropertyAccessExpression(parent) && parent.name === node)
        || (ts.isElementAccessExpression(parent) && parent.argumentExpression === node)
        || (ts.isPropertyAssignment(parent) && parent.name === node)
        || (ts.isShorthandPropertyAssignment(parent) && parent.name === node)
        || (ts.isPropertyDeclaration(parent) && parent.name === node)
        || (ts.isPropertySignature(parent) && parent.name === node)
        || (ts.isMethodDeclaration(parent) && parent.name === node)
        || (ts.isGetAccessorDeclaration(parent) && parent.name === node)
        || (ts.isSetAccessorDeclaration(parent) && parent.name === node)
        || (ts.isBindingElement(parent) && (parent.propertyName ?? parent.name) === node)
      );
      if (isStaticPropertyName) {
        fail('FLOAT_RECT_TRANSITIONAL_STATE', `${FLOAT_WRAP}#drawn`);
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
}

function assertFloatPlacementAuthority(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  const authority = `${LAYOUT_SOURCE}/floats.ts`;
  const kernel = `${LAYOUT_SOURCE}/axis-aligned-overlap.ts`;
  const wrap = `${LAYOUT_SOURCE}/float-wrap.ts`;
  const compatibility = `${LAYOUT_SOURCE}/compatibility.ts`;
  const compatibilityFacade = `${DOCX_SOURCE}/float-layout.ts`;
  const allowedKernelImports = new Map([
    [authority, [
      'AxisAlignedRect',
      'axisAlignedRectsOverlap',
      'resolveAxisAlignedOverlap',
    ]],
    [wrap, ['axisAlignedRectsOverlap']],
  ]);
  const floatCompatibilityName = /^(?:WORD_FLOAT_|WORD_PAGE_ANCHORED_TABLE_|WORD_SQUARE_LINE_|WORD_MIN_LINE_START_PT$|LINE_START_GAP_EPS_PT$|wordMinLineStartPx$)/;
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const source = sourceFile(path);
    const file = posixPath(relative(root, path));
    for (const edge of moduleEdges(path)) {
      if (!edge.literal) continue;
      const dependency = resolveLocalImport(path, edge.specifier);
      const dependencyFile = dependency ? posixPath(relative(root, dependency)) : null;
      if (dependencyFile === kernel) {
        const expected = allowedKernelImports.get(file);
        const actual = [...(edge.importedNames ?? [])].sort();
        if (edge.kind !== 'import'
          || edge.bare
          || edge.aliased
          || !expected
          || JSON.stringify(actual) !== JSON.stringify([...expected].sort())) {
          fail(
            'FLOAT_PLACEMENT_AUTHORITY',
            `${file} -> ${dependencyFile} (${edge.kind}:${actual.join(',')})`,
          );
        }
      }
      if ((edge.importedNames ?? []).some((name) => floatCompatibilityName.test(name))) {
        if ((dependencyFile !== compatibility && dependencyFile !== compatibilityFacade)
          || edge.aliased) {
          fail(
            'FLOAT_COMPATIBILITY_AUTHORITY',
            `${file} -> ${dependencyFile ?? edge.specifier}`,
          );
        }
      }
    }
    const visit = (node) => {
      if ((ts.isIdentifier(node) || ts.isStringLiteralLike(node))
        && node.text === 'resolveAxisAlignedOverlap'
        && file !== kernel
        && file !== authority) {
        fail(
          'FLOAT_PLACEMENT_AUTHORITY',
          `${file} references resolveAxisAlignedOverlap`,
        );
      }
      ts.forEachChild(node, visit);
    };
    ts.forEachChild(source, visit);
    for (const statement of source.statements) {
      if (ts.isVariableStatement(statement)) {
        const exported = statement.modifiers?.some(
          (modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword,
        );
        for (const declaration of statement.declarationList.declarations) {
          if (!ts.isIdentifier(declaration.name)) continue;
          const name = declaration.name.text;
          if (floatCompatibilityName.test(name) && file !== compatibility) {
            fail('FLOAT_COMPATIBILITY_AUTHORITY', `${file} declares ${name}`);
          }
          if (name === 'FLOAT_OVERLAP_EPS' || name === 'FLOAT_PAGE_RIGHT_SLACK') {
            const expected = name === 'FLOAT_OVERLAP_EPS' ? 0.01 : 0.5;
            if (file !== authority
              || !exported
              || !declaration.initializer
              || !ts.isNumericLiteral(declaration.initializer)
              || Number(declaration.initializer.text) !== expected) {
              fail('FLOAT_NUMERIC_POLICY', `${file} declares ${name}`);
            }
          }
          if (file === authority
            && exported
            && declaration.initializer
            && ts.isIdentifier(declaration.initializer)
            && declaration.initializer.text === 'resolveAxisAlignedOverlap') {
            fail('FLOAT_PLACEMENT_AUTHORITY', `${file} re-exports the displacement kernel`);
          }
        }
      }
      if (ts.isExportDeclaration(statement) && statement.exportClause
        && ts.isNamedExports(statement.exportClause)
        && statement.exportClause.elements.some((element) =>
          (element.propertyName?.text ?? element.name.text) === 'resolveAxisAlignedOverlap'
          || element.name.text === 'resolveAxisAlignedOverlap')) {
        fail('FLOAT_PLACEMENT_AUTHORITY', `${file} re-exports the displacement kernel`);
      }
    }
  }
}

function listFiles(root) {
  if (!existsSync(root)) return [];
  const files = [];
  for (const entry of readdirSync(root)) {
    const path = join(root, entry);
    if (statSync(path).isDirectory()) files.push(...listFiles(path));
    else files.push(path);
  }
  return files;
}

function sourceFile(path) {
  const kind = path.endsWith('.tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS;
  return ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true, kind);
}

function importIsTypeOnly(statement) {
  const clause = statement.importClause;
  if (!clause) return false;
  if (clause.isTypeOnly) return true;
  if (clause.name || !clause.namedBindings || !ts.isNamedImports(clause.namedBindings)) return false;
  return clause.namedBindings.elements.length > 0
    && clause.namedBindings.elements.every((element) => element.isTypeOnly);
}

function exportIsTypeOnly(statement) {
  if (statement.isTypeOnly) return true;
  return statement.exportClause
    && ts.isNamedExports(statement.exportClause)
    && statement.exportClause.elements.length > 0
    && statement.exportClause.elements.every((element) => element.isTypeOnly);
}

function moduleEdges(path) {
  const source = sourceFile(path);
  const edges = [];
  for (const statement of source.statements) {
    if (ts.isImportDeclaration(statement)
      && ts.isStringLiteral(statement.moduleSpecifier)) {
      const bindings = statement.importClause?.namedBindings;
      const importedNames = statement.importClause
        ? [
            ...(statement.importClause.name ? ['default'] : []),
            ...(bindings && ts.isNamespaceImport(bindings) ? ['*'] : []),
            ...(bindings && ts.isNamedImports(bindings)
              ? bindings.elements.map((element) => element.propertyName?.text ?? element.name.text)
              : []),
          ]
        : [];
      edges.push({
        kind: 'import',
        specifier: statement.moduleSpecifier.text,
        typeOnly: importIsTypeOnly(statement),
        literal: true,
        importedNames,
        aliased: !!(bindings && ts.isNamedImports(bindings)
          && bindings.elements.some((element) => element.propertyName && element.propertyName.text !== element.name.text)),
        bare: !statement.importClause,
      });
    }
    if (ts.isExportDeclaration(statement)
      && statement.moduleSpecifier
      && ts.isStringLiteral(statement.moduleSpecifier)) {
      const bindings = statement.exportClause && ts.isNamedExports(statement.exportClause)
        ? statement.exportClause.elements
        : [];
      edges.push({
        kind: 'export',
        specifier: statement.moduleSpecifier.text,
        typeOnly: exportIsTypeOnly(statement),
        literal: true,
        importedNames: bindings.length > 0
          ? bindings.map((element) => element.propertyName?.text ?? element.name.text)
          : ['*'],
        aliased: bindings.some((element) => (
          element.propertyName && element.propertyName.text !== element.name.text
        )),
      });
    }
  }
  const visit = (node) => {
    if (ts.isCallExpression(node)
      && node.expression.kind === ts.SyntaxKind.ImportKeyword) {
      const argument = node.arguments[0];
      edges.push(argument && (ts.isStringLiteral(argument) || ts.isNoSubstitutionTemplateLiteral(argument))
        ? { kind: 'dynamic-import', specifier: argument.text, typeOnly: false, literal: true }
        : { kind: 'dynamic-import', specifier: '<dynamic>', typeOnly: false, literal: false });
    }
    if (ts.isCallExpression(node)
      && ts.isIdentifier(node.expression)
      && node.expression.text === 'require') {
      const argument = node.arguments[0];
      edges.push(argument && (ts.isStringLiteral(argument) || ts.isNoSubstitutionTemplateLiteral(argument))
        ? { kind: 'require', specifier: argument.text, typeOnly: false, literal: true }
        : { kind: 'require', specifier: '<dynamic>', typeOnly: false, literal: false });
    }
    ts.forEachChild(node, visit);
  };
  ts.forEachChild(source, visit);
  return edges;
}

function resolveLocalImport(importer, specifier) {
  if (!specifier.startsWith('.')) return null;
  const clean = specifier.split('?')[0].split('#')[0];
  const unresolved = resolve(dirname(importer), clean);
  const withoutJs = unresolved.replace(/\.(mjs|cjs|js)$/, '');
  const candidates = [
    unresolved,
    `${withoutJs}.ts`,
    `${withoutJs}.tsx`,
    join(unresolved, 'index.ts'),
  ];
  return candidates.find((candidate) => existsSync(candidate) && statSync(candidate).isFile()) ?? null;
}

function dependencyGraph(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  const files = listFiles(sourceRoot).filter(isProductionTypeScript);
  const graph = new Map();
  for (const file of files) {
    graph.set(file, moduleEdges(file));
  }
  return graph;
}

function paintBoundaryViolations(root) {
  const graph = dependencyGraph(root);
  const paintRoot = resolve(root, PAINT_SOURCE);
  const layoutTypes = resolve(root, LAYOUT_SOURCE, 'types.ts');
  const layoutAffine = resolve(root, LAYOUT_AFFINE);
  const paintAffine = resolve(root, PAINT_SOURCE, 'affine.ts');
  const entries = [...graph.keys()].filter((path) => path.startsWith(`${paintRoot}${sep}`));
  const violations = [];
  const nonLiteral = [];

  for (const entry of entries) {
    const stack = [{ path: entry, chain: [entry] }];
    const visited = new Set([entry]);
    while (stack.length > 0) {
      const current = stack.pop();
      for (const edge of graph.get(current.path) ?? []) {
        if (!edge.literal) {
          nonLiteral.push(posixPath(relative(root, current.path)));
          continue;
        }
        if (edge.bare) {
          violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          continue;
        }
        if (!edge.specifier.startsWith('.')) {
          const allowedNames = SHARED_PAINT_IMPORTS.get(edge.specifier);
          const allowed = edge.kind === 'import'
            && !edge.aliased
            && allowedNames
            && edge.importedNames?.length > 0
            && edge.importedNames?.every((name) => (
              allowedNames.get(name) === (edge.typeOnly ? 'type' : 'value')
            ));
          if (!allowed) {
            violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          }
          continue;
        }
        const dependency = resolveLocalImport(current.path, edge.specifier);
        if (!dependency) {
          violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          continue;
        }
        const chain = [...current.chain, dependency];
        const insidePaint = dependency.startsWith(`${paintRoot}${sep}`);
        const allowedAffineContract = current.path === paintAffine
          && dependency === layoutAffine
          && edge.kind === 'export'
          && !edge.typeOnly
          && !edge.aliased
          && edge.importedNames?.length === LAYOUT_AFFINE_EXPORTS.size
          && edge.importedNames.every((name) => LAYOUT_AFFINE_EXPORTS.has(name));
        const allowedContract = (edge.typeOnly && dependency === layoutTypes)
          || allowedAffineContract;
        if (!insidePaint && !allowedContract) {
          violations.push(chain.map((path) => posixPath(relative(root, path))));
          continue;
        }
        if (insidePaint && !visited.has(dependency)) {
          visited.add(dependency);
          stack.push({ path: dependency, chain });
        }
      }
    }
  }
  return { violations, nonLiteral };
}

function identifierText(node) {
  if (ts.isIdentifier(node) || ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) return node.text;
  return null;
}

function assertCapabilityBoundaries(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const rel = posixPath(relative(root, path));
    const inPaint = rel.startsWith(`${PAINT_SOURCE}/`);
    const inLayout = rel.startsWith(`${LAYOUT_SOURCE}/`);
    if (!inPaint && !inLayout) continue;
    const source = sourceFile(path);
    const visit = (node) => {
      const text = identifierText(node);
      if (inPaint && text === 'measureText') fail('PAINT_CAPABILITY', `${rel} uses measureText`);
      if (inPaint && text && /^(?:resolve|merge|combine|apply|fold|compose|inherit).*(?:Style|Properties|Pr|Cascade|Format|Formatting)$/i.test(text)) {
        fail('PAINT_CAPABILITY', `${rel} uses ${text}`);
      }
      if (inLayout && text && /^(?:resolve|merge|combine|apply|fold|compose|inherit).*(?:Style|Properties|Pr|Cascade|Format|Formatting)$/i.test(text)) {
        fail('LAYOUT_STYLE_CAPABILITY', `${rel} uses ${text}`);
      }
      if (inLayout && text && /^(?:dpr|displayScale|devicePixelRatio|CanvasRenderingContext2D|OffscreenCanvasRenderingContext2D)$/.test(text)) {
        fail('LAYOUT_DISPLAY_CAPABILITY', `${rel} uses ${text}`);
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
}

function assertOccurrenceProjectionRuntimeDependencies(root) {
  const occurrence = resolve(root, LAYOUT_SOURCE, 'occurrence-projection.ts');
  const translation = resolve(root, LAYOUT_SOURCE, 'retained-geometry-translation.ts');
  const plainData = resolve(root, LAYOUT_SOURCE, 'plain-data.ts');
  // The retained-layout contract checks are development-only, so `plain-data.ts`
  // reads the policy flag at runtime. `validation-policy.ts` is admitted into
  // the sealed projection subgraph rather than exempted from it: it is guarded
  // here with an empty target set, so the seam stays a closed leaf and the
  // policy module cannot become a back door into layout or the parser model.
  const validationPolicy = resolve(root, LAYOUT_SOURCE, 'validation-policy.ts');
  const guarded = [occurrence, translation, plainData, validationPolicy];
  for (const path of guarded) {
    if (!existsSync(path)) {
      fail('OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', `missing ${posixPath(relative(root, path))}`);
    }
  }
  const allowedTargets = new Map([
    [occurrence, new Set([translation, plainData])],
    [translation, new Set()],
    [plainData, new Set([validationPolicy])],
    [validationPolicy, new Set()],
  ]);
  for (const current of guarded) {
    for (const edge of moduleEdges(current)) {
      if (edge.typeOnly) continue;
      const detail = `${posixPath(relative(root, current))} -> ${edge.specifier}`;
      if (!edge.literal || edge.kind === 'dynamic-import' || edge.kind === 'require'
        || edge.bare || edge.specifier.includes('?') || edge.specifier.includes('#')
        || !edge.specifier.startsWith('.')) {
        fail('OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', detail);
      }
      const dependency = resolveLocalImport(current, edge.specifier);
      if (!dependency || !allowedTargets.get(current)?.has(dependency)) {
        fail('OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', detail);
      }
    }
  }
}

function assertCoordinateSpaceRuntimeDependencies(root) {
  const coordinate = resolve(root, LAYOUT_SOURCE, 'coordinate-space.ts');
  const pageFactory = resolve(root, LAYOUT_SOURCE, 'page-factory.ts');
  for (const path of [coordinate, pageFactory]) {
    if (!existsSync(path)) {
      fail('COORDINATE_SPACE_RUNTIME_DEPENDENCY', `missing ${posixPath(relative(root, path))}`);
    }
  }

  for (const edge of moduleEdges(coordinate)) {
    const detail = `${posixPath(relative(root, coordinate))} -> ${edge.specifier}`;
    if (edge.kind !== 'import' || !edge.literal || edge.bare || !edge.typeOnly
      || edge.specifier !== './types.js') {
      fail('COORDINATE_SPACE_RUNTIME_DEPENDENCY', detail);
    }
  }

  const runtimeTargets = new Set([
    resolve(root, LAYOUT_SOURCE, 'coordinate-space.ts'),
    // Page finalization owns retained section decoration geometry so paint
    // cannot reconstruct section bands from layout policy.
    resolve(root, LAYOUT_SOURCE, 'column-separators.ts'),
    resolve(root, LAYOUT_SOURCE, 'page-border.ts'),
    resolve(root, LAYOUT_SOURCE, 'page-graph.ts'),
  ]);
  for (const edge of moduleEdges(pageFactory)) {
    if (edge.typeOnly) {
      if (edge.specifier.includes('?') || edge.specifier.includes('#')) {
        fail(
          'COORDINATE_SPACE_RUNTIME_DEPENDENCY',
          `${posixPath(relative(root, pageFactory))} -> ${edge.specifier}`,
        );
      }
      continue;
    }
    const detail = `${posixPath(relative(root, pageFactory))} -> ${edge.specifier}`;
    if (edge.kind !== 'import' || !edge.literal || edge.bare
      || edge.specifier.includes('?') || edge.specifier.includes('#')
      || !edge.specifier.startsWith('.')) {
      fail('COORDINATE_SPACE_RUNTIME_DEPENDENCY', detail);
    }
    const dependency = resolveLocalImport(pageFactory, edge.specifier);
    if (!dependency || !runtimeTargets.has(dependency)) {
      fail('COORDINATE_SPACE_RUNTIME_DEPENDENCY', detail);
    }
  }
}

function assertAffineRuntimeDependencies(root) {
  const affine = resolve(root, LAYOUT_AFFINE);
  if (!existsSync(affine)) {
    fail('AFFINE_RUNTIME_DEPENDENCY', `missing ${LAYOUT_AFFINE}`);
  }
  for (const edge of moduleEdges(affine)) {
    const detail = `${LAYOUT_AFFINE} -> ${edge.specifier}`;
    if (edge.kind !== 'import' || !edge.literal || edge.bare || !edge.typeOnly
      || edge.specifier !== './types.js') {
      fail('AFFINE_RUNTIME_DEPENDENCY', detail);
    }
  }
}

function assertTextRunProjectionAdapterBoundary(root) {
  const adapter = resolve(root, TEXT_RUN_PROJECTION_ADAPTER);
  if (!existsSync(adapter)) return;
  const declarationCounts = new Map();
  const source = sourceFile(adapter);
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement)) {
      fail('TEXT_RUN_PROJECTION_DECLARATION', statement.getText(source));
    }
    for (const name of declarationNames(statement)) {
      if (!TEXT_RUN_PROJECTION_ADAPTER_DECLARATIONS.has(name)) {
        fail('TEXT_RUN_PROJECTION_DECLARATION', name);
      }
      declarationCounts.set(name, (declarationCounts.get(name) ?? 0) + 1);
      const shouldExport = name !== 'projectTextRun';
      if (hasExportModifier(statement) !== shouldExport) {
        fail('TEXT_RUN_PROJECTION_DECLARATION', name);
      }
    }
  }
  for (const name of TEXT_RUN_PROJECTION_ADAPTER_DECLARATIONS) {
    if (declarationCounts.get(name) !== 1) {
      fail(
        'TEXT_RUN_PROJECTION_DECLARATION',
        `${name} count ${declarationCounts.get(name) ?? 0}`,
      );
    }
  }
  const expected = new Map([
    ['@silurus/ooxml-core', [
      { typeOnly: false, names: ['PT_TO_PX', 'canvasFontString'] },
    ]],
    [`${DOCX_SOURCE}/renderer.ts`, [
      { typeOnly: true, names: ['DocxTextRunInfo'] },
    ]],
    [`${LAYOUT_SOURCE}/affine.ts`, [
      { typeOnly: false, names: ['composeAffine', 'mapAffinePoint', 'scaleAffine'] },
    ]],
    [`${LAYOUT_SOURCE}/document-layout-variants.ts`, [
      { typeOnly: false, names: ['selectDocumentLayoutPage'] },
    ]],
    [`${LAYOUT_SOURCE}/text-index.ts`, [
      { typeOnly: false, names: ['textRunGeometryForPage'] },
      { typeOnly: true, names: ['TextRunGeometry'] },
    ]],
    [`${LAYOUT_SOURCE}/types.ts`, [
      { typeOnly: true, names: ['DocumentLayout', 'LayoutServices', 'Matrix2DData'] },
    ]],
    [`${PAINT_SOURCE}/affine.ts`, [
      { typeOnly: false, names: ['cssTransformFor'] },
    ]],
  ]);
  const actual = new Map();
  for (const edge of moduleEdges(adapter)) {
    if (edge.kind !== 'import' || !edge.literal || edge.bare || edge.aliased
      || edge.specifier.includes('?') || edge.specifier.includes('#')) {
      fail('TEXT_RUN_PROJECTION_IMPORT', `${TEXT_RUN_PROJECTION_ADAPTER} -> ${edge.specifier}`);
    }
    const key = edge.specifier.startsWith('.')
      ? (() => {
          const target = resolveLocalImport(adapter, edge.specifier);
          return target ? posixPath(relative(root, target)) : null;
        })()
      : edge.specifier;
    if (key === null || !expected.has(key)) {
      fail('TEXT_RUN_PROJECTION_IMPORT', `${TEXT_RUN_PROJECTION_ADAPTER} -> ${edge.specifier}`);
    }
    const entries = actual.get(key) ?? [];
    entries.push({
      typeOnly: edge.typeOnly,
      names: [...(edge.importedNames ?? [])].sort(),
    });
    actual.set(key, entries);
  }
  const normalize = (entries) => [...entries]
    .map(({ typeOnly, names }) => `${typeOnly ? 'type' : 'value'}:${[...names].sort().join(',')}`)
    .sort();
  for (const [key, expectedEntries] of expected) {
    const actualEntries = actual.get(key) ?? [];
    if (JSON.stringify(normalize(actualEntries)) !== JSON.stringify(normalize(expectedEntries))) {
      fail(
        'TEXT_RUN_PROJECTION_IMPORT',
        `${TEXT_RUN_PROJECTION_ADAPTER} -> ${key} expected ${normalize(expectedEntries).join('|')}, received ${normalize(actualEntries).join('|')}`,
      );
    }
  }
}

/**
 * Body paint consumes retained placements. Letting this adapter call a layout
 * entry point would silently reintroduce a second layout pass whose result can
 * diverge from pagination (especially for grouped frames). Keep the rule tied
 * to the adapter's AST instead of relying on naming conventions in paint files:
 * renderer.ts intentionally still owns legacy header/footer story layout.
 */
function assertBodyPaintConsumesRetainedLayout(root) {
  const path = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(path)) return;
  const program = ts.createProgram({
    rootNames: [path],
    options: {
      target: ts.ScriptTarget.Latest,
      module: ts.ModuleKind.ESNext,
      noResolve: true,
      skipLibCheck: true,
    },
  });
  const source = program.getSourceFile(path);
  if (!source) return;
  const checker = program.getTypeChecker();
  const declaration = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement)
      && statement.name?.text === 'renderBodyElements'
  ));
  if (!declaration?.body) return;

  const forbidden = new Set([
    'acquireParagraphLayout',
    'acquireRetainedFrameGroup',
    'buildSegments',
    'contextualSpacingAdjust',
    'estimateParagraphHeight',
    'layoutLines',
    'measureParagraph',
    'measureText',
    'paragraphGapAdjustment',
    'paragraphLayoutFromMeasurement',
    'parasShareBorderBox',
    'renderFrameParagraph',
    'renderParagraph',
    'resolveParagraphBorderEdges',
    'resolveFrameBox',
  ]);
  const retainedCanvasMethods = new Set([
    'restore',
    'save',
    'scale',
    'translate',
  ]);
  const isRetainedPropertyBoundary = (target) => (
    (target.name === 'onTextRun' && target.receiver === 'state')
    || (retainedCanvasMethods.has(target.name) && target.receiver === 'state.ctx')
  );
  const retainedImportBoundaries = new Map([
    ['./paint/canvas-drawing.js', new Set(['paintDrawingLayout'])],
    ['./paint/canvas-text.js', new Set([
      'paintParagraphLayout',
      'paintPlacedParagraphLayout',
      'paintPlacedTextBoxLayout',
      'paintTextBoxLayout',
    ])],
    ['./vertical-text.js', new Set(['verticalTextLayerPlacement'])],
  ]);

  const unwrapExpression = (expression) => {
    let current = expression;
    while (ts.isParenthesizedExpression(current)
      || ts.isAsExpression(current)
      || ts.isTypeAssertionExpression(current)
      || ts.isNonNullExpression(current)
      || ts.isSatisfiesExpression(current)) {
      current = current.expression;
    }
    return current;
  };

  const isUnshadowedGlobalIdentifier = (node, name) => {
    if (!ts.isIdentifier(node) || node.text !== name) return false;
    const symbol = checker.getSymbolAtLocation(node);
    const declarations = symbol?.declarations ?? [];
    return declarations.length > 0
      && declarations.every((item) => item.getSourceFile().isDeclarationFile);
  };

  const isCanonicalWeakMapConstruction = (expression) => {
    const value = unwrapExpression(expression);
    return ts.isNewExpression(value)
      && isUnshadowedGlobalIdentifier(value.expression, 'WeakMap')
      && (value.arguments?.length ?? 0) === 0;
  };

  const isGlobalObjectCall = (expression, method) => (
    ts.isCallExpression(expression)
    && ts.isPropertyAccessExpression(expression.expression)
    && expression.expression.name.text === method
    && isUnshadowedGlobalIdentifier(expression.expression.expression, 'Object')
  );

  const isCanonicalBodyFragmentMapInitializer = (expression) => {
    const frozen = unwrapExpression(expression);
    if (!isGlobalObjectCall(frozen, 'freeze') || frozen.arguments.length !== 1) return false;
    const assigned = unwrapExpression(frozen.arguments[0]);
    if (!isGlobalObjectCall(assigned, 'assign')
      || assigned.arguments.length !== 2
      || !isCanonicalWeakMapConstruction(assigned.arguments[0])) return false;
    const sidecars = unwrapExpression(assigned.arguments[1]);
    if (!ts.isObjectLiteralExpression(sidecars) || sidecars.properties.length !== 2) return false;
    const expectedSidecars = new Set(['framePlacement', 'sourceIndices']);
    const seen = new Set();
    for (const property of sidecars.properties) {
      if (!ts.isPropertyAssignment(property) || !ts.isIdentifier(property.name)) return false;
      const name = property.name.text;
      if (!expectedSidecars.has(name)
        || seen.has(name)
        || !isCanonicalWeakMapConstruction(property.initializer)) return false;
      seen.add(name);
    }
    return seen.size === expectedSidecars.size;
  };

  const isCanonicalBodyFragmentReceiver = (identifier) => {
    const symbol = checker.getSymbolAtLocation(identifier);
    const declarations = symbol?.declarations ?? [];
    const declaration = declarations.length === 1 ? declarations[0] : undefined;
    return declaration !== undefined
      && ts.isVariableDeclaration(declaration)
      && ts.isVariableDeclarationList(declaration.parent)
      && (declaration.parent.flags & ts.NodeFlags.Const) !== 0
      && ts.isVariableStatement(declaration.parent.parent)
      && declaration.parent.parent.parent === source
      && declaration.initializer !== undefined
      && isCanonicalBodyFragmentMapInitializer(declaration.initializer);
  };

  const isExactBodyFragmentLookup = (item) => {
    if (item.name?.text !== 'bodyFragmentFor'
      || item.parameters.length !== 1
      || !ts.isIdentifier(item.parameters[0].name)
      || item.body?.statements.length !== 1) return false;
    const statement = item.body.statements[0];
    if (!ts.isReturnStatement(statement) || !statement.expression) return false;
    const call = unwrapExpression(statement.expression);
    if (!ts.isCallExpression(call)
      || !ts.isPropertyAccessExpression(call.expression)
      || !ts.isIdentifier(call.expression.expression)
      || call.expression.expression.text !== 'bodyFlowFragments'
      || !isCanonicalBodyFragmentReceiver(call.expression.expression)
      || call.expression.name.text !== 'get'
      || call.arguments.length !== 1) return false;
    const argument = unwrapExpression(call.arguments[0]);
    return ts.isIdentifier(argument)
      && checker.getSymbolAtLocation(argument)
        === checker.getSymbolAtLocation(item.parameters[0].name);
  };

  const staticString = (expression, resolving = new Set()) => {
    const value = unwrapExpression(expression);
    if (ts.isStringLiteralLike(value)) return value.text;
    if (ts.isBinaryExpression(value)
      && value.operatorToken.kind === ts.SyntaxKind.PlusToken) {
      const left = staticString(value.left, resolving);
      const right = staticString(value.right, resolving);
      return left === null || right === null ? null : left + right;
    }
    if (!ts.isIdentifier(value)) return null;
    const symbol = checker.getSymbolAtLocation(value);
    if (!symbol || resolving.has(symbol)) return null;
    resolving.add(symbol);
    const declarations = symbol.declarations ?? [];
    const values = declarations.flatMap((item) => (
      ts.isVariableDeclaration(item) && item.initializer
        ? [staticString(item.initializer, resolving)]
        : []
    ));
    resolving.delete(symbol);
    return values.length > 0 && values.every((item) => item === values[0])
      ? values[0]
      : null;
  };

  const propertyNameText = (name) => {
    if (ts.isIdentifier(name) || ts.isStringLiteralLike(name)) return name.text;
    if (ts.isComputedPropertyName(name)) return staticString(name.expression);
    return null;
  };

  const canonicalBodyFragmentSidecarPath = (expression) => {
    const path = [];
    let current = unwrapExpression(expression);
    while (ts.isPropertyAccessExpression(current) || ts.isElementAccessExpression(current)) {
      const name = ts.isPropertyAccessExpression(current)
        ? current.name.text
        : current.argumentExpression
          ? staticString(current.argumentExpression)
          : null;
      if (name === null) return null;
      path.unshift(name);
      current = unwrapExpression(current.expression);
    }
    return ts.isIdentifier(current)
      && current.text === 'bodyFlowFragments'
      && isCanonicalBodyFragmentReceiver(current)
      && (path[0] === 'sourceIndices' || path[0] === 'framePlacement')
      ? path
      : null;
  };
  const assignmentOperators = new Set([
    ts.SyntaxKind.EqualsToken,
    ts.SyntaxKind.PlusEqualsToken,
    ts.SyntaxKind.MinusEqualsToken,
    ts.SyntaxKind.AsteriskEqualsToken,
    ts.SyntaxKind.AsteriskAsteriskEqualsToken,
    ts.SyntaxKind.SlashEqualsToken,
    ts.SyntaxKind.PercentEqualsToken,
    ts.SyntaxKind.LessThanLessThanEqualsToken,
    ts.SyntaxKind.GreaterThanGreaterThanEqualsToken,
    ts.SyntaxKind.GreaterThanGreaterThanGreaterThanEqualsToken,
    ts.SyntaxKind.AmpersandEqualsToken,
    ts.SyntaxKind.BarEqualsToken,
    ts.SyntaxKind.CaretEqualsToken,
    ts.SyntaxKind.BarBarEqualsToken,
    ts.SyntaxKind.AmpersandAmpersandEqualsToken,
    ts.SyntaxKind.QuestionQuestionEqualsToken,
  ]);
  const lateSidecarMutations = [];
  const findLateSidecarMutations = (current) => {
    if (ts.isCallExpression(current)
      && (isGlobalObjectCall(current, 'assign') || isGlobalObjectCall(current, 'defineProperty'))
      && current.arguments[0]
      && canonicalBodyFragmentSidecarPath(current.arguments[0])) {
      lateSidecarMutations.push(current.getText(source));
    } else if (ts.isBinaryExpression(current)
      && assignmentOperators.has(current.operatorToken.kind)
      && canonicalBodyFragmentSidecarPath(current.left)) {
      lateSidecarMutations.push(current.getText(source));
    }
    ts.forEachChild(current, findLateSidecarMutations);
  };
  findLateSidecarMutations(source);
  if (lateSidecarMutations.length > 0) {
    fail(
      'BODY_PAINT_LAYOUT_CAPABILITY',
      `${DOCX_SOURCE}/renderer.ts mutates canonical body fragment sidecar after initialization`,
    );
  }

  let resolveCallTarget;
  const resolveObjectProperty = (expression, propertyName, resolving) => {
    const value = unwrapExpression(expression);
    if (ts.isObjectLiteralExpression(value)) {
      const property = value.properties.find((item) => (
        'name' in item && item.name && propertyNameText(item.name) === propertyName
      ));
      if (property && ts.isPropertyAssignment(property)) {
        return resolveCallTarget(property.initializer, resolving);
      }
      if (property && ts.isShorthandPropertyAssignment(property)) {
        return resolveCallTarget(property.name, resolving);
      }
      if (property && ts.isMethodDeclaration(property) && property.body) {
        return [{ kind: 'body', body: property.body, detail: propertyName }];
      }
      return null;
    }
    if (ts.isConditionalExpression(value)) {
      const whenTrue = resolveObjectProperty(value.whenTrue, propertyName, resolving);
      const whenFalse = resolveObjectProperty(value.whenFalse, propertyName, resolving);
      return whenTrue && whenFalse ? [...whenTrue, ...whenFalse] : null;
    }
    if (!ts.isIdentifier(value)) return null;
    const symbol = checker.getSymbolAtLocation(value);
    if (!symbol || resolving.has(symbol)) return null;
    resolving.add(symbol);
    const targets = (symbol.declarations ?? []).flatMap((item) => (
      ts.isVariableDeclaration(item) && item.initializer
        ? resolveObjectProperty(item.initializer, propertyName, resolving) ?? []
        : []
    ));
    resolving.delete(symbol);
    return targets.length > 0 ? targets : null;
  };

  resolveCallTarget = (expression, resolving = new Set()) => {
    const value = unwrapExpression(expression);
    if (ts.isPropertyAccessExpression(value)) {
      return resolveObjectProperty(value.expression, value.name.text, resolving)
        ?? [{ kind: 'property', name: value.name.text, receiver: value.expression.getText(source) }];
    }
    if (ts.isElementAccessExpression(value)) {
      const name = value.argumentExpression && staticString(value.argumentExpression);
      return name === null || name === undefined
        ? [{ kind: 'unresolved', detail: value.getText(source) }]
        : resolveObjectProperty(value.expression, name, resolving)
          ?? [{ kind: 'property', name, receiver: value.expression.getText(source) }];
    }
    if (ts.isConditionalExpression(value)) {
      return [
        ...resolveCallTarget(value.whenTrue, resolving),
        ...resolveCallTarget(value.whenFalse, resolving),
      ];
    }
    if (!ts.isIdentifier(value)) {
      return [{ kind: 'unresolved', detail: value.getText(source) }];
    }
    const symbol = checker.getSymbolAtLocation(value);
    if (!symbol) return [{ kind: 'name', name: value.text }];
    if (resolving.has(symbol)) {
      return [{ kind: 'unresolved', detail: value.text }];
    }
    resolving.add(symbol);
    const targets = [];
    for (const item of symbol.declarations ?? []) {
      if (ts.isFunctionDeclaration(item) && item.body) {
        const name = item.name?.text;
        targets.push(isExactBodyFragmentLookup(item)
          ? { kind: 'local-boundary', name: 'bodyFragmentFor' }
          : { kind: 'body', body: item.body, detail: name ?? value.text });
      } else if (ts.isVariableDeclaration(item) && item.initializer) {
        const initializer = unwrapExpression(item.initializer);
        if (ts.isArrowFunction(initializer) || ts.isFunctionExpression(initializer)) {
          targets.push({ kind: 'body', body: initializer.body, detail: value.text });
        } else {
          targets.push(...resolveCallTarget(initializer, resolving));
        }
      } else if (ts.isImportSpecifier(item)) {
        let parent = item.parent;
        while (parent && !ts.isImportDeclaration(parent)) parent = parent.parent;
        targets.push({
          kind: 'import',
          name: item.propertyName?.text ?? item.name.text,
          specifier: parent && ts.isStringLiteral(parent.moduleSpecifier)
            ? parent.moduleSpecifier.text
            : null,
        });
      } else if (ts.isBindingElement(item)) {
        const name = !item.propertyName
          ? ts.isIdentifier(item.name) ? item.name.text : null
          : propertyNameText(item.propertyName);
        const variable = ts.isObjectBindingPattern(item.parent)
          && ts.isVariableDeclaration(item.parent.parent)
          ? item.parent.parent
          : null;
        const propertyTargets = name !== null && variable?.initializer
          ? resolveObjectProperty(variable.initializer, name, resolving)
          : null;
        targets.push(...(propertyTargets ?? [name === null
          ? { kind: 'unresolved', detail: item.getText(source) }
          : {
              kind: 'property',
              name,
              receiver: variable?.initializer?.getText(source) ?? '<destructured>',
            }]));
      } else if (ts.isParameter(item)) {
        targets.push({ kind: 'unresolved', detail: value.text });
      }
    }
    resolving.delete(symbol);
    return targets.length > 0
      ? targets
      : [{ kind: 'unresolved', detail: value.text }];
  };

  const directCallTargets = (node) => {
    const targets = [];
    const visit = (current) => {
      const isNamedLocalCallable = current !== node && (
        ts.isFunctionDeclaration(current)
        || ((ts.isFunctionExpression(current) || ts.isArrowFunction(current))
          && ts.isVariableDeclaration(current.parent)
          && ts.isIdentifier(current.parent.name))
      );
      if (isNamedLocalCallable) return;
      if (ts.isCallExpression(current)) {
        targets.push(...resolveCallTarget(current.expression));
      }
      ts.forEachChild(current, visit);
    };
    visit(node);
    return targets;
  };
  const paragraphBranchOf = (statement) => {
    if (!ts.isBinaryExpression(statement.expression)) return null;
    const { left, operatorToken, right } = statement.expression;
    const isTypeAccess = (node) => ts.isPropertyAccessExpression(node)
      && node.name.text === 'type';
    const isParagraph = (node) => ts.isStringLiteralLike(node)
      && node.text === 'paragraph';
    if (!((isTypeAccess(left) && isParagraph(right))
      || (isParagraph(left) && isTypeAccess(right)))) return null;
    if (operatorToken.kind === ts.SyntaxKind.EqualsEqualsEqualsToken
      || operatorToken.kind === ts.SyntaxKind.EqualsEqualsToken) {
      return statement.thenStatement;
    }
    if (operatorToken.kind === ts.SyntaxKind.ExclamationEqualsEqualsToken
      || operatorToken.kind === ts.SyntaxKind.ExclamationEqualsToken) {
      return statement.elseStatement ?? null;
    }
    return null;
  };

  const entryCalls = new Set();
  let foundParagraphBranch = false;
  const findParagraphBranches = (node) => {
    if (ts.isIfStatement(node)) {
      const branch = paragraphBranchOf(node);
      if (branch) {
        foundParagraphBranch = true;
        for (const target of directCallTargets(branch)) entryCalls.add(target);
        return;
      }
    }
    ts.forEachChild(node, findParagraphBranches);
  };
  findParagraphBranches(declaration.body);
  if (!foundParagraphBranch) {
    fail(
      'BODY_PAINT_LAYOUT_CAPABILITY',
      `${DOCX_SOURCE}/renderer.ts#renderBodyElements has no statically auditable paragraph branch`,
    );
  }

  const violations = new Set();
  const visitedBodies = new Set();
  const pending = [...entryCalls];
  while (pending.length > 0) {
    const target = pending.pop();
    if (!target) continue;
    if (target.kind === 'unresolved') {
      violations.add(`unresolved call ${target.detail}`);
      continue;
    }
    if (target.kind === 'local-boundary') continue;
    if (target.kind === 'import') {
      if (retainedImportBoundaries.get(target.specifier)?.has(target.name)) continue;
      violations.add(forbidden.has(target.name)
        ? target.name
        : `unresolved call ${target.name} from ${target.specifier ?? '<unknown import>'}`);
      continue;
    }
    if (target.kind === 'property') {
      if (isRetainedPropertyBoundary(target)) continue;
      violations.add(forbidden.has(target.name)
        ? target.name
        : `unresolved call ${target.name}`);
      continue;
    }
    if (target.kind === 'name') {
      violations.add(forbidden.has(target.name)
        ? target.name
        : `unresolved call ${target.name}`);
      continue;
    }
    if (visitedBodies.has(target.body)) continue;
    visitedBodies.add(target.body);
    for (const called of directCallTargets(target.body)) pending.push(called);
  }

  if (violations.size > 0) {
    fail(
      'BODY_PAINT_LAYOUT_CAPABILITY',
      `${DOCX_SOURCE}/renderer.ts#renderBodyElements reaches ${[...violations].sort().join(', ')}`,
    );
  }
}

function isExactLayoutParserModelGatewayImportEdge(currentRel, edge, dependency, parserModel) {
  return currentRel === LAYOUT_PARSER_MODEL_GATEWAY
    && dependency === parserModel
    && edge.kind === 'import'
    && edge.specifier === LAYOUT_PARSER_MODEL_GATEWAY_IMPORT
    && edge.typeOnly === false
    && edge.aliased === false
    && edge.bare === false
    && edge.importedNames?.length === 1
    && edge.importedNames[0] === LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL;
}

function hasExactLayoutParserModelGatewayProjection(path) {
  const source = sourceFile(path);
  let bindingReferences = 0;
  const countBindingReferences = (node) => {
    if (ts.isIdentifier(node) && node.text === LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL) {
      bindingReferences += 1;
    }
    ts.forEachChild(node, countBindingReferences);
  };
  countBindingReferences(source);
  if (bindingReferences !== 3) return false;

  const projections = source.statements.filter((statement) => (
    ts.isFunctionDeclaration(statement)
    && statement.name?.text === 'documentMathOccurrences'
  ));
  if (projections.length !== 1) return false;
  const projection = projections[0];
  const exported = projection.modifiers?.some((modifier) => (
    modifier.kind === ts.SyntaxKind.ExportKeyword
  ));
  if (!exported || !projection.body || projection.parameters.length !== 1) return false;
  const parameter = projection.parameters[0];
  if (!ts.isIdentifier(parameter.name) || projection.body.statements.length !== 1) return false;
  const returned = projection.body.statements[0];
  if (!ts.isReturnStatement(returned)
    || !returned.expression
    || !ts.isArrayLiteralExpression(returned.expression)
    || returned.expression.elements.length !== 1) return false;
  const spread = returned.expression.elements[0];
  if (!ts.isSpreadElement(spread)
    || !ts.isPropertyAccessExpression(spread.expression)
    || spread.expression.name.text !== 'mathOccurrences') return false;
  const call = spread.expression.expression;
  const exactMathProjection = ts.isCallExpression(call)
    && ts.isIdentifier(call.expression)
    && call.expression.text === LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL
    && call.arguments.length === 1
    && ts.isIdentifier(call.arguments[0])
    && call.arguments[0].text === parameter.name.text;
  if (!exactMathProjection) return false;

  const productionInputs = source.statements.filter((statement) => (
    ts.isFunctionDeclaration(statement)
    && statement.name?.text === 'productionDocumentInput'
  ));
  if (productionInputs.length !== 1) return false;
  const productionInput = productionInputs[0];
  const productionExported = productionInput.modifiers?.some((modifier) => (
    modifier.kind === ts.SyntaxKind.ExportKeyword
  ));
  if (!productionExported || !productionInput.body || productionInput.parameters.length !== 1
    || productionInput.body.statements.length !== 1) return false;
  const productionParameter = productionInput.parameters[0];
  const productionReturn = productionInput.body.statements[0];
  if (!ts.isIdentifier(productionParameter.name)
    || !ts.isReturnStatement(productionReturn)
    || !productionReturn.expression
    || !ts.isCallExpression(productionReturn.expression)
    || !ts.isIdentifier(productionReturn.expression.expression)
    || productionReturn.expression.expression.text !== LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL
    || productionReturn.expression.arguments.length !== 1
    || !ts.isIdentifier(productionReturn.expression.arguments[0])) return false;
  return productionReturn.expression.arguments[0].text === productionParameter.name.text;
}

function layoutParserModelBoundaryViolations(root) {
  const graph = dependencyGraph(root);
  const parserModel = resolve(root, PARSER_MODEL);
  const entries = [...graph.keys()].filter((path) => (
    posixPath(relative(root, path)).startsWith(`${LAYOUT_SOURCE}/`)
  ));
  const violations = [];
  const nonLiteral = [];

  for (const entry of entries) {
    const stack = [{ path: entry, chain: [entry] }];
    const visited = new Set([entry]);
    while (stack.length > 0) {
      const current = stack.pop();
      const currentRel = posixPath(relative(root, current.path));
      for (const edge of graph.get(current.path) ?? []) {
        if (!edge.literal) {
          nonLiteral.push(current.chain.map((path) => posixPath(relative(root, path))));
          continue;
        }
        if (!edge.specifier.startsWith('.')) continue;
        const dependency = resolveLocalImport(current.path, edge.specifier);
        if (!dependency) continue;
        const chain = [...current.chain, dependency];
        if (dependency === parserModel) {
          // The parser-model gateway permits exactly one projection edge. Only
          // that edge is terminal; every other resources.ts dependency remains
          // part of the transitive runtime graph and is inspected normally.
          if (isExactLayoutParserModelGatewayImportEdge(currentRel, edge, dependency, parserModel)) {
            if (hasExactLayoutParserModelGatewayProjection(current.path)) continue;
            violations.push([
              ...chain.map((path) => posixPath(relative(root, path))),
              `invalid use of ${LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL}`,
            ]);
            continue;
          }
          violations.push(chain.map((path) => posixPath(relative(root, path))));
          continue;
        }
        // A type-only edge is erased and cannot create a runtime parser-model
        // dependency through the referenced contract. A direct parser-model
        // type import was rejected above so layout stays parser-model-free.
        if (edge.typeOnly) continue;
        if (graph.has(dependency) && !visited.has(dependency)) {
          visited.add(dependency);
          stack.push({ path: dependency, chain });
        }
      }
    }
  }
  return { violations, nonLiteral };
}

function assertLayoutParserModelBoundaries(root) {
  const { violations, nonLiteral } = layoutParserModelBoundaryViolations(root);
  if (nonLiteral.length > 0) {
    fail(
      'NON_LITERAL_LAYOUT_MODULE_EDGE',
      nonLiteral.map((chain) => chain.join(' -> ')).join('\n'),
    );
  }
  if (violations.length > 0) {
    fail(
      'LAYOUT_PARSER_MODEL_DEPENDENCY',
      violations.map((chain) => chain.join(' -> ')).join('\n'),
    );
  }
}

function assertBodyLayoutAdapterBoundary(root) {
  const adapter = resolve(root, BODY_LAYOUT_ADAPTER);
  if (!existsSync(adapter)) return;
  const source = sourceFile(adapter);
  const seenImports = new Map();
  const declarations = [];
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement) || ts.isExportAssignment(statement)) {
      fail('BODY_LAYOUT_ADAPTER_EXPORT', statement.getText(source));
    }
    if (statement.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.DefaultKeyword)) {
      fail('BODY_LAYOUT_ADAPTER_EXPORT', statement.getText(source));
    }
    if (ts.isImportDeclaration(statement)) {
      if (!ts.isStringLiteralLike(statement.moduleSpecifier)) {
        fail('BODY_LAYOUT_ADAPTER_IMPORT', '<dynamic>');
      }
      const dependency = resolveLocalImport(adapter, statement.moduleSpecifier.text);
      const dependencyRelative = dependency ? posixPath(relative(root, dependency)) : null;
      const reviewedBindings = dependencyRelative
        ? BODY_LAYOUT_ADAPTER_IMPORT_BINDINGS.get(dependencyRelative)
        : undefined;
      const clause = statement.importClause;
      if (!reviewedBindings || !clause || clause.name || !clause.namedBindings
        || !ts.isNamedImports(clause.namedBindings)) {
        fail('BODY_LAYOUT_ADAPTER_IMPORT', statement.getText(source));
      }
      if (seenImports.has(dependencyRelative)) {
        fail('BODY_LAYOUT_ADAPTER_IMPORT', `duplicate:${dependencyRelative}`);
      }
      const actualBindings = new Map();
      for (const element of clause.namedBindings.elements) {
        const importedName = element.propertyName?.text ?? element.name.text;
        const kind = clause.isTypeOnly || element.isTypeOnly ? 'type' : 'value';
        if (element.name.text !== importedName || reviewedBindings.get(importedName) !== kind) {
          fail('BODY_LAYOUT_ADAPTER_BINDING', `${dependencyRelative}#${importedName}:${kind}`);
        }
        actualBindings.set(importedName, kind);
      }
      if (actualBindings.size !== reviewedBindings.size
        || [...reviewedBindings].some(([name, kind]) => actualBindings.get(name) !== kind)) {
        fail('BODY_LAYOUT_ADAPTER_IMPORT', `incomplete:${dependencyRelative}`);
      }
      seenImports.set(dependencyRelative, actualBindings);
    }
    for (const name of declarationNames(statement)) {
      declarations.push(name);
      if (!BODY_LAYOUT_ADAPTER_DECLARATIONS.has(name)) {
        fail('BODY_LAYOUT_ADAPTER_DECLARATION', name);
      }
    }
  }
  if (declarations.length !== 1 || declarations[0] !== 'createBodyLayoutInput') {
    fail('BODY_LAYOUT_ADAPTER_DECLARATION', declarations.join(','));
  }
  if (seenImports.size !== BODY_LAYOUT_ADAPTER_IMPORT_BINDINGS.size
    || [...BODY_LAYOUT_ADAPTER_IMPORT_BINDINGS.keys()].some((path) => !seenImports.has(path))) {
    fail('BODY_LAYOUT_ADAPTER_IMPORT', 'exact-import-set-required');
  }
  const declaration = source.statements.find(ts.isFunctionDeclaration);
  const returned = declaration?.body?.statements.length === 1
    && ts.isReturnStatement(declaration.body.statements[0])
    ? declaration.body.statements[0].expression
    : null;
  const isAcquisitionCall = returned && ts.isCallExpression(returned)
    && ts.isIdentifier(returned.expression)
    && returned.expression.text === 'projectBodyLayoutInput'
    && returned.arguments.length === 1
    && ts.isCallExpression(returned.arguments[0])
    && ts.isIdentifier(returned.arguments[0].expression)
    && returned.arguments[0].expression.text === 'bodyLayoutAcquisitionInput'
    && returned.arguments[0].arguments.length === 1
    && ts.isIdentifier(returned.arguments[0].arguments[0])
    && returned.arguments[0].arguments[0].text === 'document';
  if (!declaration
    || declaration.name?.text !== 'createBodyLayoutInput'
    || declaration.parameters.length !== 1
    || !ts.isIdentifier(declaration.parameters[0].name)
    || declaration.parameters[0].name.text !== 'document'
    || !isAcquisitionCall) {
    fail('BODY_LAYOUT_ADAPTER_BODY', declaration?.getText(source) ?? '<missing>');
  }
}

function assertParagraphAnchorFrameAdapterBoundary(root) {
  const adapter = resolve(root, PARAGRAPH_ANCHOR_FRAME_ADAPTER);
  if (!existsSync(adapter)) return;
  fail('FINAL_PARAGRAPH_ANCHOR_ADAPTER', PARAGRAPH_ANCHOR_FRAME_ADAPTER);
}

function assertBodyKernelServiceOwner(root) {
  const runtime = resolve(root, LAYOUT_RUNTIME_ADAPTER);
  if (!existsSync(runtime)) {
    fail('BODY_KERNEL_SERVICE_OWNER', `${LAYOUT_RUNTIME_ADAPTER} missing`);
  }
  const source = sourceFile(runtime);
  const owner = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement) && statement.name?.text === 'createLayoutServices'
  ));
  const implementation = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement) && statement.name?.text === 'createConcreteBodyLayoutKernel'
  ));
  if (!owner?.body || !implementation?.body) {
    fail('BODY_KERNEL_SERVICE_OWNER', 'missing owner or implementation');
  }
  const calls = [];
  const identifiers = [];
  const visit = (node) => {
    if (ts.isIdentifier(node) && node.text === 'createConcreteBodyLayoutKernel') identifiers.push(node);
    if (ts.isCallExpression(node)
      && ts.isIdentifier(node.expression)
      && node.expression.text === 'createConcreteBodyLayoutKernel') calls.push(node);
    ts.forEachChild(node, visit);
  };
  visit(source);
  const call = calls[0];
  const parentCall = call?.parent;
  const owned = call
    && parentCall
    && ts.isCallExpression(parentCall)
    && ts.isIdentifier(parentCall.expression)
    && parentCall.expression.text === 'attachBodyLayoutKernel'
    && parentCall.arguments.length === 2
    && ts.isIdentifier(parentCall.arguments[0])
    && parentCall.arguments[0].text === 'services'
    && parentCall.arguments[1] === call
    && call.arguments.length === 3
    && call.arguments.every((argument, index) => (
      ts.isIdentifier(argument) && argument.text === ['source', 'context', 'localMetrics'][index]
    ));
  let insideOwner = false;
  for (let node = parentCall; node; node = node.parent) {
    if (node === owner) insideOwner = true;
  }
  if (calls.length !== 1 || identifiers.length !== 2 || !owned || !insideOwner) {
    fail('BODY_KERNEL_SERVICE_OWNER', `calls:${calls.length};identifiers:${identifiers.length}`);
  }
}

const MIGRATION_IDENTIFIER = /(?:legacy|(?:use|enable|prefer|require)[a-z0-9]*(?:old|previous|alternate)[a-z0-9]*(?:engine|layout|path|algorithm)|(?:reuse|paint)enabled|requireslegacy|dryrun)/i;

function matchingIdentifierCounts(root, predicate) {
  const counts = {};
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const source = sourceFile(path);
    const file = posixPath(relative(root, path));
    const visit = (node) => {
      if (ts.isIdentifier(node) && predicate(node.text)) {
        const key = `${file}#${node.text}`;
        counts[key] = (counts[key] ?? 0) + 1;
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
  return Object.fromEntries(Object.entries(counts).sort(([left], [right]) => left.localeCompare(right)));
}

function assertNoDeletedPageProducerIdentifiers(root) {
  const matches = {};
  const sourceRoot = resolve(root, DOCX_SOURCE);
  const record = (file, name) => {
    const key = `${file}#${name}`;
    matches[key] = (matches[key] ?? 0) + 1;
  };
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const source = sourceFile(path);
    const file = posixPath(relative(root, path));
    const visit = (node) => {
      if (ts.isIdentifier(node) && DELETED_PAGE_PRODUCER_IDENTIFIERS.has(node.text)) {
        record(file, node.text);
      } else if ((ts.isIdentifier(node) || ts.isStringLiteralLike(node))
        && DELETED_LEGACY_STAMP_PROPERTIES.has(node.text)) {
        const parent = node.parent;
        const isStaticPropertyName = (
          (ts.isPropertyAccessExpression(parent) && parent.name === node)
          || (ts.isElementAccessExpression(parent) && parent.argumentExpression === node)
          || (ts.isPropertyAssignment(parent) && parent.name === node)
          || (ts.isShorthandPropertyAssignment(parent) && parent.name === node)
          || (ts.isPropertyDeclaration(parent) && parent.name === node)
          || (ts.isPropertySignature(parent) && parent.name === node)
          || (ts.isMethodDeclaration(parent) && parent.name === node)
          || (ts.isGetAccessorDeclaration(parent) && parent.name === node)
          || (ts.isSetAccessorDeclaration(parent) && parent.name === node)
          || (ts.isBindingElement(parent) && (parent.propertyName ?? parent.name) === node)
        );
        if (isStaticPropertyName) record(file, node.text);
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
  if (Object.keys(matches).length > 0) {
    fail('FORBIDDEN_PAGE_PRODUCER_IDENTIFIER', stableJson(matches).trim());
  }
}

function callsNamed(node, name) {
  const calls = [];
  const visit = (current) => {
    if (ts.isCallExpression(current)
      && ts.isIdentifier(current.expression)
      && current.expression.text === name) calls.push(current);
    ts.forEachChild(current, visit);
  };
  visit(node);
  return calls;
}

function unwrapStaticExpression(expression) {
  let current = expression;
  while (ts.isParenthesizedExpression(current)
    || ts.isAsExpression(current)
    || ts.isTypeAssertionExpression(current)
    || ts.isNonNullExpression(current)
    || ts.isSatisfiesExpression(current)) {
    current = current.expression;
  }
  return current;
}

function callOf(expression, name) {
  const value = expression && unwrapStaticExpression(expression);
  return value
    && ts.isCallExpression(value)
    && ts.isIdentifier(value.expression)
    && value.expression.text === name
    ? value
    : null;
}

function hasExactInvariantImports(source) {
  return source.statements.some((statement) => {
    if (!ts.isImportDeclaration(statement)
      || !ts.isStringLiteral(statement.moduleSpecifier)
      || statement.moduleSpecifier.text !== './invariants.js'
      || !statement.importClause
      || statement.importClause.isTypeOnly
      || statement.importClause.name
      || !statement.importClause.namedBindings
      || !ts.isNamedImports(statement.importClause.namedBindings)) return false;
    const elements = statement.importClause.namedBindings.elements;
    return elements.length === 1
      && elements.every((element) => !element.isTypeOnly && !element.propertyName)
      && elements[0]?.name.text === 'assertAndDeepFreezeDocumentLayout';
  });
}

function isCanonicalSelectedLayoutFunction(statement, source) {
  if (!ts.isFunctionDeclaration(statement)
    || statement.name?.text !== 'selectDocumentLayoutPage'
    || !hasExportModifier(statement)
    || !statement.body
    || statement.parameters.length !== 3
    || statement.parameters.some((parameter) => !ts.isIdentifier(parameter.name))
    || statement.body.statements.length !== 3) return false;
  const [services, input, pageIndex] = statement.parameters.map((parameter) => parameter.name.text);
  const [storeDeclaration, missingStore, selectedPage] = statement.body.statements;
  if (!ts.isVariableStatement(storeDeclaration)
    || (storeDeclaration.declarationList.flags & ts.NodeFlags.Const) === 0
    || storeDeclaration.declarationList.declarations.length !== 1) return false;
  const [store] = storeDeclaration.declarationList.declarations;
  const storeCall = callOf(store.initializer, 'layoutVariantStoreOf');
  if (!ts.isIdentifier(store.name)
    || !storeCall
    || storeCall.arguments.length !== 1
    || storeCall.arguments[0].getText(source) !== services) return false;
  const storeName = store.name.text;
  if (!ts.isIfStatement(missingStore)
    || missingStore.elseStatement
    || !ts.isPrefixUnaryExpression(missingStore.expression)
    || missingStore.expression.operator !== ts.SyntaxKind.ExclamationToken
    || missingStore.expression.operand.getText(source) !== storeName
    || !ts.isThrowStatement(missingStore.thenStatement)) return false;
  if (!ts.isReturnStatement(selectedPage) || !selectedPage.expression) return false;
  const selection = unwrapStaticExpression(selectedPage.expression);
  if (!ts.isCallExpression(selection)
    || !ts.isPropertyAccessExpression(selection.expression)
    || selection.expression.expression.getText(source) !== storeName
    || selection.expression.name.text !== 'selectPage'
    || selection.arguments.length !== 2) return false;
  const normalized = callOf(selection.arguments[0], 'layoutOptionsForRender');
  return normalized !== null
    && normalized.arguments.length === 1
    && normalized.arguments[0].getText(source) === input
    && selection.arguments[1].getText(source) === pageIndex;
}

function objectProperty(object, name) {
  return object.properties.find((property) => (
    (ts.isPropertyAssignment(property)
      || ts.isShorthandPropertyAssignment(property)
      || ts.isMethodDeclaration(property))
    && (ts.isIdentifier(property.name) || ts.isStringLiteralLike(property.name))
    && property.name.text === name
  ));
}

function isCanonicalWorkerVariantAttachment(call, source) {
  if (call.arguments.length !== 1) return false;
  const value = unwrapStaticExpression(call.arguments[0]);
  if (!ts.isObjectLiteralExpression(value)) return false;
  const retainedSource = objectProperty(value, 'source');
  const services = objectProperty(value, 'services');
  const buildLayout = objectProperty(value, 'buildLayout');
  if (!retainedSource || !services || !ts.isPropertyAssignment(buildLayout)) return false;
  const sourceExpression = ts.isShorthandPropertyAssignment(retainedSource)
    ? retainedSource.name : retainedSource.initializer;
  const servicesExpression = ts.isShorthandPropertyAssignment(services) ? services.name : services.initializer;
  const builder = unwrapStaticExpression(buildLayout.initializer);
  if (!ts.isArrowFunction(builder)
    || builder.parameters.length !== 1
    || !ts.isIdentifier(builder.parameters[0].name)) return false;
  const layoutCall = callOf(builder.body, 'layoutDocument');
  return layoutCall !== null
    && layoutCall.arguments.length === 3
    && layoutCall.arguments[0].getText(source) === sourceExpression.getText(source)
    && layoutCall.arguments[1].getText(source) === servicesExpression.getText(source)
    && layoutCall.arguments[2].getText(source) === builder.parameters[0].name.text;
}

function exactObjectPropertyIdentities(object, identities, source) {
  if (object.properties.length !== identities.size) return false;
  for (const [name, expected] of identities) {
    const property = objectProperty(object, name);
    const value = ts.isShorthandPropertyAssignment(property)
      ? property.name
      : ts.isPropertyAssignment(property)
        ? property.initializer
        : null;
    if (value?.getText(source) !== expected) return false;
  }
  return true;
}

function returnedObjectLiteral(body) {
  const returns = body.statements.filter(ts.isReturnStatement);
  if (returns.length !== 1 || !returns[0].expression) return null;
  let expression = unwrapStaticExpression(returns[0].expression);
  if (ts.isCallExpression(expression)
    && ts.isPropertyAccessExpression(expression.expression)
    && expression.expression.expression.getText() === 'Object'
    && expression.expression.name.text === 'freeze'
    && expression.arguments.length === 1) {
    expression = unwrapStaticExpression(expression.arguments[0]);
  }
  return ts.isObjectLiteralExpression(expression) ? expression : null;
}

function workerRenderCallIsCanonical(call, source) {
  if (call.arguments.length !== 4) return false;
  let block = call.parent;
  while (block && !ts.isBlock(block)) block = block.parent;
  const sourceDeclaration = block?.statements.flatMap((statement) => (
    ts.isVariableStatement(statement)
      ? statement.declarationList.declarations.filter((declaration) => (
          ts.isIdentifier(declaration.name) && declaration.name.text === 'source'
        ))
      : []
  ));
  const sourceLookup = sourceDeclaration?.length === 1
    ? callOf(sourceDeclaration[0].initializer, 'layoutSourceStoreOf')
    : null;
  if (sourceLookup?.arguments.length !== 1
    || sourceLookup.arguments[0].getText(source) !== 'doc.layoutServices') return false;
  const options = unwrapStaticExpression(call.arguments[3]);
  if (!ts.isObjectLiteralExpression(options)) return false;
  const services = objectProperty(options, 'layoutServices');
  const defaultDate = objectProperty(options, 'defaultCurrentDateMs');
  return call.arguments[0].getText(source) === 'source'
    && call.arguments[1].getText(source) === 'canvas'
    && call.arguments[2].getText(source) === 'req.pageIndex'
    && ts.isPropertyAssignment(services)
    && services.initializer.getText(source) === 'doc.layoutServices'
    && ts.isPropertyAssignment(defaultDate)
    && defaultDate.initializer.getText(source) === 'doc.defaultCurrentDateMs';
}

function workerRetentionCallIsCanonical(call, source) {
  return call.arguments.length === 3
    && call.arguments[0].getText(source) === 'source'
    && call.arguments[1].getText(source) === 'layoutServices'
    && call.arguments[2].getText(source) === 'req.defaultCurrentDateMs';
}

function workerRetentionSeamIsCanonical(source) {
  const declarations = source.statements.flatMap(declarationNames).sort();
  const exactDeclarations = [...WORKER_LAYOUT_RETENTION_DECLARATIONS].sort();
  if (declarations.length !== exactDeclarations.length
    || declarations.some((name, index) => name !== exactDeclarations[index])) return false;
  const retention = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement)
    && statement.name?.text === 'retainRenderWorkerDocumentLayout'
  ));
  if (!retention?.body
    || retention.parameters.length !== 3
    || retention.parameters.some((parameter) => !ts.isIdentifier(parameter.name))
    || retention.parameters.map((parameter) => parameter.name.getText(source)).join(',')
      !== 'source,layoutServices,defaultCurrentDateMs') return false;
  const attachments = callsNamed(retention.body, 'attachDocumentLayoutVariants');
  if (attachments.length !== 1
    || callsNamed(retention.body, 'layoutDocument').length !== 1
    || !isCanonicalWorkerVariantAttachment(attachments[0], source)) return false;
  const attachment = unwrapStaticExpression(attachments[0].arguments[0]);
  if (!ts.isObjectLiteralExpression(attachment)) return false;
  const retainedSource = objectProperty(attachment, 'source');
  const services = objectProperty(attachment, 'services');
  const defaultDate = objectProperty(attachment, 'defaultCurrentDateMs');
  const sourceExpression = ts.isShorthandPropertyAssignment(retainedSource)
    ? retainedSource.name
    : ts.isPropertyAssignment(retainedSource)
      ? retainedSource.initializer
      : null;
  const servicesExpression = ts.isShorthandPropertyAssignment(services)
    ? services.name
    : ts.isPropertyAssignment(services)
      ? services.initializer
      : null;
  const defaultDateExpression = ts.isShorthandPropertyAssignment(defaultDate)
    ? defaultDate.name
    : ts.isPropertyAssignment(defaultDate)
      ? defaultDate.initializer
      : null;
  const returned = returnedObjectLiteral(retention.body);
  return sourceExpression?.getText(source) === 'source'
    && servicesExpression?.getText(source) === 'layoutServices'
    && defaultDateExpression?.getText(source) === 'defaultCurrentDateMs'
    && returned !== null
    && exactObjectPropertyIdentities(returned, new Map([
      ['layoutServices', 'layoutServices'],
      ['layoutVariants', 'variants.store'],
      ['defaultCurrentDateMs', 'defaultCurrentDateMs'],
    ]), source);
}

function variableDeclarationsNamed(node, name) {
  const declarations = [];
  const visit = (current) => {
    if (ts.isVariableDeclaration(current)
      && ts.isIdentifier(current.name)
      && current.name.text === name) declarations.push(current);
    ts.forEachChild(current, visit);
  };
  visit(node);
  return declarations;
}

/**
 * The worker's parse metadata must describe the variant the load actually
 * selected, and exactly one spelling of it.
 *
 * `showTrackedChanges` and an explicit `currentDate` each paginate the document
 * differently, so a metadata route reading the DEFAULT variant would report a
 * page count and page sizes belonging to a layout nobody is going to paint.
 * Pinning both the selection (`layoutFor(layoutOptions)`) and the options it is
 * selected by (normalized from the request's own fields) keeps the reported
 * geometry, the primed progressive prefix and the painted pages on one key.
 */
function functionDeclarationNamed(source, name) {
  return source.statements.filter((statement) => (
    ts.isFunctionDeclaration(statement) && statement.name?.text === name
  ));
}

function workerMetadataAuthorityIsCanonical(source) {
  const projectors = functionDeclarationNamed(source, 'projectRenderWorkerLayoutMeta');
  const selectors = functionDeclarationNamed(source, 'renderWorkerLayoutMeta');
  if (projectors.length !== 1 || selectors.length !== 1) return false;

  const projector = projectors[0];
  const projectorReturns = projector.body?.statements.filter(ts.isReturnStatement) ?? [];
  const value = projectorReturns.length === 1
    ? projectorReturns[0].expression && unwrapStaticExpression(projectorReturns[0].expression)
    : null;
  if (!value || !ts.isObjectLiteralExpression(value)) return false;
  const pageCount = objectProperty(value, 'pageCount');
  const pageSizes = objectProperty(value, 'pageSizes');
  const bookmarks = objectProperty(value, 'bookmarkPages');
  const comments = objectProperty(value, 'commentAnchorRanges');
  const revisions = objectProperty(value, 'revisionAnchorRanges');
  const pageSizesCall = ts.isPropertyAssignment(pageSizes)
    ? unwrapStaticExpression(pageSizes.initializer)
    : null;
  const bookmarkValue = ts.isPropertyAssignment(bookmarks)
    ? unwrapStaticExpression(bookmarks.initializer)
    : null;
  const bookmarkSpread = bookmarkValue && ts.isArrayLiteralExpression(bookmarkValue)
    && bookmarkValue.elements.length === 1
    && ts.isSpreadElement(bookmarkValue.elements[0])
    ? bookmarkValue.elements[0]
    : null;
  const commentValue = ts.isPropertyAssignment(comments)
    ? unwrapStaticExpression(comments.initializer)
    : null;
  const revisionValue = ts.isPropertyAssignment(revisions)
    ? unwrapStaticExpression(revisions.initializer)
    : null;
  const commentCall = commentValue && ts.isConditionalExpression(commentValue)
    ? callOf(unwrapStaticExpression(commentValue.whenTrue), 'collectLayoutSourceCommentRangesIfPresent')
    : null;
  const revisionCall = revisionValue && ts.isConditionalExpression(revisionValue)
    ? callOf(unwrapStaticExpression(revisionValue.whenTrue), 'collectLayoutSourceRevisionRangesIfPresent')
    : null;
  const reviewIndexCalls = callsNamed(projector.body, 'reviewProjectionIndexForDocument');
  if (!ts.isPropertyAssignment(pageCount)
    || pageCount.initializer.getText(source) !== 'layout.pages.length'
    || !pageSizesCall
    || !ts.isCallExpression(pageSizesCall)
    || !ts.isPropertyAccessExpression(pageSizesCall.expression)
    || pageSizesCall.expression.name.text !== 'map'
    || pageSizesCall.expression.expression.getText(source) !== 'layout.pages'
    || !bookmarkSpread
    || callOf(bookmarkSpread.expression, 'buildBookmarkPageMap')?.arguments[0]?.getText(source) !== 'layout'
    || reviewIndexCalls.length !== 1
    || reviewIndexCalls[0].arguments.map((argument) => argument.getText(source)).join(',') !== 'layout'
    || !commentValue
    || !ts.isConditionalExpression(commentValue)
    || !revisionValue
    || !ts.isConditionalExpression(revisionValue)
    || commentValue.condition.getText(source) !== 'hasComments'
    || commentValue.whenFalse.getText(source) !== '[]'
    || revisionValue.condition.getText(source) !== 'hasRevisions'
    || revisionValue.whenFalse.getText(source) !== '[]'
    || commentCall?.arguments.map((argument) => argument.getText(source)).join(',')
      !== 'review.comments,source,reviewIndex!.renderedRunIndex,reviewProjection'
    || revisionCall?.arguments.map((argument) => argument.getText(source)).join(',')
      !== 'review.revisions,source,reviewIndex!.renderedRunIndex,reviewProjection') return false;

  const selector = selectors[0];
  const sources = variableDeclarationsNamed(selector, 'source');
  const layouts = variableDeclarationsNamed(selector, 'layout');
  const selectorReturns = selector.body?.statements.filter(ts.isReturnStatement) ?? [];
  const layoutFor = layouts[0]?.initializer
    && unwrapStaticExpression(layouts[0].initializer);
  const normalized = layoutFor
    && ts.isCallExpression(layoutFor)
    && layoutFor.arguments[0]
    && callOf(unwrapStaticExpression(layoutFor.arguments[0]), 'normalizeLayoutOptions');
  const delegated = selectorReturns.length === 1
    && selectorReturns[0].expression
    && callOf(unwrapStaticExpression(selectorReturns[0].expression), 'projectRenderWorkerLayoutMeta');
  return sources.length === 1
    && sources[0].initializer?.getText(source) === 'layoutSourceStoreOf(doc.layoutServices)'
    && layouts.length === 1
    && layoutFor
    && ts.isCallExpression(layoutFor)
    && layoutFor.expression.getText(source) === 'doc.layoutVariants.layoutFor'
    && normalized?.arguments.map((argument) => argument.getText(source)).join(',')
      === 'currentDateMs,doc.defaultCurrentDateMs,showTrackedChanges'
    && delegated?.arguments.map((argument) => argument.getText(source)).join(',')
      === 'layout,source,review';
}

function workerMetadataRouteIsCanonical(source, metadataSource) {
  const layouts = variableDeclarationsNamed(source, 'layout');
  const metadata = variableDeclarationsNamed(source, 'meta');
  const viewOptions = variableDeclarationsNamed(source, 'layoutOptions');
  if (layouts.length !== 1
    || layouts[0].initializer?.getText(source) !== 'doc.layoutVariants.layoutFor(layoutOptions)'
    || viewOptions.length !== 1
    || metadata.length !== 1) return false;
  const viewInitializer = viewOptions[0].initializer
    && unwrapStaticExpression(viewOptions[0].initializer);
  const viewCall = viewInitializer && callOf(viewInitializer, 'normalizeLayoutOptions');
  if (!viewCall
    || viewCall.arguments.length !== 3
    || viewCall.arguments.map((argument) => argument.getText(source)).join(',')
      !== 'req.currentDateMs,req.defaultCurrentDateMs,req.showTrackedChanges') return false;
  const value = metadata[0].initializer && unwrapStaticExpression(metadata[0].initializer);
  if (!value || !ts.isObjectLiteralExpression(value)) return false;
  const projection = value.properties.find(ts.isSpreadAssignment);
  const projectionCall = projection
    ? callOf(unwrapStaticExpression(projection.expression), 'projectRenderWorkerLayoutMeta')
    : null;
  const runtimeCalls = callsNamed(source, 'renderWorkerLayoutMeta');
  return projectionCall?.arguments.map((argument) => argument.getText(source)).join(',')
      === 'layout,source,reviewIndexInput'
    && runtimeCalls.length === 1
    && runtimeCalls[0].arguments.map((argument) => argument.getText(source)).join(',')
      === 'doc,reviewIndexInput,req.currentDateMs,req.showTrackedChanges'
    && workerMetadataAuthorityIsCanonical(metadataSource);
}

/**
 * The eager route must be nothing but a drain of the canonical generator, so
 * time-sliced and straight-through pagination cannot diverge and the eager
 * route cannot reach a layout that skipped the validate-and-freeze boundary.
 */
function eagerPaginationRouteDrainsCanonicalSteps(producer) {
  if (producer.body?.statements.length !== 1) return false;
  const returned = producer.body.statements[0];
  if (!ts.isReturnStatement(returned)) return false;
  const drain = callOf(returned.expression, CANONICAL_PAGINATION_DRAIN);
  return drain?.arguments.length === 1
    && callOf(drain.arguments[0], CANONICAL_PAGINATION_PRODUCER) !== null;
}

function assertCanonicalCutoverBoundaries(root) {
  const paginatorPath = resolve(root, LAYOUT_SOURCE, 'body-paginator.ts');
  const producerLabel = `${LAYOUT_SOURCE}/body-paginator.ts#${CANONICAL_PAGINATION_PRODUCER}`;
  if (!existsSync(paginatorPath)) {
    fail('CANONICAL_LAYOUT_PRODUCER', producerLabel);
  } else {
    const source = sourceFile(paginatorPath);
    const exportedValues = source.statements.filter((statement) => (
      hasExportModifier(statement)
      && (ts.isFunctionDeclaration(statement)
        || ts.isVariableStatement(statement)
        || ts.isClassDeclaration(statement)
        || ts.isEnumDeclaration(statement))
    ));
    const runtimeExportForms = source.statements.filter((statement) => (
      ts.isExportAssignment(statement)
      || (ts.isExportDeclaration(statement) && !exportIsTypeOnly(statement))
    ));
    const exportedFunctions = new Map(exportedValues
      .filter((statement) => ts.isFunctionDeclaration(statement) && statement.name)
      .map((statement) => [statement.name.text, statement]));
    // Pagination is time-sliced: `paginateBodySteps` is the single producer that
    // validates and freezes, `drainPagination` runs it straight through, and
    // `paginateBody` is the eager route spelled as exactly that drain. Pinning
    // the exported set to those three keeps one validated exit from the module.
    const producer = exportedFunctions.get(CANONICAL_PAGINATION_PRODUCER);
    const eager = exportedFunctions.get(CANONICAL_PAGINATION_EAGER_ROUTE);
    if (exportedValues.length !== CANONICAL_PAGINATOR_EXPORTS.length
      || CANONICAL_PAGINATOR_EXPORTS.some((name) => !exportedFunctions.get(name)?.body)
      || !producer.asteriskToken
      || runtimeExportForms.length !== 0
      || !eagerPaginationRouteDrainsCanonicalSteps(eager)) {
      fail('CANONICAL_LAYOUT_PRODUCER', producerLabel);
    }
    const returned = producer.body.statements.at(-1);
    const frozenCall = returned && ts.isReturnStatement(returned)
      ? callOf(returned.expression, 'assertAndDeepFreezeDocumentLayout')
      : null;
    if (!hasExactInvariantImports(source)
      || callsNamed(source, 'assertAndDeepFreezeDocumentLayout').length !== 1
      || frozenCall?.arguments.length !== 1) {
      fail('RETAINED_LAYOUT_IMMUTABILITY', producerLabel);
    }
  }

  const variantsPath = resolve(root, LAYOUT_SOURCE, 'document-layout-variants.ts');
  if (!existsSync(variantsPath)) {
    fail('SELECTED_LAYOUT_VARIANT', `${LAYOUT_SOURCE}/document-layout-variants.ts#selectDocumentLayoutPage`);
  } else {
    const source = sourceFile(variantsPath);
    const selection = source.statements.find((statement) => (
      ts.isFunctionDeclaration(statement)
      && statement.name?.text === 'selectDocumentLayoutPage'
    ));
    if (!selection || !isCanonicalSelectedLayoutFunction(selection, source)) {
      fail('SELECTED_LAYOUT_VARIANT', `${LAYOUT_SOURCE}/document-layout-variants.ts#selectDocumentLayoutPage`);
    }
  }

  const workerPath = resolve(root, DOCX_SOURCE, 'render-worker.ts');
  const workerRetentionPath = resolve(root, DOCX_SOURCE, 'render-worker-layout.ts');
  if (!existsSync(workerPath)) {
    fail('WORKER_LAYOUT_SELECTION', `${DOCX_SOURCE}/render-worker.ts`);
  } else if (!existsSync(workerRetentionPath)
    || !workerRetentionSeamIsCanonical(sourceFile(workerRetentionPath))) {
    fail('WORKER_LAYOUT_SELECTION', `${DOCX_SOURCE}/render-worker-layout.ts`);
  } else {
    const source = sourceFile(workerPath);
    const workerMetadataPath = resolve(root, DOCX_SOURCE, 'render-worker-metadata.ts');
    const topLevelPages = source.statements.some((statement) => (
      ts.isVariableStatement(statement)
      && statement.declarationList.declarations.some((declaration) => (
        ts.isIdentifier(declaration.name) && declaration.name.text === 'pages'
      ))
    ));
    let duplicateSelection = false;
    const visit = (node) => {
      if (ts.isIdentifier(node) && node.text === 'selectDocumentLayoutPage') {
        duplicateSelection = true;
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
    const variantAttachments = callsNamed(source, 'attachDocumentLayoutVariants');
    const retentionCalls = callsNamed(source, 'retainRenderWorkerDocumentLayout');
    const rendererCalls = callsNamed(source, 'renderLayoutSourceToCanvas');
    if (topLevelPages
      || duplicateSelection
      || variantAttachments.length !== 0
      || callsNamed(source, 'layoutDocument').length !== 0
      || retentionCalls.length !== 1
      || !workerRetentionCallIsCanonical(retentionCalls[0], source)
      || rendererCalls.length !== 1
      || rendererCalls.some((call) => !workerRenderCallIsCanonical(call, source))
      || !existsSync(workerMetadataPath)
      || !workerMetadataRouteIsCanonical(source, sourceFile(workerMetadataPath))) {
      fail('WORKER_LAYOUT_SELECTION', `${DOCX_SOURCE}/render-worker.ts`);
    }

    const rendererPath = resolve(root, DOCX_SOURCE, 'renderer.ts');
    if (existsSync(rendererPath)
      && callsNamed(sourceFile(rendererPath), 'selectDocumentLayoutPage').length !== 1) {
      fail('SELECTED_LAYOUT_VARIANT', `${DOCX_SOURCE}/renderer.ts#selectDocumentLayoutPage`);
    }
  }
}

function identifierCounts(root) {
  const counts = {};
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const fileCounts = Object.fromEntries(LEGACY_SYMBOLS.map((symbol) => [symbol, 0]));
    const source = sourceFile(path);
    const visit = (node) => {
      if (ts.isIdentifier(node) && Object.hasOwn(fileCounts, node.text)) fileCounts[node.text] += 1;
      ts.forEachChild(node, visit);
    };
    visit(source);
    const file = posixPath(relative(root, path));
    for (const [symbol, count] of Object.entries(fileCounts)) {
      if (count > 0) counts[`${file}#${symbol}`] = count;
    }
  }
  return Object.fromEntries(Object.entries(counts).sort(([left], [right]) => left.localeCompare(right)));
}

function bindingNames(name, names = []) {
  if (ts.isIdentifier(name)) names.push(name.text);
  else for (const element of name.elements) {
    if (!ts.isOmittedExpression(element)) bindingNames(element.name, names);
  }
  return names;
}

function declarationNames(statement) {
  if ((ts.isFunctionDeclaration(statement)
      || ts.isClassDeclaration(statement)
      || ts.isInterfaceDeclaration(statement)
      || ts.isTypeAliasDeclaration(statement)
      || ts.isEnumDeclaration(statement)
      || ts.isModuleDeclaration(statement))
    && statement.name) {
    return [statement.name.text];
  }
  if (ts.isVariableStatement(statement)) {
    return statement.declarationList.declarations.flatMap((declaration) => bindingNames(declaration.name));
  }
  return [];
}

function rendererImportEdges(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) return [];
  return [...new Set(moduleEdges(renderer)
    .filter((edge) => edge.literal)
    .map((edge) => resolveLocalImport(renderer, edge.specifier))
    .filter((path) => path && LEGACY_RENDERER_IMPORTS.has(basename(path)))
    .map((path) => `${DOCX_SOURCE}/renderer.ts -> ${posixPath(relative(root, path))}`))]
    .sort();
}

function currentLegacyInventory(root) {
  return {
    version: 3,
    legacySymbolCounts: identifierCounts(root),
    migrationIdentifierCounts: matchingIdentifierCounts(root, (name) => MIGRATION_IDENTIFIER.test(name)),
    rendererImportEdges: rendererImportEdges(root),
  };
}

function stableJson(value) {
  return `${JSON.stringify(value, null, 2)}\n`;
}

function assertPaintBoundaries(root) {
  const { violations, nonLiteral } = paintBoundaryViolations(root);
  if (nonLiteral.length > 0) {
    fail('NON_LITERAL_MODULE_EDGE', nonLiteral.join('\n'));
  }
  if (violations.length > 0) {
    fail('FORBIDDEN_PAINT_EDGE', violations.map((chain) => chain.join(' -> ')).join('\n'));
  }
}

function hasExportModifier(statement) {
  return statement.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword) ?? false;
}

function rendererExports(path) {
  const names = [];
  const source = sourceFile(path);
  for (const statement of source.statements) {
    if (hasExportModifier(statement)) {
      const declared = declarationNames(statement);
      names.push(...(declared.length > 0 ? declared : ['default']));
    }
    if (ts.isExportDeclaration(statement) && statement.exportClause && ts.isNamedExports(statement.exportClause)) {
      names.push(...statement.exportClause.elements.map((element) => element.name.text));
    }
  }
  return [...new Set(names)].sort();
}

function rendererImportBindings(source) {
  const bindings = new Set();
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement) || !statement.importClause) continue;
    if (statement.importClause.name) bindings.add(statement.importClause.name.text);
    const named = statement.importClause.namedBindings;
    if (named && ts.isNamespaceImport(named)) bindings.add(named.name.text);
    if (named && ts.isNamedImports(named)) {
      for (const element of named.elements) bindings.add(element.name.text);
    }
  }
  return bindings;
}

function unwrapAdapterExpression(expression) {
  if (ts.isAwaitExpression(expression)
    || ts.isParenthesizedExpression(expression)
    || ts.isAsExpression(expression)
    || ts.isSatisfiesExpression(expression)) {
    return unwrapAdapterExpression(expression.expression);
  }
  return expression;
}

function adapterValueIsAllowed(expression, callable) {
  const value = unwrapAdapterExpression(expression);
  if (ts.isIdentifier(value)
    || ts.isStringLiteralLike(value)
    || ts.isNumericLiteral(value)
    || value.kind === ts.SyntaxKind.TrueKeyword
    || value.kind === ts.SyntaxKind.FalseKeyword
    || value.kind === ts.SyntaxKind.NullKeyword) return true;
  if (ts.isPropertyAccessExpression(value)) return adapterValueIsAllowed(value.expression, callable);
  if (ts.isElementAccessExpression(value)) {
    return adapterValueIsAllowed(value.expression, callable)
      && (!value.argumentExpression || adapterValueIsAllowed(value.argumentExpression, callable));
  }
  if (ts.isArrayLiteralExpression(value)) {
    return value.elements.every((element) => !ts.isSpreadElement(element)
      ? adapterValueIsAllowed(element, callable)
      : adapterValueIsAllowed(element.expression, callable));
  }
  if (ts.isObjectLiteralExpression(value)) {
    return value.properties.every((property) => {
      if (ts.isPropertyAssignment(property)) return adapterValueIsAllowed(property.initializer, callable);
      if (ts.isShorthandPropertyAssignment(property)) return true;
      if (ts.isSpreadAssignment(property)) return adapterValueIsAllowed(property.expression, callable);
      return false;
    });
  }
  if (ts.isCallExpression(value)) {
    const callee = unwrapAdapterExpression(value.expression);
    const callableName = ts.isIdentifier(callee)
      ? callee.text
      : ts.isPropertyAccessExpression(callee) && ts.isIdentifier(callee.expression)
        ? callee.expression.text
        : null;
    return callableName != null
      && callable.has(callableName)
      && value.arguments.every((argument) => adapterValueIsAllowed(argument, callable));
  }
  return false;
}

function adapterBodyIsAllowed(body, callable) {
  return body.statements.every((statement) => {
    if (ts.isVariableStatement(statement)) {
      return statement.declarationList.declarations.every((declaration) => (
        ts.isIdentifier(declaration.name)
        && (!declaration.initializer || adapterValueIsAllowed(declaration.initializer, callable))
      ));
    }
    if (ts.isExpressionStatement(statement)) return adapterValueIsAllowed(statement.expression, callable);
    if (ts.isReturnStatement(statement)) {
      return !statement.expression || adapterValueIsAllowed(statement.expression, callable);
    }
    return false;
  });
}

function exactArgumentTexts(call, expected, source) {
  return call !== null
    && call.arguments.length === expected.length
    && call.arguments.every((argument, index) => argument.getText(source) === expected[index]);
}

function sourceRendererBodyIsCanonical(declaration, source) {
  if (!declaration?.body || declaration.body.statements.length !== 2) return false;
  const [first, second] = declaration.body.statements;
  if (!ts.isVariableStatement(first)
    || first.declarationList.declarations.length !== 1) return false;
  const retained = first.declarationList.declarations[0];
  if (!ts.isIdentifier(retained.name) || retained.name.text !== 'normalized') return false;
  if (!exactArgumentTexts(
    callOf(retained.initializer, 'normalizeRenderOptions'),
    ['source', 'canvas', 'pageIndex', 'opts'],
    source,
  )) return false;
  if (!ts.isReturnStatement(second)) return false;
  return exactArgumentTexts(
    callOf(second.expression, 'renderSelectedDocumentPage'),
    [
      'normalized.selection.layout',
      'normalized.selection.page',
      'canvas',
      'normalized.paintOptions',
    ],
    source,
  );
}

function publicRendererBodyIsCanonical(declaration, source) {
  if (!declaration?.body || declaration.body.statements.length !== 1) return false;
  const returned = declaration.body.statements[0];
  if (!ts.isReturnStatement(returned)) return false;
  const render = callOf(returned.expression, 'renderLayoutSourceToCanvas');
  if (render?.arguments.length !== 4) return false;
  const sourceCall = callOf(render.arguments[0], 'layoutSourceStore');
  return exactArgumentTexts(sourceCall, ['doc'], source)
    && render.arguments[1].getText(source) === 'canvas'
    && render.arguments[2].getText(source) === 'pageIndex'
    && render.arguments[3].getText(source) === 'opts';
}

function assertFinalRendererAdapter(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) fail('FINAL_ADAPTER_MISSING', `${DOCX_SOURCE}/renderer.ts`);
  const source = sourceFile(renderer);
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement)
      && (!statement.exportClause || !ts.isNamedExports(statement.exportClause))) {
      fail('FINAL_ADAPTER_EXPORT', statement.getText(source));
    }
    for (const name of declarationNames(statement)) {
      if (!FINAL_RENDERER_DECLARATIONS.has(name)) fail('FINAL_ADAPTER_DECLARATION', name);
    }
  }
  for (const name of rendererExports(renderer)) {
    if (!FINAL_RENDERER_EXPORTS.has(name)) fail('FINAL_ADAPTER_EXPORT', name);
  }
  const callable = rendererImportBindings(source);
  callable.add('createLayoutServices');
  callable.add('normalizeRenderOptions');
  callable.add('renderLayoutSourceToCanvas');
  const sourceRenderer = source.statements.find((statement) => ts.isFunctionDeclaration(statement)
    && statement.name?.text === 'renderLayoutSourceToCanvas');
  const publicRenderer = source.statements.find((statement) => ts.isFunctionDeclaration(statement)
    && statement.name?.text === 'renderDocumentToCanvas');
  if (!sourceRendererBodyIsCanonical(sourceRenderer, source)) {
    fail('FINAL_ADAPTER_BODY', 'renderLayoutSourceToCanvas');
  }
  if (!publicRendererBodyIsCanonical(publicRenderer, source)) {
    fail('FINAL_ADAPTER_BODY', 'renderDocumentToCanvas');
  }
  for (const edge of moduleEdges(renderer)) {
    if (!edge.literal) fail('FINAL_ADAPTER_IMPORT', '<dynamic>');
    if (edge.bare) fail('FINAL_ADAPTER_IMPORT', edge.specifier);
    if (!edge.specifier.startsWith('.')) continue;
    const target = resolveLocalImport(renderer, edge.specifier);
    if (!target) fail('FINAL_ADAPTER_IMPORT', edge.specifier);
    const rel = posixPath(relative(root, target));
    const allowed = rel.startsWith(`${LAYOUT_SOURCE}/`)
      || rel.startsWith(`${PAINT_SOURCE}/`)
      || rel === LAYOUT_RUNTIME_ADAPTER
      || rel === `${DOCX_SOURCE}/layout-source-model-adapter.ts`
      || rel === TEXT_RUN_PROJECTION_ADAPTER
      || (edge.typeOnly && rel === `${DOCX_SOURCE}/types.ts`);
    if (!allowed) fail('FINAL_ADAPTER_IMPORT', `${DOCX_SOURCE}/renderer.ts -> ${rel}`);
  }
}

function assertFinalRendererAssetImports(root) {
  const fixture = resolve(root, CONFORMANCE_FIXTURE);
  if (!existsSync(fixture)) fail('FINAL_RENDERER_ASSET_MISSING', CONFORMANCE_FIXTURE);
  const html = readFileSync(fixture, 'utf8');
  const rendererImports = /import\s*\{([\s\S]*?)\}\s*from\s*['"]\/src\/renderer\.ts['"]/g;
  for (const match of html.matchAll(rendererImports)) {
    const importedNames = match[1].split(',').map((binding) => (
      binding.trim().split(/\s+as\s+/u, 1)[0]
    ));
    const removedBinding = importedNames.find((name) => (
      name === 'createLayoutServices' || name === 'layoutDocument'
    ));
    if (removedBinding) {
      fail('FINAL_RENDERER_ASSET_IMPORT', `${CONFORMANCE_FIXTURE} -> ${removedBinding}`);
    }
  }
}

function parseArguments(argv) {
  const options = {
    root: process.cwd(),
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--root') options.root = resolve(argv[++index]);
    else if (arg === '--final') continue;
    else fail('UNKNOWN_ARGUMENT', arg);
  }
  return options;
}

export function checkDocxLayoutBoundaries(options) {
  const root = resolve(options.root);
  const baselinePath = resolve(root, BASELINE_PATH);
  const baselineExists = existsSync(baselinePath);
  assertNoProductionTestSupportImports(root);
  assertAcquisitionContextBoundary(root);
  assertProductionBodyAcquisitionAuthority(root);
  assertRendererAcquisitionProjectionBoundary(root);
  assertFloatRectTransportBoundary(root);
  assertFloatPlacementAuthority(root);
  assertNoDeletedPageProducerIdentifiers(root);
  assertPaintBoundaries(root);
  assertCapabilityBoundaries(root);
  assertAffineRuntimeDependencies(root);
  assertTextRunProjectionAdapterBoundary(root);
  assertCoordinateSpaceRuntimeDependencies(root);
  assertOccurrenceProjectionRuntimeDependencies(root);
  assertBodyPaintConsumesRetainedLayout(root);
  assertLayoutParserModelBoundaries(root);
  assertBodyLayoutAdapterBoundary(root);
  assertParagraphAnchorFrameAdapterBoundary(root);
  assertBodyKernelServiceOwner(root);
  assertCanonicalCutoverBoundaries(root);

  if (baselineExists) fail('FINAL_BASELINE_PRESENT', BASELINE_PATH);
  assertFinalRendererAdapter(root);
  assertFinalRendererAssetImports(root);
  const actual = currentLegacyInventory(root);
  if (Object.keys(actual.legacySymbolCounts).length > 0
    || Object.keys(actual.migrationIdentifierCounts).length > 0
    || actual.rendererImportEdges.length > 0) {
    fail('FINAL_LEGACY_BOUNDARY', stableJson(actual).trim());
  }
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    checkDocxLayoutBoundaries(parseArguments(process.argv.slice(2)));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
