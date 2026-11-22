// Things that were deleted on purpose and must not come back. Data only: the
// checker is scripts/check-tombstones.mjs. Add an entry here when a migration
// finishes; never remove one to make a change pass.

/** Files that must not exist. `path` is repository-relative; `basename` bans
 * that file name anywhere under `under`. */
export const DELETED_PATHS = [
  // Migration ratchets. Both guards run in final mode only; a baseline file
  // would be a way to accept regressions again.
  { path: 'scripts/docx-layout-boundary-baseline.json' },
  { path: 'scripts/docx-compatibility-observation-baseline.json' },
  // Replaced by retained paragraph anchor frames.
  { path: 'packages/docx/src/paragraph-anchor-frame-adapter.ts' },
  // Pre-retained-layout DOCX producers. The retained table pipeline owns row
  // heights; fragment arrays are no longer a paint input.
  { under: 'packages/docx/src', basename: 'table-geometry.ts' },
  { under: 'packages/docx/src', basename: 'table-fragments.ts' },
  { under: 'packages/docx/src', basename: 'layout-fragments.ts' },
];

/** Identifiers that must not appear in production TypeScript under `root`.
 * `includeTests` extends an entry to test files as well. */
export const DELETED_IDENTIFIERS = [{
  root: 'packages/docx/src',
  // Any identifier with one of these names.
  names: [
    // Renderer-owned layout and paint of the pre-retained architecture.
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
    // Deleted page producers, raw page routes and their test adapters.
    'bodyFragmentFor',
    'bodyLayoutFallback',
    'paginateDocument',
    'paginateWithHeaderFooterReserve',
    'PaginatedBodyElement',
    'physicalPageSizeForPage',
    'prebuiltPages',
    'retainedLayout',
    'computeTableRowHeights',
    'estimateTableHeight',
    'measureRetainedCellContentHeightPt',
    'measureCellParagraphWindow',
    'measureCellElementHeight',
    'trimTrailingStructuralMarker',
    'paragraphSegsStateSensitive',
    'rescaleLayoutLines',
  ],
  // Any identifier matching this pattern: migration switches and silent
  // alternate-engine fallbacks.
  pattern: /(?:legacy|(?:use|enable|prefer|require)[a-z0-9]*(?:old|previous|alternate)[a-z0-9]*(?:engine|layout|path|algorithm)|(?:reuse|paint)enabled|requireslegacy|dryrun)/i,
  // `legacyDoc` is the public opt-in key for the optional binary DOC reader
  // (an input format, like `math` or `chartEx`), not a layout migration switch.
  patternExceptions: ['legacyDoc'],
  // Legacy stamp fields: banned as a property name (access, declaration,
  // object literal key or destructured key), free as a local variable name.
  propertyNames: [
    'colIndex',
    'colGeom',
    'colTopPt',
    'sectionHF',
    'sectionGeom',
    'sectionPageNumType',
    'sectionTextDirection',
  ],
}, {
  // Node DOCX probes once declared the type of their dynamic renderer import by
  // hand, which let a deleted name keep type-checking until a gated probe ran.
  // Select pages from `layoutDocument(...).pages` instead.
  root: 'packages/node/src',
  includeTests: true,
  names: ['paginateDocument', 'prebuiltPages'],
}];
