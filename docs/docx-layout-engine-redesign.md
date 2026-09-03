# DOCX Layout Engine Redesign

## Status

Implemented. Issue #1037 delivered the redesign as a series of independently
reviewed pull requests while preserving the public API. The final architecture
and its CI gates are now the maintained source of truth.

This design completes and supersedes the transitional boundaries in
`docs/docx-layout-context-fragments-design.md`. The existing fragment work is the
starting point, not the final architecture.

The implementation history is retained in Issue #1037. The migration roadmap
and per-series execution plans were removed when the final audit completed.

## Problem

The DOCX renderer supports a broad set of WordprocessingML features, but layout
policy, measurement, pagination, and painting still overlap:

- body paragraphs and tables can use either fragment paint or legacy paint;
- pagination can measure content and paint can measure it again;
- runtime layout data is stamped onto parsed model objects and continuation clones;
- page, column, story, float, font, and paint state share one mutable render state;
- some table-cell content is measured twice because row sizing does not retain the
  measured line geometry needed by fragments;
- headers, footers, notes, text boxes, and floating objects do not yet share one
  complete layout result;
- browser smoke tests prove that a canvas received ink, but do not prove page,
  line, table, or margin geometry.

These paths can disagree even when each path is locally reasonable. The result is
an architecture in which a new document can expose page overlap, margin invasion,
missing borders, or different line partitions between pagination and paint.

## Goals

1. Establish one pipeline from parsed OOXML facts to an immutable layout result
   and then to Canvas paint.
2. Measure and paginate in point space exactly once for a given placement.
3. Make paint a projection of completed layout geometry, with no line, row, or
   page remeasurement.
4. Cover body, headers, footers, notes, text boxes, tables, and floating drawings
   through the same layout contracts.
5. Preserve all existing public APIs and worker-mode behavior.
6. Remove legacy layout stamps, production migration flags, and silent fallback
   to a second layout algorithm.
7. Separate specification-defined behavior from documented Office compatibility
   behavior and unsupported-feature diagnostics.
8. Enforce layout invariants and main/worker parity in CI.
9. Finish with the structural quality of a clean implementation without requiring
   a risky one-shot rewrite.

## Non-Goals

- Changing `DocxViewer`, `DocxScrollViewer`, `DocxDocument`, or current render
  option signatures.
- Making DOCX pagination semantics artificially identical to XLSX or PPTX.
- Adding fixture-specific branches, empirical scaling factors, or filename-based
  behavior.
- Releasing intermediate migration states.
- Replacing validated shared DrawingML, image, color, chart, or font primitives
  when their current contracts already serve all three formats.

## Normative and Compatibility Policy

Implementation decisions are classified in this order:

1. ECMA-376 / ISO/IEC 29500 normative behavior.
2. Microsoft implementation notes and documented Office extensions.
3. Controlled Office-output observations where the standard is silent.
4. Unsupported or unresolved behavior.

Normative behavior records the relevant element, attribute, or section in tests
and explanatory comments. Office-observed behavior belongs in an isolated
compatibility module with its evidence boundary and synthetic tests. It must not
be hidden inside generic geometry code.

Unsupported or unresolved behavior uses a deterministic fallback and emits an
internal diagnostic. It never switches silently to a legacy painter. A new
compatibility heuristic requires an explicit design decision and evidence; the
absence of a specification rule is not permission to fit a private fixture.

## Target Data Flow

```text
Rust OOXML parser
  -> resolved document facts
  -> private parser-wire projection into immutable plain inputs
  -> immutable layout contexts
  -> point-space layout engine
  -> immutable DocumentLayout
  -> Canvas paint adapter
```

### Parser

Rust owns OOXML parsing and style-cascade resolution. It preserves authored,
inherited, explicitly-cleared, and absent values when those distinctions affect
layout. It does not decide browser line breaks, page breaks, Canvas scale, or DPR.

Parser-to-TypeScript model changes remain optional and backward compatible. A
parser fact that is needed for layout is preserved rather than reconstructed by a
renderer heuristic. Private wire-only facts are projected exactly once at the
parser-model boundary into parser-independent plain data; layout and paint never
dereference the private parser representation.

### Layout contexts

TypeScript resolves immutable document, section, story, container, paragraph,
run, and font contexts. These contexts contain WordprocessingML policy in points.
They do not contain Canvas state, scale, DPR, mutable page cursors, or paint
callbacks.

Story and container remain orthogonal. A table cell is a container inside a body,
header, footer, note, or text-box story; it is not a separate story kind.

### Layout engine

The layout engine receives resolved contexts plus explicit services:

```ts
export interface LayoutServices {
  readonly text: TextLayoutService;
  readonly images: ImageMetadataService;
  readonly math: MathMetadataService;
}

export interface LayoutOptions {
  readonly currentDateMs: number;
}
```

The text service is responsible for font selection, shaping, and measurement.
It preserves the four independent WordprocessingML `rFonts` axes, applies theme
precedence within each cascaded slot, and retains authored theme-attribute
presence even when its theme reference cannot resolve, because the matching
direct attribute remains suppressed in that case. It returns per-scalar spans
plus legal grapheme boundaries, which also constrain emergency overlong-word
splits. Each span carries one immutable core `CanvasFontRoute`;
measurement, kashida/vertical probes, and paint serialize that same complete CSS
family list. Exact registered tuples win first. Otherwise an authored family is
an engine-scoped native CSS request with a fontTable-derived generic tail, then
the shared bounded face-name classifier for faces absent from the table. When no
OOXML font slot resolves at any level, the DOCX consumer policy uses a serif
generic default for Latin and complex-script slots while preserving the existing
East Asian sans fallback, as permitted by §17.3.2.26; none of these fallback
routes claims that a local face was discovered.
Non-content glyphs use the same authority. Effective numbering-level run facts
and paragraph-mark run facts remain private parser-wire metadata. The
parser-model boundary projects numbering into a plain immutable shape input;
the parser-free layout module shapes mixed-script body and text-box markers and
the parser-free paint module serializes the retained spans. This follows
ECMA-376 §17.9.6 (`w:lvl`) and the §17.3.2.26 (`w:rFonts`) font-slot rules.
Paragraph-mark run properties follow §17.3.1.29 (`w:pPr`), without
reconstructing effective formatting in layout. Empty and anchor-only paragraph-
mark metrics also use the text service, so main and worker execution share one
font authority.

Auto-fit text enters through the ordinary `buildSegments` pipeline. Each atom
retains its complete service request and route, including the four independent
`w:rFonts` slots and theme-presence semantics from ECMA-376 §17.3.2.26. The
region resolver sums the actual shaped advances across grapheme, kinsoku,
font-route, and `joinPrev` boundaries instead of replacing them with one leading
run font. Table auto-fit selection follows §17.4.52 (`w:tblLayout`).
The image service resolves intrinsic dimensions and metadata. Services are inputs
only; functions, Canvas objects, and resource handles never enter the layout
result.

Layout uses points at scale 1. It owns line selection, row sizing, page and column
placement, float exclusion, note reservation, story placement, and page-dependent
field convergence. Geometry-affecting render options are normalized before layout
and participate in a stable layout cache key together with fingerprints derived
from the injected service instances; callers cannot provide service fingerprints
independently. Paint-only width, DPR, and color do not participate.

Math resource addressing is normalized once at the parser-model boundary. That
pure traversal shallow-clones only math-bearing ancestry and attaches typed
plain-data `SourceRef`/resource keys to internal math runs; neither parser-array
identity nor a WeakMap participates in layout.

Static enforcement walks the transitive runtime module graph from every
production layout module and rejects any path to `parser-model.ts`, including
bridge modules, aliases, literal or non-literal dynamic imports, and CommonJS
`require`. Only the exact unaliased named runtime import
`{ normalizeInternalDocumentModel }` from `../parser-model.js` in
`layout/resources.ts` is a terminal parser projection edge. Its binding has one
permitted use: the exported `documentMathOccurrences` function returns
`[...normalizeInternalDocumentModel(doc).mathOccurrences]`. Static enforcement
rejects local re-export, alias, leak, or extra references. Every other gateway
edge is traversed normally; erased type-only contracts do not create runtime
paths. This prevents parser ownership from leaking back into layout without
false positives for type-only contracts.

### Document layout

`DocumentLayout` is self-contained for paint:

```ts
export interface DocumentLayout {
  readonly pages: readonly LayoutPage[];
  readonly diagnostics: readonly LayoutDiagnostic[];
}

export interface LayoutPage {
  readonly geometry: PageGeometry;
  readonly section: SectionLayoutContext;
  readonly layers: PageLayers;
  readonly readingOrder: readonly LayoutNodeId[];
}
```

Page layers explicitly own background, behind-text drawings, header, body, notes,
front drawings, and footer content. Their `paintOrder` is derived from normative
anchor/story stacking rules and isolated Office compatibility evidence, then
stored as a layout result rather than assumed by paint or implemented as a
mutable callback side channel.

The primary node kinds are:

- `ParagraphLayout`: lines, baselines, advances, text placements, paragraph
  borders, and shading;
- `TableLayout`: columns, rows, cells, continuation ranges, vertical merges, and
  resolved shared border segments;
- `DrawingLayout`: bounds, wrap geometry, transforms, and z-order;
- `TextBoxLayout`: a nested story/container layout, including block content;
- `NoteLayout`: note region geometry and its relationship to the body reference.

Nodes retain only an opaque source reference for diagnostics, search, selection,
and navigation. Paint does not read `DocParagraph`, `DocTable`, or other parser
objects. Text, font descriptors, styles, coordinates, borders, and resource keys
needed by paint are already resolved.

Flow nodes store `flowBounds`, `inkBounds`, optional `clipBounds`, and
`advancePt` separately. Ordinary-flow allocation ownership is constrained by
page margins and non-overlap. Negative-spacing ink, frames, clipped overhang,
and explicitly allowed floating overlap are not incorrectly rejected as flow
ownership violations.

### Paint

Paint consumes a page from `DocumentLayout` and resource caches. It may set fonts,
transforms, colors, and draw images or text, but it may not:

- call text measurement or shaping;
- construct line segments;
- select line, row, or page breaks;
- calculate table columns or row heights;
- resolve styles or OOXML compatibility rules;
- mutate the parser model or layout result.

Scale and DPR are applied only by the paint adapter. Changing scale, DPR, target
canvas, or repaint count does not change layout geometry.

## Module Boundaries

The final implementation is organized by responsibility:

```text
packages/docx/src/layout/
  context.ts
  flow.ts
  text.ts
  numbering-marker.ts
  resources.ts
  convergence.ts
  compatibility.ts
  error-page.ts
  paragraph.ts
  table.ts
  floats.ts
  stories.ts
  paginator.ts
  invariants.ts
  diagnostics.ts
  types.ts

packages/docx/src/paint/
  canvas-page.ts
  canvas-text.ts
  numbering-marker.ts
  canvas-table.ts
  canvas-drawing.ts
```

`paginator.ts` coordinates placement but delegates paragraph, table, float, and
story-specific calculations. Page transitions are explicit state-machine events,
not closures over a renderer-wide mutable state.

`renderer.ts` is a thin internal adapter that prepares resources, calls layout,
and calls paint. It does not retain a second layout implementation. The final
boundary checker rejects hidden algorithms, transitional exports, and imports
outside the reviewed adapter surface.

DrawingML geometry, images, colors, fonts, charts, fills, and effects are shared
through `packages/core` or `packages/ooxml-common` where the OOXML concept is truly
common. DOCX page flow, WordprocessingML stories, table pagination, and paragraph
adjacency remain DOCX-specific.

## Convergence and Errors

Page-dependent fields, note reservation, and mutually-dependent floating
placement use explicit layout convergence. Each iteration produces a stable
fingerprint. The solver stops when the relevant inputs and geometry stabilize,
detects repeated fingerprints as a cycle, and applies a resource safety bound.

A resource safety bound is not a compatibility heuristic. Hitting it produces a
diagnostic and failure rather than returning an overlapping or partially stale
layout.

Recoverable unsupported decoration may omit only that decoration and record a
diagnostic. NaN geometry, invalid ownership, non-convergence, or broken layout
invariants use the existing render error path. Public error callback signatures do
not change.

Failure containment follows the retained-layout ownership boundary:

- unsupported or malformed optional content is recoverable only when its layout
  node can be replaced by a deterministic retained no-op without inventing
  geometry or changing surrounding flow;
- the retained node owns an immutable diagnostic with a stable source reference,
  and pagination collects that diagnostic into `DocumentLayout`;
- valid content before and after the recoverable node continues through the same
  acquisition, pagination, and paint pipeline;
- a recognized image or chart with a missing package resource is recoverable
  only when its authored DrawingML extent is valid; inline flow retains that
  exact advance and line height, floating content retains its anchor and wrap
  geometry, paint receives a diagnostic no-op, and no synthetic resource key or
  public image/chart model is created;
- corrupt required package parts, non-finite or negative geometry, invalid
  ownership, non-convergence, and invariant violations remain fatal;
- a recognized `nextColumn` transition that cannot map to a non-overlapping
  physical column remains fatal on the first page; after at least one page has
  completed, pagination retains only pages strictly before the rejected
  transition, discards the uncommitted current-page suffix, and attaches a
  source-bound `UNSUPPORTED_FEATURE` error diagnostic;
- a recoverable failure never invokes a legacy renderer, retries with guessed
  defaults, or weakens the validation of fatal state.

This classification is made at the narrowest layer that understands whether the
feature is optional. Catching arbitrary exceptions at the viewer boundary would
erase that distinction and could paint a structurally invalid layout.

## Worker Mode

Main-thread and worker rendering call the same layout and paint modules. Worker
mode performs parsing, font registration, layout, and paint inside the worker,
retains `DocumentLayout`, paints requested pages to `OffscreenCanvas`, and
transfers `ImageBitmap` results.

`DocumentLayout` contains only structured-clone-safe plain data: arrays, objects,
numbers, strings, booleans, and resource keys. It does not contain Canvas
contexts, functions, `WeakMap`, DOM nodes, `FontFace`, `ImageBitmap`, or live
archive handles.

Transforms and clips use numeric/discriminated plain-data records, construction
deep-freezes the result in development/tests, and clone tests cover every node
kind. Source references include a stable story-instance key (body, header/footer
part, note ID, or enclosing text-box path) plus the block path.

Search and selection geometry is projected from layout rather than collected by a
second dry render. Main and worker modes must produce the same normalized layout
fingerprint for services with identical text, image, and math resource snapshots.

The worker retains layouts by the structured-clone-safe string returned from
`layoutOptionsKey(options, services)`, including service-owned fingerprints. The
load-time default key backs synchronous page metadata. A per-call `currentDate`
selects a keyed variant without mutating the default metadata or changing public
method signatures; NUMPAGES is converged within that variant.

The optional built-in math engine is selected by a serializable stable identity
and lazy-imported in worker mode. It uses MathJax's DOM-independent adaptor,
rasterizes MathJax's path-and-rule SVG vocabulary through the same auxiliary
Canvas painter in Window and Worker, and supplies the same measured resources to
the layout service boundary used by main mode.

## Migration Strategy

The implementation uses incremental replacement, but the final state is equivalent
in structure to a clean implementation. Intermediate pull requests keep `main`
working, while releases remain blocked until all series are complete.

For each migrated feature class:

1. characterize current non-target behavior;
2. add a failing invariant or contract test;
3. route the feature exclusively through the new layout contract;
4. delete its legacy measurement, runtime stamps, fallback gate, and migration
   flag in the same pull request;
5. verify no intended visual change;
6. obtain an independent critical review and resolve every valid finding.

The migration must not add a new production switch between equivalent algorithms.
Temporary algorithm-selection adapters may preserve an internal test or caller
shape only within the pull request that removes their remaining consumer. A
permanent input-normalization boundary is allowed when unchanged public model
compatibility requires it, provided every normalized input feeds the sole layout
algorithm and the boundary contains no measurement, pagination, or paint logic.

## Delivery Series

### Series A: Body, tables, and page flow

- establish invariant and paint-purity CI;
- establish final font/text/image/math services, layout option keys, and one
  convergence primitive before feature migration;
- migrate numbered, vertical, floating-wrap, and state-dependent body paragraphs;
- remove paragraph line stamps and fragment paint gates;
- derive row sizing and cell blocks from one table measurement;
- migrate nested and floating tables;
- remove table width/height stamps and table paint gates;
- extract the page/column state machine from the renderer;
- establish main/worker layout parity.

### Series B: Stories and page layers

- migrate headers, footers, footnotes, and endnotes to shared story layout;
- preserve complete `txbxContent` block structure in the parser model;
- migrate text boxes, including nested tables and supported floating content;
- materialize behind/front drawing order in `PageLayers`;
- derive search and selection geometry from layout;
- remove callback-based z-order and dry-render geometry collection.

### Series C: Compatibility infrastructure and final consolidation

- express float placement as explicit constraints using the shared convergence
  primitive;
- isolate Office-observed compatibility rules;
- propagate unsupported-feature diagnostics from parser to layout;
- add a redistributable synthetic conformance corpus and browser geometry CI;
- reduce `renderer.ts` to the adapter boundary;
- prove that no legacy measurement, runtime layout stamp, fallback gate, or
  migration flag remains.

## Tests and Invariants

### Parser conformance

Committed tests build minimal OOXML for direct formatting, inheritance, explicit
clearing, and absence. Parser fields used by layout are tested at the Rust model
and TypeScript wire boundary.

### Layout geometry

Normalized geometry snapshots cover page count, line ranges, coordinates, table
continuations, page layers, z-order, and font resolution. They are the primary
layout fidelity gate; pixel tests are supplementary.

Required invariants include:

1. ordinary in-flow allocation bounds do not enter the bottom margin;
2. ordinary flow fragments on different pages cannot own the same allocation
   region; explicitly overlapping floats and ink overhang are modeled separately;
3. paginator cursor advancement equals fragment advance ownership;
4. measured and painted line partitions are identical;
5. measured and painted table row heights are identical;
6. scale, DPR, repaint count, and target canvas do not change pagination;
7. main and worker layout fingerprints match;
8. parser and layout objects remain unchanged during paint;
9. identical input and layout services produce identical layout output;
10. unsupported behavior produces a diagnostic rather than a silent algorithm
    switch.

### Synthetic conformance corpus

Redistributable generated fixtures cover pairwise combinations of:

- body, header, footer, note, and text-box stories;
- paragraphs, tables, and nested tables;
- inline and floating objects;
- horizontal, vertical, and bidirectional flow;
- automatic, exact, and at-least spacing;
- direct, style, inherited, and explicitly-cleared formatting;
- embedded, local, substitute, and missing fonts;
- page-, margin-, column-, paragraph-, and line-relative anchors.

Private documents remain local visual evidence and never appear in public test,
commit, pull-request, or issue text.

### Paint purity and static enforcement

Paint tests use a Canvas stub whose `measureText` throws. Static analysis rejects:

- measurement/layout algorithm imports in paint modules while allowing type-only
  imports from the plain layout result contract;
- direct or transitive runtime paths from layout to the private parser model,
  including bridge, dynamic-import, alias, and CommonJS paths, except the exact
  `layout/resources.ts` named normalization import and its AST-frozen projection;
- display-scale/DPR state in layout modules without rejecting authored OOXML text
  scale properties;
- runtime layout stamps on parser model types;
- duplicate OOXML property merge implementations;
- new production migration flags or legacy layout fallbacks.

### Browser and visual verification

Node/Skia geometry tests are the deterministic primary gate. Chrome, Firefox, and
WebKit verify browser execution and main/worker parity. Visual regression tests
remain supplementary. Native Canvas shaping may differ across browser engines,
so cross-browser checks compare authored fixed geometry after deterministic
normalization and apply semantic invariants to shaped text; text coordinates and
partitions are not numerically compared across engines unless a deterministic
shaping engine and font corpus are supplied. Main/worker parity within one browser
remains exact. Behavior-preserving
migration pull requests require no intentional visual difference.

## Pull Request Review Gate

Every pull request is reviewed by an independent agent after implementation and
local verification. The review must be critical and cover:

- OOXML specification-first behavior and cited evidence;
- single responsibility and readable module boundaries;
- duplicate layout, measurement, style, or paint logic;
- correct sharing with DOCX, XLSX, PPTX, core, and ooxml-common;
- avoidance of false cross-format abstraction;
- public API and worker-mode compatibility;
- model and layout immutability;
- tests of behavior and invariants rather than implementation details.

Valid findings are fixed, all affected verification is rerun, and review is
repeated when the correction materially changes the design. The pull request is
merge-committed only after findings and required checks are clear. Squash merge
and direct pushes to `main` are prohibited.

## Release Gate

No package release or version tag is created while any delivery series remains
incomplete. A release becomes eligible only after the final architecture audit
confirms:

1. one production layout algorithm per supported feature class;
2. no paint-time measurement or pagination;
3. no runtime layout stamps on parser objects;
4. no production migration flags or legacy fallback gates;
5. all stories and supported floating content participate in `DocumentLayout`;
6. main/worker parity and required invariants pass in CI;
7. public APIs remain backward compatible;
8. every series and its independent review are complete.

Public compatibility is proved by comparing the generated package declaration
surface with the baseline captured before migration, not by manually inspecting
a subset of source files. A transitive import-graph check proves that no paint
entry reaches measurement/layout/style-resolution code and that layout/pagination
algorithms exist only in the focused allowlisted modules.

## Acceptance Criteria

The redesign is complete when all goals, invariants, delivery series, review
gates, and release gates in this document are satisfied. Reducing the number of
known defects without deleting the dual architecture is not completion.
