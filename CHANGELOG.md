# Changelog

All notable changes to @silurus/ooxml are documented here. The project follows
semantic versioning. While the major version is zero, minor releases may contain
explicitly documented breaking changes; patch releases remain compatible with
the corresponding minor release.

## Unreleased

## 0.83.0 — 2026-08-28

Minor release making large Word documents and PowerPoint presentations useful
sooner, while improving Word review output and PowerPoint font fidelity.

- **progressive Word and PowerPoint viewing:** add opt-in `progressiveLayout`
  to the single-canvas Viewers, virtualized ScrollViewers, and document engines
  in main and worker modes. Opening pages or slides can render while the rest of
  the file is prepared, with a shared completion lifecycle and
  `waitUntilLayoutComplete()` for operations that require the final layout.
- **stable progressive interaction:** preserve the user's scroll position and
  visible content while Word pagination grows, keep the final PowerPoint slide
  count and scroll extent stable from first paint, and show loading states for
  slides that are not ready yet.
- **Word layout performance:** eliminate repeated acquisition and serialization
  work in large documents, especially for tables spanning many pages, without
  introducing a second or approximate pagination path.
- **Word tracked changes:** add an opt-in `showTrackedChanges` markup view with
  author-coloured underlines, strikethroughs, and margin change bars. The
  accepted final state remains the default.
- **PowerPoint fidelity and interaction:** use presentation-embedded fonts in
  both rendering modes and keep overflowing slide text selectable outside its
  authored shape bounds.
- **Word correctness and viewer stability:** prevent empty continuous-section
  transitions from skipping an authored page number, keep the visible page
  anchored when zoom settles, and preserve a page's logical size when worker
  bitmap resolution is constrained. Long web addresses now wrap within the
  page instead of leaving unnatural gaps or being clipped at the edge.
- **site and viewer polish:** preserve live viewers across browser back/forward
  navigation, use worker-progressive DOCX and PPTX viewers in Try Yours, and
  keep equations and other enabled optional renderers available there; also
  avoid showing a default focus outline around the XLSX canvas.
- **compatibility:** no existing option is removed or renamed, and progressive
  layout and tracked-change markup remain opt-in. No source migration is
  required. Presentations that contain embedded fonts now use them
  automatically, improving fidelity for those files.

## 0.82.1 — 2026-08-26

Compatible patch release improving Word and PowerPoint layout reliability,
comment-margin scrolling, and the bundled VS Code viewing experience without
changing the 0.82 public integration contract.

- **Word table pagination:** defer a trailing partial row that has not emitted
  content when a vertically merged row group outgrows the remaining page band,
  preventing repeated admission of the same over-page fragment. (#1373)
- **PowerPoint baseline text:** render non-zero DrawingML baseline runs at the
  Office-observed glyph scale while retaining authored line geometry, avoiding
  citation-only lines near a wrap boundary. (#1376)
- **comment margins:** omit the inter-card gap after the final built-in comment
  card so a bottom-aligned card does not create a small nested scroll range.
- **VS Code extension:** enable the first-party ChartEx, 3-D chart, and offline
  Region Map renderers for DOCX, XLSX, and PPTX webviews.
- **site and integration guidance:** preserve WASM-backed viewer initialization
  in the production site, bring framework examples to the current 0.82 line,
  and document host-controlled light and dark XLSX viewer chrome.
- **package declarations:** align the published Node entry's declaration build
  aliases with its existing typecheck boundary so the ESM package emits the
  complete Node declaration graph.
- **compatibility:** no application or API migration is required from 0.82.0.

## 0.82.0 — 2026-08-26

Minor release adding read-only review information across Word, Excel, and
PowerPoint while keeping application-owned workflows outside the Viewer.

- **built-in comments:** add opt-in DOCX and PPTX comment margins and retain the
  established XLSX cell-anchored presentation. Cards, markers, resolved-thread
  filtering, author colors, zoom, RTL placement, and keyboard access share one
  cross-format interaction model.
- **application-owned comment UI:** expose format-scoped comment retrieval,
  target geometry, selection context, and `goToComment()` navigation for DOCX
  pages, XLSX sheets, and PPTX slides. Applications can build independent
  lists or fully custom overlays without adopting the built-in cards.
- **Word tracked changes:** retain body-story insertion, deletion, and move
  records as detached data while rendering the accepted final state. Tracked
  changes remain separate from comments and have no built-in markup UI.
- **integration boundary:** keep comment parsers, models, and primitive APIs in
  the base format entries while loading the built-in DOM presentation only when
  needed. Stable classes and CSS custom properties cover ordinary visual
  customization; custom structure remains application-owned.
- **viewer composition:** share the focused main/worker paint boundary between
  the single-canvas and virtualized DOCX/PPTX viewers, preserving the primitive
  Viewer foundation for custom scroll or presentation shells.
- **performance and reliability:** cache virtual-scroll geometry, bound sparse
  DOCX anchor resolution, deduplicate PPTX target queries, and make asynchronous
  XLSX/PPTX comment navigation fail closed and latest-request-wins.
- **documentation:** publish live built-in and custom-list examples for all
  three formats, symmetric review-data API guidance, and current production
  bundle measurements.

## 0.81.0 — 2026-08-25

Minor release expanding Microsoft ChartEx support and improving chart fidelity
across Word, Excel, and PowerPoint. Applications that display ChartEx charts
must opt in to the new module; classic charts remain built in.

- **ChartEx migration:** move waterfall, histogram, Pareto, funnel,
  box-and-whisker, treemap, and sunburst rendering to the tree-shakeable
  `@silurus/ooxml/chart-ex` entry. Pass `chartEx` to DOCX, XLSX, or PPTX viewers
  that need these chart families.
- **chart fidelity:** preserve more authored Chart Style, Chart Color, fill,
  line, marker, axis, legend, label, data-table, visible-source, and plot-area
  semantics through the shared classic and ChartEx pipeline.
- **3-D charts:** improve Bubble3D material and series precedence, axis and
  gridline ownership, Surface wireframes, and projected picture fills,
  including crop, tile, stack, wall, and reversed-axis behavior.
- **host integration:** apply the shared chart behavior consistently in DOCX,
  XLSX, and PPTX main and worker rendering, while keeping ChartEx outside the
  default main-mode import closure.
- **Word layout:** exclude truly empty unnumbered cell paragraphs from AutoFit
  content width while retaining authored spaces and non-breaking spaces, and
  end horizontal snap-grid underlines at the terminal glyph ink.
- **documentation:** add the ChartEx migration announcement and maintain current
  bundle-size measurements on the independent bundle-size page.

## 0.80.3 — 2026-08-25

Patch release improving Word list resilience and production worker compatibility
without changing the 0.80 public integration contract.

- **Word numbering resilience:** keep list bodies at the authored paragraph
  start when `nothing` or `space` suffixes remain inside a hanging-indent region,
  avoiding a non-negative layout invariant failure. (#1366)
- **production worker compatibility:** resolve the optional MathJax asset in the
  main realm so emitted render-worker assets contain no inner `import.meta`,
  including when rebundled by Next.js 14 and webpack. (#1367)
- **compatibility:** no application or API migration is required from 0.80.2.

## 0.80.2 — 2026-08-18

Compatible patch release improving Office document fidelity and navigation
without changing the 0.80 public integration contract.

- **Excel chart fidelity:** improve authored date axes, grouped and stacked
  bar scales, 3-D Surface geometry and materials, classic tick placement,
  theme-derived axis widths, and multi-level category separators.
- **Excel navigation:** follow internal hyperlinks to direct cell references,
  ranges, and in-scope defined names, switching sheets when necessary.
- **Word pagination resilience:** preserve completed pages when an unsupported
  later section flow cannot be placed, while retaining a source-bound
  diagnostic for the unsupported content.
- **compatibility:** no application or API migration is required from 0.80.1.

## 0.80.1 — 2026-08-17

Patch. Corrects numeric axis-label spacing across authored chart families
without changing the 0.80 public integration contract.

- **scatter charts:** place numeric labels beyond authored outward tick marks
  instead of allowing major ticks to enter or touch adjacent glyphs.
- **3-D charts:** keep value and category labels clear of projected tick marks
  while preserving the established Office-compatible axis spacing.
- **ChartEx axes:** use an axis-relative label offset for histogram,
  box-and-whisker, and Pareto value axes instead of scaling the offset with the
  label font size.
- **compatibility:** no application or API migration is required from 0.80.0.

## 0.80.0 — 2026-08-16

Compatible minor release. Completes built-in renderer parity between the
existing main-thread and worker rendering modes across DOCX, XLSX, and PPTX.

- **worker renderer parity:** equations, authored 3-D charts, and country-level
  Region Maps now use the same `math`, `threeD`, and `regionMap` injection
  options in main and worker modes.
- **production-safe workers:** publish self-contained render workers and carry
  consumer-resolved assets such as MathJax across the worker boundary, including
  applications that rebundle the packages with Vite.
- **equation quality:** use one bounded Canvas raster path in Window and Worker
  contexts, with staged high-resolution reduction for cleaner equation output.
- **integration guidance:** add worker file-upload examples and document how to
  choose between main and worker rendering, including responsiveness, download,
  frame-transfer, browser-support, and custom-renderer trade-offs.
- **compatibility:** main mode remains the default and existing applications do
  not require migration. DOCX content requiring browser-only vertical-glyph
  selection automatically falls back to effective main mode for correct shaping.

## 0.79.1 — 2026-08-15

Patch. Corrects authored chart labels, legend frames, and automatic worksheet
row heights without changing the 0.79 public integration contract.

- **data labels:** keep series-local label settings on their authored series,
  and apply value-axis display units consistently to generated value labels in
  classic, ChartEx, and optional 3-D chart paths.
- **trendline labels:** align automatic equation and R² blocks with Excel's
  plot-relative placement while preserving authored manual layouts.
- **legend frames:** retain explicit solid legend fills and outlines, including
  DrawingML theme color transforms, across DOCX, XLSX, and PPTX chart hosts.
- **worksheet row heights:** auto-fit rows that omit an authored `ht` to wrapped,
  rich, rotated, or larger text while preserving explicit heights and excluding
  merged cells, matching Excel's display behavior. (#1260)
- **compatibility:** no application or API migration is required from 0.79.0.

## 0.79.0 — 2026-08-14

Compatible minor release. Adds opt-in advanced chart renderers and completes a
specification-first fidelity audit across charts hosted by Word, Excel, and
PowerPoint.

- **optional 3-D charts:** add the tree-shakeable `@silurus/ooxml/three-d`
  entry, using one homogeneous model-space camera for chart walls, axes, grids,
  labels, and bounded meshes. It supports cartesian and pie 3-D families plus
  authored box, cylinder, cone, cone-to-max, pyramid, and pyramid-to-max bar
  shapes. Applications opt in on the main thread; existing integrations retain
  the 2-D fallback.
- **offline Region Maps:** add the tree-shakeable
  `@silurus/ooxml/region-map` entry for country-level ChartEx world maps using
  pinned public-domain Natural Earth geometry. Authored projections, theme
  colors, and two/three-stop value ramps are retained; unsupported cached or
  sub-country views fail closed without network access.
- **shared chart fidelity:** preserve more authored axis, title, rich-label,
  legend, point, line, fill, and Chart Style properties through the shared
  parser/model/renderer pipeline. Automatic linear axes, explicit bounds,
  major/minor ticks, plot margins, secondary axes, and legend layout now use
  common bounded policies across classic and ChartEx families.
- **specialized chart families:** improve waterfall connectors and labels,
  Pareto and histogram axes, box-and-whisker geometry, treemap and sunburst
  hierarchy labels, funnel, bubble, line, area, combo, and chart-sheet layout.
  XLSX chart sheets and absolute chart anchors are retained.
- **bounded work and package boundaries:** cap expanded 3-D primitives,
  hierarchy/data-label work, map rows, ticks, and parser caches. The 3-D camera,
  mesh code, and geographic asset remain outside ordinary DOCX/XLSX/PPTX entry
  graphs unless explicitly imported.
- **site and compatibility:** the API reference and Try Yours demo expose the
  optional renderers consistently across all three host formats. No existing
  option is removed or renamed; add-ons are required only for their authored
  chart families.

## 0.78.1 — 2026-08-12

Patch. Fixes a DOCX pagination regression introduced in v0.78.0 without
changing the 0.78 public API.

- **Word table pagination:** keep text in narrow table cells on the same pages
  as Word when full-width spaces participate in East Asian line layout. This
  prevents the end of a label from moving unexpectedly to the next page.
- **authored paragraph spacing:** preserve intentional blank space created by
  consecutive full-width spaces instead of collapsing it during line layout.
- **compatibility:** no migration or application changes are required when
  upgrading from v0.78.0. (#1226)

## 0.78.0 — 2026-08-12

Compatible minor release. Improves chart rendering across Office formats and
adds Excel-style multiple-area selection to worksheets without requiring
application changes.

- **chart fidelity:** render authored axes, ticks, chart and plot backgrounds,
  labels, legends, pattern fills, percentage stacking, and series geometry more
  like Office across DOCX, XLSX, and PPTX. (#1196, #1219)
- **modern charts:** improve waterfall, box-and-whisker, treemap, bubble, pie,
  scatter, area, and mixed chart layouts, including linked Chart Styles and
  Chart Colors. (#1196, #1219)
- **XLSX multiple selection:** use Ctrl on Windows/Linux or Command on macOS to
  add cell, row, or column areas while preserving the active cell and configured
  selection color without doubled borders between adjacent areas. (#1194, #1219)
- **worksheet interaction:** keep overflowing cell text visible across the
  virtualized viewport and improve outline controls, collapse scrolling, frozen
  pane boundaries, and focus presentation. (#1219)
- **compatibility:** no migration is required; existing viewer setup,
  single-area selection, and public APIs remain supported.

## 0.77.1 — 2026-08-11

Patch. Improves shared DrawingML fidelity and PowerPoint rendering without
changing the 0.77 public API.

- **shared DrawingML geometry and styles:** evaluate custom-geometry guides and
  quadratic Bézier paths consistently in DOCX, XLSX, and PPTX; preserve theme
  fill and line recipes, chart color-map overrides, and complete linear/path
  gradient geometry instead of flattening them during parsing. Theme fragment,
  relationship, raster-tile, and projected-effect work remains bounded for
  untrusted documents. (#1209, #1210, #1212, #1213)
- **PowerPoint themes, charts, and tables:** resolve theme backgrounds, image
  fills, tint/shade transforms, hidden layout graphics, chart series and axis
  defaults, table banding, text roles, bullets, borders, gradients, and vertical
  cell content more like PowerPoint. (#1188–#1202, #1204, #1206, #1212, #1213)
- **PowerPoint effects and 3-D:** compose shadows, reflections, soft edges,
  callout leaders, arrowheads, picture effects, bevels, extrusion colors, and
  projected 3-D silhouettes from the complete painted object. Effect surfaces
  use conservative bounds and allocation limits, and floor reflections keep a
  sharp contact edge while blur increases with distance. (#1203, #1205,
  #1207–#1211, #1214–#1217)
- **layout correctness and verification:** ignore terminal paragraph whitespace
  that would create a spurious continuation line, preserve inherited bullet and
  auto-number formatting, and verify the cross-format renderer changes against
  the complete private regression corpus using v0.77.0 as the baseline.
  (#1193, #1198, #1204, #1213)

## 0.77.0 — 2026-08-10

Breaking minor release. Unifies read-only selection and integration context
across DOCX, XLSX, and PPTX; lets applications identify selected document
elements and context-menu targets; and connects the active VS Code preview to
the OOXML MCP server. See the
[v0.77 migration guide](https://ooxml.silurus.dev/announcements/v077-migration-guide/)
for every removed or renamed API.

- **Excel-compatible XLSX selection:** replace the lossy range-only
  compatibility API with `setSelection()`, `selectionState`, and
  `onSelectionStateChange()`. The canonical state keeps selected areas,
  ActiveCell, and the Shift-extension anchor separate, supports multiple areas,
  and preserves row, column, whole-sheet, RTL, clipboard, keyboard, pointer,
  and programmatic behavior. (#1176)
- **cross-format selection context:** expose bounded, detached text, cell-range,
  and chart/picture/shape context through `getSelectionContext()` and the common
  `onSelectionContextChange` callback. Optional `enableElementSelection` adds
  read-only object selection with a visible outline in DOCX, XLSX, and PPTX;
  `onContextMenu` provides the original browser event plus a lazy asynchronous
  target lookup. The official site includes matching three-format demos and
  browser-integration guidance. (#1181, #1182)
- **VS Code and MCP integration:** add an authenticated loopback bridge from the
  active OOXML preview to one `ooxml_get_active_context` tool, including local
  document identity and bounded selected content. Consolidate strict-subset
  DOCX, XLSX, and PPTX tools, validate the bridge payload and lifecycle, and
  document the Copilot-specific active-context path separately from standalone
  path-based MCP use. (#1182)
- **predictable asynchronous errors:** Promise-returning Viewer and headless
  operations now reject their Promise on failure even when `onError` is set;
  `onError` remains for later Viewer-managed work that has no Promise to await.
  Initial scroll-window and presentation-media work follows the same rule.
  (#1182)
- **public API cleanup:** remove the legacy XLSX selection names, no-op rendering
  options, internal worker transport types, and exact aliases that duplicated a
  canonical type. Narrow borrowed and operation-specific options, export every
  reachable public type, forward Viewer passwords correctly, and align the Node
  and Markdown input contracts with their documented binary inputs. Migration
  tables cover each source change without temporary duplicate APIs. (#1182)
- **package guidance and regression evidence:** clarify npm unpacked size versus
  the assets reachable from each format entry, keep the optional MathJax engine
  separately lazy-loaded, and verify the complete private DOCX/XLSX/PPTX corpus
  against the previous renderer with pixel-identical output. (#1180, #1182)

## 0.76.2 — 2026-08-08

Patch. Improves Word-compatible text and table layout, extends spreadsheet
selection beyond the visible viewport, and hardens packaging and public demos
without changing the 0.76 public API.

- **docx typography and tabs:** implement document-grid character pitch,
  tab-stop, decimal-tab fallback, and right-indent adjustment through the
  canonical retained layout pipeline; retain Normal-style typography and font
  alternate-name metadata; and honor `balanceSingleByteDoubleByteWidth` for the
  specification-defined single/double-byte ratio plus narrowly registered
  Office behavior for ASCII/CJK grid deltas and authored space sequences.
  Measurement, wrapping, intrinsic sizing, selection geometry, RTL placement,
  and paint now consume the same retained width authority. (ECMA-376
  §§17.3.1.1, 17.3.1.37, 17.6.5, 17.8.3.1 and 17.15.3.3; MS-OI29500
  §§2.1.534 and 2.1.556; MS-OE376 §2.1.236; #1177, #1178)
- **docx retained layout:** finalize bookmark metadata only after every page
  story joins the retained graph, and preserve signed table-indent spill while
  exact row height clips only the block axis. (§§17.4.50, 17.4.80 and
  17.13.6.2; #1174, #1175)
- **xlsx interaction:** show overflowing worksheet scrollbars by default, allow
  pointer drag selections to auto-scroll beyond the visible viewport, preserve
  RTL and row/column selection semantics, and bind deferred selection updates to
  the initiating pointer. The multi-window public demo now follows the parent
  light/dark theme. (#1170, #1172, #1173)
- **VS Code packaging:** exclude stale source maps at the VSIX package boundary
  so reused development output cannot unexpectedly double the published
  extension size. (#1171)

## 0.76.1 — 2026-08-07

Patch. Improves Word-compatible legacy layout, stabilizes public-site demos,
and repairs the release workflows without changing the 0.76 public API.

- **docx fidelity:** retain section properties on break-only paragraphs; resolve
  the complete DrawingML style-matrix color choice; preserve absolute VML image
  anchors, CSS declaration precedence, wrap distances, coordinate ownership,
  and signed stacking; render legacy form-checkbox state and sizing; and allow
  authored fixed nested tables to overflow a containing cell when Word does.
  (ECMA-376 §§17.16.29–30, 17.18.87, 19.1.2.19, 20.1.4.1.7 and
  20.1.4.1.30; MS-OE376 §§2.1.505, 2.1.1692, 2.1.1701 and 2.1.1710;
  #1167, #1168)
- **viewer and website stability:** keep the first-page leading padding while
  paginated DOCX/PPTX demos start at scroll position zero, restore homepage
  viewers after browser history navigation, normalize external-link arrows on
  mobile, prevent announcement pages from causing horizontal overflow, and
  document that loading ownership and main/worker execution mode are independent
  choices. (#1164, #1165, #1166)
- **release infrastructure:** isolate VS Code extension package builds from stale
  workspace output and update the GitHub Pages artifact upload action for the
  current runner runtime. (#1162, #1163)

## 0.76.0 — 2026-08-06

Breaking minor release. Unifies Browser Viewer ownership, replaces the legacy
synchronous Node parser helpers with one bounded asynchronous pipeline, and adds
an independently mountable XLSX sheet surface without changing the ordinary
Browser `new Viewer(target)` plus `await viewer.load(source)` flow. See the
[0.75 to 0.76 migration guide](docs/migration-0.76.md).

- **browser Viewer architecture:** add `XlsxSheetViewer`, a caller-canvas active
  worksheet viewport with selection, search, zoom, logical scrolling, and
  independent sheet navigation. The existing `XlsxViewer` retains its complete
  workbook UI and now composes the same acquisition, geometry, viewport,
  selection, render-dispatch, and surface implementation. DOCX and PPTX single
  and scroll Viewers likewise share the same Canvas lifecycle and static render
  dispatch mechanics without merging their format-specific layout semantics.
  (#1155, #1159)
- **explicit engine reuse:** replace Viewer constructor-option injection with
  synchronous named factories: `fromDocument()`, `fromPresentation()`, and
  `fromWorkbook()`. The normal `viewer.load(source)` path remains the simple,
  Viewer-owned default; factories provide an explicit caller-owned path for
  master/detail, multi-pane, and multi-window rendering. Factory return types
  omit `load()`, conflicting load options are rejected, and borrowed engines are
  never destroyed by their Viewers. A declaration-level symmetry check guards
  these ownership and naming contracts across DOCX, PPTX, and XLSX. (#1160)
- **node parser API:** remove the synchronous `parseDocx()`, `parsePptx()`,
  `parseXlsx()`, `parseXlsxSheet()`, `parseXlsxAllSheets()`, and standalone PPTX
  media extraction compatibility paths. New `materialize*()` helpers preserve
  convenient caller-owned results, while `openDocxDocument()`,
  `openPptxPresentation()`, and `openXlsxWorkbook()` provide bounded sequential
  sessions with cancellation, metrics, archive reuse, deterministic cleanup,
  and runtime recovery through the same canonical acquisition pipeline used by
  Browser rendering. (#1134, #1155)
- **resource governance:** add the shared `resourceLimits.maxArchiveEntries`
  admission policy across Browser main mode, workers, and Node sessions. The
  calibrated default is 4,096 entries, `null` disables only the configurable
  limit, and the non-configurable 20,000-entry structural ceiling remains in
  force. Violations use the existing typed resource-limit error and metrics
  vocabulary. (#1133, #1155)
- **xlsx fidelity and interaction:** localize the built-in short-date format,
  reuse date formatters and streamed worksheet metrics, center wrapped values
  that remain one rendered line, and mirror text overflow, footer controls, and
  sheet-tab ordering/alignment for RTL workbooks. The public site demonstrates
  one parsed workbook rendered as independently scrollable sheets, including
  same-origin popup windows. (#1156, #1157, #1158, #1159)
- **tooling and documentation:** build the libraries with TypeScript 7 while
  isolating the temporary TypeScript 6 Compiler API dependency used by
  declaration tooling. Publish English and Japanese architecture baselines, a
  complete migration guide and site announcement, JavaScript Office Viewer SEO
  metadata, sitemap and robots endpoints, refreshed public API references, and
  updated README screenshots. (#1154, #1155, #1160)

## 0.75.5 — 2026-08-05

Patch. Improves Word-compatible paragraph layout and makes framework examples
usable as both embedded source references and top-level live projects.

- **docx pagination:** carry an explicit `splittable` / `indivisible`
  fragmentation contract from paragraph acquisition through page selection.
  Marker-only and otherwise indivisible paragraphs now paginate atomically,
  while incomplete retained line boundaries fail closed instead of being
  guessed by the renderer. (#1151)
- **docx text:** resolve numbering marker fonts and colours once from numbering
  level and paragraph-mark properties for both body text and text boxes; retain
  the selected font geometry for run highlights; and give positive authored
  character spacing precedence over document-level punctuation compression.
  (§§17.3.2.15, 17.3.2.23, 17.9.24, 17.15.1.18; observed Word inheritance and
  pitch precedence; #1151)
- **website / examples:** present embedded StackBlitz projects as code browsers
  without auto-starting WebContainers, keep top-level links runnable, and use
  the standard shared footer on the framework chooser, framework guides, and
  format pages. (#1150)

## 0.75.4 — 2026-08-04

Patch. Improves Word DrawingML fidelity, safely handles malformed style
inheritance, and adds runnable TypeScript integration guides.

- **docx DrawingML:** preserve inline WordprocessingShape objects through
  parsing, immutable layout, pagination, and paint; retain authored image fills
  through the shared DrawingML fill model; and carry signed run positioning and
  resolved page-field metrics without paint-time inference. (§§20.4.2.7–8,
  20.1.8.14; #1146)
- **docx styles:** detect circular `basedOn` inheritance by style ID and apply
  every unique style once before terminating a malformed revisiting edge. Valid
  paragraph, character, and table style chains remain unbounded by arbitrary
  depth caps, while invalid input can no longer overflow the WASM stack.
  (§17.7.4.3; MS-OI29500 §2.1.233; #1148)
- **website / examples:** publish isolated TypeScript projects and independent,
  search-oriented StackBlitz guides for React, Vue, Svelte, and Solid. The
  public navigation now points to one framework chooser, redundant per-format
  snippets and Storybook links are removed, DOCX teardown guidance uses the
  current `destroy()` lifecycle, and npm discovery includes `office-viewer` and
  `office` keywords. (#1146, #1147)

## 0.75.3 — 2026-08-04

Patch. Restores Word-compatible DOCX pagination at signed `atLeast` spacing
boundaries and keeps scheduled ZIP extraction fuzzing buildable.

- **docx layout:** preserve Word's signed `atLeast` boundary for content-less
  paragraph marks when Far East layout is enabled on an active document grid.
  Negative values retain their authored magnitude, zero follows ordinary
  `atLeast` spacing, and positive values continue to allocate whole grid cells,
  without changing non-grid or `snapToGrid=false` layout. (§§17.3.1.33,
  17.6.5; observed Word behavior; #1144)
- **fuzzing:** align the independent ZIP extraction fuzz target with the
  resource-governed package API, retain the package handle for entry reads, and
  add ordinary CI format and Clippy gates so parser API drift is detected before
  scheduled campaigns. (#1143)

## 0.75.2 — 2026-08-04

Patch. Improves Word-compatible DOCX text metrics, anchored drawing flow,
table pagination, and compound border rendering.

- **docx text:** retain underline decoration across tab advances, preserve CJK
  fallback for fixed-pitch East Asian faces, and resolve Far East layout-grid
  metrics from the East Asia font axis for empty paragraph marks and visible
  Latin or CJK text. (§§17.3.1.37, 17.6.5; observed Word `useFELayout`
  behavior; #1141)
- **docx tables / drawings:** keep overlap-permitted, non-wrapping in-cell
  drawings out of automatic row containment, include winning collapsed-border
  half-rules in non-exact row tracks, and reserve page-local boundary rules
  when selecting split rows. (§§17.4.43, 17.4.66, 17.4.80, 20.4.2.3;
  observed Word row-footprint behavior; #1141)
- **docx borders:** retain compound `ST_Border` rails and paint table outer
  borders as joined closed frames, including continuous corner connections,
  without inferring table structure during the paint pass. (§17.18.2; #1141)

## 0.75.1 — 2026-08-03

Patch. Improves PowerPoint chart and text layout fidelity, keeps spreadsheet
footer navigation stable across zoom levels, and refines the website preview
experience.

- **pptx / charts:** honor authored inner manual plot rectangles instead of
  clamping them to automatic chart gutters, and use PowerPoint's default
  automatic scale target for a deleted value axis whose ticks are not visible.
  This keeps stacked horizontal bars aligned with separately authored labels
  and other slide content. (§21.2.2.89; observed PowerPoint scale behavior;
  #1136, #1138)
- **pptx text:** inherit placeholder text insets from layouts and masters
  attribute by attribute, and exclude a final paragraph's trailing spacing from
  the multi-column overflow decision. This preserves authored page margins,
  wrapping, and column placement. (§21.1.2.1.1; Annex L.3.2.3; #1136)
- **xlsx:** keep the sheet-tab previous/next navigation area at its 100% row
  header width, independently of the worksheet cell zoom. (#1137)
- **website:** retain preloaded home-page viewers across format switches, add
  selectable DOCX/PPTX text layers, keep preview frames and wide content
  responsive, avoid unrelated font downloads during parser prewarming, and
  improve progress feedback and mobile alignment. (#1135) The stacked mobile
  hero icon also retains its full colour in light mode instead of inheriting
  the tablet-layout translucency.

## 0.75.0 — 2026-08-03

Minor. Makes bounded resource use, typed limit failures, and observable usage
part of the public DOCX, XLSX, and PPTX contract while retaining the existing
Viewer construction and `load()` interfaces. It also publishes a redesigned
documentation site and restores two specification-driven DOCX layout cases.

- **core / docx / xlsx / pptx:** add one shared `resourceLimits` policy with
  conservative defaults of 128 MiB for one inflated package part and 256 MiB
  across distinct parts read during a package session. Proven violations reject
  with structured `OoxmlResourceLimitError` details; non-configurable hard
  ceilings continue to guard structural models, caches, renderer indexes, XML
  content, serialization, and decoded images. The former positive
  `maxZipEntryBytes` option remains a deprecated compatibility alias. This also
  incorporates the retained-session budget, poisoning, and actual-inflation
  invariants proposed in #1119. (#1102, #1120, #1130)
- **core / diagnostics:** expose content-free `OoxmlResourceMetrics` through
  `onResourceMetrics` for initial loads and `getResourceMetrics()` for a fresh
  snapshot after lazy work. `debug: true` prints the same measurements as a
  single typography-safe console card; collection does not require debug mode.
  Recognized residual WASM traps remain conservatively classified as
  `parser-crashed`, because panic, allocation failure, stack overflow, and
  explicit `unreachable` can lose their distinct causes at the current WASM
  boundary. (#1102, #1120)
- **xlsx:** stream worksheet XML as acknowledged, complete-row units instead of
  retaining a full-part XML tree. Existing materializing APIs remain compatible
  and are protected by separate model and renderer ceilings. (#1102, #1120)
- **docx / pptx:** stream bounded document units and slides through shared pull
  sessions while preserving DOCX's required sequential pagination and compact
  PPTX bootstrap facts. Raw parts, decoded images, derived images, and slide or
  page caches now have explicit owners and bounded lifecycles. (#1120, #1130)
- **node:** publish uniformly owned `openDocxDocument()`,
  `openXlsxWorkbook()`, and `openPptxPresentation()` sessions with explicit
  `close()` lifecycles, cancellation, resource policy, and metrics. Streaming
  iterators reuse one retained package; canvas rendering continues to require a
  caller-supplied backend. (#1120, #1130)
- **workers:** destroy partially constructed engines after a rejected load, or
  terminate the worker directly when no engine exists, across DOCX, XLSX, and
  PPTX while preserving the original failure. This incorporates and generalizes
  the leak fix proposed in #1124. (#1120, #1130)
- **docx:** measure nested AutoFit tables from intrinsic cell content before
  fitting the parent grid, retain parent-cell coordinates across page slices,
  and account for downward run-position overflow when advancing drop-cap flow.
  (ECMA-376 §§17.18.87, 17.3.1.11, 17.3.2.24; #1129)
- **xlsx:** stop drawing a renderer-owned outside border around the worksheet;
  host applications can frame the target container without a clipped or doubled
  edge around the sheet controls. (#1131)
- **website / docs:** replace the former showcase with a responsive light/dark
  documentation system, migration announcements, complete error-handling and
  deprecation references, clearer API descriptions, production resource
  guidance, and styled inline code. The published Viewer examples and Try Yours
  surfaces retain local-only file processing. (#1131)

## 0.74.6 — 2026-08-01

Patch. Prevents PowerPoint slides from becoming vertically compressed in
continuous-scroll previews after rapid zooming and subsequent scrolling.

- **pptx / vscode:** clear transient preview sizing when recycling a pooled
  slide canvas, so a canvas reused after a pinch or wheel zoom returns to the
  renderer's natural aspect ratio instead of retaining a stale preview height.
  Add regression coverage for zoom-burst recycling and remounting. (#1127)

## 0.74.5 — 2026-07-31

Patch. Restores inherited PowerPoint placeholder geometry and authored text
spacing, enables presentation media in VS Code previews, and improves find and
loading feedback across the viewers.

- **pptx parser:** inherit a placeholder's position and extent from the matching
  layout or master placeholder when the slide-level shape omits its transform,
  while preserving any geometry authored directly on the slide. (#1121)
- **pptx / vscode:** make embedded audio and video interactive in VS Code
  presentation previews, with bounded media lifecycles and graceful handling of
  unsupported or failed media sources. (#1122)
- **pptx / xlsx:** honor explicitly authored DrawingML percentage line spacing
  independently of substituted-font design-height floors. Implicit single
  spacing retains the compatibility floor for tall-metric fonts such as Meiryo,
  while `spcPct` and absolute point spacing follow the document value.
  (§21.1.2.2.5, §21.1.2.2.11; #1123)
- **viewers / vscode:** expose separate CSS colors for ordinary and active find
  matches across DOCX, XLSX, and PPTX. VS Code previews map the colors to
  `ooxmlViewer.findMatchBackground` and
  `ooxmlViewer.findActiveMatchBackground` theme contributions. (#1125)
- **website:** show an accessible progress indicator in Try Yours while a file
  is being parsed and rendered, without replacing the preview surface or
  allowing an older load to overwrite a newer one. (#1125)

## 0.74.4 — 2026-07-31

Patch. Corrects Word measure and table-border semantics, and completes the
continuous-scroll, zoom, selection, and find experience in the website and
VS Code previews.

- **docx parser:** accept absolute-unit members such as `12pt` and `3.6mm` in
  half-point and twips measures, while keeping unsigned font-size and kerning
  values distinct from signed text-position values. Invalid direct values now
  preserve style inheritance instead of collapsing the effective size to zero.
  (§17.18.42, §17.18.81, §22.9.2.14; #1110, #1116)
- **docx:** resolve `nil` and `none` shared-cell borders according to the
  ECMA-376 conflict algorithm, evaluate each segment beside a horizontally or
  vertically merged cell independently, and use the terminal continuation
  cell for a vertical merge's bottom edge. (§17.4.66; #1114)
- **website:** use `DocxScrollViewer` and `PptxScrollViewer` directly in Try
  Yours, retaining selectable text and browser-native find while adding
  trackpad pinch / Ctrl-or-Cmd-wheel zoom. The preview starts at a fixed 90%
  natural scale with a 50% minimum, and PPTX media remains interactive only
  near the real viewport. (#1115)
- **vscode:** preserve a readable natural DOCX scale, fit PPTX slides inside
  the preview width with side gutters, and capture webview trackpad pinch
  gestures. Add a case-insensitive Ctrl+F / Cmd+F popup for DOCX, XLSX, and
  PPTX with next/previous navigation and Escape / close-button dismissal.
  (#1112, #1113, #1117)

## 0.74.3 — 2026-07-30

Patch. Restores equations in VS Code previews and adds consistent continuous
scrolling and zoom controls.

- **vscode:** retain the bundled MathJax and STIX Two Math engine so equations
  render in DOCX, XLSX, and PPTX previews instead of being removed as an
  apparently side-effect-free import. DOCX and PPTX now use their virtualized
  continuous-scroll viewers, while all three formats share zoom-out, current
  percentage, and zoom-in controls. (#1109)

## 0.74.2 — 2026-07-27

Patch. Prevents contextually kerned Japanese closing punctuation from collapsing
onto the following glyph.

- **docx:** constrain document-level punctuation compression by the contextual
  shaped-cluster advance as well as adjacent tight ink. When Canvas has already
  kerned a closing punctuation pair to the retained half-cell, the layout engine
  no longer subtracts the isolated glyph's sidebearing a second time. Paint
  remains measurement-free and consumes the corrected immutable layout.
  (ECMA-376 §17.15.1.18; [MS-OE376] §2.1.562; #1107)

## 0.74.1 — 2026-07-27

Patch. Restores authored rotation for descendants of reflected DrawingML groups.

- **ooxml:** reverse child rotation direction when composing a parent group
  with exactly one reflection, including nested groups and leaf shapes. This
  restores mirrored decorations and other rotated descendants across PPTX,
  DOCX, and XLSX while preserving the shared Annex L transform path.
  (ECMA-376 Part 1 Annex L §L.4.7.4–§L.4.7.6; #1105)

## 0.74.0 — 2026-07-27

Minor. Exposes saved XLSX pivot-table metadata for read-only inspection and
improves authored chart, slicer, and DrawingML fidelity across XLSX and PPTX.
The release also strengthens DOCX punctuation-collision regression coverage and
documents the project's read-only scope.

- **xlsx:** expose immutable saved pivot-table identity, placement, field,
  cache-source, status, and diagnostic facts through
  `Worksheet.pivotTables` / `pivotDiagnostics`. Saved worksheet cells and
  styles remain authoritative; refresh, recalculation, filtering,
  restructuring, and interactivity remain out of scope. Validate package
  relationship facts strictly and report malformed or unsupported cache links
  without trusting suffix-matched relationship types. (ECMA-376 §18.10; #997)
- **xlsx:** resolve custom slicer frame, header, and item-state styles from
  SpreadsheetML table-style DXFs and the Office 2010 slicer-style extension.
  (#1101)
- **charts:** preserve authored series styling, percent-axis display units and
  label sizing, and manual-layout axis-label gutters instead of clipping labels
  to the plot frame. (#1099–#1101)
- **pptx:** render gradient shape strokes and preserve the projected extent of
  DrawingML 3D camera scenes. (#1101)
- **docx:** add an end-to-end paint regression that protects the retained
  layout clearance between compressed closing punctuation and a following
  glyph with left overhang. (#1103)
- **docs:** define the repository as a read-only viewer project and clarify the
  boundary between supported inspection interactions and out-of-scope editing
  or round-tripping. (#1098)

## 0.73.2 — 2026-07-24

Patch. Restores DOCX list-marker, table-pagination, and grouped-drawing geometry
that regressed during the retained-layout migration.

- **docx:** anchor right-to-left numbering markers to the retained aligned line
  instead of the page edge, so bullets and item numbers stay attached to their
  paragraph text.
- **docx:** admit a merged table when its resolved row tracks fit the remaining
  page, avoiding continuation pages that contain only overflowed gauge or bar
  rows.
- **docx:** reproduce Word's observed non-uniform scaling for directly grouped
  leaves at exact odd quarter turns while keeping nested groups and the shared
  DOCX/XLSX/PPTX DrawingML transform on the normative Annex L path. The
  compatibility rule is deliberately scoped to the evidence available in Word.
  (ECMA-376 Part 1 Annex L §L.4.7.4; [MS-OE376] §2.1.1360; #1037, #1096)

## 0.73.1 — 2026-07-24

Patch. Prevents compressed full-width punctuation from colliding with the
following glyph in horizontal DOCX text.

- **docx:** bound document-level full-width punctuation compression by both
  adjacent glyph ink extents. A following glyph that extends left of its
  advance origin can no longer overlap a closing parenthesis or other eligible
  punctuation, including across source-run formatting boundaries. The
  collision constraint is resolved once during layout with a linear grapheme
  pass; paint remains measurement-free and vertical text retains its separate
  geometry path. (§17.15.1.18, §17.18.7; [MS-OE376] §2.1.562; #1094)

## 0.73.0 — 2026-07-24

Minor. Completes the DOCX migration to one immutable layout-to-paint pipeline,
then restores the full fidelity corpus through that architecture rather than
reintroducing renderer-side measurement or legacy fallbacks. The release also
adds bounded partial rendering for unsupported optional drawings, exposes stable
source identities in text-run callbacks, and moves normal type-checking to
TypeScript 7.

**docx — retained layout architecture and fidelity**

- Replace renderer-owned pagination and reduced-state fallback paths with one
  retained, point-space `DocumentLayout`: immutable page/column transitions,
  occurrence-owned paragraphs/tables/drawings, canonical page stories and
  layers, deterministic convergence, and identical main-thread/worker layout
  inputs. Final architecture and compatibility gates now fail closed if an
  alternate layout path is introduced. Public DOCX declarations and worker
  contracts remain compatible. (#1037, #1038–#1077)
- Restore the parser and layout facts needed for Word-compatible pagination and
  paint after the cutover: field/tab boundaries, numbering-indent provenance,
  table/frame/border geometry, text-box and page-owned anchors, inherited shape
  text color, section decorations, East Asian line metrics, and selected-font
  ink. Canvas paint consumes retained geometry without re-measuring it.
  (#1080, #1083, #1084)
- Keep OpenType `vert` acquisition and painting on the same Canvas/font route;
  retain vertical glyph advances and origins; avoid applying negative pitch
  twice to brackets, punctuation, question/exclamation marks, long marks,
  middle dots, and ellipses. Preserve Japanese punctuation/kana compression per
  grapheme through wrapping, continuation, intrinsic measurement, and paint.
  (§17.3.2.10, §17.6.20, §17.15.1.18; #1086, #1091, #1092)
- Reduce repeated retained-layout work through immutable paragraph projection
  caches and dense wrap-polygon memoization. Shared deeply frozen polygons now
  compile once at O(V²) plus O(PV) validation instead of O(PV²), eliminating
  the former O(V³) aggregate case when paragraph acquisitions scale with
  polygon vertices. (#1076)

**docx — failure containment**

- Preserve surrounding document content when an optional VML text path is
  unsupported or an image/chart relationship is unavailable but its authored
  extent is valid. The canonical layout retains a diagnostic no-op node and its
  geometry; invalid numeric geometry and structural failures remain fatal.
  This is a bounded recovery policy, not catch-all exception suppression or a
  legacy renderer fallback. (#1090, #1091; progresses #1088)

**public callbacks and viewer fixes**

- Text-run callbacks now expose stable source identities:
  `DocxTextRunInfo.paragraphId` (`w14:paraId`),
  `PptxTextRunInfo.shapeId` (`cNvPr@id`), and required
  `XlsxTextRunInfo.sheetName` / `cellRef` values. (#1078, #1079)
- XLSX built-in zoom controls explicitly use `type="button"`, preventing them
  from submitting an ancestor HTML form. (#1087)

**tooling and packaging**

- Run project type-checks and declaration emission with the native TypeScript 7
  CLI while isolating TypeScript 6 for tools that still require the legacy
  JavaScript Compiler API. Remove duplicate CI package type-checks, replace
  `vite-plugin-dts` / API Extractor with `rolldown-plugin-dts`, and preserve all
  five published declaration entries and their TypeScript 5.9/6/7 compatibility.
  On the measured Apple Silicon baseline, cold `pnpm typecheck` fell from
  7.41 s to 1.51 s and cold `pnpm build` from 8.47 s to 1.79 s. (#1075)

## 0.72.2 — 2026-07-13

Patch (fixes on v0.72.1). Word-compatible docx layout geometry plus xlsx
cacheless chart and RTL-anchor correctness.

- **docx:** preserve floating-anchor host line metrics independently of the
  drawing payload (picture / chart / shape / group); resolve installed normal
  faces for eligible local metrics while keeping authored styled faces and
  prioritizing embedded regular faces (§17.8.3.3), applied consistently to text,
  floating-anchor hosts, and empty paragraph marks; include justified text and
  picture list markers in paragraph border bounds (§17.9.7, §17.3.1.24); resolve
  each emitted table page slice by the borders it actually paints without
  assuming monotonic row heights (§17.4.66, §17.4.80); honor an explicit
  `atLeast` grid minimum for a tall first line (observed Word behavior). No
  sample-specific thresholds. (#1035)
- **xlsx:** resolve cacheless legacy chart formulas through a shared DrawingML
  resolver contract (category / value / series-name / scatter / bubble sources,
  multi-level categories), with a bounded sparse worksheet-reference session
  shared across charts and sparklines; mirror anchored objects and clipping
  rectangles on right-to-left sheets without flipping content. (#1034)

## 0.72.1 — 2026-07-13

Patch (fix on v0.72.0). Preserves authored OOXML layout and grouped-drawing
geometry that had been discarded or reconstructed from incomplete state before
rendering, which shifted anchored content, lost legacy VML form shapes, and
introduced pagination/border artifacts.

- **docx:** preserve omitted-grid table borders, floating-layout constraints,
  and line metrics consistently across measurement, pagination, and painting;
  implement the legacy VML stroke/fill cascade (default colors, units, dashes,
  endcaps, arrows, geometry) so authored form shapes render (ECMA-376 §17.4.14,
  §17.4.15, §17.4.66, §17.4.85; Part 4 VML §19.1.2.21).
- **DrawingML (docx / xlsx / pptx):** centralize group transforms in the shared
  layer and propagate Annex L scale / rotation / flip semantics consistently
  across all three formats (Annex L §L.4.7.4–L.4.7.6); carry pptx table, chart,
  and media frame transforms through the parser contracts into rendering. No
  document-path or sample-specific thresholds. (#1032)

## 0.72.0 — 2026-07-13

Minor. A large docx fidelity cycle centered on **vertical writing (縦書き)** and
**line-layout correctness**: full tbRl/btLr section support with per-section text
direction, vertical text boxes and vertical ruby, UAX#50 vertical glyph forms
routed through the font's real `vert` glyphs, and Word-adjudicated docGrid /
`contextualSpacing` / table pagination. International text gains dictionary line
breaking for Thai/Lao/Khmer, grapheme-fill breaking for Myanmar/Tibetan, an
extended UAX#14 no-break predicate, and true kashida justification. A per-font
advance model closes the Canvas-vs-Word measurement gap, and the viewers add an
`enableHyperlinks` switch and a shared zoom ladder on the xlsx chrome. ~140 merged
PRs since 0.71.0.

**docx — vertical writing (縦書き, §17.6.20 / §20.1.10.83)**

- **Full vertical-section support.** Section-level `tbRl` and `btLr` now paginate
  and paint through a single vertical path with per-section text direction and
  per-page rotation, mixed vertical + final-horizontal sections, physical-margin
  horizontal headers/footers (§17.10.1), upright block tables with horizontal
  cells (§17.4.80), and physical-page anchor resolution (§20.4.3.x).
- **Vertical text boxes.** `<wps:bodyPr vert>` parses into `ShapeRun.textVert`;
  `vert` / `vert270` / `eaVert` text boxes render with the correct orientation,
  including mongolianVert (column left→right), vertical ruby on the right half,
  upright inline images, and cross-axis autofit grow/shrink for eaVert `spAutoFit`
  boxes (#988, #1000).
- **Real vertical glyphs.** The full Tu/Tr repertoire (brackets, long-vowel ー,
  wave dashes, ：；〖〗) is drawn from the font's OpenType `vert` glyphs when the
  surface can reach the feature, with measure == paint; the earlier manual
  rotate/mirror/shear approximations are dropped where real glyphs exist (#969).

**docx — line layout & pagination fidelity**

- **docGrid cell counting (§17.6.5).** Character-grid line pitch is counted from
  a deterministic design height rather than the substituted font's measured box,
  fixing environment-dependent and scale-unstable page breaks (untabled East
  Asian faces use the Word FE 1.3 × em design height); run character-grid
  participation and the table line-grid compatibility setting are honored.
- **`contextualSpacing` (§17.3.1.9).** Per-side subtraction semantics between
  same-style paragraphs, now also honored inside text boxes.
- **`fitText` (§17.3.2.14).** Parsed as a composite spec and laid out as fit
  regions, with justify/overlay propagation and RTL residual-pad placement.
- **Tables.** Over-tall `vMerge` spans split across pages with per-fragment
  vAlign (§17.4.6/§17.4.85); shared and mid-row page-cut borders resolve as
  §17.4.66 conflicts drawn at the winning cell extent; cell fragments become the
  paint authority for sliced pieces.
- Ruby carries uniform line height into continuations and honors
  `w:rubyPr/w:hpsRaise` (§17.3.3); vanished-mark paragraphs collapse to zero
  height (§17.3.2.41); a trailing empty paragraph grazing the bottom margin is
  kept (#981); numbering markers color from the level and paragraph-mark rPr.

**International text & line breaking**

- **Dictionary line breaking** for Thai/Lao/Khmer, plus breaks at no-space
  SEA ↔ Latin/CJK/digit transitions and mixed CJK+SEA runs (#959, #960).
- **Grapheme-fill breaking** for Myanmar and Tibetan (#961).
- The **UAX#14 no-break pair predicate** is extended beyond LB28 (#966).
- **Kashida justification** via U+0640 tatweel insertion with a Word-parity
  insertion-priority model (§17.18.44, #954, #724); `thaiDistribute` spreads
  across Thai grapheme clusters and Word's SEA justified line fit is matched
  (#991).
- RTL fixes: independent `bCs`/`iCs` complex-script toggles (#937), numbered
  hanging first line at the start indent, and trailing-whitespace placement.

**Fonts & list numbering**

- **Per-font advance model** — condensed-face profiles plus a Canvas-vs-Word
  bias correction, closing measurement drift for present fonts (#794, #698).
- **Cross-font metric emulation reverted** (PR #979 "Regime B"): a missing font
  now falls back to a substitute with natural reflow (page counts may differ from
  the authoring host, by design); the same-font Canvas-vs-Word bias correction
  ("Regime A") is kept.
- `decimalEnclosedCircle` and `aiueo(FullWidth)` list numbering formats.

**pptx**

- Absolute bullet size `buSzPts` (§21.1.2.4.10); bullet sub-properties cascade
  property-wise, including `buClr` on auto-numbered bullets (§21.1.2.4.4).
- UAX#50 vertical glyph orientation in the `eaVert` / stacked paths.
- Duotone applied to picture-fill blips (§20.1.8.23, #889); RTL `tabLst` mirroring
  (#831).

**xlsx**

- UAX#50 vertical glyph orientation in stacked cells; formula vertical-position
  fix (#877); shared image cache across sheets (#781).

**Viewers / interaction**

- **Disable hyperlink interactivity:** every hyperlink-supporting viewer
  (`DocxViewer`, `DocxScrollViewer`, `PptxViewer`, `PptxScrollViewer`,
  `XlsxViewer`) now accepts an `enableHyperlinks?: boolean` constructor option
  (default `true`). When `false`, hyperlink interactivity is disabled entirely:
  no overlay/cell hit-testing, no pointer cursor over links, no default
  navigation (external new-tab / internal bookmark / slide / sheet jump), and
  `onHyperlinkClick` is never called. Links still render exactly as authored but
  are inert. The switch gates the wiring at a single point per viewer rather than
  scattering flag checks through the hit-test/click paths, so the existing
  `onHyperlinkClick` behaviour is unchanged when the option is omitted or `true`.
  (#891, IX1)

- **XlsxViewer built-in +/- zoom buttons now walk the shared zoom ladder:**
  the tab-bar steppers step through the IX9 `ZoomableViewer` presets
  (25 → 33 → 50 → 67 → 75 → 90 → 100 → 110 → 125 → 150 → 175 → 200 → 250 →
  300 → 400 %), exactly like the contract's `zoomIn()`/`zoomOut()`, instead of
  the previous linear ±10 %. Behaviour change for existing users of the
  built-in chrome: from 100 %, "+" still lands on 110 %, but subsequent steps
  now follow the presets (125 %, 150 %, …) and an off-preset wheel-zoomed scale
  snaps onto the ladder on the first step. The slider and `setScale` are
  unchanged. (#842, IX9)

- find / hyperlink / selection overlays track the canvas's actual CSS size for
  responsive layouts.

**Tooling & packaging**

- The markdown package resolves its standalone WASM asset correctly (#730).

## 0.71.0 — 2026-07-06

Minor. The 2026-07 second improvement cycle lands: in-document **find**
(`findText`) and a unified **zoom contract** across all five viewers, clickable
**hyperlinks** with internal-anchor navigation, a rebuilt xlsx number-format
grammar, xlsx **outline grouping** and **furigana**, docx **vertical-text**
(縦書き) fidelity (UAX#50 vertical forms + 縦中横 tate-chu-yoko), docx page
decorations (page borders, line numbers, section vAlign), a large chart
correctness + chartEx pass, WordArt text warps in pptx, and worker-mode parity
for selection / find. New public surface: `findText` / `findNext` / `findPrev`
/ `clearFind`, `getScale` / `setScale` / `fitWidth` / `fitPage`, and
`onHyperlinkClick` on every viewer.

**Viewers / interaction**

- **In-document find:** `findText(query, opts?)` / `findNext()` / `findPrev()`
  / `clearFind()` on `DocxViewer`, `PptxViewer` and `XlsxViewer` — searches the
  whole document, highlights every hit, and tags each match with its natural
  location (docx page / pptx slide / xlsx sheet+cell). Pure string/index math
  lives in `@silurus/ooxml-core`. (#783, IX2)
- **Unified zoom contract:** all five viewers share one zoom API — `getScale()` /
  `setScale(scale)` plus `fitWidth()` / `fitPage()` (fit persists across
  reflow), backed by `core/interaction/zoomable.ts`. Previously zoom was
  inconsistent (single-canvas docx/pptx had no API, scroll viewers only had
  absolute `setScale`, xlsx only had the slider). (#837, IX9 + PD12) The wheel /
  pinch gesture now anchors on the raw pointer, so the point under the cursor
  stays put across all viewers (the padding-scaled horizontal lead-in was
  double-counted). (#882)
- **Clickable hyperlinks + internal navigation:** links in all three formats are
  now clickable via an overlay/hit-test layer (zero canvas repaint — file-specified
  colour/underline are untouched), sanitized through a shared
  `core/interaction/hyperlink.ts`, with an `onHyperlinkClick(target)` callback.
  Internal anchor targets (bookmarks / slide jumps / defined names) resolve and
  navigate. (#775, #778, IX1)
- **Worker-mode selection / find parity:** in `mode: 'worker'`, docx/pptx text
  selection (empty overlay + one-time warn) and `findText` (guard returning `[]`)
  were silently dropped because per-run `onTextRun` geometry could not cross the
  worker boundary. Run geometry now crosses the boundary, so selection and find
  behave identically in main and worker mode. (#841, IX6)

**docx**

- **Vertical text (縦書き) fidelity:** UAX#50 vertical-form substitution provenance
  corrected and dead entries dropped (#789); vertical-text glyphs centred on the
  column by font metrics rather than ink (#792); vertical 、。 kept in the cell's
  upper-right instead of ink-centred low (#793); 縦中横 (tate-chu-yoko) runs
  (e.g. "9月29日") no longer overlap (#834); find/selection overlay over a
  縦中横 run clamped to its one-em cell (§17.3.2.10, #861, #836); pinned by a
  sample-26 regression test. (§17.3.2 vertical text family)
- **International page-number formats (`w:pgNumType` + field switches):** the PAGE
  field honours `<w:pgNumType>` (§17.6.12 restart / format) and `\*` format
  switches — decimal / upperRoman / lowerRoman / upperLetter / lowerLetter native,
  plus per-section restart. (#800, WD2) Extended with hex, numberInDash and
  decimalZero, hebrew2 (appends a ת run, not the repeat-letter scheme), and
  koreanLegal native-Korean numbering; locale-independent switches route through
  core. (§17.18.59 ST_NumberFormat, #860, #795)
- **TIME / DATE fields:** `TIME` / `DATE` fields evaluated against their `\@`
  date-time picture. (§17.16.5.72 / .16, #828)
- **Section decorations:** page borders (`w:pgBorders`, §17.6.10 — standard line
  styles, offsetFrom/display/zOrder defaults; art borders documented as
  unsupported), line numbering (`w:lnNumType`), and section vertical alignment
  (`w:vAlign`) now parse and render — previously ignored. (#807, WD6)
- **Multilevel heading numbering:** a `<w:lvlOverride>` now substitutes the full
  overriding `<w:lvl>` definition (§17.9.7), and `lvl` / `pStyle` numbering-level
  backlinks resolve so a heading chain picks up the right list level rather than
  collapsing to level 0 (§17.9.23 / §17.9.8). (#870)
- **Newspaper-column band geometry:** the multi-column band uses per-section page
  geometry, so a section that changes page size / margins lays its columns out
  against its own page rather than the document's first section. (#865)
- **Hard break opening a paragraph:** a paragraph that BEGINS with a hard
  `<w:br w:type="page"/>` / `column` break used to drop the break together with
  its empty leading chunk, so it never advanced the page/column. The opening
  break is now honoured. (§17.3.1.20, #871)
- **Floating tables — cell isolation + band deferral:** table-cell content is
  isolated from the page's floating objects (a float can no longer reflow text
  inside an unrelated cell), and a page-anchored floating table whose band
  intersects another float table defers a row at a time instead of overlapping.
  (#867)
- **Adjacent cell border conflicts (§17.4.66):** shared internal gridlines were
  double-drawn and neighbour background fills could re-cover already-painted
  borders; conflicts now resolve by border weight (not paint order). (#811, WD7)
- **Table cell theme colour + auto contrast (§17.3.2.6):** runs whose colour is
  never applied route through the auto-contrast pick, and table-cell shading folds
  into the auto text-colour contrast decision. (#848)
- **Table indent (`w:tblInd`, §17.4.50):** parsed and applied. (#827, #819)
- **VML watermark / imagedata (`w:pict`):** bare `<v:imagedata r:id>` inline VML
  images (§19.1.2.11) map to inline image runs; VML text watermarks
  (`v:textpath`, §19.1.2.23) render. (#809, WD3)
- **Run character metrics (`w:w` / `w:spacing` / `w:position` / `w:kern`):**
  the four previously-unparsed run properties parse and apply; `w:w` character
  scale now applies at **paint** under active docGrid / justify so measured and
  painted widths agree (§17.3.2.43, measure == paint). (#813 WD4, #858, #816)
- **Footnotes inside tables:** footnotes referenced from inside table cells are
  drawn at the bottom of the page. (§17.11.10, #849)
- **Paragraph border spacing (§17.3.1.7):** the bottom paragraph-border extent is
  reserved in the vertical flow and mirrored in the cell content measurer. (#847)
- **Textbox default text colour:** `wps:style` `fontRef` honoured as the textbox
  default text colour. (#826, #821)
- **RTL (bidi) tab stops:** tab stops mirror in RTL paragraphs and resolve in
  reading order against the text margin. (#830, #835)
- **RTL floating tables / header images / table width:** column-relative anchors
  resolve against the text column (§20.4.3.4) and the measure pass seeds contentX
  like the paint pass (#844); page/margin-anchored floating tables taller than the
  text region split a row at a time (#845); a floating table's width caps at the
  page, not the text-column band (#852).
- **sample-12 caption pagination:** the 1-inch float line-start rule is scoped to
  CONTENT lines; empty paragraph marks keep the pilcrow threshold (regression from
  #676). (#791)
- **Emphasis mark on fields:** `w:em` wired onto `FieldRun` (PAGE / NUMPAGES). (#758)
- **Blip duotone + alphaModFix on images:** DrawingML `<a:duotone>` recolour
  (§20.1.8.23) and `<a:alphaModFix>` opacity (§20.1.8.6) apply to inline / anchored
  pictures, pinned by a node pixel-probe render test. (#888)

**pptx**

- **WordArt text warps (`a:prstTxWarp`, §20.1.9.19 ST_TextShapeType):** all 40
  presets implemented — Canvas 2D cannot deform glyph outlines, so a per-glyph
  approximation stretches each glyph onto the warp envelope (box-fit semantics).
  (#838, PP4) Single-edge "Follow Path" warps follow the path at natural width
  (followPathUScale, #856, #846); a local envelope-slope shear is applied to each
  warped glyph so the strip edges stay parallel to the envelope (#866); high-curvature
  envelopes subdivide into deviation-driven piecewise-affine strips (#872); and the
  strips draw as clipped vector `fillText` rather than resampled bitmap slices, so
  no seams appear between strips (#876).
- **SmartArt fallback (S → M):** a `<p:graphicFrame>` DrawingML diagram (§21.4)
  with no pre-baked drawing part previously drew nothing; a staged spec-compliant
  fallback (frame → text list) now renders it. (#798, PP3)
- **Table border single-draw:** shared table borders draw once (documented
  span-vs-subdivided-neighbour limitation, #824). (#817)
- **Blip duotone + alphaModFix on pictures:** `<a:duotone>` (§20.1.8.23) and
  `<a:alphaModFix>` (§20.1.8.6) recolour / fade pictures, pinned by a node
  pixel-probe test. (#888)

**xlsx**

- **Number-format grammar rebuild (§18.8.30 / §18.8.31):** the number-format engine
  moves from a regex + `toFixed` approximation to a full tokenizer → AST → renderer.
  New coverage: comma scaling (`#,##0,` = thousands, `0.0,,` = millions), fractions
  (`# ??/??`, Stern-Brocot), conditional sections (`[>…]`), and condition-selected
  negative sections format the magnitude. (#799, XL3)
- **Outline grouping:** Excel row/column outline grouping in both main and worker
  modes — gutter (group bracket + summary elbow), +/− collapse toggle, and numbered
  level buttons. (`<row>` / `<col>` outlineLevel/collapsed/hidden §18.3.1.73 / .13,
  `<sheetPr><outlinePr>`.) Non-shared gutter canvases attach only while an outline
  exists. (#843, hotfix #853, XL4)
- **Furigana (rPh / phoneticPr):** East-Asian furigana implemented — `<rPh sb/eb>`
  (§18.4.6) + `<phoneticPr>` (§18.4.3) parse, drawn only when a cell opts in via
  `ph="1"` (§18.3.1.4). Row-level `<row ph>` toggle inheritance added. (#810 XL5,
  #859, #814)
- **Sparkline implicit references:** `extract_range_values` reconstructs cells that
  omit the optional `@r` attribute (§18.3.1.4 / §18.3.1.73) by tracking row groups
  and a running column, so a sparkline over an r-less worksheet no longer renders
  blank. (#864) Implicit reference fill order documented as convention (not
  spec-mandated). (#851)
- **Shape outline `lnRef idx=0`:** a `<a:lnRef idx="0">` on an xlsx drawing shape is
  treated as "no outline" (§20.1.4.2.19), fixing a stray border around embedded
  equation shapes. (#875)
- **Blip duotone + alphaModFix on pictures:** `<a:duotone>` (§20.1.8.23) and
  `<a:alphaModFix>` (§20.1.8.6) recolour / fade pictures. (#885)

**charts**

- **chartEx renderers (CH15):** funnel / histogram → treemap → sunburst →
  boxWhisker recognised and rendered; chartEx parse centralised in
  `ooxml-common::parse_chartex_part`; boxWhisker outlines stroked in
  accent × lumMod 80%. (#750, #751, #764)
- **3D flatten + stock + auto-title (CH13):** 3D charts flattened, stock charts and
  a synthesized single-series auto-title added. (#753)
- **docx inline charts (CH12):** embedded DrawingML charts render via the shared
  core `renderChart` (new `DocxColorResolver`); line/area data labels honour
  `dLblPos` (default right, §21.2.2.48). (#746, #745)
- **Pie data-label radius:** solid-pie `ctr` data labels placed near the rim
  (§21.2.2.48); series-line `noFill` and per-point `dLbl` overrides honoured. (#808,
  #802)
- **Scatter legend key:** markers-only scatter series draw a **marker** legend key
  instead of a line swatch. (#857, #803)
- **Combo primary value axis:** the primary value axis spans both the bar and line
  series of a combo chart, so a line series no longer rescales the axis on its own. (#874)
- **Gridline z-order + area minor gridlines:** the area renderer now strokes value /
  category major gridlines UNDER the series fills (matching every other cartesian
  renderer and PowerPoint, #881) and honours `<c:minorGridlines>` (§21.2.2.129),
  which it previously ignored (#884).
- **CT_Boolean default audit:** bare chart flag elements honour `val` default=true
  across 10 sites (dml-chart.xsd). (#862, #806)
- **Axis follow-ups:** ST_FixedAngle boundary corrected (open type, closed UI range),
  plus the earlier axis/gridline/title-position work from the cycle. (#786)

**core**

- **Antiqua serif classification:** Antiqua faces classify as serif. (#825, #822)
- **Decimal half-up rounding:** a shared `roundDecimalHalfUp` display formatter
  fixes `toFixed`'s 2.675 → 2.67 problem (Office display rounding). (#863, #801)

**tests / CI**

- **VRT sample matrix:** sample-27..31 added to the VRT matrix with 71 reference
  PNGs (#854); a docx sample-28 page-count regression test pins the hard-break
  pagination fix (#869, #873).

## 0.70.2 — 2026-07-03

Patch (hotfix on v0.70.1). Fixes the packaging of the parser WASM asset
reference, which broke webpack 5 and Next.js consumers in 0.70.x.

- **build:** dist bundles shipped a double-nested
  `new URL(new URL("….wasm", import.meta.url).href, import.meta.url).href`.
  webpack 5 compiled the outer call into an empty critical-dependency
  ContextModule that threw `MODULE_NOT_FOUND` the moment any entry was
  imported, and Turbopack rewrote the outer `import.meta.url` to `file://`,
  breaking the worker's WASM fetch in Next.js (dev and build alike). The
  reference is now emitted as the bare single-level
  `new URL("…", import.meta.url).href`. Verified end-to-end: webpack 5,
  Next.js 16 (Turbopack dev/build), Vite 8 dev/build, plain
  `<script type="module">`. (#715)
- **README:** bundler note rewritten to the verified compatibility matrix —
  Vite 7 dev servers need `optimizeDeps: { exclude: ['@silurus/ooxml'] }`
  (fixed upstream in Vite 8); esbuild / Angular CLI use the `wasmUrl` load
  option. (#715)

## 0.70.1 — 2026-07-03

Patch. Documentation-only: correct the Angular / esbuild guidance for loading
the parser WASM assets (#709). No code changes.

- **README:** the bundler note wrongly claimed esbuild detects the
  `new URL('…', import.meta.url)` asset reference — it does not
  (evanw/esbuild#795), and neither does the Angular CLI's esbuild-based
  application builder (angular/angular-cli#22388). The Angular 19 example now
  documents the verified two-step setup: copy `*_parser_bg.wasm` with an
  `angular.json` assets glob **and** pass `wasmUrl` — the copy alone fixes only
  production builds, because `ng serve` resolves the reference into Vite's
  dep-optimizer cache. Retires the stale `@angular-builders/custom-webpack`
  note. (#710)

## 0.70.0 — 2026-07-03

Minor. Continuous-scroll viewers (`DocxScrollViewer` / `PptxScrollViewer`),
hidden slide / sheet display modes, a docx text-box layout unification, a large
performance pass (parse throughput, wasm chunk size, image-bitmap reuse), and the
math engine reshaped into a lazily-fetched standalone asset. New public surface:
the scroll viewers, `hiddenSlideMode` / `hiddenSheetMode`, and the `wasmUrl` load
option.

**Viewers**

- **Continuous scroll viewers:** `DocxScrollViewer` / `PptxScrollViewer` render
  the whole document as one vertically-scrolling, PDF-reader-style surface. They
  take a container `<div>` (not a canvas), virtualize the page/slide list (only the
  visible window + a small overscan is mounted), recycle canvases, and expose a
  styleable "desk" (`background` / `gap` / `padding*` / `pageShadow`), flicker-free
  Ctrl/⌘ + wheel (and trackpad-pinch) zoom, an optional per-page selectable text
  overlay, and headless-engine injection to share one parse across panes. (#650,
  #651, #652, #653, #654, #655, #656)
- **destroy():** hardened to return the caller-owned canvas to its original DOM
  position and to tolerate stale sibling refs / headless documents. (#659)

**docx**

- **Text boxes on the body engine:** text-box paragraphs now run through the same
  line-layout engine as body text, so kinsoku 行頭/行末禁則 (§17.15.1.58–60), UAX#9
  bidi (`w:bidi`, §17.3.1.6), justification (§17.18.44) and tab stops (§17.3.1.37)
  all apply inside a box. (#697)
- **Floating tables:** split a floating table across pages a row at a time, and
  clamp a page/margin-anchored floating table into its container per Word's
  observed behaviour (§17.4.57). (#691, #694)
- **Text frames:** a page-overflowing text frame keeps with its anchor
  (`w:framePr`, §17.3.1.11). (#690)
- **Float wrap:** a line beside a floating object starts at Word's measured
  one-inch rule rather than a computed offset. (#696)
- **Shrink-to-fit:** draw shrink-fitted lines with the same compression the fit
  judgment assumed, so the measured and painted widths agree. (#703)
- **Paragraph shading / borders:** a paragraph's shading box paints its bottom edge
  by replaying the painted paragraph height (#645-series follow-up, §17.3.1.7).
- **WMF metafiles:** render `META_STRETCHDIBITS` (embedded DIB) via a shared
  `core/image/dib` decoder instead of a black box (#639, cross-package with emf).
- **Layout robustness:** page/line splitting is scale-invariant — the same line
  breaks at any zoom (line rescaling is total over anchored-image segments). (#689)

**pptx**

- **Hidden slides:** `hiddenSlideMode` (`'show'` (default) / `'skip'` / `'dim'`)
  controls how hidden slides (`<p:sld show="0">`, §19.3.1.38) are presented — skip
  keeps absolute indices and honors an explicit goToSlide; dim draws a translucent
  overlay. `setHiddenSlideMode()` / `visibleSlideCount` accompany it. (#637)
- **Callout line ends:** arrow / oval line-end decorations now render on every
  callout series (`ST_LineEndType`, §20.1.8.3), not just callout-1 geometries.
  (#634)
- **Preset engine:** more shapes routed through the shared spec-correct preset
  engine (math operators, matched flowchart shapes, spec-correct micro-geometry —
  the A2 batches). (#687, #695, #704)
- **Parser modularization:** `lib.rs` split into eight focused modules (text,
  shapes, …) with no behaviour change. (#686)

**xlsx**

- **Hidden sheets:** `hiddenSheetMode` (`'show'` (default) / `'skip'` / `'dim'`)
  controls how hidden / very-hidden sheets (`<sheet state>`, §18.2.19) appear in
  the tab bar — skip removes the tab and jumps navigation over it; dim renders it at
  reduced opacity. `setHiddenSheetMode()` accompanies it. (#638)
- **Scroll:** wheel/scroll renders are coalesced onto a single rAF, and stale
  worker bitmaps are dropped with a render-generation counter. (#667)
- **Sheet switching:** parse results and sheet switches cross the worker boundary
  as transferable JSON bytes. (#662, #666, #671)

**perf**

- **pptx parse:** a 15 MB deck parses in ~11 ms (was ~30 ms) — the model is parsed
  once and the worker→main hand-off is a transferable ArrayBuffer decoded a single
  time. (#666, #669)
- **wasm chunks:** the three parser WASM modules ship as real streaming-compiled
  assets (−85–89 % vs the old base64 data-URL inline), with a new `wasmUrl` load
  option to serve them from a CDN or self-hosted path. (#681)
- **math chunk:** the MathJax + STIX Two Math engine (~3 MB) is emitted as a
  standalone sibling asset fetched lazily the first time a document has an equation;
  the opt-in `math.mjs` chunk shrinks from 4.13 MB to a ~1.1 KB loader. (#702)
- **image bitmaps:** a per-document decoded-bitmap LRU cache makes revisiting a
  docx page ~9× faster (no re-decode). (#668)
- **docx cells:** cell-paragraph line layout is stamped at measure and reused at
  paint. (#705)
- **misc:** stateful ZIP archive handles for repeat work, effect hot-path trims,
  and an opt-in `workerTimeoutMs` for wedged-worker detection. (#663, #670, #672)

**internal**

- **Layout contract (B2):** the docx renderer's measure and paint paths are unified
  behind a single line-layout contract, with `line-layout.ts` extracted and the
  compute-once reuse gate hardened. (#684, #693, #700)
- **Shared consolidation:** DrawingML color-node extraction, theme/rels parsing, and
  TypeScript shared types consolidated into `ooxml-common` / `core`. (#677, #680,
  #682)
- **CI:** a rendering smoke test on the browser worker path, and the npm publish is
  gated on tests + publint + attw + a tarball smoke test. (#658, #664)

## 0.69.0 — 2026-06-30

Minor. A CJK text-layout fidelity pass across docx / pptx, plus indent, tab-stop,
and paragraph-border accuracy. No public API changes; the Feature Support table is
unchanged (all affected features were already supported — this release sharpens
their rendering).

**docx**

- **CJK justify / 約物連続:** East Asian runs under a document grid and on
  fully-distributed justified lines are now drawn as one contiguous `fillText`
  with canvas letter-spacing, instead of one isolated draw per glyph. This honours
  the browser's JIS X 4051 約物連続 packing (a closing-class punctuation followed
  by an opening bracket — `：［`, `、［`, `）（` — packs ~half-em), so an opening
  bracket no longer overruns the following glyph (§17.6.5, §17.18.44). (#626, #630)
- **Justify line ends:** a justified (`both`) line terminated by a manual `<w:br/>`
  is left-aligned (a logical line end), while `distribute` still fills it
  (§17.18.44 + §17.3.3.1). A breakable CJK run glued after a Latin word is no longer
  measured as atomic. (#618, #623)
- **Indents:** a numbered paragraph's direct `<w:ind>` overrides the numbering
  level per attribute (§17.9.22); a numbering level's `firstLine` keeps its sign
  (§17.3.1.12); text-box paragraphs honour `<w:ind>` (§17.3.1.12). (#611, #612, #621)
- **Tab stops:** parse `<w:defaultTabStop>` and compute a spec-correct,
  margin-relative tab advance (§17.3.1.37 / §17.15.1.25). (#628)
- **Paragraph borders / shading:** honour a paragraph's *direct* `pPr`
  borders/shading (an `apply_direct_para` drift dropped them); merge borders
  per-edge across the style hierarchy; an explicit nil/none edge clears an
  inherited edge (§17.3.1.7). (#613, #617, #619)
- **Robustness / internals:** cap paragraph paint at the laid-out line count so a
  continuation slice never indexes a phantom line; unify `apply_direct_run` with
  the canonical `apply_run`. (#615, #616)

**pptx**

- **CJK justify / `@spc` / 約物連続:** `@spc` (rPr letter-spacing) justify pieces,
  `@spc` tab-stop segments, and fully-distributed justified runs are drawn
  contiguously with canvas letter-spacing, so 約物連続 opening brackets no longer
  overlap the next glyph (§21.1.2.3.x, §20.1.10.59). (#627, #629, #631)
- **Justify line ends:** a justified (`just`) line ended by a manual `<a:br>` is
  left-aligned, while `dist` / `thaiDist` still fill it (§20.1.10.59 + §21.1.2.2.1). (#624)
- **First-line indent:** unified across the wrap measurement, wrap budget, and the
  draw offset; a positive first-line indent narrows the first line's wrap width;
  list levels inherit `marL` / `marR` / `indent` from the lstStyle cascade
  (§21.1.2.2.7, §21.1.2.4.13). (#614, #622, #625)

**xlsx**

- **Shape text:** honour paragraph indent (`marL` / `marR` / `indent`) on
  shape / text-box paragraphs (§21.1.2.2.7). (#620)

## 0.68.0 — 2026-06-29

Minor. XLSX viewer interactions — drag-to-resize columns/rows, mouse-wheel /
trackpad-pinch zoom, and a customizable selection color — all **view-only**: the
loaded file is never modified. Plus a spec-fidelity and consolidation pass.

- **xlsx (resize):** drag a column or row header border to resize it (`resizable`
  option, default on; set `false` to disable). The dragged size is written into
  the in-memory worksheet model only — it never modifies the loaded file. Column
  widths convert per ECMA-376 §18.3.1.13. (#567, #607, #609)
- **xlsx (zoom):** Ctrl/⌘ + mouse-wheel and trackpad-pinch zoom, in addition to
  the Excel-style slider. The step is exponential in the wheel delta, so the
  total zoom tracks the scroll distance rather than the number of events — a
  trackpad pinch and a mouse wheel covering the same distance zoom by the same
  amount. (#607, #609)
- **xlsx (selection):** customizable cell-selection accent color (`selectionColor`
  option / `setSelectionColor()`; default Google blue, fill at 8% opacity). (#607)
- **xlsx (internals):** the column-width px conversion is now ECMA-376
  §18.3.1.13-exact; the resize hit-test is extracted to a pure, unit-tested
  helper; the local `hexToRgba` is folded into core's shared version (used by
  docx/pptx); and `XlsxViewerOptions` extends the shared core `LoadOptions`
  instead of re-declaring its fields. (#609)

## 0.67.0 — 2026-06-29

Minor. A large DOCX fidelity pass — journal templates (sample-9/12/13) now match
Word across headers/footers, line breaking, justification, and text metrics —
plus cross-format CJK/Latin line-breaking and XLSX rich-text fixes.

- **docx (headers / footers):** headers and footers are now resolved per active
  section per page, including first-page (`w:titlePg`, §17.10.6) and even/odd
  (`w:evenAndOddHeaders`, §17.10.1); a header or footer taller than its page
  margin reserves space so the body never overlaps it (§17.6.11), across every
  column of a multi-column section.
- **docx (line breaking / justification):** a justified (`both` / `distribute`)
  line spreads slack across every inter-CJK boundary including the line's last
  segment (§17.18.44); a CJK run followed by a glued non-starter ("。" / "、" in its
  own run) splits to fill the line instead of being pushed down whole; and a
  UAX#14 LB13 non-starter authored across a run boundary (a leading "," / "。")
  stays glued to its word — fixed consistently in docx, pptx, and xlsx with the
  shared predicate centralized in core.
- **docx (text metrics / layout):** Times New Roman / Arial line height comes from
  the hhea design metrics (matching Word's line box); small caps are sized per
  character so a capital initial stays full size (§17.3.2.33); anchored text-box
  content renders with correct paragraph spacing, style alignment, and z-order; an
  empty paragraph mark reserves a full line; heading numbering shares one counter
  per abstract list (§17.9); a `numId=0` de-listed paragraph drops the inherited
  list indent; a negative `w:pgMar` top/bottom places the body's `|margin|` inside
  the page edge (§17.6.11); and continuous-section spacer paragraphs collapse /
  suppress spacing-before to match Word.
- **docx (columns / figures):** continuous two-column sections balance at line
  granularity; EMF figures crop to their `rclFrame`.
- **xlsx (rich text):** empty shape / cell paragraphs reserve a line; non-wrapped
  rich text honors `\n`; wrapped super/subscript sizes correctly; and a wrapped
  rich cell re-resolves the bidi base direction per LF paragraph.
- **cross-format (core):** `<a:srcRect>` image cropping is unified in core and
  shared by docx, pptx, and xlsx (including the metafile gate via `isMetafileMime`).

## 0.66.3 — 2026-06-24

Patch. DOCX cover-page fix — resolves the 0.66.1 hotfix trade-off (sample-5 and
sample-13 are now both correct).

- **docx (cover page):** sample-5's novel body overprinted the 夢十夜 title page,
  and the 0.66.1 hotfix that fixed it pushed sample-13's intro off the title page.
  The cover is a Word "Cover Pages" building block (ECMA-376 §17.5.2
  `docPartGallery="Cover Pages"`) whose text flow is empty — the page is filled by
  page-anchored cover graphics that don't advance the text cursor — so Word places
  it on its own page via an implicit break. The parser now emits a page break after
  a Cover Pages block, and the section-break reading is restored to the upcoming
  section's start type (§17.6.22). The cover stands alone (sample-5: 7 pages) while
  a `continuous` body after a title section stays on the title page (sample-13: 5
  pages); sample-12 unchanged (3 pages). A redundant cover break adjacent to an
  explicit page/section break is dropped to avoid a spurious blank page.

## 0.66.2 — 2026-06-23

Patch. DOCX RTL / table fidelity fixes.

- **docx (table auto-fit):** an `AutoFit to Contents` table (`tblW=auto`) was sized
  from its saved `<w:tblGrid>` (often the full text column), so it spanned the page
  and overrode its own `w:tblPr/w:jc` placement — sample-7's cover page lost the
  right-aligned and left-aligned tables. `tblW=auto` now sizes from per-cell `tcW`
  /content (ECMA-376 §17.4.63 / §17.18.87); a preferred-width table (`tblW=dxa/pct`)
  still trusts the grid Word baked its layout into (sample-3 unchanged).
- **docx (RTL numbers):** a decimal inside an Arabic run drew reversed
  (`1234.56` → `56.1234`). A single common separator between two numbers now joins
  them into one left-to-right number (UAX#9 W4), so `1234.56` / `1,234.56` / `12:34`
  stay intact while hyphen-separated dates still reorder.
- **docx (RTL text boxes):** a rich (multi-format) text box drew its words in
  logical order, scrambling RTL/mixed text. It now runs the same UAX#9 visual
  reorder (rule L2) body paragraphs use, so RTL words read right-to-left while
  Latin/number islands stay left-to-right (sample-8).

## 0.66.1 — 2026-06-23

Patch. Hotfix for a 0.66.0 regression.

- **docx:** a cover page's `nextPage` section break was read as continuous, so the
  body flowed up and overprinted the title page (sample-5). Use the section
  marker's own break kind again (§17.6.x). A continuous body after a title
  section (sample-13) goes back to the next page until the cover's vAlign
  page-fill is modelled.

## 0.66.0 — 2026-06-23

Minor. A DOCX continuous-section newspaper-column overhaul plus EMF figure
support — dense, multi-page journal templates (sample-12/13) now match Word.

- **docx (multi-column sections):** continuation columns start at the section's
  region top, not the page top (§17.6.4); a section break is governed by the
  upcoming section's start type, so a `nextPage` title section before a
  `continuous` body no longer page-breaks (§17.6.22); header/footer references
  inherit across sections so a running header declared on the first section
  shows on every page (§17.10.1); non-final continuous multi-column sections are
  balanced across their columns, matching Word's page fill (§17.6.4).
- **docx (text & images):** an over-long unbreakable token (e.g. a long URL in a
  narrow column) breaks at the character level instead of bleeding past the
  column (overflow-wrap); consecutive paragraphs with identical `<w:pBdr>` borders
  merge into one box instead of ruling every line (§17.3.1.7); contextualSpacing
  collapses the gap between same-style paragraphs to zero, so a code listing's
  lines sit tight (§17.3.1.9).
- **core (images):** Enhanced Metafile (EMF) parts now rasterize via a new
  [MS-EMF] player (world transform, polylines/polygons, text, DIBs), so EMF
  charts/figures render instead of dropping to blank. (Sub-rectangle cropping of
  metafiles — `<a:srcRect>` — is deferred: the figure is shown in full for now.)

## 0.65.0 — 2026-06-23

Minor. A large PowerPoint chart-fidelity push — blank charts now render, combo
charts gain a secondary value axis, and pie/doughnut colours and labels match
PowerPoint — on top of an extensive DOCX table / floating-layout / list batch
and continued cross-format core unification.

### charts (pptx)

- **Charts that silently rendered as a blank area now render.** Two root causes:
  a relationship `Target` that is a package-root-absolute part name (leading
  `/`, e.g. `/ppt/charts/chart5.xml`; OPC / ECMA-376 Part 2 §9.3) was joined onto
  the source part's directory and the chart graphicFrame was dropped; and a
  `<c:cat>` backed by a `<c:multiLvlStrCache>` (multi-level category axis,
  §21.2.2.114) collapsed every series to a single point — an area chart of one
  point is a zero-width sliver. Both fixed, plus the same absolute-`Target` fix
  for slide parts. (PR #557)
- **Combo charts** (`<c:barChart>` + `<c:lineChart>` in one plot) now plot the
  line series against an independent **secondary value axis** drawn on the right
  — its own nice scale, ticks, labels and rotated title (`axPos="r"` /
  `<c:crosses val="max">`, §21.2.2.33). The auto value-axis major unit is now
  axis-length-aware (a longer axis gets finer gridlines, matching PowerPoint's
  roughly constant gridline spacing), the category baseline draws on top of the
  bars, and value-axis numbers clear the tick marks. (PR #560)
- **Pie / doughnut**: per-slice `<c:dPt>` fill colours render correctly (the
  point index is a child `<c:idx>` element, §21.2.2.84, not an attribute, so
  every slice previously fell back to the series colour), and `<c:showPercent>`
  slice labels (`54% / 27% / …`) now draw. (PR #559, #561)
- Bar / column / area charts gained `<c:majorTickMark>` ruler ticks (default
  `out`, ST_TickMark §21.2.3.48, shared with xlsx); area honours the category
  axis `crossBetween` (half-band inset), draws the value-axis rule, and labels
  every category. (PR #559)

### docx

- **Tables**: cell and column widths size from `<w:tblGrid>` / `<w:tcW>` spec
  geometry via a shared two-pass measure, `tblCellMar` inherits through the
  `TableNormal` style chain (§17.4.41), exact-height row clipping is Y-only so
  nested-table borders survive (§17.4.80), `tblOverlap` is scoped correctly
  (§17.4.56), and border render order / double-border rails and conditional
  (`cnf`) row/column formatting are honoured. (PR #513, #518, #527–#531, #538,
  #542–#547, #549, #553, #555)
- **Floating layout**: page- and margin-anchored floats are pre-scanned so
  earlier paragraphs wrap around them (§20.4.3.x), floating tables, frame
  `vAnchor` banding (§22.9.2.20), drop-cap `framePr`, and image-anchor
  `relativeFrom` / alignment are threaded through (§20.4.3.5 / §20.4.3.2).
  (PR #515, #520, #521, #526, #543, #550, #554)
- **Lists / text**: picture bullets and bullet/inline symbol-font resolution,
  table-of-contents styling (including the in-field hyperlink colour),
  per-section column layout, and sample-10/-11 multi-column + inline-formatting
  fixes. (PR #507, #514, #516, #519, #522, #524, #525, #534, #540)

### core

- Continued cross-format unification: a shared font generic classifier and
  WMF/metafile decoder, symbol-font (Wingdings/Webdings) handling, dash-pattern
  and auto-contrast colour helpers, complex-script code-point segmentation, and
  shared chart axis-line / `crossBetween` helpers. (PR #509–#512, #528,
  #535–#537, #541, #562)

### xlsx

- Merged-cell border z-order and the shared two-pass table-cell sizing path.
  (PR #517, #546)

## 0.64.3 — 2026-06-21

Patch. Crisp 1-px rendering of Excel grid/border lines at devicePixelRatio = 1, header/data alignment, and toned-down homepage copy.

- **xlsx**: Cell borders now render as a crisp single device row at `devicePixelRatio = 1` instead of an anti-aliased ~2-px band. A thin (1-logical-px) stroke drawn on an integer cell edge straddled two device rows at 50% coverage each; an odd-device-width stroke is now nudged half a device pixel onto a pixel midpoint, while even device widths (e.g. a thin border at dpr=2) are left untouched. Border widths stay in logical px — the orchestrator's `ctx.scale(dpr,dpr)` maps them to device px (ECMA-376 §18.18.3 ST_BorderStyle).
- **xlsx**: Row- and column-header separators now land on the exact device pixel of the matching data-grid line (they previously sat one device pixel to the left / above at dpr=1), and the freeze-pane divider lines use a device-correct half-pixel offset (they used a hardcoded logical `+0.5` that blurred at `dpr > 1`). The odd-device-width crisp offset is factored into a shared `crispOffset` helper.
- **site**: Toned down the homepage hero — replaced the "most faithful … on the web" / "pixel-faithful" superlatives with an accurate description and added a brief coverage caveat.

## 0.64.2 — 2026-06-18

Patch. PowerPoint text-rendering fidelity fixes — placeholder alignment now resolves through the slide master, letter spacing counts whole code points, and CJK justification shares the docx code path.

- **pptx**: Placeholder paragraph alignment now inherits through the slide master's `<p:txStyles>` (`titleStyle`/`bodyStyle`/`otherStyle` `@algn`) and matches the layout placeholder by `idx`, so a master-defined alignment is no longer dropped to the `"l"` default nor overridden by an unrelated typeless placeholder's centering. Adds an idx-aware `lookup_alignment` and folds the master `txStyles` alignment into the inheritance chain (ECMA-376 §19.3.1.36 idx matching; §19.3.1.52 → §19.3.1.49/.5/.35). (PR #501)
- **pptx**: Letter spacing (`spc`) is applied per Unicode code point instead of per UTF-16 code unit, so runs containing astral characters (emoji, rare CJK) are spaced correctly. (PR #500)
- **pptx / core**: pptx CJK justification (`just`/`dist`, §20.1.10.59) now runs through the same shared `@silurus/ooxml-core` helper as docx, eliminating the last divergent justify path. (PR #499)
- **site**: The showcase "Try yours" page prewarms the docx/xlsx/pptx WASM engines on an idle tick for a faster first parse, and the homepage browser-tab title now reads "Office Open XML Viewer". (PR #502)

## 0.64.1 — 2026-06-18

Patch. Fixes glyph overlap on justified CJK lines containing punctuation adjacent to Latin text in both docx and pptx.

- **docx / pptx**: `both`/`distribute` (docx §17.18.44) and `just`/`dist` (pptx §20.1.10.59) justification now anchor each sliced CJK piece to the whole-string cumulative advance, preserving the browser's contextual 約物半角 half-width collapse. Previously, summing isolated piece advances overran the segment box and painted the following run on top of the trailing CJK glyph. (PR #497)
- **core**: `justifiedPiecePositions` / `JustifiedPiece` promoted to `@silurus/ooxml-core` so both docx and pptx renderers share the same helper.

## 0.64.0 — 2026-06-17

Minor. Two headline changes plus several rendering features. **(1)** Embedded SVG
images (Microsoft `asvg:svgBlip` extension, MS-ODRAWXML) now render across Word
and Excel as well as PowerPoint, through shared blip helpers in
`@silurus/ooxml-core` and the common Rust crate; PowerPoint additionally renders
pictures that carry *only* the svgBlip (no raster fallback). **(2)** Embedded
images are extracted lazily by zip path instead of being inlined as base64 at
parse time, cutting the parsed-model payload ~90% on image-heavy files. Also:
PowerPoint text highlight, custom-geometry arrow heads, and custom-geometry arc
fixes across all three formats.

### Shared

- **Lazy, byte-on-demand image pipeline.** Parsers now emit each embedded image's
  zip path + MIME instead of a base64 `data:` URL, and the renderers fetch the
  bytes on demand through a `fetchImage(path, mime)` callback (twin of the
  existing audio/video extraction). New size-guarded
  `ooxml_common::zip::extract_zip_entry` plus an `extract_image` WASM export per
  parser. The parsed-model payload drops ~93% (pptx) / ~89% (docx) on
  image-bearing files. Decoded bitmaps and SVGs are cached **per document** —
  keyed by byte source, not by zip path, so two open documents that reuse the
  same internal path (e.g. `ppt/media/image1.png`) never swap images — and are
  released on `destroy()`.
- New `ooxml_common::blip` Rust helpers (`svg_blip_rid`, `blip_embed_rid`,
  `mime_from_ext`) and a shared path-keyed `getCachedSvgImageByPath` decoder in
  `@silurus/ooxml-core`, replacing three divergent `mime_from_ext` copies and the
  previously pptx-only svgBlip logic. `.svg` parts map to `image/svg+xml`
  everywhere; all three renderers prefer the vector original and fall back to the
  raster on decode failure.
- New `highlightBox` core helper: a single source of truth for text-highlight /
  marker extents, shared by the docx and pptx renderers.

### pptx

- Render pictures whose `<a:blip>` carries only the `asvg:svgBlip` extension with
  no raster `r:embed` (e.g. icons inserted as pure SVG) instead of dropping them.
- An SVG picture with an `a:srcRect` crop no longer routes the SVG through
  `createImageBitmap` (which cannot rasterize SVG in every browser); it decodes
  via the shared `<img>` path.
- Honor `a:highlight` (text highlight / marker) on runs (§21.1.2.3.4).
- Draw arrow heads on custom-geometry (freeform / curve) lines, not just preset
  connectors; degenerate arcs are skipped when locating the line endpoints.

### docx

- Honor the `asvg:svgBlip` extension on inline, anchored and grouped pictures:
  draw the vector original, fall back to the raster. `.svg` parts are no longer
  mislabeled `image/png`. (§20.1.8.13 + MS-ODRAWXML SVG extension)
- Custom-geometry `ArcTo` angle fields (`stAng` / `swAng`, §20.1.9.3) now
  deserialize, so arcs in freeform shapes render instead of being dropped.

### xlsx

- Honor the `asvg:svgBlip` extension on `xdr:twoCellAnchor` and grouped pictures;
  `.svg` media parts are no longer excluded from collection.
- Custom-geometry `ArcTo` angle fields (`stAng` / `swAng`) and `ShapeTextRun`
  font fields (`fontFace` / `fontSize`) now serialize as camelCase, fixing
  dropped shape arcs and unstyled shape text.

### node

- `renderSlideNode` accepts an optional `sourceBuffer`; when supplied, headless
  renders extract embedded images for real via `makeSourceBufferFetchImage` (the
  lazy `extract_image` path) instead of skipping them.

## 0.63.0 — 2026-06-17

Minor: Japanese line breaking (kinsoku) and justified-text alignment shared
across PowerPoint, Word and Excel through new `@silurus/ooxml-core` text
kernels; embedded SVG images and more faithful per-slide theme / colour-map
resolution in PowerPoint.

### pptx

- Resolve each slide's theme and master through its own layout→master→theme
  chain and honor the slide-master `<p:clrMap>`, so decks that reuse one master
  across differently-themed slides pick up the right scheme colours. (#471)
- Honor per-slide / per-layout `<p:clrMapOvr><a:overrideClrMapping>`
  (§19.3.1.7), including for colours inherited from the master. (#474)
- Render embedded SVG pictures from the `asvg:svgBlip` Microsoft-2016 blip
  extension as vectors (with the PNG rasterisation as fallback), warming the SVG
  decode cache and cropping against the PNG basis. (#472)
- Recognise embedded video declared via `p14:media` and always draw a paused
  play badge over the poster frame. (#473)
- Justify text alignment (`a:pPr algn="just"` / `"dist"`, §20.1.10.59):
  inter-word slack is distributed across the line and the decoration / text-
  selection layers span the widened gaps; tab runs are guarded. (#481)
- Apply kinsoku in the CJK wrap loop, honoring `a:pPr@eaLnBrk` (§21.1.2.2.7 —
  行頭/行末禁則) via the shared core engine. (#479)

### docx

- Place a list item's first line at `indentLeft` instead of the hanging point,
  and parse `<w:suff>` so the bullet-to-body gap is not hardcoded to a tab. (#476)
- Keep the 字下げ (first-line) indent fixed under justification and apply
  kinsoku across run boundaries when laying out a line. (#477)
- Spread CJK `both` / `distribute` justification by inter-character pitch
  (§17.18.44) rather than only at word gaps, and match the fixed segment exactly
  (not `>=`) so RTL / bidi justified lines no longer drift. (#483)

### xlsx

- Apply kinsoku (行頭/行末禁則) when wrapping CJK text inside a cell, via the
  shared core engine. (#479)

### core

- Hoist the default scheme-slot colour map (dk1/lt1/accent1–6 fallbacks) into
  `ooxml-common` so every parser resolves unmapped scheme slots identically. (#475)
- Extract the kinsoku line-breaking engine into a shared `text/kinsoku` module
  and unify the duplicated CJK break-range predicates into one
  `isCjkBreakChar`, consumed by docx / pptx / xlsx. (#479, #482)
- Extract the justified-line slack kernel (`distributeLineSlack`) into core,
  shared by the docx and pptx justifiers. (#484)

### docs / site

- Add the VS Code Marketplace badge to the README and use Title Case headings;
  surface SVG images and kinsoku in the showcase-site capability columns. (#478)

## 0.62.0 — 2026-06-16

Minor: chart axis titles, borders and value-axis scaling across Excel and
PowerPoint charts; PowerPoint symbol-font runs; plus soft-edge and
embedded-media fixes.

### charts (xlsx + pptx)

- Render chart axis titles at their XML font size / weight / colour, anchored in
  a reserved gutter so they no longer collide with the tick labels, and default
  axis and chart titles to bold (ECMA-376 ST_Style). Extract the scatter
  bottom-axis (X) title, which had been dropped for every scatter chart. (#460, #464)
- Draw a chart-space border when `<c:chartSpace><c:spPr><a:ln>` specifies one,
  and scale axis-line and border widths by `ptToPx`. (#460, #463)
- Extend the value axis to Excel's "nice" maximum (a one major-unit margin above
  the data) and share a single rounding rule for the axis bounds and gridline
  step across bar / line / area / radar / scatter (`valueAxisScale()`). (#465, #467)
- Hoist the chart title / axis-title / border extractors into `ooxml-common`,
  shared by the xlsx and pptx parsers. (#464)

### pptx

- Render `a:sym` symbol-font runs (e.g. Wingdings arrows) instead of tofu. (#462)
- Stop copying large embedded media a second time when building the playback
  handle, so navigating to a slide with a ~200 MB embedded video no longer
  duplicates it on the main thread. (#468)

### core

- Match PowerPoint's soft-edge feather (`softEdge`, §20.1.8.31): build an opaque
  edge-clamped colour layer and replace its alpha with a blurred silhouette
  (σ = rad/3), so the perimeter dissolves symmetrically instead of leaving a
  hard outer step. (#469)

### xlsx

- Scale multi-line cell text line-height and decoration offsets by the cell
  display scale, so they stay proportional at any zoom. (#466)

### site

- Reset the file input so re-selecting the same file on the “try yours” page
  re-renders the preview. (#461)

## 0.61.0 — 2026-06-15

Minor: docx pagination fidelity — floating-object displacement, line-level page
breaks with widow/orphan control, and document-grid line heights — plus shared
arc shape rendering.

### docx

- Account for floating-object vertical displacement in pagination. Body text
  pushed below a full-width anchor-float band (ECMA-376 §20.4.2.x) now consumes
  page space in the paginator exactly as it does in the renderer, so a paragraph
  no longer spills past the bottom margin and is clipped (sample-9 page 4). (#457)
- Split an overflowing paragraph at a line boundary instead of relocating it
  whole, honoring `keepLines` (§17.3.1.14) and `widowControl` (§17.3.1.44 — keep
  ≥2 lines together across a page break). Replaces a `h > ½ page` heuristic that
  suppressed ordinary line-level breaks. (#457)
- Round document-grid line heights up to whole grid cells for East Asian lines
  (§17.6.5 / §17.3.1.32): a CJK line taller than the pitch reserves two cells; a
  Latin line keeps its natural height. Fixes compressed CJK title blocks that
  cascaded every page boundary off Word. (#458)

### shapes (docx + pptx)

- Render the `arc` preset through the shared preset engine so docx and pptx draw
  it identically, and gate connector arrow heads on line geometry. (#455)

## 0.60.2 — 2026-06-15

Patch: docx connector arrow-head correctness.

### docx

- Honor shape `<a:xfrm flipH/flipV>` (§20.1.7.6). docx previously dropped flip
  entirely, so a near-horizontal `straightConnector1` kept the right line but
  swapped its start/end — drawing the arrow head on the wrong tip. Flip is now
  parsed and applied as a canvas transform composed with rotation, mirroring
  the pptx renderer (sample-9 figure 1 leader arrows now match Word).
- Gate connector arrow-head drawing on line/connector geometry instead of any
  preset. `getConnectorAnchors` resolves `path[0]` of any preset, so a filled
  shape carrying an `<a:ln>` head/tail end could otherwise get spurious arrow
  heads at its first subpath's endpoints.

## 0.60.1 — 2026-06-15

Patch: docx shape rendering now goes through the shared spec-driven preset
engine, with connector arrow heads and dash patterns.

### docx

- Route `<a:prstGeom>` shapes through core's `renderPresetShape` engine
  (presets.json, 186 ECMA-376 geometries) instead of the legacy
  `buildShapePath` switch. This recovers 37 preset families that previously
  fell back to a plain rectangle — actionButton (×18), callout2/3 and the
  border/accent callout variants, gear6/9, leftCircularArrow,
  leftRightCircularArrow, leftRightRibbon, lineInv, nonIsoscelesTrapezoid,
  swooshArrow, corner/plaque/squareTabs, flowchartOfflineStorage/Or,
  upDownArrowCallout, chartPlus/Star/X. `arc` keeps its bespoke fallback and
  `custGeom` still uses `buildCustomPath` (PR #450).
- Root cause of the rectangle fallback: core `buildShapePath` matched preset
  names case-sensitively, but docx passed the raw camelCase `presetGeometry`
  — now matched via `toLowerCase()`.
- Connector line ends: parse and render `<a:ln>` `headEnd` / `tailEnd` arrow
  decorations (§20.1.8.3) and `prstDash` dash patterns (§20.1.8.48); the
  shared `drawArrowHead` helper moved from the pptx renderer into
  `@silurus/ooxml-core`.
- Export the `LineEnd` type from the docx public entry point (reachable from
  `DocxDocument` via `ShapeRun.headEnd` / `tailEnd`).

## 0.60.0 — 2026-06-14

Minor: docx floating-object layout (overlap avoidance, below-float line flow)
plus grouped-drawing and shape-fill fixes.

### docx

- Floating-object layout: reposition overlapping floats per `allowOverlap`
  (ECMA-376 §20.4.2.3 mandates it for `allowOverlap="false"`; cross-paragraph
  avoidance for the `"true"` default is Word-parity), flow captions/lines and
  empty / anchor-only paragraph-mark rows (§17.3.1.29) below full-width floats
  (§20.4.2.16/.17), and move a paragraph-anchored float that overflows the page
  bottom to the next page — so side-by-side photos and their captions render
  like Word.
- Shape fidelity: honor an explicit `<a:noFill/>` over the shape style's
  `fillRef` (§20.1.8.44), keep degenerate line connectors (`prstGeom="line"`
  with a zero extent, §20.1.9.18), and compose nested `<wpg:grpSp>` group
  transforms recursively for both shapes and grouped images (§20.1.7.5/.6).
- Emit the paragraph-mark line for anchor-only paragraphs so consecutive
  image-only paragraphs no longer collapse onto each other (fixes shifted
  figure captions).

### xlsx

- Keep merged-cell text rendered when the cell's top-left anchor scrolls out of
  view, by wrapping it in the off-screen-anchor pre-pass.

### docs

- Drop the retired shields.io VS Code Marketplace badges from the README.

## 0.59.3 — 2026-06-14

Patch: shrink the published bundle ~36% by inlining each parser's WASM only once.

- **build**: each parser's WASM was inlined as a base64 data URL twice per format
  — once for the main thread (the canonical `?url` import, posted to the worker)
  and again inside every worker bundle, because the wasm-bindgen glue's
  `new URL('*_bg.wasm', import.meta.url)` fallback was inlined by vite-plugin-wasm
  even though workers always init from the URL the main thread passes. Setting
  `omit-default-module-path` in each parser's wasm-pack metadata regenerates the
  glue without that fallback, so the duplicate copies disappear naturally in the
  build. `dist` drops from 12.69 MB to 8.06 MB; the VS Code extension webview
  shrinks too. No API or behavior change — WASM is still delivered as an inline
  data URL (zero-config), and the glue `init()` is internal (never part of the
  public API; every caller already passes an explicit module).

## 0.59.2 — 2026-06-14

Patch: shrink the VS Code extension bundle.

- **vscode-extension**: the extension renders on the main thread only, but each
  viewer package ships a render worker as a dynamically-imported chunk that
  base64-inlines a second copy of the renderer + WASM. esbuild's `iife` webview
  output cannot code-split that import into a lazy chunk (as the npm build does),
  so it inlined ~2 MB of dead worker code. An esbuild plugin now stubs the
  `render-worker-host` import to a no-op (never reached at runtime here), cutting
  `webview.js` from 10.6 MB to 8.4 MB raw (4.1 MB → 3.2 MB gzip). No behavior
  change — worker rendering was never used in the extension.

## 0.59.1 — 2026-06-14

Patch: fix a first-paint regression introduced in 0.59.0. With `useGoogleFonts: true`,
`load()` force-loaded a fixed set of script-fallback Noto families before first
paint — including all eight CJK families (Noto Sans/Serif KR/SC/TC/JP) requested
with no `text=` subset, i.e. the full multi-MB families — regardless of the
document's content, so a pure-Latin document showed a blank viewer until every
CJK font downloaded. (The 0.59.0 preload fix unmasked this: before it, `.load()`
silently never fired.)

- **fonts**: preload only the Noto families whose script the document's text
  actually contains. A new pure `scriptPreloadNamesForText(text, cjkLang)` scans
  the parsed model's text by Unicode block (early-exit) — Latin-only documents
  preload no script fonts at all; CJK/Arabic/Thai/Hebrew/Devanagari documents
  still preload their faces. The main-thread `load()` and the render worker
  derive the set from the same parsed model, so worker/main rendering stays
  pixel-equivalent.

## 0.59.0 — 2026-06-14

The off-main-thread release: an opt-in `mode: 'worker'` that parses **and**
renders pptx / docx / xlsx entirely inside a Web Worker, returning each frame as
a transferable `ImageBitmap` so document rendering never blocks the UI thread.
Plus a broader script→Noto font-fallback chain under `useGoogleFonts`, and a fix
that makes that web-font preload actually load multi-word families.

core:

- `mode: 'worker'` plumbing shared across all three formats: worker-safe render
  guards (`isHTMLCanvas` / `defaultDpr`), a FontFaceSet-agnostic Google Fonts
  preloader that works off `self.fonts` in a worker, and per-format render
  workers that keep the parsed model worker-side (#427).
- Shared script→Noto fallback definitions — CJK (KR/SC/TC/JP, ordered by the
  document's language), Cyrillic, Thai, Devanagari and Hebrew — appended to the
  canvas font stack under `useGoogleFonts` so non-Latin runs resolve to a real
  web font instead of tofu (#426).
- Fix: `preloadGoogleFonts` loaded nothing for multi-word families. It
  re-selected which faces to load by matching `FontFace.family`, but Chrome
  serializes a multi-word family back with quotes (`"Nunito Sans"`), so the
  match found nothing and the font silently never loaded (worker text then fell
  back to a default face, changing line wrapping). It now loads the FontFace
  objects it created by reference (#436).
- Centralized each format's font-preload name collection so main-thread and
  worker modes always preload an identical set (#428).

pptx / docx / xlsx:

- Headless `mode: 'worker'` with `renderSlideToBitmap` / `renderPageToBitmap` /
  `renderViewportToBitmap` (both modes; worker mode runs off the main thread).
  docx paginates worker-side; xlsx parses sheets on demand in the worker.
  Equations require `mode: 'main'` (#427).
- `mode: 'worker'` on the interactive `PptxViewer` / `DocxViewer` / `XlsxViewer`
  too: the whole viewer — scroll, sheet tabs, frozen panes, zoom, sheet
  switching, media playback — renders off the main thread, painting worker
  bitmaps via a `bitmaprenderer` context. xlsx cell selection still works
  (geometry-based); the pptx/docx text-selection overlay is unavailable in
  worker mode (`onTextRun` can't cross the worker boundary) (#432).
- pptx: worker-mode `presentSlide` composites main-thread video over a
  worker-rendered base; concurrent picture/poster bitmap prefetch shaves the
  serial-await latency off first paint (#427).
- xlsx: anchored images decode via `createImageBitmap` and pattern-fill tiles
  build on `OffscreenCanvas`, so the render path is worker-safe (#427).

docx:

- Honor `m:oMathParaPr/m:jc` for display-math cell layout (§22.1.2.88) (#431).

pptx:

- Don't draw bullet / number markers on empty paragraphs (#425).



The sp3d release: full DrawingML 3-D shape shading — bevels, extrusion side
walls, and a calibrated light rig — layered on the scene3d camera from 0.57.0.
Plus xlsx list data-validation dropdowns, pie/doughnut charts, and an opt-in
Google Fonts toggle in the VS Code preview.

core:

- Bevel shading primitives for `sp3d` (§20.1.5.9 `bevelT` / §20.1.5.5
  `bevelB`): the lip is a distance-field raster whose azimuth is sourced from
  the gradient of a *blurred* distance field (not raw coverage), so the lit
  band is scale-invariant and crack-free at HiDPI — anti-facet passes remove
  the terminator facet on curved silhouettes (#414, #415, #416, #417).
- Bevel band geometry no longer "flat-cuts" elliptical silhouettes; the lip
  width is locked across render scales (#417).
- Light rig (§20.1.5.12 `lightRig`): `rot` reversal + three-point fill light,
  calibrated so matte/plastic materials read correctly (#421).

pptx:

- `sp3d` bevel shading now renders on shapes and pictures — `bevelT` /
  `bevelB` lips lit by `lightRig`, with `matte` / `plastic` materials
  (§20.1.5.9 / .5, #410, #414–#418).
- Extrusion side walls (§20.1.5.3 `extrusionH` / `extrusionClr`) swept under
  the `scene3d` camera, plus `p:sp` scene3d projection so shape text is drawn
  in the projected frame (#410).
- The bevel/scene3d offscreen no longer clips centre-aligned shape borders
  (#420).

xlsx:

- List-type data validation (§18.3.1.32): the selected cell shows a dropdown
  arrow whose click opens a read-only panel listing the allowed values,
  resolved from inline lists or range references
  (`XlsxWorkbook.resolveValidationList`) (#411).
- Charts: pie and doughnut series now render (#396).

vscode:

- Opt-in `ooxmlViewer.useGoogleFonts` setting surfaces the library's
  metric-compatible font substitution in the preview, widening the webview
  CSP to the Google Fonts CDN only while enabled (off by default, and
  force-disabled in untrusted workspaces) (#414).

## 0.57.0 — 2026-06-11

The effects & 3D release: DrawingML shape/picture effects, scene3d camera
perspective, and the shared preset-geometry engine now serving both pptx
and xlsx. Includes two breaking API changes ahead of v1.0 (below).

core:

- Preset-geometry engine (186 presets, ECMA-376 §20.1.9 / §19.5.31.3) moved
  from pptx into `@silurus/ooxml-core` so every format shares it (#392).
- scene3d camera projection: §20.1.5.5/.11 rotation model (left-handed,
  lat/lon/rev), calibrated perspective lens, and a scale-invariant
  supersampled homography warp — mesh error is measured in raster space so
  large / HiDPI renders stay crack-free (#393).
- Canvas effect helpers for inner shadow (§20.1.8.40), soft edge
  (§20.1.8.53) and reflection (§20.1.8.50) (#385).

pptx:

- Shape effects now render: `innerShdw`, `softEdge`, `reflection` (#385);
  the same effectLst applies to pictures (`p:pic`, §19.3.1.37) including
  `outerShdw` and glow (#389).
- Pictures clip to any preset geometry (`prstGeom` as clip silhouette,
  §20.1.9.18) — previously roundRect/custGeom only; the omitted-`avLst`
  default adj is honored (#391, #393).
- 3D camera perspective (`scene3d` camera + `rot`) on pictures, with
  effects applied after projection; `sp3d` contour edge rendered as a flat
  approximation (bevel shading planned) (#393).
- Picture borders: `a:ln` on `p:pic` stroked along the clip silhouette
  (#393).

xlsx:

- Custom table styles (`<tableStyles>`, §18.8.83) no longer synthesize
  borders the file does not define; all seven renderable
  `tableStyleElement` roles are honored (#388).
- Drawing shapes render through the shared preset engine with `avLst`
  adjust handles — parallelograms, callouts, triangles et al. no longer
  fall back to rectangles (#392).
- RTL sheets (`rightToLeft`, §18.3.1.87) keep their top-right start
  position when the host is laid out late or resized (#390).

API (breaking, pre-1.0):

- `XlsxViewerOptions.onSheetChange` is now `(index, total)` matching the
  docx/pptx viewers; read the sheet name via `sheetNames[index]` (#387).
- `PresentationHandle.dispose()` renamed to `destroy()` (#387).
- All types reachable from the public barrels are exported, guarded by a
  compile-time completeness test (#387).

infra:

- CI gates Rust: `cargo fmt --check`, `clippy -D warnings`, and the full
  test suite run on every PR (#386).

## 0.56.0 — 2026-06-11

The right-to-left release: full RTL (Arabic / Hebrew) rendering across all
three formats (issue #366), plus Japanese kinsoku line breaking and Word
table auto-layout.

core:

- From-scratch UAX#9 bidirectional engine behind a swappable `BidiEngine`
  seam — passes the full Unicode conformance suites (BidiCharacterTest
  91,707 + BidiTest 490,852 lines), with a §4.3 HL1 class-override hook
  (#367, #375).
- Google-font preload now awaits every substitute `FontFace` before first
  paint — deterministic rendering across reloads (#375).

docx:

- Right-to-left text: per-line UAX#9 reorder, `w:bidi` / run `w:rtl`
  (ambiguous punctuation → RTL, §17.3.2.30), complex-script formatting axis
  (`w:szCs` / `w:bCs` / `rFonts@cs`, §17.3.2.26/.38/.39 with direct-value
  mirroring per §17.3.2.18), Word-compatible AN digit ordering for dates,
  RTL lists / logical indents (Part 4 §14.11.2), `w:bidiVisual` column
  order (§17.4.1), `w:tcBorders` start/end edges (§17.4.66–67), and theme
  cs-font fallback via `w:themeFontLang` (§17.15.1.88) (#368, #374–#378,
  #380).
- Word table auto-layout: columns sized by preferred widths with a minimum
  content-width floor (`w:tblLayout` autofit, §17.4.52/.63) (#372, #374).
- Japanese kinsoku line breaking (`w:kinsoku` default-on, custom
  `noLineBreaksBefore/After` sets, §17.15.1.58–.60) (#383).
- Line-metric fidelity: substituted fonts use the document font's OS/2 win
  metrics (two-regime floor/shrink), `hRule="exact"` rows clip (§17.4.81),
  zero exact line spacing (#371), and a trailing `<w:br/>` keeps its empty
  line (§17.3.3.1) (#379, #382).

pptx:

- Right-to-left text: intra-line bidi, RTL bullets, `bodyPr@rtlCol` column
  order, `tblPr@rtl` tables, content-driven table row heights (`a:tr@h` as
  minimum, §21.1.3.18) (#369, #375).
- Arabic fallback web fonts are now gated to Arabic-script faces — Latin
  text in substituted fonts resolves sans-serif again (#381).

xlsx:

- Right-to-left sheets (`sheetView rightToLeft`, §18.3.1.87): mirrored
  grid, headers and frozen panes, rich-text cell bidi with `readingOrder`
  (#370, #375).
- Selection overlay and pointer hit-testing share the grid's RTL mirror —
  the selection frame now follows horizontal scrolling (#376).

## 0.55.0 — 2026-06-09

API (all formats) — **breaking, opt-in math only:**

- **The `math` engine is now injected once at construction / load, never
  per-render.** Previously docx took `math` at load while pptx and xlsx took it
  per-render (`renderSlide` / `renderViewport` options) — three subtly different
  shapes for the same dependency. Now it is uniform: pass `math` to a viewer
  constructor (`new DocxViewer/PptxViewer/XlsxViewer(target, { math })`) or to a
  headless `.load()` (`DocxDocument` / `PptxPresentation` / `XlsxWorkbook`), and
  every render reuses it. `math` is a dependency injected once, mirroring the
  viewer ↔ headless symmetry, and render methods now carry only layout options.
- **Breaking:** `RenderSlideOptions.math` (pptx) and `RenderViewportOptions.math`
  (xlsx) were removed; `math` was added to the shared `LoadOptions` and removed
  from the shared `RenderOptions`. Code that passed `math` to `renderSlide` /
  `renderViewport` must move it to the viewer constructor or `.load()`. Callers
  that already inject via the viewer/headless constructor (the documented path)
  are unaffected. docx behaviour is unchanged. The opt-in tree-shaking guarantee
  is unchanged: omit `math` and the ~3 MB engine never enters the bundle.
- The VS Code extension now also renders equations in **xlsx** (it already did
  for docx/pptx); the showcase site's "try" page injects `math` at load.

## 0.54.0 — 2026-06-09

xlsx:

- **A custom `numFmt` with `formatCode="General"` now renders the cell value
  instead of the literal text "General".** LibreOffice Calc writes a custom
  numFmt (id ≥ 164) with `formatCode="General"` for every workbook it saves; the
  formatter only reached the General path via the builtin `numFmtId=0`, so these
  cells tokenised "General" as a literal pattern and every numeric cell showed the
  word "General". ECMA-376 §18.8.30 reserves "General" as the General number
  format regardless of `numFmtId`, so a trimmed, case-insensitive
  `formatCode === "general"` is now normalised to it (#358).

charts:

- **Chart axis ticks and data labels with an explicit `formatCode="General"` now
  render the value.** `formatChartValWithCode` only fell back to the General
  formatter for null/empty codes, so the `<c:numFmt formatCode="General">` that
  LibreOffice emits on axes and data labels was tokenised as a literal pattern.
  Shared by pptx/docx/xlsx charts via `@silurus/ooxml-core` (#358).

docs:

- The README "Demo (Storybook)" link is relabelled "Live demo" — it points to the
  project site (`ooxml.silurus.dev`), not Storybook — and the headless-engine
  `Examples` reference now links to `/storybook/`, where those stories live.

## 0.53.0 — 2026-06-09

pptx:

- **Table styles now apply cell fill, text colour and bold from `tcStyle` /
  `tcTxStyle`.** The cell fill is wrapped in `<a:fill>` (ECMA-376 §20.1.4.2.27),
  which the parser skipped, so `wholeTbl` / `firstRow` / banding fills never
  resolved (transparent cells); and `<a:tcTxStyle>` (per-role default colour and
  `b="on"` bold) was ignored, so header text rendered black/non-bold. Table
  `<a:tint>` now uses the literal ECMA-376 §20.1.2.3.34 formula (a 20% band is a
  near-white wash) instead of the SmartArt linear lerp. Built-in "Medium Style 1
  - Accent 2" tables now match PowerPoint (orange header + white bold text,
  light-pink banding).
- **List bullets are inherited from the layout / master `lstStyle`.** Paragraphs
  with no explicit `<a:buChar>`/`<a:buAutoNum>` now resolve their marker from the
  per-level bullet cascade (master `bodyStyle`, layout placeholder lstStyles) by
  `lvl` (§19.7.10 / §21.1.2.4), with the matching hanging-indent metrics — body
  placeholders that previously rendered with no bullets now show the master `•`
  or a layout auto-number.
- **Line arrowheads enlarged to match PowerPoint.** The sm/med/lg width/length
  multipliers were ~half PowerPoint's size (§20.1.10.32/.33 define the steps only
  as relative); a lg/lg triangle now measures ≈8× line width, matching the PDF.
- **Stale slide renders are cancelled to stop canvas ghosting.** `renderSlide`
  is async (it awaits image/equation decode), so navigating faster than a render
  completes interleaved multiple slides onto one canvas. A per-canvas render
  token now makes superseded renders bail, so the latest slide always wins.

## 0.52.0 — 2026-06-09

pptx:

- **Placeholder body text now inherits the master `txStyles` colour even when
  bound by `idx`.** A slide body placeholder bound by `idx` (e.g.
  `<p:ph idx="35"/>`) whose matching layout shape declared size-but-not-colour
  rendered black instead of the master `bodyStyle` colour. Colour resolution was
  idx-strict and returned early on a missing layout-idx colour, never falling
  through to the master default. The idx-strict rule (ECMA-376 §19.7.16) only
  applies to the layout tier — it stops a sibling layout placeholder from leaking
  its colour. The master `txStyles` tier (titleStyle/bodyStyle/otherStyle) is a
  document-wide default keyed by placeholder *type* (§21.1.2.4 / §19.3.1) and is
  inherited regardless of `idx`. Fixes white-on-dark body text (master
  `schemeClr bg1`) rendering black on themed decks.

## 0.51.0 — 2026-06-09

xlsx:

- **OMML equations now render in shapes / text boxes (opt-in).** Excel stores
  "Insert > Equation" as OMML inside the shared DrawingML `<xdr:txBody>` grammar
  (ECMA-376 §22.1), exactly like PowerPoint. The xlsx renderer now parses those
  `m:oMath` / `m:oMathPara` runs and renders them through the same `MathRenderer`
  DI used by docx/pptx — pass `math` (from `@silurus/ooxml/math`) to `XlsxViewer`
  (or via `XlsxWorkbook.renderViewport` opts); omit it and the ~3 MB engine is
  tree-shaken away and equations are skipped. `ShapeTextRun` became a tagged union
  (`text` / `break` / `math`) mirroring the pptx text-run model.
- Equations are parsed in `oneCellAnchor` shapes and in shapes wrapped in
  `mc:AlternateContent` / `a14:m`, not just `twoCellAnchor` (parser `drawing.rs`).
- Fixed shape colour resolution for equation boxes: `<a:solidFill>` now applies
  DrawingML transforms (`lumMod` / `lumOff` / `shade` / `tint` / `satMod`) and
  resolves `<a:sysClr lastClr>`; an `<a:ln>` with a fill but no `w` defaults to
  0.75 pt; equation glyph colour reads the math run (`m:r`) and ignores the
  structural `<m:ctrlPr>` fill.
- Shape text and equations now scale with `cellScale` (the Excel zoom slider),
  matching how cell text scales — previously the shape box grew with zoom while
  its contents stayed a fixed pixel size and drifted out of alignment.

VS Code extension / docs:

- Fixed the VS Code Marketplace homepage link (now points to the project homepage)
  and ran a documentation-staleness audit (PR #352).

## 0.50.1 — 2026-06-04

VS Code extension:

- **Equations now render in the webview.** Two causes: since 0.49.0 the math
  engine is opt-in (the webview wasn't passing a `math` engine), and the
  engine's lazy `<script>` injection is blocked by the webview's nonce CSP. The
  extension now bundles the self-contained MathJax + STIX Two Math engine into
  the webview (sets `globalThis.__ooxmlStix2` on load — no script injection,
  CSP-safe) and passes a `MathRenderer` adapter to the docx/pptx viewers.

(No functional change to the npm library; version bumped to keep the single
release series.)

## 0.50.0 — 2026-06-04

docx:

- **Tables split across pages.** A table taller than the page was paginated as
  one atomic block and clipped at the bottom margin; it now splits row by row
  (ECMA-376 table pagination). Breaks land only on vMerge-safe boundaries
  (§17.4.85) and leading `w:tblHeader` rows (§17.4.78) repeat at the top of each
  continuation page.
- **Inline math reserves a full line box.** A math run's line height was taken
  from the MathJax SVG ink extents, so a short equation (a lone "−", a single
  operator) collapsed its line — and its table row — to near-zero and pinned the
  glyph to the top of the cell. The line box is now floored to the run font's
  natural ascent/descent (tall math keeps its larger ink box).

xlsx:

- Activating a sheet tab no longer scrolls the page. `XlsxViewer` kept the
  active tab visible with `scrollIntoView`, which also scrolled every ancestor
  (including the document); it now scrolls the tab strip horizontally only.

Docs site:

- "Try yours" renders pptx with interactive audio/video playback (only on-screen
  slides hold a live handle). Its slides render at display width so the media
  controls aren't shrunk. The per-format hero copy no longer references the
  cryptic "sample-1.xlsx".

## 0.49.0 — 2026-06-04

Packaging: **the math engine is now opt-in.** ⚠️ Breaking change.

- The MathJax + STIX Two Math engine (~3 MB) was statically imported by the
  docx/pptx renderers, so every consumer shipped it in their initial bundle even
  with no equations. It now lives behind a separate entry point,
  `@silurus/ooxml/math`, exposing a named `math` engine. Pass it to a viewer to
  render equations; omit it and a bundler tree-shakes the ~3 MB away entirely.

  ```ts
  import { DocxViewer } from '@silurus/ooxml/docx';
  import { math } from '@silurus/ooxml/math';
  new DocxViewer(canvas, { math });
  ```

  Works for `PptxViewer` and the headless `DocxDocument` / `PptxPresentation`
  APIs (all take `math` in their options). xlsx never references the engine.
- **Migration**: equations in docx/pptx no longer render automatically — add the
  `@silurus/ooxml/math` import and pass `math`. No change needed if you don't use
  equations (and you no longer pay the ~3 MB for them).
- Docs site: the "Try yours" page enables equation rendering; the pptx
  master-detail demo's large preview now fills its pane.
- Repo hygiene: internal design/plan/dev-note docs (`docs/superpowers/`,
  `docs/dev-notes/`) are no longer tracked.

## 0.48.1 — 2026-06-04

Docs: correct the README "Bundle size note" — the package became ESM-only in
0.47.0, so the note no longer describes a CJS output or "single module format"
tree-shaking; it now reflects the ESM-only bundle and the shared, lazily-loaded
math engine.

## 0.48.0 — 2026-06-04

Packaging: **smaller math engine.**

- The bundled STIX Two Math font now ships only the math-relevant glyph ranges;
  non-math ranges (Cyrillic, phonetics, dingbats, accented-Latin variants) are
  omitted. Math italic variables, Greek, blackboard-bold, script, fraktur,
  sans-serif/monospace variants, operators, arrows, accents and stretchy
  delimiters are all retained. Engine asset 4.2 MB → 3.0 MB; package unpacked
  11.3 MB → 9.6 MB, tarball 4.3 MB → 3.7 MB.
- Edge case: an equation using a non-Latin/non-math alphabet (e.g. a Cyrillic
  variable) would now show a missing glyph; such usage is vanishingly rare in
  OOXML math.

## 0.47.0 — 2026-06-04

Packaging: **ESM-only distribution.**

- The published `@silurus/ooxml` bundle inlines a large math-typesetting engine;
  emitting a duplicate CommonJS copy of every chunk roughly doubled the package.
  We now ship **`.mjs` only** (no `.cjs`), halving the package: unpacked
  22.3 MB → 11.3 MB, tarball 8.4 MB → 4.3 MB.
- **Breaking** for `require('@silurus/ooxml/…')` (CommonJS) consumers — use an
  `import` instead. Every modern bundler (Vite / webpack / Rollup / esbuild /
  Next) and Node ≥ 20 consume ESM, and the documented examples already use
  `import`.

## 0.46.0 — 2026-06-04

PowerPoint equation rendering, and a new shared math font.

### pptx

- **OMML equations** (ECMA-376 §22.1) now render in PptxViewer — inline and
  display math, including PowerPoint's `a14:m` / `mc:AlternateContent`
  wrappers and bare `a14:m`. Fractions, n-ary operators, radicals, matrices,
  accents/overbars, norms, sub/superscripts, scripts, blackboard-bold, etc.
- Equations inherit their run colour (e.g. a purple title) and font size.
- **Font fidelity**: pptx now loads the metric-compatible Office substitutes
  (Calibri → Carlito, Cambria → Caladea) like docx/xlsx, so text matches
  PowerPoint's advance widths and no longer overflows `wrap="none"` boxes.
- **Shapes**: a wedge callout whose tail tip is dragged inside the body now
  renders as the plain (rounded) base shape — no spurious inward notch.

### core

- **Math engine** moved to MathJax v4 with the **STIX Two Math** font baked in
  statically — a Times-like face close to PowerPoint's Cambria Math. Fully
  offline (DOM-free, zero network, zero cross-origin); works out of the box on
  `npm install`. The OMML parser is hoisted into the shared `ooxml-common`
  crate and reused by docx and pptx.

### build

- Viewer packages no longer copy sample fixtures into `dist/`, and the math
  engine is tree-shaken out of viewers that don't render equations (xlsx).

## 0.45.1 — 2026-06-04

Patch fixes for docx math (OMML) rendering.

### docx

- **Cases / piecewise**: a delimiter with an explicit empty closing char
  (`m:d` `begChr="{"`, `endChr=""`) no longer renders a spurious right
  parenthesis — an empty delimiter char now produces an invisible fence rather
  than falling back to `)`.
- **Group-char arrows**: `m:groupChr` arrows (`\to\above` / `\to\below`) render
  at a fixed accent size instead of being stretched to a narrow base (which
  looked cramped / clipped). Only brace-like group chars (overbrace / underbrace)
  stretch.

## 0.45.0 — 2026-06-04

A consistency pass over the public API ahead of a future 1.0, plus a
concurrency bug fix in the parser workers. **This release contains breaking
API changes** (renamed types and a changed `XlsxWorkbook` construction path);
see "Breaking changes" below.

### Fixed

- **Worker response correlation (pptx / docx / xlsx)**: the docx and xlsx
  parse clients matched worker responses by message `type` only, so with two
  requests in flight (a thumbnail grid, concurrent `parseSheet` calls) the
  first response of a matching type could resolve the wrong promise, and an
  unrelated message left a promise pending forever. All three packages now
  correlate by a per-request id via a shared `WorkerBridge` in
  `@silurus/ooxml-core` (covered by unit tests). pptx was already correct.
- **`DocxViewer` worker leak**: `DocxViewer` had no teardown path, leaking its
  parser worker and injected DOM wrapper. It now exposes `destroy()`.

### Breaking changes

- **`XlsxWorkbook` construction** now matches `PptxPresentation` /
  `DocxDocument`: use `await XlsxWorkbook.load(source, opts)` instead of
  `new XlsxWorkbook()` followed by `wb.load(source)`.
- **Public type renames** to remove cross-package name collisions and a DOM
  global shadow:
  - docx data model `Document` → `DocxDocumentModel`
  - docx `TextRun` → `DocxTextRun`
  - xlsx `Fill` → `CellFill`, `Font` → `CellFont`
  - the text-overlay callback payload `TextRunInfo` → `PptxTextRunInfo` /
    `XlsxTextRunInfo` (docx was already `DocxTextRunInfo`)
  - the four per-package `LoadOptions` are now a single shared type re-exported
    from each package (`maxZipEntryBytes` moved into the core `LoadOptions`).

### Added

- `DocxViewer`: `onError` / `onPageChange` callbacks and a `canvasElement`
  accessor. `XlsxViewer`: `goToSheet` / `nextSheet` / `prevSheet`,
  `sheetIndex` / `sheetCount`, and a `canvasElement` accessor — navigation and
  accessors now parallel `PptxViewer`.
- Unified `load()` error contract across all three viewers: if an `onError`
  callback is set it is invoked and `load` resolves; otherwise the error is
  rethrown (no more silent swallowing).

### Internal

- New `@silurus/ooxml-core` building blocks: `WorkerBridge` (+ `decodeDataUrl`)
  for worker request/response correlation, and `units.ts` centralizing the
  pt→px and EMU constants (`PT_TO_PX`, `EMU_PER_PT`, `EMU_PER_PX`,
  `EMU_PER_INCH`) that were previously re-spelled four different ways.
- xlsx VRT references regenerated from the current render (xlsx has no
  Excel-export ground truth, so the renderer is its own baseline).

## 0.44.0 — 2026-06-04

Adds OOXML math (OMML) rendering to the docx viewer via bundled MathJax, plus
table-style fidelity (shading / banding / borders / alignment) and table-of-
contents rendering (dot leaders, right-aligned page numbers). Verified against
the Word PDF export of the test document.

### docx

- **Math equations (OMML)**: `m:oMath` / `m:oMathPara` are extracted in the Rust
  parser into a shared AST (fractions, sub/superscripts, n-ary with correct
  limit placement, radicals, delimiters, matrices / `eqArr`, `limLow` / `limUpp`,
  group chars, bars, accents) and rendered via **MathJax** — converted to MathML,
  then SVG, then rasterized onto the canvas. MathJax (Apache-2.0) is **bundled and
  loaded same-origin** (no cross-origin request); `setMathJaxUrl` overrides the
  source. Equations size to the surrounding text and center as display math.
- **Table styles**: resolve `w:style type="table"` cell shading, borders, and
  conditional formatting (`tblStylePr` firstRow / band1Horz / band2Horz),
  honoring each row's `w:cnfStyle` bitmask (§17.4.7) — so banded tables paint
  correctly. Table `w:jc` centering and display-equation cell centering added.
- **Table of contents**: tab **leaders** (dot / hyphen / underscore, §17.3.1.37)
  and **right / center / decimal tab stops** (measured from the text margin), so
  TOC entries render `heading … page` on one line with flush right-aligned page
  numbers. Complex multi-paragraph fields (TOC) now render their result content
  (only PAGE / NUMPAGES are recomputed), fixing a dropped first entry.
- Internal-document links (TOC entries, cross-references) render as plain body
  text instead of the Hyperlink style's blue / underline, matching Word; external
  URL links stay blue + underlined.

## 0.43.0 — 2026-06-03

Adds Excel-style sheet-tab colors and a zoom slider to the xlsx viewer, plus a
PowerPoint chart-axis fidelity fix verified against the PDF exports of the test
decks.

### xlsx

- **Sheet tab colors**: `<sheetPr><tabColor>` (ECMA-376 §18.3.1.93; theme +
  tint / indexed / rgb) now renders as a color bar along each sheet tab's bottom
  edge. The color is surfaced on the workbook sheet list via a bounded
  worksheet-head read (`<sheetPr>` is the first child, so `<sheetData>` is never
  inflated), so every tab paints up front without eagerly parsing each sheet
  (#315).
- **Zoom slider**: an Excel-style zoom control pinned to the right end of the
  sheet-tab bar (10%–400%, with 100% at the slider's center via a
  piecewise-linear position→scale map). Gated by the new `showZoomSlider`
  viewer option (default on); `zoomMin` / `zoomMax` are configurable (#315).

### pptx

- **Chart axes**: horizontal bar charts now draw the category-axis line
  PowerPoint renders. The left rule was previously misattributed to the value
  axis — in a bar chart the category axis is the vertical/left one — so a chart
  whose value axis is `<c:delete val="1">` (sample-2 slide-16) drew no axis line
  at all. Axis tick labels now honor the file's `<c:txPr>` text color and
  `<c:spPr><a:ln>` line color (ECMA-376 §21.2.2.*) instead of a hardcoded gray,
  resolved through the lumMod/lumOff path (tx1 15%/85% → ~#D9D9D9, bg1 75% →
  #BFBFBF); this also fixes slide-7's column charts (#314).

## 0.42.0 — 2026-06-03

Rendering-fidelity release, verified against the Word / PowerPoint / Excel PDF
exports of the test decks. Continues replacing sample-fit heuristics with
spec-faithful behavior and expands the canvas-free unit-test coverage.

### Rendering fidelity

- **pptx**: SmartArt shapes now honor the explicit text frame PowerPoint stores
  per shape in the fallback drawing (`<dsp:txXfrm>`), so labels land where
  PowerPoint puts them — e.g. a process arrow's bullet list starts past the
  overlapping circle node instead of behind it, and a roundRect label clears the
  badge rect overlapping its bottom (#312). Preset-geometry text also lays out
  inside the geometry's text rectangle before insets apply (ECMA-376 §20.1.9.21),
  fixing centered arrow text drifting into the arrowhead and roundRect corner
  insets.
- **pptx**: per-list-level default font sizes (`lvl1pPr`..`lvl9pPr` `defRPr sz`,
  §21.1.2.4) are inherited from the slide master / layout, so a 2nd-level bullet
  shrinks (28→20pt) as in PowerPoint instead of staying at the level-1 size (#311).
- **pptx**: media playback time renders with tabular figures (each digit in a
  fixed slot), removing the layout jitter as the clock ticks (#310).
- **docx**: `auto` / single line spacing is sized from the intended font's
  Windows metrics (§17.3.1.33), fixing vertical drift where overlapping title
  lines accumulated down the page (#306).
- **docx**: an explicit `<w:color w:val="auto"/>` is honored as the automatic
  (black) color and overrides an inherited style color (§17.3.2.6), so
  `auto`-colored runs under a gray character style render black, not gray (#307).
- **docx**: per-cell `<w:tcMar>` margins override the table-level `<w:tblCellMar>`
  default (§17.4.42 over §17.4.41), and empty paragraphs contribute their font
  metrics — restoring missing space before cell content (#309).
- **xlsx**: per-series chart data-label colors (`<c:ser><c:dLbls><c:txPr>`,
  §21.2.2.216) are resolved individually instead of collapsing every series to
  the first series' color — parity with the pptx chart fix (#308).
- **xlsx**: a number format with a dedicated negative section formats the value's
  magnitude in it (§18.8.30), so `0;(0)` renders -5 as `(5)`, not `(-5)` (#301).

### Charts

- The automatic value-axis maximum restores Excel's documented headroom — the
  first major unit above `Ymax + (Ymax - Ymin)/20` — so the tallest series sits
  below the top gridline rather than flush, fixing auto-scaled bars that had
  become ~10% too tall (#304).
- Stacked bar-chart data labels render with the correct per-series color, comma
  grouping, and visibility vs PowerPoint (§21.2.2.216) (#305).

### Internals

- Axis-scaling math (`niceStep` / `niceAxisMax` / `niceAxisMin`) is extracted to
  a canvas-free `axis-scale.ts` module with unit tests locking the Excel-faithful
  auto-max behavior (#302).
- Added vitest unit tests for the xlsx formula engine and number formats
  (ABS/INT/MOD/CEILING/FLOOR, ISBLANK/EXACT, COUNTIF, DATE, date format codes)
  (#301, #302).

## 0.41.0 — 2026-06-01

Rendering-fidelity and performance release. Replaces several sample-fit
heuristics with spec-faithful behavior (verified against the Word / PowerPoint
PDF exports of the test decks), and removes per-frame work from the hot render
paths. Also lands the engineering foundation for a stable release: VRT
regression detection, a modular xlsx parser/renderer, and PR-gating CI.

### Rendering fidelity

- **pptx**: titles and captions now honor `cap="all"` / `cap="small"` inherited
  from the slide master/layout placeholder style (ECMA-376 §21.1.2.3.13), not
  just from the run — e.g. a template `titleStyle` with `cap="all"` upper-cases
  mixed-case text exactly as PowerPoint does.
- **pptx**: normAutofit now applies PowerPoint's stored `fontScale` /
  `lnSpcReduction` (§21.1.2.1.3) instead of re-deriving the shrink with a search,
  reproducing PowerPoint's exact text layout.
- **xlsx**: cell indentation is `indent × 3 × MDW` of the normal-style font
  (§18.8.1), replacing an ungrounded font-size factor.
- **charts**: the automatic value-axis maximum is the smallest major-unit
  multiple ≥ the data max — no extra "headroom" step — matching Excel; overlap
  is handled by honoring `<c:plotArea><c:manualLayout>` (§21.2.2.32), which the
  waterfall renderer now respects too.
- **docx**: `<w:lastRenderedPageBreak/>` (Word's layout cache, §17.3.1.20) is
  ignored uniformly — pagination is computed from our own layout — removing a
  ruby-only special case (verified byte-identical on the ruby sample).

### Performance

- **xlsx**: per-sheet render lookups (the full-sheet cell map, conditional-format
  compile, table/sparkline maps) are memoized per worksheet instead of rebuilt
  every scroll frame.
- **xlsx**: scroll/click cell lookup is O(log n) via cumulative-offset axes with
  binary search, replacing a linear scan that could walk ~1M rows.
- **pptx**: decoded image bitmaps are cached instead of re-decoding the inlined
  base64 on every render.

### Internals

- VRT regression mode + snapshot capture for local pixel-diff gating.
- The xlsx Rust parser (`lib.rs`, −75%) and TypeScript renderer (−31%) are split
  into focused modules.
- CI now validates every PR (WASM build, typecheck across all packages,
  build, ast-grep lint) and gates the npm publish on a clean typecheck.

## 0.40.0 — 2026-05-30

xlsx viewer UX release. Adds Excel-style sheet-tab navigation so workbooks with
more sheet tabs than fit the container width stay reachable. No parser or
rendering changes; docx / pptx are untouched.

### Features

- **xlsx (viewer)**: the sheet-tab bar hides its scrollbar, which left plain-mouse users (no trackpad / Shift+wheel) unable to reach tabs that overflow the container width. Added fixed Excel-style prev / next triangle buttons at the left of the tab bar. They scroll the tab strip one clipped tab per click — they do **not** change the active sheet — and grey out (disabled) at each end and when there is no overflow. Tabs now keep their natural width (`flex:none`) so the strip genuinely overflows, while `overflow-x:auto` is retained so trackpad / Shift+wheel scrolling still works. The two buttons span the row-header width so the tab strip starts in line with the data columns (offset by one inter-tab gap). Covered by a new Playwright interaction test (`packages/xlsx/tests/visual/tab-nav.spec.ts`).

### Notes

VS Code extension is bumped to 0.40.0 to stay in lockstep with the npm packages (no extension-specific changes this release).

## 0.39.0 — 2026-05-24

xlsx fidelity release. Six correctness fixes for chart and row-height
rendering surfaced by `demo/sample-1`'s Forest Inventory / Carbon & Growth
/ Biodiversity Index sheets. The chart-related fixes also extend the
shared `@silurus/ooxml-core` chart renderer, so pptx / docx consumers see
the same improvements (pptx demo VRT 9/9 still passes byte-identically).

### Fixes

- **xlsx**: the `customHeight="1"` gate on `<row ht>` added in 0.37.0 was too strict — `demo/sample-1` sheets 2-5 store `ht="36.95"` on row 2 without `customHeight`, and Excel renders that row at ~49 px (36.95 pt × 4/3), not the workbook default. Always honor `ht` and `defaultRowHeight` when present; `customHeight` is metadata about *how* the height was set, not a gate on whether to honor it. The same gate is removed from `sheetFormatPr@defaultRowHeight`.
- **chart**: `<c:catAx|valAx><c:spPr><a:ln><a:noFill>` (line-only hide, distinct from `<c:delete val="1"/>` which removes the entire axis) is now honored across line / bar / area / scatter / waterfall renderers. New `catAxisLineHidden` / `valAxisLineHidden` fields on `ChartData` / `ChartModel`; parser detects `<a:noFill>` under `<a:ln>` and each renderer gates its axis rule on it. `demo/sample-1` "Carbon & Growth" uses this to suppress the value-axis vertical line on the embedded line chart.
- **chart (line)**: series color resolution now falls back to `<a:ln><a:solidFill>` when `<c:ser><c:spPr>` has no direct `<a:solidFill>`. Line / scatter / radar series carry their color on the stroke fill; the previous lookup only checked the area fill, so explicit `<a:srgbClr val="2D6A4F">` overrides fell through to the theme accent rotation (the "Carbon & Growth" "Year" series rendered as accent1 #156082 blue).
- **chart**: legend swatch is now a horizontal line segment (instead of a filled rectangle) for line / stackedLine / stackedLinePct / radar / scatter chart types. Bar / area / pie keep the rectangle. Matches Excel's per-chart-type legend convention.
- **chart (radar)**: `<c:radarChart><c:radarStyle val>` (ECMA-376 §21.2.3.10) is parsed and threaded onto `ChartModel.radarStyle`. The renderer now only fills the polygon when `radarStyle === "filled"`; "standard" / "marker" / absent draw the line only. Without this every radar series was washed with a 25 % alpha fill.
- **chart (radar)**: per-series `<c:marker><c:symbol val="none"/>` is honored even when the chart-type style is `radarStyle="marker"` (series-level override wins, ECMA-376 §21.2.2.33). Missing data points (sparse `<c:val>` cache where ptCount > the supplied pts) are now skipped instead of synthesized as 0 — the polyline breaks on holes and only closes when every point is present. `demo/sample-1` "Biodiversity Index" omits idx 0 on every series; Excel draws an open polyline from Northridge to Hollowvale without bridging back through the top spoke.

### Notes

VRT references for `demo/sample-1` sheets 1 and 2 are now visibly stale relative to the corrected layouts (cumulative effect of the 0.37.0 pt→px row-height change and this release's `<row ht>` gate fix). The new renders match Excel; references need `UPDATE_REFS=1` once confirmed.

## 0.38.0 — 2026-05-23

Single pptx text-layout fix that affects any text body whose layout inherits a
non-zero `<a:spcBef>`. Demo VRT is unchanged (9/9 demo/sample-1 slides still
match), but the slide-5 chart in that deck now lines up the way PowerPoint
exports it instead of stacking the caption on top of the chart title.

### Fixes

- **pptx**: `<a:spcBef>` is the gap *between* paragraphs (ECMA-376 §21.1.2.2.6); PowerPoint suppresses it on the first paragraph of any text body because there is no preceding paragraph to space against. The renderer was applying it unconditionally for the first line of every paragraph including paragraph 0, pushing the text down by the inherited spcBef. `public/demo/sample-1.pptx` slide-5 surfaces this: the "Figure 1. Canopy cover index …" caption placeholder inherits a layout-level `spcBef=1000` (10 pt) from the slide layout's body lstStyle. With the unconditional spcBef the caption rendered ~10 px below the placeholder top and collided with the chart title "Canopy Cover Index" sitting just below (chart `graphicFrame@y` is only 93026 EMU = ~10 px below the placeholder). The fix gates `topGap` on `paraIdx > 0` in `buildLayout`.

### Author metadata

- npm package `author` switched from the GitHub `noreply` placeholder introduced in 0.37.0 to a dedicated `silurus.dev@gmail.com` address so consumers have a real reply path that isn't the maintainer's personal inbox. Applied across all 7 publishable packages (`root`, `core`, `pptx`, `xlsx`, `docx`, `node`, `markdown`); `vscode-extension` keeps using `publisher: silurus` with no author email.

## 0.37.0 — 2026-05-23

xlsx-focused fidelity release. Four small spec-correctness fixes that together
align the picture-bearing samples (calendars, dashboards with embedded clip art)
with what Excel actually renders — positions, proportions, italic flags — using
ECMA-376 §20.5.2.33 / §18.3.1.73 / §22.9.2 as the source of truth instead of
the empirical observations the renderer was previously holding onto. Demo VRT
references for `private/sample-10` (H7 sun emoji, G13 parasol+waves group) and
the right-side L–N emoji column now line up with their Excel exports; the
`demo/sample-1` Evergreen dashboard re-screenshot in the README reflects the
new row-height interpretation (rows render ~33% taller, so fewer fit per
1200×720 viewport — same as Excel). pptx / docx renderers untouched, but their
README images were re-captured at the same time so all three views come from
the same release.

### Fixes

- **xlsx**: `<xdr:twoCellAnchor editAs="oneCell">` pictures and shape groups now honour the saved EMU extent (`<xdr:spPr><a:xfrm><a:ext>` or `<xdr:grpSpPr><a:xfrm><a:ext>`) from the parser instead of deriving the rect from the from/to cell anchors. ECMA-376 §20.5.2.33 + Excel's "Move but don't size with cells" behaviour: with `editAs="oneCell"` Excel preserves the picture's saved size regardless of cell resizing, and the to anchor is only updated to track that fixed size. The previous from/to-derived rect would amplify any column-width / row-height discrepancy into a visible aspect-ratio drift (`private/sample-10` H7 sun rendered at 0.877 aspect vs Excel's 0.888; G13 parasol+waves at 2.06 vs 1.75). Parser now captures `editAs` + `nativeExtCx/Cy` on `ImageAnchor` / `ShapeAnchor`; renderer falls back to from/to for the default `twoCell` mode and any anchor missing the native ext.
- **xlsx**: column-width MDW lookup table for Meiryo UI 10/11 pt. When the host doesn't ship Meiryo UI (macOS), Canvas2D's `measureText` falls back to a narrower sans-serif and returns MDW≈7 for Japanese faces where Excel uses 8. The miscalculation cascaded: every column rendered ~12 % narrow, shifting every drawing anchor inside the sheet roughly one cell to the right of where Excel renders it. The new `MDW_TABLE` in `renderer.ts` overrides the measurement for the families where the fallback systematically under-measures (Meiryo UI / Meiryo) and lets other fonts (Yu Gothic 12 pt, where Canvas happens to land on 8) keep using the measurement.
- **xlsx**: row height `pt → px` conversion. ECMA-376 §18.3.1.73 specifies `<row ht>` in points; the renderer had been treating the value as already-resolved pixels because of an empirical observation on `private/sample-27`. That observation conflated two things — sample-27's rows lacked `customHeight="1"`, so Excel was ignoring the `ht` attribute entirely and falling back to the workbook default, not that `ht` was stored in pixels. The parser now honours `customHeight` (only records `ht` when the attribute is set), and the renderer applies the spec-correct `×4/3` conversion. The intrinsic default in the parser moves from 20 px to 15 pt (same display result) so both code paths share the same units.
- **xlsx**: ST_OnOff `val` attribute respected when parsing `<b>` / `<i>` / `<strike>` in font definitions. ECMA-376 §22.9.2: `<i val="0"/>` is an explicit *clear* override, not implicit "on". The parser previously treated the mere presence of the element as turn-on, which inverted the meaning of a differential format like `<dxf><font><b/><i val="0"/></font></dxf>` (which Excel reads as set-bold + clear-italic). `private/sample-10`'s calendar month-transition CF carries exactly that shape, and `C5` "6月7日" / `F11` "7月1日" were rendering as bold italic when Excel shows them as plain bold. The fix is a new `parse_st_on_off` helper applied across the three font-parsing sites in `styles.xml` (`<fonts>`, `<dxfs>`, rich-string `<rPr>`). DrawingML's attribute-form `<a:rPr b="1" i="1">` path is unaffected.

## 0.36.0 — 2026-05-19

Single security-surface change: the 512 MiB per-entry ZIP decompression cap that
backed the "zip-bomb safe" guarantee is now caller-configurable. Default is
unchanged, so renderers / VRT references / `demo/sample-*` outputs are
byte-identical to 0.35.0 — README screenshots were not refreshed for that
reason.

### Features

- **core / pptx / docx / xlsx**: viewer constructor option `maxZipEntryBytes` now overrides the per-entry ZIP decompression cap that backs the "zip-bomb safe" guarantee. Default stays at 512 MiB; raise it for legitimate decks with large embedded media, or lower it to tighten the budget for untrusted input. Plumbed through `PptxViewer` / `DocxViewer` / `XlsxViewer` constructors → worker `parse_*` messages → new `Option<u64>` argument on the `#[wasm_bindgen]` entry points (`parse_pptx`, `parse_docx`, `parse_xlsx`, `parse_sheet`, `extract_media`, plus the `*_to_markdown` variants). The Rust parsers consult the cap via a new shared `ooxml-common::zip` thread-local, replacing the three near-identical `const MAX_ZIP_ENTRY_BYTES: u64 = 512 * 1024 * 1024;` duplicates. Zero / negative values fall back to the default. Existing callers that omit the option see no behavior change.

## 0.35.0 — 2026-05-17

Two new companion packages (`@silurus/ooxml-node`, `@silurus/ooxml-markdown`) and an inline render of DOCX track-changes. Browser viewer renderers are unchanged for pptx / xlsx; for docx, the markup overlay only fires on runs that sit inside `<w:ins>` / `<w:del>` blocks, so `demo/sample-1.docx` (no tracked changes) renders byte-identically. README screenshots not updated for this reason. Zero runtime dependencies preserved.

### Features

- **node**: new `@silurus/ooxml-node` workspace package exposing Node-side parsers for all three formats (`parsePptx`, `parseDocx`, `parseXlsx`, `parseXlsxAllSheets`). Loads the workspace WASM artifacts via `fs.readFileSync` + `WebAssembly.Module` so there's no Web Worker or DOM dependency — works in CI, serverless, and CLI contexts. Includes a `renderSlideNode` helper that targets any user-supplied `Canvas` implementation (recommended: `skia-canvas`) and an `installImageBitmapShim` polyfill for the renderer's raster picture path. A first-pass `ooxml-thumbnail` CLI ships under `packages/node/bin/` (pptx-only; docx/xlsx pending font and image polyfill work).
- **pptx / docx / xlsx**: each format package now declares `./wasm` and `./wasm-binary` subpath exports so Node-side consumers (and bundlers) can locate the wasm-pack JS shim and the underlying `.wasm` file without reaching into `src/wasm/` paths.
- **docx**: ECMA-376 §17.13.5 track-changes (`<w:ins>` / `<w:del>`) now render inline. The Rust parser tags each `TextRun` produced inside a tracked block with `{ kind: 'insertion' | 'deletion', author, date }`; the renderer overlays the author-derived colour (8-hue stable palette hashed by author name), an underline for insertions, and a strikethrough for deletions. The new `RenderPageOptions.showTrackChanges` flag (default `true`) toggles the markup off for a "Final / No Markup" view. Deletion text inside `<w:delText>` is now surfaced (was silently dropped) alongside insertion text inside `<w:t>`. Public TS surface now also exposes the previously parser-only `Document.revisions` / `comments` / `footnotes` / `endnotes` fields (CHANGELOG 0.32.0) with proper TS types (`DocRevision`, `DocComment`, `DocNote`, `RunRevision`).
- **markdown**: new `@silurus/ooxml-markdown` workspace package exposing `pptxToMarkdown` / `docxToMarkdown` / `xlsxToMarkdown` as pure WASM calls (no JSON round-trip). The Rust `*_to_markdown` functions used to be native-only (mcp-server); they're now also `#[wasm_bindgen]`-exported so the same projection runs in browser, Node, or via the new `ooxml-md` CLI. Markdown output matches the v0.33.0 MCP projections byte-for-byte (titles via `placeholderType`, bullets via `<w:outlineLvl>`, pipe-table XLSX per sheet). Includes a node20-based GitHub Action under `packages/markdown/action/` for bulk-converting an entire repository's OOXML files in CI.
- **storybook**: new `PptxViewer/Markdown`, `DocxViewer/Markdown`, `XlsxViewer/Markdown` stories that load demo/sample-1 (or a user file), invoke the WASM-backed conversion, and display the markdown output alongside size / compression ratio / latency.

## 0.34.0 — 2026-05-16

Single xlsx-focused change. pptx / docx renderers are byte-identical to 0.33.2, and the xlsx grid render with no active selection is unchanged — header colors only diverge once the viewer holds a selection. No README screenshot updates because of this.

### Features

- **xlsx**: `XlsxViewer` now highlights the row / column headers that the current selection belongs to, matching Excel's two-tier indicator. Cell selection (`cells` mode) paints the row & column headers of the selected range in a slightly darker grey (`#e8eaed`). Whole-row / whole-column selection (`rows` / `cols` mode) paints the selection-axis headers in light blue (`#caddf6`) with a blue border (`#5b9bd5`) and the other axis in the grey accent. Select-all paints every header in light blue. Two new `RenderViewportOptions` fields (`selectedRowRange` / `selectedColRange`, each carrying `start` / `end` / `strong`) flow from the viewer through `XlsxWorkbook.renderViewport` into `renderer.renderHeaders`, where `drawColHeader` / `drawRowHeader` pick fill and stroke colors per index. Canvas re-render is now triggered alongside every `updateSelectionOverlay()` call so headers repaint in lock-step with the overlay border.

## 0.33.2 — 2026-05-12

### Features

- **mcp-server**: every tool now declares MCP `ToolAnnotations` hints (`readOnlyHint=true`, `idempotentHint=true`, `openWorldHint=false`). Clients that honour these hints (VS Code Copilot, Claude Desktop, …) can auto-approve calls without showing the per-call "このセッションで許可する / スキップ" confirmation that has been firing on every read since 0.32.0. Every tool we expose is a pure read of a local OOXML file — no filesystem mutation, no network, no external state — so the hints are unconditional.

## 0.33.1 — 2026-05-12

### Fixes

- **vscode-extension**: unlink the cached `ooxml-mcp-server` binary before overwriting it. The version-pin fix introduced in 0.32.1 first triggered an in-place rewrite on the 0.32.x → 0.33.0 upgrade, and macOS `amfid` caches the kernel's code-signing decision for an executable by path. Writing a new ad-hoc-signed binary (different content hash) over the existing file leaves the stale decision attached; the next exec is silently refused, dyld blocks before reaching `main()`, and VS Code logs `Waiting for server to respond to "initialize" request...` forever. `fs.promises.rm(dest, { force: true })` immediately before the writeFile severs the cache association so amfid evaluates the new file from scratch.

**Affected users on 0.33.0**: until you upgrade to 0.33.1, manually delete the cached binary and reload VS Code so the extension can prompt to (re)install:

```
rm ~/Library/Application\ Support/Code/User/globalStorage/silurus.office-open-xml-viewer/bin/ooxml-mcp-server*
```

## 0.33.0 — 2026-05-11

MCP-server-focused release introducing a **text-focused projection** tier alongside the rich structured tools. Renderer is unchanged from 0.32.1 (demo VRT 20/20 at 100% match).

### Features

- **mcp-server**: `pptx_to_markdown`, `docx_to_markdown`, `xlsx_to_markdown` — convert OOXML files to GitHub-flavoured markdown, designed for AI agents that need to *read* content efficiently rather than reason about layout. Each parser exposes a native `to_markdown_native(data)` projection that walks the already-parsed model and emits semantic markdown:
  - **PPTX**: slide titles via `placeholderType ∈ {title, ctrTitle}`, body shapes as nested bullets at the paragraph `lvl` depth, tables, chart summaries, speaker notes, comments. Auto-generated metadata placeholders (`sldNum`/`dt`/`ftr`/`hdr`) and decorative pictures/connectors are dropped.
  - **DOCX**: headings via `<w:outlineLvl>` → `#`-`######`, numbered/bullet paragraphs honour the resolved abstractNum format, tables with vMerge continuation, footnotes/endnotes collated as `[^N]: …`, comments as `> **author**: text`. Per-run `**bold**` / `*italic*` / `~~strikethrough~~` / `[link](url)` preserved with whitespace pulled outside the wrappers.
  - **XLSX**: `## SheetName` per sheet followed by a pipe table of the populated bbox, merged-cell continuation cells rendered empty, fully-empty middle rows trimmed, IEEE-754 ULP noise masked so `702.6` no longer renders as `702.5999999999999`.
- For demo/sample-1.pptx the markdown is ~21× smaller than the raw XML, ~3× smaller than the structured tools, and only ~8% bigger than the flat-text extractor.

### Fixes

- **mcp-server/xlsx**: `xlsx_get_cell_range`, `xlsx_get_formulas`, and `xlsx_search_cells` were matching the wrong CellValue serde tag (PascalCase `"Text"`/`"Number"` vs. the camelCase `"text"`/`"number"` actually emitted), silently returning empty strings for every cell value since these tools were introduced in 0.32.0. Same class of bug as the v0.32.0 `pptx_extract_text` fix. Caught while testing the new `xlsx_to_markdown` against demo/sample-1.xlsx; the regression window is exactly one release.

## 0.32.1 — 2026-05-11

Two bug fixes that surfaced shortly after 0.32.0. Renderer is unchanged from 0.32.0.

### Fixes

- **vscode-extension**: redownload the bundled `ooxml-mcp-server` binary when the extension upgrades. `resolveBinaryPath` previously kept any cached binary as long as it existed on disk, so after `silurus.office-open-xml-viewer` updated 0.31.0 → 0.32.0 the workspace kept running the old 0.31.0 binary indefinitely — silently missing the v0.32.0 fixes (the `pptx_extract_text` empty-text bug being the most visible). Now writes a sibling `<binary>.version` pin file at download time and forces a redownload on extension-version mismatch. Explicit user override paths are still honored unchanged. PATH fallback only applies on a clean install (no cache yet) so a stale globally-installed `cargo install`'d binary can no longer mask the version-pin check.
- **stories (xlsx / docx / pptx)**: the "Debug – raw parse JSON" story silently returned from the file-change handler when wasm `init()` had not yet resolved, leaving the placeholder visible forever. The change handler now `await`s the same init promise (idempotent + cached, so subsequent awaits resolve instantly) and shows a "Parsing `<filename>`…" status as soon as a file is picked, regardless of init latency.

## 0.32.0 — 2026-05-11

This release is **MCP-server-focused**. The renderer (xlsx / docx / pptx viewers) is byte-identical to 0.31.0 — confirmed by running the full demo VRT against the rebuilt WASM (xlsx 5/5, docx 6/6, pptx 9/9 at 100% match). No README screenshot updates because of this.

### Features

- **mcp-server**: 30 new tools, bringing the total from 14 to 41. The previous toolkit was limited to text extraction and cell reads; the new tools expose chart data, named ranges, conditional formats, sheet layout, table / picture / shape drill-down, paragraph run formatting, document outlines, comments, footnotes, track changes, data validations, slide notes, and inferred shape relations on PPTX slides.
- **mcp-server**: `pptx_get_shape_relations` infers connector hookups (with arrow direction when `headEnd` / `tailEnd` are arrows), bbox containment, overlap (with IoU), and axis-aligned alignment groups on a slide. Detection is purely spatial — `confidence: "inferred"` flags this — and uses the new shape `id` / `name` parser fields for stable referencing.
- **pptx-parser**: `<p:cNvPr @id @name>` and `<p:nvPr><p:ph @type @idx>` are now serialized on every `ShapeElement` as `id` / `name` / `placeholderType` / `placeholderIdx`. The placeholder fields fix `slide_title` resolution: agents can now filter for `placeholderType ∈ {"title", "ctrTitle"}` instead of returning whichever shape happens to be first in z-order.
- **pptx-parser**: `ppt/notesSlides/notesSlideN.xml` and legacy `ppt/comments/commentN.xml` (with author resolution from `ppt/commentAuthors.xml`) are now parsed and surfaced on each `Slide` as `notes` and `comments`. Modern Office365 threaded comments still TODO.
- **docx-parser**: `<w:outlineLvl>` is surfaced on `DocParagraph` (was internal in `ParaFmt` only). `<w:ins>` / `<w:del>` track-changes events are collected with author / date / text in `Document.revisions` (a non-disturbing second pass — body parse logic and rendering unchanged). `word/comments.xml`, `word/footnotes.xml`, and `word/endnotes.xml` are now parsed into `Document.comments` / `footnotes` / `endnotes`.
- **xlsx-parser**: `<dataValidation>` rules (ECMA-376 §18.3.1.32) are now parsed into `Worksheet.dataValidations`. Comment full text and resolved author are now in `Worksheet.comments`; `comment_refs` (used by the renderer for the red-triangle indicator) is derived from this list and remains stable.

### Fixes

- **mcp-server/pptx**: `pptx_extract_text`, `pptx_search_text`, and `pptx_get_slide_structure` had three latent bugs that all silently produced empty output: the helper matched run-tag strings (`"textRun" / "run"`) that don't exist (pptx-parser emits `"text"` / `"break"` after `rename_all = "camelCase"`); `slide_title` checked a `placeholderType` field that the parser hadn't been serializing yet; and table-cell extraction looked at `cell.paragraphs` instead of `cell.textBody.paragraphs`. All three are fixed with regression tests against sample-1.pptx.

## 0.31.0 — 2026-05-10

### Fixes

- **pptx**: `<a:bodyPr><a:spAutoFit/>` + `wrap="square"` interaction now matches ECMA-376 §20.1.10.5 / §20.1.10.7 — text wraps when its natural single-line width exceeds the bbox, otherwise the shape stays auto-fit at one line. PR #242 had unconditionally suppressed wrap on spAutoFit to keep sample-2 slide-13's "20代" textbox on one line, but that broke the much more common case (slide-16's right-column callouts: "認知拡大によりインバウンド案件が増加し…") where the same XML pattern is supposed to wrap. New `naturalWidthExceedsBbox()` pre-pass picks the right behaviour per shape.
- **pptx**: chartEx waterfall callout connectors now align with the candle's lower-left corner. PR #248 had over-shrunk top padding (8% → 6%), moving bar bottoms below the absolute-slide-coord callout tips. Solving the EMU geometry from sample-2 slide-8's "CS部門増員による人件費を計上" callout (`<a:prstGeom prst="callout1" adj3="-67286" adj4="95815">`) shows PowerPoint uses padT ≈ 12% with padB ≈ 14% — set those values.
- **pptx**: chartEx waterfall now honors `<cx:catScaling gapWidth>` (sample-2 slide-8: `gapWidth="0.8"`). Previously the bar/catGap ratio was hardcoded 0.55; now the parser reads the chartEx fraction, converts to legacy percentage form, and the shared renderer applies `barW = catGap / (1 + gapWidth/100)` (ECMA-376 §17.18.34) — same formula `<c:barChart>` already uses. Default 150% per spec when omitted.
- **core/chart**: `niceAxisMax` adds a step of breathing room when `dataMax / niceMax ≥ 0.9`, matching PowerPoint's stacked-bar auto-axis behaviour. sample-2 slide-16 (data 9715, niceMax 10000 = 97% fill) now snaps to 12000 so the bars don't spill past the right text column.
- **pptx**: chartEx waterfall data labels honor per-data-point colour overrides (`<cx:dataLabel idx><cx:txPr><a:solidFill>`) and position negative-value labels *below* the bar instead of above. sample-2 slide-8 negative steps `△ 52`, `△ 40`, `△ 108` now render in red (accent1) below their respective candles, matching PowerPoint.
- **pptx**: callout1 family `<a:ln><a:tailEnd type="oval">` is now drawn at the line's tip. Previously arrow ends were only rendered for `CONNECTOR_GEOMS`; sample-2 slide-8 callouts have a filled oval terminator that pointed nowhere. Per ECMA-376 callout1 gd, the line endpoints are `(x1, y1) = (w·adj2/100000, h·adj1/100000)` (attach) and `(x2, y2) = (w·adj4/100000, h·adj3/100000)` (tip).
- **pptx**: `<p:pic>` honors `<a:custGeom>` clipping. sample-2 slide-12's website inset image is now trimmed to the laptop silhouette (was previously overlaying the bezel).
- **pptx**: `<a:bodyPr numCol>` collapses to a single column when total content fits in one column, matching PowerPoint. sample-2 slide-13's 従来 box has `numCol="2"` but only 4 paragraphs — PowerPoint stacks them vertically, not 2+2. Pre-existing balanced split (#243) still kicks in for the 新機能 box (9 paragraphs that overflow).
- **pptx**: chart bar data labels honor `<c:dLbls><c:txPr><a:defRPr sz>` (was clamped at 11px by a `barW * 0.6` heuristic). Vertical (column) bar labels now centered horizontally on the bar — `drawBarDataLabel` was being called with `barW`/`barH` swapped, putting labels far to the right of each bar.
- **pptx**: every `LayoutPlaceholders` lookup with a by_idx map (`fill`, `font_size`, `blip_fill`, `stroke`, `color`, `line_spacing`) is now idx-strict per ECMA-376 §19.7.16 — a slide shape with `<p:ph idx="N">` may only inherit from the layout placeholder with the matching idx, never a sibling-by-type. Closes the gap that was already known about from the PR #241 `lookup_fill` fix.
- **pptx**: bullet color follows ECMA-376 §21.1.2.4 chain `<a:buClr>` → first run's color → paraDefault → bodyDefault. Previously skipped the first-run step, so bullets without explicit `buClr` rendered in the (often white) shape default and disappeared on darker styles.

### Features

- **pptx**: theme `<a:objectDefaults>` (txDef / spDef) inheritance for `<a:bodyPr>` defaults. `parse_shape` reads `<p:cNvSpPr txBox="1"/>` to pick the right defaults slot — txDef for true text boxes, spDef for placeholders / shape-with-text. Without this, the theme's declared `<a:spAutoFit/>` (etc.) didn't propagate to slide-level shapes that left attributes blank.
- **pptx**: `<a:bodyPr numCol>` / `<a:bodyPr spcCol>` (§20.1.10.34) — multi-column text body with balanced paragraph distribution.
- **pptx**: `<c:plotArea><c:layout><c:manualLayout>` (§21.2.2.32) — explicit plot-area placement now flows through `parse_legacy_chart` and the shared renderer; the bar chart honours it the way the scatter chart already did.

## 0.30.0 — 2026-05-10

### Features

- **pptx**: `<a:bodyPr numCol>` and `<a:bodyPr spcCol>` (§20.1.10.34) — multi-column text body. Paragraphs flow across N columns with a balanced `⌈N/numCol⌉` split so the right column lines up with the left at row level. Sample-2 slide-13's "ターゲット条件" / "新機能" textbox uses this to align 利用履歴 / 価値観 / 属性詳細 with 年齢 / 性別 / 居住地.
- **pptx**: `<a:bodyPr><a:spAutoFit/>` (§20.1.10.5) now disables horizontal wrap. spAutoFit means the *shape* grows to fit the text, so the bbox stops being a wrap boundary — mixed-size sequences like "20代" or "YoY+11.9%" stay on one line.
- **pptx**: theme `<a:objectDefaults>` inheritance (§20.1.6.7). `<a:txDef>` / `<a:spDef>` bodyPr defaults now flow through to slide-level shapes that leave attributes blank. `parse_shape` reads `<p:cNvSpPr txBox="1"/>` to choose the right defaults slot — txDef for true text boxes, spDef for placeholders / shape-with-text — so theme `<a:spAutoFit/>` doesn't bleed onto wrapping shapes.
- **pptx**: `<p:pic>` `<a:custGeom>` clipping (§20.1.9.8). Pictures with a custom-geometry path are clipped to that silhouette before drawing — fixes sample-2 slide-12 where the website inset image was overlaying the laptop bezel instead of being trimmed inside it.
- **pptx**: `<c:legend><c:legendPos>`, `<c:barChart><c:gapWidth>`, `<c:overlap>`, `<c:dLblPos>`, `<c:dLbls><c:txPr>` font color, `<c:dLbls><c:numFmt>`, `<c:valAx><c:numFmt>`, and ChartEx `<cx:axis hidden>` now flow from the parser through to the renderer (sample-2 slides 7 + 8). xlsx already exposed these; pptx had been quietly dropping them.
- **pptx**: callout1 family preset (`callout1` / `bordercallout1` / `accentcallout1` / `accentbordercallout1`) attach/tip pairing fixed to the ECMA-376 (Y, X) convention from `presets.json`. Sample-2 slide-8's connector callouts now point at the correct bars.
- **core/chart**: waterfall chart honors `valAxisHidden` (skips gridlines, tick labels, and the L-frame's left segment). `dataMin = 0` when all bars are non-negative so bar bases sit on the x-axis instead of floating above it (sample-2 slide-8).
- **core/chart**: bar data labels now read `chart.dataLabelFontSizeHpt` to match the file's specified size; the previous `barW * 0.6` heuristic stayed clamped at 11px regardless of XML.
- **rust**: `ooxml_common::chart` module extracted with shared chart-XML probes (`extract_legend`, `extract_axis_min_max`, `axis_is_deleted`, `extract_chartex_axis_hidden`, `extract_data_label_font_color`, …). pptx and xlsx parsers now call into one place so future field additions stay in lockstep. New `ColorResolver` trait keeps theme-aware colour resolution per-crate while sharing the DOM walks.

### Fixes

- **pptx**: data label horizontal centering on column-style bar charts. `drawBarDataLabel` was being called with `(barW, barH)` swapped where it expected `(barL=length, barW=thickness)`, so `cx = bx + barW/2` ended up using the bar's HEIGHT and pushed every label far to the right. Vertical (column) charts now centre labels on the bar centre as PowerPoint does.
- **pptx**: `LayoutPlaceholders` lookups (`fill`, `font_size`, `blip_fill`, `stroke`, `color`, `line_spacing`) are now idx-strict per ECMA-376 §19.7.16. When the slide-level shape carries `<p:ph idx="N">` the only valid layout source is the placeholder with the matching idx; the previous fall-through-to-`by_type` was leaking sibling-placeholder values onto unrelated slots (sample-2 slide-4's header strip rendered grey because layout10 idx=13's gray fill was bleeding onto idx=12).
- **pptx**: bullet color inheritance follows ECMA-376 §21.1.2.4.4 / §21.1.2.4.10: `<a:buClr>` → first run's color → paraDefault → bodyDefault. Previously the chain skipped the first-run step and bullets without explicit `buClr` rendered in whatever `<p:style><a:fontRef>` resolved to (white on slide-13's bullet boxes — invisible).
- **pptx**: `parse_text_body` no longer breaks a non-whitespace token away from preceding non-whitespace content on the same line, matching PowerPoint's `ST_TextWrappingType` "square" behavior for sequences like "YoY+11.9%".

### Refactors

- **rust**: `pptx-parser` and `xlsx-parser` chart-XML extractors deduplicated via the new `ooxml_common::chart` module. ~150 lines of overlapping bodies removed; the same parser now feeds both pptx and xlsx ChartElement variants.

## 0.29.1 — 2026-05-10

### Fixes

- **docs**: Storybook Introduction page (GitHub Pages) showed broken screenshots because `docs/images/` are bundled via `import` in `.storybook/Introduction.mdx`. Retook all three screenshots using `canvas.toDataURL()` (canvas-only, no Storybook UI chrome) and triggered a fresh deploy.

## 0.29.0 — 2026-05-10

### Features

- **docx**: section breaks — `continuous` (no page break), `oddPage`, `evenPage` (§17.18.79 `ST_SectionMark`). Parity breaks pad a blank page when the new section would otherwise start on the wrong side.
- **docx**: `<w:ruby>` furigana annotations (§17.3.4) rendered above base text at `<w:hps>` font size. Line height expands uniformly across lines that carry ruby so annotations never overlap.
- **docx**: text inside drawing shapes (`<wps:txbx>`) with `lIns/tIns/rIns/bIns` insets (§21.1.2.1.1) and `wps:bodyPr @anchor` vertical alignment (`t / ctr / b`).
- **docx**: anchor shape positioning — `<wp:positionH/V relativeFrom="…">` full container set (page / margin / leftMargin / rightMargin / topMargin / bottomMargin / paragraph / line — §20.4.3.4); `<wp:align>` horizontal/vertical alignment (§20.4.3.1); `<wp14:pctPosH/VOffset>` percentage positioning (§20.4.2.7); `<wp14:sizeRelH/V>` percentage size overriding `<wp:extent>` (§20.4.2.18–19). `<mc:AlternateContent>/<mc:Choice Requires="wp14">` wrappers are resolved correctly.
- **docx**: wgp group shape transform — `grpSpPr/xfrm chOff/chExt → off/ext` scale applied to child shapes (§20.1.7). Each child's offset and size are converted from child-coord-space to page EMU using the group's sx/sy factors.
- **docx**: font family classification from `word/fontTable.xml` (§17.8.3.10). Parser reads `<w:font><w:family w:val="roman|swiss|modern|…"/>` and exposes it as `fontFamilyClasses` in the document JSON. The renderer uses this as the authoritative serif/sans-serif/monospace source, falling back to name-pattern matching only for fonts absent from the table.
- **docx**: preset geometry shapes via `core.buildShapePath` — shares the full catalog (ellipse / roundRect / triangles / arrows / callouts / ribbons / flowchart / …) with the pptx renderer.
- **docx**: `<a:theme>` fill references (`fillRef idx`) resolve through the theme's `fillStyleLst` / `bgFillStyleLst` with scheme-colour recolouring.

### Fixes

- **docx**: `bodyPr` inset defaults corrected to ECMA-376 §21.1.2.1.1 values — lIns=rIns=91440 EMU (7.2pt), tIns=bIns=45720 EMU (3.6pt). Previous `unwrap_or(0.0)` caused text frames to overlap their shape borders when attributes were absent.
- **docx**: ruby paragraph line height snaps to integer docGrid pitches (§17.6.5) to prevent cumulative drift across multi-line ruby paragraphs.
- **docx**: ruby paragraphs split across pages correctly when the paragraph doesn't fit on a single page.
- **docx**: `<w:lastRenderedPageBreak/>` honored in ruby paragraphs where self-paginator line-height drift is largest.
- **docx**: requested font family is preserved verbatim as the first CSS font-stack entry; fallback chain adds appropriate Japanese serif/sans-serif fonts.
- **docx**: `DocRun::Break.break_type` serialised as camelCase (`lineBreak` / `pageBreak` / `columnBreak`) to match TS consumer expectations.

### Refactors

- **core**: `buildShapePath` and shape-path helpers moved from `@silurus/ooxml-pptx` into `@silurus/ooxml-core` so DOCX and PPTX share one preset-geometry catalog.
- **rust**: `ooxml-common` crate extracted for colour transform helpers shared between the pptx and docx parsers.

## 0.28.0 — 2026-05-10

### Features

- **pptx**: parse `a:hlinkClick` on text runs and resolve to URLs via slide _rels (ECMA-376 §21.1.2.3.5). Hyperlink runs now also pick up the theme `hlink` colour and render with a forced underline. `Presentation.hlinkColor` / `RenderContext.themeHlinkColor` are exposed to consumers.
- **pptx**: full `rPr` underline-style enum (`single` / `double` / `singleAccounting` / `doubleAccounting` / `dotted` / `dotDash` / `dotDotDash` / `dash` / `dashLong` / `wavy` / `wavyDbl` and the `*Heavy` variants — ECMA-376 §21.1.2.3.16). New `drawTextDecoLine` helper handles dashed / dotted / wavy / double / heavy in both rich-text and plain-text paths.
- **pptx**: per-run `uFill` underline colour (§21.1.2.3.20) — the line can render in a separate colour from the glyph.
- **pptx**: `vertAlign` superscript / subscript (§21.1.2.3.13), caps `all` / `small` (§21.1.2.3.13), letter spacing `spc` (§21.1.2.3.5), and `strike="dblStrike"` two-line strikethrough (§21.1.2.3.10) all flow through to the renderer.
- **pptx**: `rPr > a:ea` East Asian typeface (§21.1.2.3.7) — CJK glyphs in mixed-script runs now use the theme / explicit ea font instead of falling back to the latin one.
- **pptx**: pattern fill (`a:pattFill`, §20.1.8.40 / §20.1.10.59) ships 30 preset bitmaps (`pct5`–`pct90` / `horz` / `vert` / `cross` / `diag*` / `grid` / `brick` / `check` / `trellis` etc.) tiled via `CanvasPattern`.
- **pptx**: `a:glow` (§20.1.8.17) renders as a coloured Canvas shadow with zero offset; `a:innerShdw` (§20.1.8.21), `a:softEdge` (§20.1.8.31), and `a:reflection` (§20.1.8.27) are now parsed (rendering for those three lands in a follow-up release). `apply_color_transforms` also gained `alphaModFix` / `alphaMod` / `alphaOff` so existing `outerShdw` colours emitted with `alphaModFix` honour their alpha.
- **pptx**: compound stroke styles `<a:ln cmpd="dbl|thinThick|thickThin|tri">` (§20.1.8.42) are parsed for every shape and rendered for straight lines / connectors. The base centre stroke is erased with `destination-out` before the sub-lines paint so dash / arrow heads / glow remain consistent.
- **pptx**: `formatAutoNum` (§20.1.10.61) covers all symmetric Plain / Period / ParenR / ParenBoth variants for arabic, alphaLc / Uc, romanLc / Uc, plus arabicDb (full-width digits). 8 → 22 schemes.
- **xlsx**: rPr / cell font `<vertAlign>` (§18.4.6) and `<u>` enum (§18.4.13) — superscript / subscript apply `~65%` size + baseline shift; `double` / `doubleAccounting` underline draws as two parallel lines via the new `drawTextDecoLine`.

### Fixes

- **xlsx renderer**: default `fg` / `bg` colours for `<patternFill>` when the colour children are absent (ECMA-376 §18.8.20). Excel emits `darkHorizontal` / `darkVertical` etc. without explicit colours; the prior truthy-check on `fgColor` rendered them as nothing. Same fix flows through merged-anchor cells via the new `paintCellPatternFill` helper.
- **xlsx renderer**: every preset `patternType` now ships an explicit 8×8 (or 12×12 for matched-line-count `*Horizontal` / `*Vertical` / `*Grid`) bitmap tuned by sight against Excel's actual output. Gray family (`gray0625` → `darkGray`) reads as a continuous dot-density gradient (~6% → ~75%) starting from the `gray0625` motif and adding dots at each tier; directional hatches (`*Horizontal` / `*Vertical`) match dark / light line counts at the same pitch with `dark*` using a 2-px-thick bar and `light*` a 1-px line; `darkGrid` renders as a 2×2 checker, `lightGrid` as a wider 4-px-pitch grid.
- **xlsx renderer**: `hatchPattern` now pre-bakes `pat.setTransform(scale(1/sx, 1/sy))` so each source bit lands on exactly one destination *device* pixel at integer scales (`scale=2`, retina), and tiles at native resolution at non-integer scales (`scale=1.5`) instead of triggering Canvas's bilinear pattern resampler — which previously smeared `lightHorizontal` into a uniform half-tone. Cache key includes the rounded scale so two render passes at different scales don't collide.
- **xlsx renderer**: render `<diagonal style="double"/>` borders as two parallel lines along the diagonal's perpendicular (ECMA-376 §18.18.3). Diagonals previously fell back to a single line.
- **xlsx viewer**: selection overlay aligns with cell borders at any `cellScale`. The renderer rounds each cell's scaled width per-cell; `viewer.getCellRect` / `updateSelectionOverlay` / `updateSpacerSize` now mirror the same per-cell rounding, eliminating the up-to-several-pixel drift visible at non-integer scales.

### Docs

- README feature support table reflects all of the above. PowerPoint animations / transitions are explicitly marked "Not planned" per scope.

## 0.27.0 — 2026-05-08

### Features

- **vscode-extension**: auto-install integration for `ooxml-mcp-server`. When the workspace contains an `.xlsx`, `.docx`, or `.pptx` file, the extension offers to enable the MCP server so AI agents (Copilot Agent, Claude, etc.) can read those files via dedicated tools instead of unzipping XML by hand. The binary is downloaded on demand from GitHub Releases (~5 MB, SHA256-verified) into the extension's globalStorage; existing `cargo install` / Homebrew users on `PATH` are reused as-is. New settings: `ooxmlViewer.mcpServer.enabled` (`auto` / `always` / `never`) and `ooxmlViewer.mcpServer.binaryPath`. Requires VS Code 1.101+ for the finalised MCP API.
- **mcp-server**: each release now ships prebuilt binaries for macOS (arm64 / x64), Linux (x64 / arm64), and Windows (x64) on the GitHub Release page, removing the Rust toolchain requirement for end users.

## 0.26.0 — 2026-05-08

### Fixes

- **xlsx renderer**: pick the higher-precedence ST_BorderStyle (none<hair<dotted<…<medium<thick<double) when two adjacent cells specify a border on the shared edge. Previously the lower / right cell's thinner top / left would partially erase the upper / left cell's medium or thick edge — most visible on sample-27 row 9 where the medium top was shown as roughly the same weight as the thick bottom even though the styles differed. ECMA-376 §18.3.1.4 doesn't spell out conflict resolution, but Excel renders the stronger style at a conflict.
- **xlsx renderer**: render `hair` as a 1-px on / 1-px off dashed pattern instead of a faint solid line. ECMA-376 §18.18.3 ST_BorderStyle "hair" is the finest dashing in Excel's border picker; the previous solid render was indistinguishable from `thin` (e.g. sample-27 G13 outer frame).
- **xlsx renderer**: close `double`-border corners when the lower / right cell has to redraw the over-painted segment. The inherited top / left now uses the upper / left cell's outer-vs-inner extension so the line restored after the lower cell's fill is the *outer* (extended) one — eliminating 1-px gaps at the four outer corners (visible on isolated double-bordered cells such as sample-27 E5).
- **xlsx renderer**: widen `medium` (1.5 → 2) and `thick` (2 → 3) `lineWidth` to match Excel's pt convention (thin=1pt / medium=2pt / thick=3pt). Previously medium and thick both rasterised to roughly a 2-pixel band on canvas, so a row with medium-on-top and thick-on-bottom (e.g. sample-27 row 9 driven by `thickBot` / `thickTop` row attributes) read as the same weight.

### Tooling

- **vrt**: dropped the GitHub Actions VRT workflow and rebuilt the flow as local-only. The `private/sample-*` fixtures are not redistributable so cannot be committed; running VRT on CI provided no signal beyond `demo/sample-*`. `pnpm vrt` now expects the developer to have private fixtures locally and gates reference updates behind `UPDATE_REFS=1`.

## 0.25.0 — 2026-04-30

### Breaking Changes

- **pptx**: `PptxViewer` constructor signature changed from
  `new PptxViewer(container: HTMLElement)` to `new PptxViewer(canvas: HTMLCanvasElement)`.
  Callers must now create and place the `<canvas>` element themselves; the viewer
  wraps it internally (same reparent pattern as `DocxViewer`). `XlsxViewer` keeps
  its `container: HTMLElement` argument because it owns a sheet-tab DOM bar in
  addition to the grid canvas.

  **Migration**:
  ```diff
  -const pptx = new PptxViewer(document.getElementById('pptx-container')!);
  +const canvas = document.getElementById('pptx-canvas') as HTMLCanvasElement;
  +const pptx = new PptxViewer(canvas);
  ```

### Features

- **docx**: `DocxViewer` / `DocxDocument.load` now accept a
  `useGoogleFonts?: boolean` option matching the existing PPTX / XLSX shape.
  When `true`, the viewer preloads Google Fonts substitutes for theme-declared
  typefaces (`<a:fontScheme><a:majorFont|minorFont><a:latin>`) so canvas text
  measurement matches Word's. Default `false` to avoid third-party requests.
- **core**: `LoadOptions` and `preloadGoogleFonts` are now exported from
  `@silurus/ooxml-core`. The three format packages (`docx`, `pptx`, `xlsx`)
  re-export their own `LoadOptions` from this shared shape so an application
  can pass the same options object to any viewer.
- **docx parser**: theme `majorFont` / `minorFont` (Latin axis) are now
  serialized into the parsed `Document` JSON. Existing renderer behavior is
  unchanged; the new fields are read by `DocxDocument.load(...,
  { useGoogleFonts: true })` to drive font preload.

### Bug Fixes

- **docx**: wrapper now sets `vertical-align:top` and forces
  `canvas { display: block }` to eliminate the ~6 px baseline-descender gap
  that previously let the host container's background color show through
  below the rendered page.
- **pptx**: same baseline-gap fix applied to `PptxViewer`'s wrapper and to
  the low-level `renderSlide(canvas, ...)` path so direct callers do not
  inherit the gap either.

## 0.24.3 — 2026-04-30

Documentation-only patch standardizing the `Viewer.load` examples across
all three formats.

- **Documentation** (`README.md`, Storybook):
  - **README Quick Start + framework examples now show `viewer.load(url)`
    for PPTX**: previously the README and the six framework code samples
    (React / Vue / Angular / Svelte / SolidJS / Qwik) demonstrated the
    PPTX viewer being driven through a manual
    `fetch(url).then(r => r.arrayBuffer()).then(viewer.load)` dance,
    while the DOCX and XLSX examples passed the URL string directly to
    `viewer.load(url)`. The asymmetry was misleading: `PptxViewer.load`
    has always accepted `string | ArrayBuffer` and internally calls
    `fetch + arrayBuffer` for the string form, identical to
    `DocxViewer` and `XlsxViewer`. Updated all examples to the
    URL-string form so the three viewers read the same way (PR #176).
  - **Storybook `buildViewerUI` helpers (PPTX, XLSX) pass the URL
    directly**: matched the existing DOCX `buildViewerUI` style by
    dropping the manual fetch + ArrayBuffer step and letting the viewer
    do the network call itself (PR #176).

## 0.24.2 — 2026-04-30

Patch release fixing a centerContinuous regression introduced by the
0.23.0 row-z-order fix.

- **XLSX engine** (`packages/xlsx`):
  - **centerContinuous internal verticals stay hidden** (ECMA-376
    §18.18.40 ST_HorizontalAlignment): the row-z-order repair added in
    0.23.0 — which inherits the cell-to-the-left's `xf.right` as the
    current cell's `left` when the latter is unset — also fired for the
    *empty* left edges that the centerContinuous block deliberately
    nulled, re-introducing the internal verticals between e.g. A3-D3 and
    E3-H3 in sample-27. Gate the inherit pass on
    `!suppressLeftGridCol.has(ci)` so cells participating in a
    centerContinuous run keep the run reading as one visual span. The
    top-inherit doesn't need the guard because centerContinuous only
    suppresses verticals (PR #174).

## 0.24.1 — 2026-04-30

Patch release fixing column-width pixel conversion for files whose Normal
style is not Calibri 11 pt.

- **XLSX engine** (`packages/xlsx`):
  - **Max Digit Width derived from the workbook's Normal-style font**
    (ECMA-376 §18.3.1.13): the renderer previously hard-coded `MDW = 8`
    (the value for Calibri 11 pt at 96 DPI), so files whose Normal style
    points to a different face produced columns at the wrong pixel width.
    The parser now resolves `<cellStyleXfs>[0].fontId` →
    `<fonts>[fontId].name.val` / `.sz.val` and exposes
    `Worksheet.defaultFontFamily` / `defaultFontSize`. The renderer
    measures the actual maximum digit width using Canvas2D's
    `measureText` (cached per `(family, sizePt)` key) and threads it
    through `RenderContext`. `colWidthToPx(w, mdw?)` now accepts MDW as
    a parameter; both `renderer.ts` and `viewer.ts` pass the resolved
    value. For Meiryo UI 10 pt the measured MDW is ≈ 6 px, so e.g.
    sample-10.xlsx column C (`width="21.125"` chars) now renders at
    127 px to match Excel — was ~174 px before. Calibri 11 pt files are
    unchanged (MDW measurement still ≈ 8) (PR #172).

## 0.24.0 — 2026-04-30

Minor release fixing spurious horizontal borders inside None-style Excel
tables and exposing column-level DXF references on `TableInfo`.

- **XLSX engine** (`packages/xlsx`):
  - **Skip the table-style overlay for "None"-style tables**
    (ECMA-376 §18.5.1.4): when `<tableStyleInfo>` has no `name` attribute
    the table uses the "None" style — no visual table formatting overlay
    should be applied. The previous renderer always entered the overlay
    block, whose fallback `else` branch drew a horizontal accent rule on
    every cell of the table. The Rust parser now defaults `style_name` to
    `""` (instead of `"TableStyleMedium2"`) when `name` is absent, and
    `buildTableStyleMap` skips tables with empty `styleName`. Eliminates
    the spurious horizontal borders that were visible across all 13 None-
    style tables of sample-7.xlsx (PR #167).
  - **Spec-faithful cell-border rendering**: an interim `suppressEdge`
    heuristic that tried to hide the leftmost-totals-cell xf borders for
    None-style tables (PR #168/#169) was reverted — those borders are
    legitimate user-set formatting per ECMA-376 §18.8.45 and Excel does
    render them. The B/C boundary across rows 50-60 is now a uniform
    thick teal vertical line, supplied by `B.right=thick` on most rows
    and `C.left=thick` on totals rows (same theme=5 colour at the same
    column boundary) (PR #170).
  - **Column-level DXF references** (ECMA-376 §18.5.1.3 `tableColumn`):
    the parser now exposes `TableInfo.columns: TableColumnInfo[]` with
    each column's `dataDxfId` / `headerRowDxfId` / `totalsRowDxfId`. The
    renderer does not consume them yet (Excel pre-bakes column DXF
    results into the cell `xf` for the common case), but the data is
    available for future named-style overlay logic that needs to layer
    column-specific DXFs on top of `wholeTable`/`headerRow` overlays
    (PR #170).

## 0.23.0 — 2026-04-30

Minor release adding drawing-shape text rendering for XLSX and correcting
column-width pixel conversion for Calibri 11pt.

- **XLSX engine** (`packages/xlsx`):
  - **Shape textboxes rendered**: stand-alone `<xdr:sp>` elements in
    `<xdr:twoCellAnchor>` (i.e. textboxes not inside a `<xdr:grpSp>` group)
    are now parsed and drawn. The parser extracts `<xdr:txBody>` content —
    paragraph alignment, per-run bold/italic/size/color/font — and the
    renderer draws them with word-wrap and body anchor support
    (ECMA-376 §20.5.2, §20.1.7.2).
  - **Theme fill / text-color fallback**: when a shape has no explicit
    `<a:solidFill>` in `<xdr:spPr>`, the fill color and default text color
    are now resolved from `<xdr:style>/<a:fillRef>` and `<a:fontRef>`
    (theme accent/lt/dk slots), matching Excel's style-inheritance chain.
  - **Vertical centering fix**: `drawShapeText` with `anchor="ctr"` now
    correctly centers the text block regardless of whether `blockH > innerH`
    — the previous `Math.max(0, …)` clamp caused text to appear
    top-aligned in those cases.
  - **Row-border z-order fix**: each cell now inherits the bottom edge of
    the cell directly above as its own top edge, so the uniform thick
    bottom-border spanning D2–Q2 is no longer overdrawn by the next row's
    `fillRect` pass (PR #164).
  - **MDW corrected to 8** (ECMA-376 §18.3.1.13): the Max Digit Width
    constant was updated from 7 to 8 to match Canvas2D measurements of
    Calibri/Carlito 11pt at 96 DPI and the actual column EMU widths that
    Excel 365 writes into drawing XML. This makes the default 8-char column
    64 px (matches Excel 100% zoom) and fixes drawing-anchor `colOff` values
    that exceeded the column width with MDW=7 (PR #165).
    ⚠ VRT reference images will differ — column widths are ~14 % wider.

## 0.22.1 — 2026-04-29

Patch release. Fix a visible layout shift on the Storybook demo when
the user first scrolled.

- **XLSX engine** (`packages/xlsx`):
  - **Force every Carlito / Caladea FontFace to actually load before
    paint**: `document.fonts.load("16px Carlito")` only triggers the
    sub-ranges that match characters already present in the live DOM.
    Canvas-only text didn't qualify, so on systems without Calibri the
    first paint fell back to sans-serif metrics and the moment a scroll
    re-rendered, the browser had lazy-loaded the substitute and switched
    to its real metrics — a visible reflow. Iterate `document.fonts` for
    the requested families and call `face.load()` directly, bypassing
    the unicode-range gating.

## 0.22.0 — 2026-04-29

Minor release rolling up the v0.21.x XLSX engine work and refreshing the
README screenshot.

- **Assets**:
  - Re-captured `docs/images/xlsx.png` at 1886×1064 (was 1246×735) so
    the xlsx column in the README screenshot table matches the docx /
    pptx column widths instead of rendering visibly narrower.
- **XLSX engine** (cumulative since 0.21.0):
  - `centerContinuous` runs hide their internal default gridlines and
    explicit cell borders, leaving only the outer perimeter visible —
    matching Excel's merged-cell-style appearance (PRs #152, #153,
    #155, ECMA-376 §18.18.40).
  - `defaultColWidth` derived from `baseColWidth` when the file omits
    the explicit attribute (PR #156, ECMA-376 §18.3.1.81).
  - `defaultRowHeight` honored only when `customHeight="1"` is set;
    otherwise the renderer falls back to its 20-px intrinsic baseline,
    matching Excel (PR #157).
  - Row heights treated as 1:1 display pixels rather than the nominal
    point size — Excel writes the resolved pixel value into `ht`, so
    the previous 4/3 conversion produced rows ~33 % too tall (PR #158).
  - `double` border style (ECMA-376 §18.18.3 ST_BorderStyle) now
    rendered as two parallel 1-px lines with closed corners (PRs #158,
    #159).

## 0.21.3 — 2026-04-29

Patch release. XLSX cell sizing, border-style coverage, and
`centerContinuous` run accounting fixes.

- **XLSX engine** (`packages/xlsx`):
  - **`centerContinuous` run split at value-bearing anchors**: ECMA-376
    §18.18.40 says spanned cells reference the same style id while only
    the anchor holds a value. Two adjacent anchors therefore form two
    separate runs, with the border between them visible. The renderer
    now closes the previous run and starts a fresh one whenever it
    encounters a centerContinuous cell that itself carries a value
    (PR #155).
  - **`defaultColWidth` derived from `baseColWidth`**: ECMA-376
    §18.3.1.81 specifies that when `<sheetFormatPr>` provides
    `baseColWidth` but no `defaultColWidth`, the default column width
    is `baseColWidth + 5 px / maxDigitWidth` characters (4 px margin +
    1 px gridline). The Rust parser now applies that fallback so
    sample-27.xlsx (`baseColWidth="10"`, no `defaultColWidth`) renders
    columns at 75 px instead of the previous 8.43-char baseline
    (PR #156).
  - **`defaultRowHeight` honored only when `customHeight` is true**:
    ECMA-376 §18.3.1.81 `customHeight` describes the
    `defaultRowHeight` attribute as informational unless the flag is
    set; Excel falls back to its intrinsic 20-px baseline otherwise.
    The parser now matches that behavior (PR #157).
  - **Row heights treated as display pixels (1:1)**: Excel writes row
    heights as the resolved display-pixel value rather than the
    nominal point size — `defaultRowHeight="20"` renders 20 px,
    `ht="21"` with `thickBot` renders 21 px, etc. The renderer's
    pt→px multiplier is dropped for row metrics (the 4/3 conversion
    is preserved for fonts, indent, and line-height) and the parser
    intrinsic moves from 15 → 20 (PR #158).
  - **`double` border style** (ECMA-376 §18.18.3 `ST_BorderStyle`):
    previously rendered as a single line because the dispatcher only
    handled thick / medium / dashed-family / dotted / hair. Now drawn
    as two parallel 1-px lines flanking the cell edge with closed
    corners — outer pair extended past corners by 1 px so adjacent
    doubled edges meet outside, inner pair trimmed by 1 px so the
    perpendicular inner segments meet exactly at the inner corner
    (PRs #158, #159).

## 0.21.2 — 2026-04-29

Patch release. XLSX `centerContinuous` border rendering fixes.

- **XLSX engine** (`packages/xlsx`):
  - **Default gridlines hidden inside `centerContinuous` runs**: a
    contiguous run of cells with `horizontal="centerContinuous"` is now
    rendered as a single visual span — internal default gridlines are
    suppressed so the centered text reads as one (PR #152, ECMA-376
    §18.18.40).
  - **Explicit cell borders also hidden inside `centerContinuous`
    runs**: matching Excel's behavior, internal `left`/`right` borders
    of cells inside a centerContinuous run are masked, leaving only the
    outer perimeter visible — even when the run cells carry an
    explicit thin/medium/thick box border (PR #153).

## 0.21.1 — 2026-04-29

Patch release. Re-capture `docs/images/xlsx.png` at the correct landscape
dimensions (1246×735) — the 0.21.0 update accidentally shipped a
narrow 668×1200 portrait crop.

## 0.21.0 — 2026-04-29

Patch-level release. XLSX overflow rendering fixes and refreshed project icon.

- **XLSX engine** (`packages/xlsx`):
  - **`centerContinuous` overflow**: when text was wider than the
    selection range, it was clipped at the range boundary. Excel keeps
    overflowing the text symmetrically into adjacent empty cells, so the
    renderer now extends the draw rect using the same logic as
    `center` alignment (PR #148, ECMA-376 §18.18.40).
  - **Neighbour fill no longer overpaints overflow text**: cell text is
    drawn in a deferred second pass after every cell's background, so
    e.g. a left-aligned overflow stays visible on top of an adjacent
    cell with a non-default fill — matching Excel's z-order (PR #148).
- **Assets**:
  - Refreshed `docs/images/icon.png` and the VS Code extension icon
    (PR #149).

## 0.20.0 — 2026-04-28

Minor release. Built-in PowerPoint table styles and additional XLSX cell alignments.

- **PPTX engine** (`packages/pptx`):
  - **74 built-in table style presets**: all PowerPoint ribbon table styles
    (Themed 1/2, Light 1/2/3, Medium 1/2/3/4, Dark 1/2 × accent variants)
    are now resolved from a hard-coded GUID catalog when no
    `ppt/tableStyles.xml` definition is present. Fills and borders are
    computed directly from the presentation theme (PR #146,
    ECMA-376 §19.3.1.39; GUID catalog from LibreOffice MPL 2.0 source).
- **XLSX engine** (`packages/xlsx`):
  - **Cell alignment — `centerContinuous`, `fill`, `distributed`,
    `justify`, `readingOrder`**: five additional `<xf alignment>` modes
    are now parsed and applied during cell rendering (PR #145).

## 0.19.0 — 2026-04-27

Minor release. Spec-faithful improvements to scatter chart rendering
on the XLSX side and an Excel-matching font cascade for cell text.

- **PPTX engine** (`packages/xlsx`, `packages/core`):
  - **Scatter X-axis position**: parse `<c:catAx><c:crosses>` /
    `<c:crossesAt>` and draw the X-axis line at the resolved Y
    coordinate (`autoZero` → `toY(0)`). The Vertex42 "Project
    Timeline" template puts milestones above and tasks below the
    timeline ruler — that finally renders correctly because the
    ruler now sits at y=0 instead of the bottom of the plot rect
    (PR #140).
  - **Bold and font sizes from `<c:txPr>` / `<a:defRPr>`**: chart
    title, both axis tick labels, series-level `<c:dLbls>`
    defaults (size 1200/1400 = 12 / 14 pt), and per-idx `<c:dLbl>`
    rich text now honor `b="1"` and `sz="..."` (PR #140).
  - **Axis line styling**: `<c:catAx | valAx><c:spPr><a:ln>`
    resolved color and width are applied to the axis line stroke
    (sample-26's 5 pt 50 % gray timeline ruler now renders as
    intended) (PR #141).
  - **Auto-derived axis range snaps to nice round values**: when
    `<c:scaling><c:min/max>` aren't set, we expand both ends to a
    multiple of `niceStep`. Sample-26 X-axis now spans
    2018/2/19 .. 2018/10/27 in 50-day steps, matching Excel
    (PR #141).
  - **Tick marks honor `<c:majorTickMark>` / `<c:minorTickMark>`**:
    the XLSX adapter was hard-coding `'cross'`; replaced with the
    parsed value (default `'out'` per ECMA-376 §21.2.2.49). Scatter
    renderer now actually calls `drawAxisTick` at every major tick
    on both axes; tick stroke inherits the axis line color and
    width (PR #142).
- **XLSX cell font** (`packages/xlsx`):
  - Calibri-styled cells fell back to system Arial / Helvetica on
    macOS / Linux, which is ~10–15 % wider per character at every
    weight × size — visible on `demo/sample-1` where the B2 title
    overflowed past column F instead of stopping inside column E.
    Add an opt-in `useGoogleFonts` to `XlsxWorkbook.load` (and
    forwarded from `XlsxViewer`'s constructor option) that loads
    [Carlito](https://fonts.google.com/specimen/Carlito) and
    [Caladea](https://fonts.google.com/specimen/Caladea) — Google's
    metric-compatible substitutes for Calibri and Cambria. With the
    substitutes loaded, advance widths match Excel and column
    layout shows the title fitting inside the same cells Excel
    does. Default off — zero third-party requests until the host
    opts in (PR #143).

The unified `ChartModel` and `ChartSeries` gained matching optional
fields. PPTX charts continue rendering unchanged — the PPTX parser
will pick up the same fields in a follow-up release.

## 0.18.2 — 2026-04-27

Patch release. Sparkline visual polish.

- **Core sparkline** (`packages/core/src/sparkline/renderer.ts`):
  - Vertical padding inside the cell switched from a 2 px cap to a
    proportional 20 % of the cell height (with a 2 px floor). The peak
    / trough of the line and high / low marker dots no longer overlap
    the row separators, and the breathing room stays consistent across
    zoom levels.

## 0.18.1 — 2026-04-27

Patch release. One sparkline correctness fix.

- **Core sparkline** (`packages/core/src/sparkline/renderer.ts`):
  - `computeFlagged` now marks **every** point tied for the high or low
    value, not just the first occurrence. Excel does the same. Visible
    on `private/sample-7.xlsx` Q10 (3 non-zero leading values + 9 zeros
    under `low="1"`): all 9 zero points are now dotted. Other
    highlights (`first` / `last` / `negative`) are unchanged — only
    `high` / `low` had the tied-value case.

## 0.18.0 — 2026-04-27

Minor release. Excel scatter / bubble charts gain a spec-faithful set of
features so Gantt-style "Project Timeline" templates render against the
reference. Sparkline release reused — no docx/pptx behavioral changes.

- **PPTX engine** (`packages/xlsx`, `packages/core`):
  - Per-point `<c:dPt>` overrides (color, marker shape / size / fill /
    line) — ECMA-376 §21.2.2.39.
  - `<c:marker>` symbol / size resolution with all 10 ECMA shapes
    (circle / square / diamond / triangle / x / plus / star / dot /
    dash / picture). `picture` falls back to circle pending image-marker
    resolution.
  - `<c:dLbl>` per-point custom data labels (§21.2.2.45) with rich-text
    flattening and `<a:fld type="CELLRANGE">` substitution from the
    series' `<c15:datalabelsRange>` cache. Position (`l`/`r`/`t`/`b`/
    `ctr`/`outEnd`) is honored.
  - `<c:errBars>` (§21.2.2.20) X / Y direction with all five
    `errValType` modes — `cust` (cell-range), `fixedVal`, `stdErr`,
    `stdDev`, `percentage`. Cust values can be signed (Vertex42 Gantt
    uses negative minus values to flip stems toward the X-axis).
    Dashed strokes and end caps via `<a:prstDash>` / `<c:noEndCap>`.
  - `<c:title><c:layout>` and `<c:plotArea><c:layout>` manual layout
    (§21.2.2.27) — when present, used directly for absolute placement.
  - Scatter's dual `<c:valAx>` blocks are now disambiguated by
    `<c:axPos>` so X-axis (`b`/`t`) and Y-axis (`l`/`r`) settings end
    up in their respective slots. `<c:scaling><c:min/max>` and
    `<c:numFmt>` are picked up per axis. `<c:delete val="1"/>` on
    either axis hides it correctly.
  - `formatChartValWithCode` recognises date format codes (m/d/y/h/s
    outside quotes) and routes through a new `formatExcelDate` so
    scatter X-axis tick labels for date series come out as
    `2018/4/12` instead of raw serial numbers.

The unified `ChartModel` / `ChartSeries` gained the corresponding
optional fields. PPTX charts continue rendering unchanged — the PPTX
parser hasn't been updated to populate them yet (follow-up).

## 0.17.0 — 2026-04-27

Minor release. Handwritten ink strokes now render via PowerPoint's
rasterized fallback, plus rendering accuracy fixes for the PowerPoint
engine.

- **PPTX engine** (`packages/pptx`):
  - `mc:AlternateContent` now walks `mc:Fallback` when `mc:Choice`
    produces no output (previously Choice was always taken and Fallback
    silently discarded). PowerPoint embeds ink / handwriting as
    `p:contentPart` (InkML) inside Choice with a rasterized `p:pic`
    inside Fallback; this restores those strokes.
  - When the Choice subtree is an ink `p:contentPart`, the fallback PNG
    is rendered at its natural pixel size centered in the bounding box
    rather than always stretched. Empty / single-tap strokes (whose
    fallback PNG is only a few pixels) no longer blow up into blocky
    artifacts. Visible strokes are unaffected.
  - Fix preset-shape arc visual-to-parametric angle conversion for
    non-square shape boxes (ECMA-376 §20.1.9.18 `<a:arcTo>`). The
    path-executor was using canvas-scaled radii where path-local radii
    were required, skewing every arc segment when `sx ≠ sy`. Visible
    on `cloudCallout` placed in a landscape box, where the inner cloud
    detail arcs were misaligned even though the outline looked plausible.
  - `PptxViewer` wrapper sets `vertical-align: top` so the inline-block
    line-box descender (~6 px on default font metrics) no longer leaks
    the host container's background through below the canvas.
- **VS Code extension** (`packages/vscode-extension`):
  - Downsize `icon.png` to 512×512 to match the practical Marketplace
    icon range. No functional change.

## 0.16.1 — 2026-04-26

Patch release. Project icon refresh.

- **Project icon**: replace `docs/images/icon.png` with a refreshed master
  and resync `packages/vscode-extension/icon.png` from it. README and the
  VS Code Marketplace listing now show the updated artwork.

## 0.16.0 — 2026-04-26

Minor release. Audio / video–embedded files now open in the VS Code
extension and play back interactively.

- **VS Code extension** (`packages/vscode-extension`):
  - Replace the `Array.from(bytes)` + `webview.postMessage` data path with
    `webview.asWebviewUri()` + `fetch().arrayBuffer()` in the webview. The
    previous IPC route serialized file bytes as a JSON number array, which
    hung the spinner indefinitely on media-embedded pptx (50–200 MB) and,
    after a stop-gap that sent `Uint8Array` directly, returned zero-byte
    buffers (`Could not find EOCD` for every file) because VS Code's
    webview `postMessage` does not reliably structured-clone typed arrays.
    Same approach the bundled PDF viewer uses; native binary path, no IPC
    size or type cliff.
  - Switch the scroll-stack pptx renderer from `PptxPresentation.renderSlide`
    to `presentSlide` so embedded audio / video become clickable with a
    canvas-native play / pause / progress bar.
- **PPTX engine** (`packages/pptx`):
  - `presentSlide` now forwards `opts.onTextRun` to its inner `renderSlide`
    call so the transparent text-selection layer keeps working when
    interactive playback is enabled. The same bug existed when calling
    `PptxViewer` with both `enableMediaPlayback: true` and
    `enableTextSelection: true`; fixed at the engine level so both paths
    benefit.
  - `createPresentationHandle` now skips its `requestAnimationFrame` loop
    and pointer wiring for slides without media, so a 50-slide deck no
    longer spawns 50 idle animation loops.

## 0.15.2 — 2026-04-26

Patch release. Project icon refresh.

- **Project icon**:
  - Adopt a new high-resolution master at `docs/images/icon.png` (2048×2048)
    and reference it from the root README.
  - **VS Code extension** (`packages/vscode-extension`):
    - Replace the previous 128×128 Marketplace icon with the new master so the
      Marketplace listing renders crisply on retina displays.
    - Show the icon at the top of the extension README (via the GitHub raw
      URL — Marketplace ignores relative image paths).
    - Add a `cp ../../docs/images/icon.png ./icon.png` step to
      `vscode:prepublish` so the bundled `.vsix` icon is always re-synced from
      the master at publish time and never drifts.

## 0.15.1 — 2026-04-26

Patch release. Mobile UX fix for the XLSX viewer plus Storybook tidy-up.

- **XLSX viewer** (`packages/xlsx`):
  - Distinguish tap from swipe on touch / pen input. `pointerdown` no longer
    commits a cell selection for non-mouse pointers; the gesture is buffered
    and only commits on `pointerup` if the pointer stayed within an 8 px slop.
    A swipe to scroll on a phone or tablet now leaves the selected cell
    alone. Mouse input is unchanged so drag-to-extend keeps working.
- **Storybook**:
  - Drop the `Selectable — file upload` / `Selectable — sample-1.xlsx`
    stories. The public demo already exercises cell selection, so the stories
    were redundant.

## 0.15.0 — 2026-04-25

VS Code extension polish + selection overlay accuracy fix. No new format
support compared to 0.14.1; library packages are bumped to 0.15.0 so the
tag-driven CI keeps the npm versions in sync with the VS Code Marketplace
release.

- **VS Code extension** (`packages/vscode-extension`):
  - Add Marketplace icon (`icon.png`, 128×128) and wire it up via `package.json#icon`.
  - Shorten `displayName` from `Office Viewer — DOCX, XLSX, PPTX` to `Office Viewer`; supported formats remain in the description and feature list.
  - Replace the plain text loading status with a CSS-only spinner (#107).
  - Center the loading spinner and error status on the viewport.
  - Open documents in the currently focused column instead of forcing a split (#114).
- **Viewer / selection overlay** (`packages/docx`, `packages/pptx`, VS Code webview):
  - Carry the canvas `ctx.font` shorthand (font-family / weight / style) through `TextRunInfo` / `DocxTextRunInfo` and apply it on the transparent selection `<span>` so its width tracks the drawn glyphs. Previously the overlay relied on a fallback font, which drifted at the trailing edge of European text. Kerning / ligatures are intentionally left at the browser default to match canvas behavior.
- **Docs**:
  - Add a forward-looking note explaining the dual-layer (canvas + transparent DOM) selection architecture as a deliberate stop-gap, with a reference to the WICG `html-in-canvas` `drawElement` API as the planned unified replacement.
  - VS Code extension README: 3-column screenshot table for the Marketplace listing.
  - Standardize format ordering across READMEs as DOCX → XLSX → PPTX.
  - Add Marketplace badges to the root README.

## 0.14.1 — 2026-04-25

VS Code Marketplace metadata fix. The Marketplace `vsce publish` of 0.14.0
failed because the extension `name` (`ooxml-viewer`) collided with an existing
listing. This release renames the extension and broadens the displayName so
that users searching the Marketplace for "Office", "XLSX", "DOCX", or "PPTX"
discover it.

- **Rename** (`packages/vscode-extension`):
  - `name`: `ooxml-viewer` → `office-open-xml-viewer`
  - `displayName`: `OOXML Viewer` → `Office Viewer — XLSX, DOCX, PPTX`
  - `description`: emphasizes Office file support and the local-only privacy
    posture for Marketplace searchers.
- Library packages (`@silurus/ooxml{,-pptx,-xlsx,-docx,-core}`) are bumped to
  0.14.1 to keep tag-driven CI in sync; **no code changes for the libraries.**

## 0.14.0 — 2026-04-25

VS Code extension UX overhaul. The `.docx` and `.pptx` editors switch from a
prev/next pager to a **continuous scroll-stack** that renders every page or
slide at once with a transparent text layer (PDF.js style). The Webview chrome
now follows the active VS Code theme (light / dark / high-contrast) via
`--vscode-editor-background` and `--vscode-foreground`. The Marketplace README
gains screenshots and a privacy statement asserting zero network access.

### vscode-extension

- **Scroll-stack viewer** (`packages/vscode-extension`) — replaces the
  page-by-page navigation for docx/pptx. Every page/slide is rendered
  vertically with its own transparent text layer; selection and copy work
  across the whole document.
- **Theme-aware backgrounds** — body/foreground driven by VS Code CSS
  variables; the chrome around documents follows the active theme without
  hardcoded fallbacks.
- **CSP + handshake** — workers accept a `data:`-URL wasm asset (decoded
  inside the worker) so the Webview CSP can stay strict; the editor waits for
  a `webview-ready` ping before posting the file payload, fixing an init
  ordering race.
- **Marketplace README** — adds screenshots (absolute raw URLs so they render
  on the Marketplace page), a "Privacy & Security" section, and reflects the
  scroll-view UX.

### tooling

- **pptx wasm script** (`packages/pptx`) — switch to `wasm-pack build
  --out-dir`, matching xlsx/docx, so the generated `pptx_parser.d.ts` is
  written into `src/wasm/`. Resolves the CI `Build library packages` failure
  where per-package `tsc --build` errored on `worker.ts` with TS7016.
- `.gitignore`: exclude `*.vsix` build output.

## 0.13.0 — 2026-04-25

UX and tooling release. The core viewer packages gain **text and cell
selection** (PDF.js-style transparent overlay so the browser's native
selection/copy work on top of Canvas). Two new companion packages ship
alongside: a **VS Code extension** (`ooxml-viewer`) that registers custom
editors for `.xlsx` / `.docx` / `.pptx`, and a **Rust MCP server**
(`ooxml-mcp-server`) that exposes the parsers as structured tools for AI
agents. No rendering-fidelity changes.

### viewer UX

- **Text selection overlay (pptx/docx/xlsx)** — each viewer now emits an
  `onTextRun` stream from the renderer and mounts an absolute-positioned
  `<span>` per text run above the canvas with `color: transparent`. The
  browser's native selection, copy, and `::selection` styling all work
  against the overlay, so users can select, Ctrl+C copy, or drag text
  exactly as they would in a DOM-rendered document. pptx handles rotated
  and vertical text via `transform: rotate(...)` on the overlay spans.
- **xlsx cell selection** (`packages/xlsx`):
  - `getCellAt(clientX, clientY)` on `XlsxViewer` hit-tests canvas
    coordinates to row/col addresses (respects merged cells and freeze
    panes).
  - Four selection modes: single cell, range, row (click row header),
    column (click col header), all (corner click). Drag to extend.
    Shift+click extends from the current anchor.
  - `Ctrl+C` copies the selected range as tab-separated text (TSV) to the
    clipboard, mode-aware (full row → entire row; range → block).
  - `onSelectionChange` callback on `XlsxViewerOptions`; `selection`
    getter on the viewer. New exports: `CellAddress`, `CellRange`,
    `SelectionMode`, `TextRunInfo`.
  - Selection overlay clamps to the header/freeze-pane boundaries so the
    highlight doesn't bleed over the sticky row/column bands.

### New package: VS Code extension (`packages/vscode-extension`)

- Registers `CustomEditorProvider` for `.xlsx`, `.docx`, and `.pptx`, so
  double-clicking an Office file in the VS Code explorer opens it in the
  same Canvas viewer used by the Storybook demo.
- Webview bundles the existing `XlsxViewer` / `DocxViewer` / `PptxViewer`
  classes; selection events can be relayed to the extension host via
  `acquireVsCodeApi().postMessage()`.

### New package: Rust MCP server (`packages/mcp-server`)

- Exposes the existing xlsx/docx/pptx parsers as an MCP server so agents
  (Claude, Copilot, Codex, …) can query OOXML files without shelling out
  to `unzip` + ad-hoc Python. Structured tools include
  `xlsx_get_cell_range`, `xlsx_get_formulas`, `docx_get_structure`,
  `docx_get_tables`, `pptx_get_slide_structure`, and format-specific
  search helpers.
- Built natively from the same Rust crates (the `rlib` output of
  `packages/{xlsx,docx,pptx}/parser`), so the parser logic is shared
  with the browser build one-to-one.

## 0.12.0 — 2026-04-25

xlsx fidelity release focused on sample-1 ("Holiday shopping budget") and
sample-10 ("Calendar"): static pivot/table **slicers** now render from the
Office 2010 extension, **chart data labels** honor per-series `txPr`
(white-on-bar) and the `<c:dLblPos>` / `<c:numFmt>` / `<c:gapWidth>` /
`<c:overlap>` chart attributes, and several conditional-formatting / text
layout fixes land for the calendar sample.

### xlsx

- **Static slicers** (Office 2010 extension `x14:slicerList`, §A.5): parse
  `xl/slicers*.xml` + `slicerCaches/` and render the slicer button array
  with its header and theme-resolved accent fill. Slicers for pivot and
  table sources both lay out correctly; "in-filter" vs "out-of-filter"
  button colors come from the slicerStyle dxfs.
- **Chart bar gap + overlap** (§21.2.2.13 `c:gapWidth`, §21.2.2.25
  `c:overlap`): bar cluster geometry now uses the spec formula
  `clusterWidth = barW · (1 + (N-1)·(1-overlap/100) + gapWidth/100)` so
  paired bars in a two-series chart show the expected gap between
  category clusters instead of flush-packed bars.
- **Chart data-label position and number formats** (§21.2.2.16
  `c:dLblPos`, §21.2.2.21 / .35 / .37): value-axis tick labels honor
  `<c:valAx><c:numFmt>`, data labels honor `<c:dLbls><c:numFmt>` with a
  per-series `<c:val><c:numRef><c:formatCode>` fallback, and labels
  render at the requested position (`inBase` / `inEnd` / `ctr` /
  `outEnd`) with collision-safe placement on horizontal bars.
- **Per-series data-label font color** (§21.2.2.47 `c:ser/c:dLbls`):
  Excel frequently writes the label `txPr` (including `schemeClr
  val="bg1"` → white) on each series rather than on the chart-level
  `<c:dLbls>`. The parser now falls back to the first series's dLbls
  when the chart-level block omits the color, fixing white-on-bar
  labels.
- **Horizontal bar series ordering** (§21.2.2.28 `c:order`,
  §21.2.2.40 `c:delete`): series are sorted by their declared order
  and the visual stack is reversed for horizontal bars so the first
  series appears on top (matches Excel). `<c:catAx><c:delete val="1"/>`
  and `<c:valAx><c:delete>` hide the corresponding axis band, freeing
  padding for the chart itself.
- **Pie/doughnut, radar, waterfall data-label formats**: the same
  `valAxisFormatCode` / `dataLabelFormatCode` plumbing flows through
  non-bar renderers, so value labels on those types pick up the file's
  Excel number-format code (e.g. `¥#,##0.00`).
- **Transparent chart space + theme-palette series colors**
  (§21.2.2.39 `c:chartSpace/c:spPr`): `<a:noFill>` on the chart space
  keeps the underlying cell grid visible behind the chart (Excel's
  default). Series `<c:spPr>` with `<a:schemeClr val="accent1"/>` etc.
  now resolves against the file theme instead of falling through to
  the renderer's built-in palette.
- **Legend manual layout** (§21.2.2.31 `c:legend/c:manualLayout`):
  absolute `x`/`y`/`w`/`h` placement fractions override the default
  side-of-plot legend rectangle while `legendPos` still chooses which
  side of the plot gets the reserved band.
- **dxf numFmt override from conditional formatting** (§18.3.1.10
  `dxf/numFmt`): `cellIs` / `top10` / `aboveAverage` etc. CF rules that
  point to a dxf with a `<numFmt>` now apply that format code when the
  rule matches, not just the fill and font overrides.
- **dxf patternType=none as explicit fill clear**: treats `<patternFill
  patternType="none"/>` inside a dxf as an explicit override that
  unsets the base cell fill, not as "inherit base fill". Matches Excel
  UI where the CF explicitly removes the background.
- **4th format-section (text) honored** (§18.8.30, `;;;` idiom):
  `#,##0;[Red](#,##0);0;@` now applies the fourth section to text-typed
  cells; a `;;;` code correctly hides both numeric *and* text cells.
- **`notContainsBlanks` conditional formatting** (§18.18.15
  `ST_CfType`): the opposite of `containsBlanks`; rules of this type
  now paint non-empty cells instead of silently skipping all cells.
- **`<xdr:grpSp>` custom geometry** (§20.5.2.17): group-shape children
  with `<a:custGeom>` inherit the group's frame transform, so grouped
  freeform icons draw at the correct position and scale.
- **Japanese calendar date format `ge.m.d`** (§18.8.30): era-prefixed
  numeric dates (`R7.4.25`) render alongside the existing era name /
  era year / weekday codes landed in 0.10.0.
- **Image in grouped anchor** (`<xdr:grpSp>` + `<xdr:pic>`): pictures
  nested inside a group anchor no longer drop; the group transform is
  applied to the embedded image frame before rendering.
- **CJK wrap on wrapText cells**: break opportunities between Kanji /
  Hiragana / Katakana characters are recognized when `wrapText="1"`
  is set, matching Excel's line-break behavior for Japanese text.
- **CF over empty cells**: rules that previously required a cached
  `<v>` now also evaluate empty cells against the `containsBlanks` /
  `notContainsBlanks` / text operators.
- **Scroll-flicker fix**: the virtual-scroll frame awaits the next
  animation tick before clearing the canvas, eliminating the flash
  of blank cells during fast scroll.
- **`ShapeGeom::Image` JSON field**: serialized as `dataUrl` (camelCase)
  to match the rest of the parser's JSON surface; fixes shape-image
  rendering in downstream renderers that weren't converting the
  snake_case variant.

## 0.11.0 — 2026-04-22

xlsx fidelity release focused on sample-9 ("Gift budget and tracker"):
stacked combo charts keep their stacking, chart series honor theme
accent colors, custom `<tableStyle>` elements actually style their
cells, and `cellIs` conditional-formatting rules match text operands.

### xlsx

- **Stacked combo charts** (§21.2.2.17): locking `grouping` once a
  non-line series sets it prevents a trailing `lineChart grouping=
  "standard"` from overwriting the bar's `stacked` / `percentStacked`,
  so bar+line combos keep stacked bars.
- **Chart series `<a:schemeClr>` resolution** (§21.2.2.35 `c:spPr`):
  series colors declared as `accent1`..`accent6` / `dk*` / `lt*` are
  resolved against the file's theme color table instead of falling
  back to palette defaults.
- **Custom `<tableStyle>` elements** (§18.8.40): parse `wholeTable`
  and `headerRow` dxf indices from `xl/styles.xml/tableStyles`, then
  overlay the resolved dxf fill, font color, and horizontal / vertical
  borders on top of cells. Built-in style names keep the existing
  accent-based renderer unchanged. `Border` gains `horizontal` /
  `vertical` to carry the inner-rule edges emitted only by tableStyle
  dxfs.
- **Text operands in `cellIs` CF rules** (§18.18.15 `ST_CfOperator`):
  `cellIs` previously only evaluated numeric cells, so text rules like
  `equal "Birthdays"` silently skipped every non-numeric row. Now
  parses each `<formula>` as a quoted string literal or number, and
  compares case-insensitively for equal / notEqual / containsText /
  notContains / beginsWith / endsWith / between / notBetween.

## 0.10.0 — 2026-04-22

xlsx number-format and volatile-function release. Cells with `TODAY()` /
`NOW()` formulas now show today's date at render time instead of the
cached `<v>` from when the file was last saved, and the format-code
renderer gains Japanese weekday / imperial era support plus several
internationally important codes (elapsed time, literal preservation,
scientific notation).

### xlsx

- **Volatile formula recompute** (§18.3.1.40): the parser now carries
  each cell's `<f>` text, and the renderer detects `TODAY()` / `NOW()`
  and substitutes the live serial before formatting. Dates no longer
  appear frozen to the file's last-save date.
- **Japanese weekday format codes** (§18.8.30): `aaa` → 水, `aaaa` →
  水曜日. Detected as date formats even without a `y`/`d` specifier.
- **Japanese imperial era format codes** (§18.8.30): `g` / `gg` / `ggg`
  render the era name (R / 令 / 令和) and `e` / `ee` / `r` / `rr` render
  the era year. Era table covers Meiji through Reiwa; no runtime
  dependency added.
- **Elapsed-time brackets** `[h]` / `[m]` / `[s]` (§18.8.30): render the
  full duration instead of wrapping at 24h / 60m / 60s, so a 54-hour
  value formatted `[h]:mm` reads `54:00`.
- **Literal text preservation in number formats**: quoted strings
  (`"$"#,##0.00`) and backslash-escaped characters (`\$#,##0`), as well
  as non-placeholder currency glyphs like `¥` / `€`, are now kept around
  the formatted number instead of being stripped.
- **Scientific notation** `0.00E+00` / `0.00E-00`: honors the exponent
  sign placeholder and pads the exponent to at least two digits.

## 0.9.0 — 2026-04-22

Focused xlsx release: conditional formatting now evaluates formula-based
rules, resolves defined names, overlays `<dxf>` borders per edge, and
honors Excel's `x14:dataBar@gradient="0"` for solid bars. The CF formula
evaluator is broadened to cover the functions most commonly used in
`expression` rules.

### xlsx

- **Conditional formatting — `expression` rules** (§18.3.1.10): the
  formula is tokenized, references are shifted by the sqref anchor, and
  the AST is walked to a boolean. `stopIfTrue` and rule priority are
  honored so later rules can't mask earlier hits.
- **Defined-name resolution** (§18.2.5): sheet-scoped names used inside
  CF expressions (e.g. `task_start`, `today`) are resolved by inlining
  the formula and shifting embedded relative refs from A1.
- **CF `<dxf>` borders** (§18.8.17): per-edge overlay — a CF rule can
  draw a red left/right stripe without erasing the cell's existing
  top/bottom border.
- **Data-bar gradient flag**: `x14:dataBar@gradient="0"` (living in a
  separate worksheet-level `<extLst>` linked by GUID) now produces a
  solid fill. Previously bars always rendered with a gradient.
- **Data-bar / color-scale theme colors**: `<color theme="…" tint="…">`
  inside `<dataBar>`/`<colorScale>` is now resolved through the workbook
  theme (was srgb/indexed only).
- **`sheetView showGridLines`** (§18.3.1.83): when unchecked in Excel's
  View tab, the default `#d0d0d0` grid lines are no longer drawn.
- **Formula evaluator broadening** (for CF `expression` rules): `A1:B5`
  ranges; `&` concatenation; IFERROR/IFS; type checks (ISTEXT, ISERROR,
  ISNA, …); math (TRUNC, CEILING, FLOOR, MOD, POWER, SQRT, SIGN, EXP,
  LN, LOG10); aggregates (AVERAGE, COUNT/COUNTA/COUNTBLANK, COUNTIF,
  SUMIF, AVERAGEIF with operator-prefixed criteria); text (LEN, LEFT,
  RIGHT, MID, UPPER, LOWER, TRIM, EXACT, FIND, SEARCH, CONCATENATE, T,
  N, VALUE); ROW/COLUMN; date (TODAY, NOW, DATE, YEAR, MONTH, DAY,
  WEEKDAY, with the 1900 leap-year serial compensation).

## 0.8.1 — 2026-04-21

### Infrastructure

- **Demo URL changed to `https://ooxml.silurus.dev`** — custom domain updated
  from `demo.silurus.dev` to `ooxml.silurus.dev` for scalability across future libraries.

## 0.8.0 — 2026-04-21

### Infrastructure

- **Demo URL changed to `https://demo.silurus.dev`** — GitHub Pages now served
  from a custom domain. README and npm `homepage` field updated accordingly.
- Storybook build base path simplified to `/` (was `/office-open-xml-viewer/`
  on CI); `CNAME` file is now written into the artifact on every deploy.

## 0.7.0 — 2026-04-21

Quality pass across pptx shape rendering and chart legends — no new
feature categories, but several existing ✅ features now match the
Office output more faithfully.

### pptx

- **`cxnSp` connectors honor `<p:style><a:lnRef idx="N">`** as a stroke
  fallback (#74). Previously a connector that only declared
  `headEnd` / `tailEnd` on `<a:ln>` (no `solidFill`) rendered invisible;
  the style-level stroke now fills in color and width.
- **`<p:style><a:lnRef>` stroke width resolves from the theme's
  `fmtScheme > lnStyleLst`** for both `<p:cxnSp>` (#74) and `<p:sp>`
  (#76). The previous hard-coded 9525 EMU (0.75 pt) under-weighted
  every idx ≥ 2 stroke — idx=2 is 19050 EMU (1.5 pt) and idx=3 is
  25400 EMU (2 pt) in the Office default theme. Brackets, braces, and
  arcs that inherited the style line now render at the thickness
  PowerPoint shows.
- **`<a:tint>` mixes in linear sRGB** (IEC 61966-2-1) rather than
  straight sRGB (#74). Sampling the PDF export of the reference
  SmartArt arrow (#156082 + tint=60000) yields ~#D1D6DB, which the
  linear-sRGB lerp now reproduces pixel-for-pixel.
- **`bentConnector{2-5}` / `curvedConnector{2-5}` routed through the
  ECMA-376 preset path evaluator** (#74), and `getConnectorAnchors()`
  walks the preset cmd list so arrow heads sit on the true tangent
  angle instead of the bounding-box diagonal.
- **`rtTriangle` prstGeom** (right-angle at bottom-left) gained a
  proper path (#74); previously fell back to `rect`.
- **`adj5`–`adj8` threaded through parser → renderer → preset
  evaluator** (#74) for callouts whose gdLst references them
  (e.g. `accentBorderCallout3`).

### charts

- **`c:legendPos` and marker visibility** now drive legend placement
  and series point rendering across the chart families (#72); radar
  charts also honor the value-axis scale instead of defaulting to
  `0–max`.

### xlsx

- **Data bar conditional formatting** renders with the Excel 2010+
  gradient fill instead of the flat solid color (#73), matching the
  in-cell gradient Excel draws.

### Docs

- README screenshots refreshed for the release.
- CLAUDE.md codifies two workflow rules: squash merges to `main` are
  forbidden (use `--merge` or `--rebase`), and the release process
  (README screenshots + support table + CHANGELOG + version bump)
  is documented as a single PR procedure.

## 0.6.0 — 2026-04-21

### docx

Layout improvements driven by cross-referencing Word's PDF export of
demo/sample-1 with our paginator / line-layout output. Unless noted, the
work lands as strict ECMA-376 reading of the relevant sections — empirical
tolerance knobs were deliberately avoided per the project's spec-first
rule.

- **Line spacing, explicit vs inherited (ECMA-376 §17.6.5 + §17.3.1.33).**
  `line_spacing_explicit` now flows through the style cascade. A paragraph
  whose `w:spacing/@w:line` is inherited only from docDefault snaps to one
  grid pitch per line in a `w:docGrid`-enabled section; a paragraph that
  sets `w:line` on its own pPr or a named style multiplies against the
  pitch. Fixes body labels like ESSAY / BY THE EDITORS advancing at
  `pitch × 1.15` instead of the `pitch` Word uses.
- **Paragraph margin collapsing.** The gap between two paragraphs is now
  `max(prev.spaceAfter, this.spaceBefore)` rather than the sum (CSS-style
  collapsing margins). Matches Word's observed 18 pt gap between
  `after=360` → `before=240` paragraphs.
- **spaceAfter may overflow the bottom margin.** A paragraph fits when
  `y + (h − spaceAfter) ≤ contentH`; the trailing whitespace is suppressed
  at page boundaries. Lets a closing paragraph with a large `after` land
  flush against the bottom margin.
- **Knuth-Plass-style shrink tolerance on wrap-fit.** ECMA-376 doesn't
  prescribe a line-breaking algorithm; we adopt the standard typographic
  policy used by TeX / InDesign / Word — each inter-word space may
  compress by up to 25 % of its natural width when testing fit. Absorbs
  the ~0.1–0.3 px/glyph advance difference between Chromium's canvas
  and Word's internal text layout, so long paragraphs wrap like Word's.
- **Implicit `w:keepNext` on heading paragraphs (w:outlineLvl 0–8).** Word's
  built-in Heading 1–9 styles carry an implicit keepNext even when
  styles.xml omits it; parser now sets `keep_next=true` when a paragraph's
  effective style declares `w:outlineLvl`.
- **Table style `w:pPr` cascade (§17.7.6).** A table's `w:tblStyle` now
  contributes its paragraph formatting to every cell paragraph, resolved
  between docDefault and the paragraph's own style. For the default
  "Table Grid" style (`line=240 auto`, `after=0`), this tightens cell
  line spacing from ~28 pt to ~18 pt per line, matching Word.
- **docGrid per-grid-line computation (§17.6.5).** Parsing
  `w:docGrid/@w:type` and `@w:linePitch` on the section now feeds into
  the line-box formula. Headings authored with oversized `lineRule="auto"`
  values (e.g. `line="1040"` on a 56 pt title) no longer blow up into
  ~300 pt tall lines — they snap to the section's grid pitch times the
  multiplier.
- **Inter-word compression on justified lines.** When canvas measurement
  forces a line slightly over `availW`, the final render compresses
  inter-word spaces (capped at ~¼ of the line's ascent) instead of
  overflowing the right margin.

### Stories / samples

- xlsx viewer: active sheet tab is now visually smaller than inactive
  tabs, which matches the project's layout preference.
- pptx interactive playback: media play / pause badge style unified; the
  story now explicitly opts into `presentSlide` so static rendering and
  playback paths share identical chrome.

### Known limitation

- Word chains `w:keepNext` transitively through "heading cluster"
  paragraphs (kicker label → title) that are not themselves marked with
  `w:keepNext` or `w:outlineLvl`. ECMA-376 §17.3.1.15 defines keepNext as
  1-hop only, so we follow the spec and accept that a kicker paragraph
  can land at the bottom of page N while the matching title heads page
  N+1 in layouts like demo/sample-1 page 3's "FIELD NOTES · CHAPTER THREE".

## 0.5.0 — 2026-04-21

### xlsx

- **Pattern fills (ECMA-376 §18.8.20).** `gray125` and hatch patterns
  (`darkGrid`, `lightGrid`, `darkHorizontal`, `lightHorizontal`, etc.) now
  render as a blend of the cell's `fgColor` and `bgColor`, and the hatch
  varieties draw a repeating `CanvasPattern` tile instead of a flat blended
  color.
- **Gradient fills (§18.8.24).** Parse `<gradientFill>` (linear `degree` +
  path-style bounding box) and render via `createLinearGradient` /
  `createRadialGradient` with multi-stop color interpolation.
- **Comment indicators (§18.7.3).** Commented cells get a small red triangle
  in the top-right corner to mirror Excel's visual cue. Parsed from
  `xl/comments*.xml` and exposed as `Worksheet.commentRefs`.

### docx

- **Page margins respected in pagination.** The paginator's per-paragraph
  height estimator now builds the same `WrapLayoutCtx` as the renderer when
  anchor-image floats are active, so float-aware line wrapping and
  estimation agree. Pages whose content wraps around floats no longer
  overshoot or undershoot the bottom margin.
- **Line-box metrics and vertical centering (§17.3.1.33).** Replace
  per-character `actualBoundingBoxAscent` with font-metric
  `fontBoundingBoxAscent` / `fontBoundingBoxDescent`, so every line in a
  paragraph — and every paragraph that shares a font/size — sits on a
  consistent baseline. `lineBoxHeight` now reads: `auto` = natural × value,
  `exact` = pt × scale, `atLeast` = max(natural, pt × scale). Glyphs are
  centered within the line box (extra spacing split above and below),
  fixing text that previously rendered top-aligned inside wide auto-spaced
  paragraphs.

## 0.4.0 — 2026-04-21

### pptx

- **Audio/video playback.** `PptxPresentation.presentSlide(canvas, index, opts?)`
  returns a disposable `PresentationHandle` that layers the current video
  frame and self-drawn play/pause + seek chrome on top of the statically
  rendered slide. Click on a media element to toggle playback; drag the
  progress handle to scrub. Audio gets a capsule-shaped pill with time on
  the left and a thin seekable bar on the right. `renderSlide` stays pure
  and stateless.
- **Lazy media extraction.** The parse output no longer inlines poster images
  as base64 data URLs; `PptxPresentation.getMedia(path)` fetches bytes on
  demand via a new `extract_media` WASM export. Sample decks with 200 MB
  video now have a <1 KB parse JSON instead of a 16 MB one.
- **Charts, spec-faithful.**
  - Chart space vs. plot area fill distinguished (`<c:chartSpace><c:spPr>`
    separate from `<c:plotArea><c:spPr>`). Transparent outer chart lets the
    slide background show through.
  - Legend visibility driven by `<c:legend>` presence.
  - `<c:crossBetween>` honored (0.5-step category padding for "between").
  - Value-axis line + `majorTickMark` drawn.
  - Title / axis / data-label sizes come from XML `<c:txPr>` `sz` in hpt,
    scaled via `ptToPx = 12700 × slideScale` so fonts track the viewport.
  - Line width and marker radius also scale by pt-per-px.
- **Text rendering.**
  - Line-height maximum is computed from actual run sizes, not placeholder
    `defRPr sz="30000"` prompt markers — fixes 24pt text rendering ~260px
    below its anchor in demo/sample-1 slide 8.
  - Bullet size derives from the first run's font size (ECMA-376
    §21.1.2.4.13) instead of the layout default — fixes em-dash bullets
    overlapping text in demo/sample-1 slide 7.
  - **Text overflow no longer clipped** at shape bounds per §20.1.2.3.6.
- SmartArt connector shapes with `cy=0` no longer inflated by the body-text
  auto-height fallback (horizontal timelines now render horizontal).

### docx

- **Pagination properties (ECMA-376 §17.3.1.14 / .15 / .44).** Parse and
  honor `w:keepNext`, `w:keepLines`, `w:widowControl` through the full
  style cascade (docDefaults → style → pPr). `widowControl` defaults to
  true when absent. `keepNext` now causes a page break before the current
  paragraph when the chain of kept-together paragraphs wouldn't fit.
- **Justified / distribute alignment (§17.18.44).** Inter-word whitespace
  is expanded so `<w:jc w:val="both"/>` paragraphs fill the content width;
  `distribute` also stretches the last line. Previously everything outside
  `right`/`center` collapsed to left-aligned — big visual change for docs
  whose docDefaults declare `both`, which includes most Word-authored
  documents.
- **Line spacing (§17.3.1.33).** `lineSpacingMultiplier` now respects the
  pt value on `atLeast` and `exact` rules (previously both collapsed to
  1.2× font). Decorative titles that encode `lineRule="auto"` with very
  large values (~720+) now render with correct line height.
- **Indentation (§17.3.1.12).** Accept logical `start`/`end` aliases in
  addition to `left`/`right` on `<w:ind>`. `hanging` still wins over
  `firstLine` per spec.
- **rFonts theme references (§17.3.2.26 / §20.1.4.1.14).** theme.xml's
  `<a:fontScheme>` is parsed and rFonts `asciiTheme` / `hAnsiTheme` /
  `eastAsiaTheme` refs are resolved against it at run assembly. Direct
  typeface attributes still take precedence per spec.
- **Default paragraph style.** Fall back to the document's `w:default="1"`
  style ID (e.g. `a`, `標準`) instead of the hardcoded literal `Normal`.
  Matters for contextualSpacing grouping on non-English templates.
- **ST_OnOff (§17.3.2.22).** `bool_prop` now recognises `off` — previously
  interpreted as `true`.
- **Footnote / endnote markers (§17.11.16 / §17.11.7).** Render the
  reference number as a superscript marker inline (previously dropped).
  Full page-bottom footnote layout is deferred.
- **Table cell widths (§17.18.87).** `w:type` now defaults to `dxa` when
  absent; non-dxa types fall back to grid allocation.

### Stories / samples

- Interactive pptx sample story auto-disposes the `PresentationHandle` +
  `PptxPresentation` when its root detaches from the DOM. Storybook
  story-swap no longer leaks playing audio.
- CSS spinner overlay (`createCanvasSpinner`) shows while a sample is
  loading. Both the opinionated `buildViewerUI` and the interactive
  `presentSlide` story use the same helper.

### Known follow-up

- Word's auto-rule rendering for very large multipliers (e.g. `line="640"
  lineRule="auto"` on 28pt headings) still diverges from spec — Word
  Desktop and Word Web themselves disagree here, so we stick to the letter
  of the spec instead of empirical tuning.
- Full bottom-of-page footnote layout.
- Tab alignment variants beyond `pos` (center / right / decimal).
- `cstheme` font axis (only `ascii` / `hAnsi` / `eastAsia` resolve today).

## 0.3.0 — 2026-04-20

DOCX shape rendering (solid / gradient fill, lumMod/lumOff, z-order),
anchor image text wrap (Square + TopAndBottom), default paragraph style
fallback. PPTX placeholder image alpha (`a:alphaModFix`) and master
`txStyles` bold/italic inheritance. Shape helpers extracted to
`@silurus/ooxml-core`.

## 0.2.0 and earlier

See git history.
