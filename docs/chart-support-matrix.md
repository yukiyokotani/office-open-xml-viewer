# Chart support matrix

This document is a summarized implementation view for DrawingML charts. The
specification-derived source of truth is
[#1276](https://github.com/yukiyokotani/office-open-xml-viewer/issues/1276).
A chart family can be generally available while individual authored properties
remain partial; the issue therefore tracks the finer-grained audit and work
items that are intentionally collapsed here.

## Inventory method

The audit starts from the Strict and Transitional `dml-chart.xsd` contracts in
ECMA-376, then follows every referenced chart type and group into the shared
parser, wire model, and renderer. DrawingML paint/text contracts, MS-ODRAWXML
ChartEx and linked style parts, extension-list contracts, and host package
relationships are audited as separate workstreams. Office-produced output is
used only to establish a bounded compatibility rule where the specification
leaves behavior application-defined.

The matrix is not considered exhaustive until every specification workstream
in #1276 is classified. New implementation findings must be recorded in that
issue before this summary is promoted to `Supported`.

## Status and completion criteria

| Status | Meaning |
| --- | --- |
| Supported | The parser preserves the authored property, the shared model exposes it, the shared renderer consumes it, and focused tests cover the supported boundary. |
| Partial | A documented subset is implemented. The Notes column states the exact boundary. |
| Missing | Valid authored markup is discarded or retained without a renderer. |
| Unverified | The implementation exists, but its Office compatibility boundary has not been established. |
| Not applicable | The property does not affect browser rendering or is intentionally outside the product scope. |

A row becomes **Supported** only when all of the following are present:

1. parser-to-model contract coverage in `packages/ooxml-common`;
2. shared rendering coverage in `packages/core` where the concept is common to
   DOCX, XLSX, and PPTX;
3. a focused geometry/style regression test;
4. an Office-produced fidelity comparison for behavior left application-defined
   by ECMA-376 or MS-ODRAWXML.

Specification-defined behavior does not require a compatibility heuristic.
Application-defined behavior must stay **Partial** or **Unverified** until its
observed input boundary is recorded in
[`chart-compatibility-evidence.md`](chart-compatibility-evidence.md).

## Classic chart families

| ID | Family / property | Parser | Model | Renderer | Status | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| C-LINE-001 | `lineChart` standard, stacked, percent-stacked lines | Yes | Yes | Yes | Supported | Includes markers, smoothing, blank-cell policy, data labels, error bars, and trendlines. |
| C-LINE-002 | `lineChart/dropLines` | Yes | Yes | Yes | Supported | One owning-group envelope per category spans the effective category-axis crossing and all plotted group points. Interior crossings and paint order are Office-verified. |
| C-LINE-003 | `lineChart/hiLowLines` | Yes | Yes | Yes | Supported | Valid on ordinary line charts as well as stock charts. |
| C-LINE-004 | `lineChart/upDownBars` | Yes | Yes | Yes | Partial | Direct paint is supported. Empty-paint automatic white/black styling is limited to the retained legacy Style 2 observation. |
| C-LINE-005 | Multiple `lineChart` groups in one plot area | Yes | Yes | Partial | Partial | Decoration ownership is retained and consumed; other group-level line properties still require individual provenance rows. |
| C-LINE-006 | Group-level `marker` and `smooth` defaults | Yes | Yes | Yes | Supported | Group defaults are retained; an explicit series-level marker or smooth value remains authoritative. |
| C-AREA-001 | `areaChart` standard, stacked, percent-stacked areas | Yes | Yes | Yes | Supported | Series fill, labels, axes, and stacking are shared across hosts. |
| C-AREA-002 | `areaChart/dropLines` | Yes | Yes | Yes | Supported | `EG_AreaChartShared` ownership and direct line paint are preserved. One envelope line per category spans the axis crossing and all standard or cumulative stacked points, with Office-verified paint order. |
| C-BAR-001 | `barChart` direction, grouping, overlap, and gap | Yes | Yes | Yes | Supported | Each bar group retains its own provenance and geometry. |
| C-BAR-002 | `barChart/serLines` | Yes | Yes | Yes | Supported | All group-owned line paints are preserved. A single authored element joins adjacent points at facing value-end edges for both directions. Excel rejects multiple authored elements, so the renderer preserves them but fails closed instead of inventing a style-association rule. |
| C-STOCK-001 | Stock drop lines, high-low lines, and up/down bars | Yes | Yes | Yes | Partial | Stock drop/high-low lines preserve direct and linked line paint. Up/down bars use the shared solid/gradient/pattern fill model and preserve full preset outline formatting. Empty-paint Office defaults are theme-aware for the observed omitted/legacy Style 1, 2, and 10 boundary; other legacy styles remain unresolved rather than guessed. |
| C-SCATTER-001 | Scatter style, X/Y values, markers, and smoothing | Yes | Yes | Yes | Supported | Numeric axes and the six `scatterStyle` modes are represented. |
| C-RADAR-001 | Standard, marker, and filled radar styles | Yes | Yes | Yes | Supported | Direct series paint and marker controls are consumed. |
| C-PIE-001 | Pie/doughnut point explosion | Yes | Yes | Yes | Supported | Per-point explosion is retained. |
| C-PIE-002 | Pie/doughnut series-level explosion | Yes | Yes | Yes | Supported | `CT_PieSer/explosion` supplies the default; a point-level `dPt/explosion` overrides it. |
| C-PIE-003 | First-slice angle and doughnut hole size | Yes | Yes | Yes | Supported | Authored schema bounds are preserved. |
| C-OFPIE-001 | Pie-of-pie/bar-of-pie split, sizing, and connector geometry | Yes | Yes | Yes | Partial | Position, value, percent, and custom splits follow ECMA-376. An omitted `splitType` uses the bounded Office rule in MS-OE376 §2.1.1596(b); the Office-prohibited explicit `auto` value fails closed. Exact connector and gap geometry remains under compatibility audit. |
| C-BUBBLE-001 | Bubble size, scale, negative bubbles, and size representation | Yes | Yes | Yes | Supported | Resource-bounded and shared across hosts. |
| C-BUBBLE-002 | `bubble3D` | Yes | Yes | Yes | Partial | Group, series, and point provenance is retained. Current Excel paints every point from the series value, falling back to the owning group and then false; point-level values do not alter the visible material. Three bounded neutral material layers overlay automatic/solid/gradient/pattern/picture fills without changing authored alpha, preserve no-fill and outline independence, apply to series legend/data-label keys, and participate in the chart-wide paint budget. Negative bubbles follow MS-OE376's unconditional inversion rule; the automatic 3-D negative style is white material with a black outline. The implementation intentionally does not reuse Surface camera or lighting constants. |
| C-SURFACE-001 | Surface/surface3D mesh, bands, camera, and authored band formatting | Yes | Yes | Yes | Partial | Both surface families retain their distinct type and render through the bounded projected source-grid mesh. Filled bands consume direct/linked structured paint. A wireframe uses the first Surface series outline as its mesh default, falls back to a fixed-index linked `dataPointWireframe` line, and otherwise uses the automatic band colours. Direct `bandFmt` outlines split the source-grid edges at value-band boundaries and override that default only inside their band; first-series and band no-line remain authoritative. Width, preset/custom dash, cap, and join merge property-wise. Later-series outlines do not become mesh defaults. Relative linked palettes and compound wireframe lines remain fail-closed. Authored floor/wall thickness uses the shared projected `CT_Surface` slab geometry. Planar and positive-thickness floor/wall picture fills consume the observed stretch, projected-reference plain `stack`, bounded face-local DrawingML tile grids, and bounded `stackScale` semantics. On thick plain-stack slabs, front and lateral side faces repeat while end faces map one complete source. Picture faces independently honor `applyToFront`, `applyToSides`, and `applyToEnd`. Stretch consumes bounded signed `srcRect` source rectangles and `fillRect` destination rectangles with a positive visible intersection. Tiling is limited to the observed `pictureFormat=stretch` boundary and consumes explicit physical DPI, scale, alignment, offset, schema-defaulted or explicit flip, and a bounded signed `srcRect` inside every tile before projectively mapping the completed grid to each face. A tile/stretch-invalid `fillRect` combination remains fail-closed. Floor ignores the stack unit. Office-observed category/value major and minor gridline rules continue over their corresponding planar or positive-thickness camera-visible surface faces while retaining authored line style; horizontal Bar3D swaps the surface pairs with its axes. Material-dependent slab shading remains unpainted. |
| C-3D-001 | Classic bar/column 3-D shapes and camera projection | Yes | Yes | Yes | Partial | Box, cylinder, cone, and pyramid geometry is bounded; compatibility evidence limits camera/material approximations. |
| C-3D-002 | Classic line, area, and pie 3-D projection | Yes | Yes | Yes | Partial | The authored 3-D group is retained and dispatched through the shared renderer. Pie `hPercent` scales thickness per MS-OE376; structured direct/linked series paint is consumed without introducing unmeasured material constants. Office-observed direct line dPt formatting owns only the incoming segment. Excel retained but did not paint direct 3-D area point formatting, so area bodies remain series-painted without invented face segmentation. |
| C-COMBO-001 | Multiple classic chart families and primary/secondary axes | Yes | Yes | Yes | Partial | Every classic plot group is retained in source order with its exact family, contiguous series range, group-local settings, and authored axis IDs. Axis identifiers remain bounded opaque 32-bit keys, including the observed signed-decimal view. Axis ownership is resolved only when the referenced pair is unambiguous; same-position duplicates remain unresolved. Observed bar/line or bar/area, scatter/bubble, stock-then-line, and the observed shared-primary mixed bar-direction pair render without sharing stack or percentage state between groups. Distinct secondary category axes on line/area groups, mixed-direction secondary axes, mixed 2-D/3-D, and unobserved family combinations fail closed instead of being coerced into the legacy top-level family. Arbitrary family layering remains Partial because OOXML does not define it. |

## Shared axes, labels, legends, and chart-space properties

| ID | Property | Parser | Model | Renderer | Status | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| C-AXIS-001 | Linear/date/log axes, authored bounds and units | Yes | Yes | Yes | Partial | Explicit fractional month/year units in the retained `1 <= value < 4` compatibility boundary are rendered. Other fractional and application-defined automatic date intervals remain unsupported. |
| C-AXIS-002 | Category `lblAlgn` and `lblOffset` | Yes | Yes | Yes | Supported | Category/date and auxiliary category axes retain both properties; the shared renderer applies interval alignment and the specified percentage of its family-specific default label gap. |
| C-AXIS-003 | Axis crossing (`crosses`, `crossesAt`, and `crossBetween`) | Yes | Yes | Partial | Partial | Bar/column, line, area, and Surface use one effective crossing for axis geometry, ticks, labels, and group decorations. Remaining numeric-X and arbitrary combination boundaries are still under audit. |
| C-LABEL-001 | Value/category/series/percent labels, separators, leader lines, and manual layout | Yes | Yes | Yes | Supported | Per-point and series-level overrides are retained. |
| C-LABEL-002 | `showBubbleSize` | Yes | Yes | Yes | Supported | Series and point-level flags compose the authored bubble-size cache value into the label. |
| C-LABEL-003 | `showLegendKey` | Yes | Yes | Yes | Supported | Series- and point-level flags use the effective series/point key paint in bounded, rich, callout, and manually positioned labels. |
| C-LABEL-004 | `showDLblsOverMax` | Yes | Yes | Yes | Supported | Labels beyond the effective primary or secondary value-axis maximum are suppressed consistently, including stacked endpoints and classic 3-D Cartesian charts. |
| C-LEGEND-001 | Position, text, fill, line, and manual layout | Yes | Yes | Yes | Supported | |
| C-LEGEND-002 | `legend/overlay` and per-entry delete/style | Yes | Yes | Yes | Supported | Overlay legends do not reserve plot space; source-indexed deletion and text overrides share one measured legend model in 2-D and 3-D paths. |
| C-SPACE-001 | `roundedCorners` | Yes | Yes | Yes | Supported | The authored flag applies one bounded Office-compatible outer geometry to fill, clipping, and border; structured chart-space fill remains authoritative. |
| C-SPACE-002 | `plotVisOnly` hidden-source behavior | Yes | Yes | Yes | Supported | Worksheet row/column visibility is projected over caches before shared layout, stacking, labels, legends, and extent planning. |
| C-TABLE-001 | Plot-area data table content and borders | Yes | Yes | Yes | Supported | One measured layout serves column, horizontal bar, line, area, stock, and combination charts. Scatter's authored table is ignored according to the retained Office boundary. |

## Chart Style roles

The shared parser retains the paint-bearing `CT_ChartStyle` roles defined by
MS-ODRAWXML §2.8.3.1 for both classic and ChartEx parts. The three hosts resolve
the same linked `styleN.xml` and `colorsN.xml` relationships; the remaining
status differences below are renderer-consumption gaps, not lost package data.

| ID | Role group | Status | Notes |
| --- | --- | --- | --- |
| S-STYLE-001 | Shared role parsing and package wiring | Supported | Paint recipes, fixed/relative Chart Colors indices, `NoStyle`, and bounded palette expansion are retained through DOCX, XLSX, and PPTX. |
| S-STYLE-002 | `chartArea`, `plotArea`, `legend` | Partial | Direct paint, `noFill`, and `noLine` remain authoritative. Linked solid, gradient, and pattern fill/outline recipes are consumed through shared frame painters across classic 2-D, optional 3-D, ChartEx, and offline Region Map paths, including host rotation and partial manual-layout dimensions. Preset and bounded custom dash, width, cap, join, and the Office-observed compound rails are supported. Explicit pen alignment and miter-limit geometry remain unpainted. |
| S-STYLE-003 | `categoryAxis`, `valueAxis`, tick labels, `seriesAxis` | Supported | Linked line and text defaults are consumed for primary, secondary, and 3-D series axes. Direct chart formatting remains authoritative. |
| S-STYLE-004 | `seriesLine`, `dropLine`, `hiLoLine`, `upBar`, `downBar` | Supported | Direct formatting wins; the linked role supplies only omitted fill/line properties. Decoration lines retain visibility, color, width, preset dash, cap, and join. |
| S-STYLE-005 | `errorBar`, `leaderLine` | Supported | Direct line paint and `noFill` win; the linked role supplies omitted color, width, dash, and visibility in shared 2-D and optional 3-D paths. |
| S-STYLE-006 | `trendline` | Supported | Direct trendline paint and `noFill` win; the linked role supplies omitted color, width, dash, and visibility to the plot and legend. |
| S-STYLE-007 | `dataTable` | Supported | The linked line role supplies omitted grid color, width, dash, and `noFill`; direct table formatting remains authoritative. |
| S-STYLE-008 | `gridlineMajor`, `gridlineMinor` | Supported | Linked line fallback applies to enabled primary and secondary gridlines; direct paint and `noFill` retain precedence. |
| S-STYLE-009 | `dataLabelCallout`, `trendlineLabel` | Partial | Direct point, series, and trendline formatting remains authoritative; linked roles fill omitted text, body, and shape properties in classic 2-D labels, pie/doughnut callouts, the implemented ChartEx label paths, and optional 3-D data labels. Rich runs, paragraph alignment, bounded rotation/wrapping, signed insets, anchors, manual layout, structured shape paint, and text/shape no-fill are retained on those paths. Optional 3-D trendline labels, per-glyph `eaVert`, `just`/`dist` expansion, language-specific shaping, and rich-run wrapping remain unsupported. |
| S-STYLE-010 | `dataPoint3D`, `dataPointWireframe`, `floor`, `wall`, `plotArea3D` | Partial | `dataPoint3D` direct/linked solid, gradient, and pattern paint is consumed by classic 3-D bar, pie, line, and area paths with direct no-fill precedence and one paint resolution per datum, series, or observed incoming line segment. A Surface wireframe consumes the first series outline as the mesh default, then a fixed-index linked `dataPointWireframe` line, then automatic band colours; direct `bandFmt` outlines and no-line split and override that default per value band. Width, preset/custom dash, cap, and join merge property-wise. Later-series outlines are not promoted to mesh defaults; relative linked palettes and compound lines fail closed. Structured floor/wall and `plotArea3D` paint is consumed with direct surface formatting authoritative. Authored thickness uses the shared bounded `CT_Surface` slab projection; planar and positive-thickness floor/wall images consume stretch, projected-reference plain `stack`, bounded face-local DrawingML tile grids, and bounded `stackScale` face mappings. Thick plain-stack front/side faces repeat while end faces map one complete source. Stretch consumes bounded signed source/destination rectangles; each tile consumes its own bounded signed source rectangle. Authored category/value major and minor gridline styles continue across the measured planar and camera-visible positive-thickness slab faces, including the horizontal Bar3D axis swap. Tile/stretch-invalid destination rectangles, direct 3-D area point body paint, and material-dependent slab shading remain fail-closed or unpainted. |
| S-STYLE-011 | `dataPointMarker` and marker layout | Partial | Direct and linked solid, gradient, pattern, and relationship-backed picture fills share the DrawingML paint model across implemented classic marker-bearing series and ChartEx box markers. Picture markers retain raster/SVG twins, source cropping, stretch, chained `alphaModFix`, and the supported two-color duotone effect. Tiling is drawn only when alignment, positive scale, positive authored DPI, and rotation semantics are explicit; omitted embedded-DPI defaults, unsupported blip effects, and unresolved relationships fail closed rather than inventing an automatic marker color. |

The S-STYLE-002 compound ratios and their deliberately excluded boundaries are
recorded in [Chart compatibility evidence and scope](chart-compatibility-evidence.md).

## ChartEx layouts

The ChartEx family renderer is an optional main/worker capability imported from
`@silurus/ooxml/chart-ex`. Classic DrawingML 2-D chart families remain in each
format entry. Without the optional renderer, recognized ChartEx models retain
their parsed data but paint the deterministic unsupported-chart placeholder.

MS-ODRAWXML §2.24.4.19 defines exactly eight `ST_SeriesLayout` values:
`boxWhisker`, `clusteredColumn`, `funnel`, `paretoLine`, `regionMap`,
`sunburst`, `treemap`, and `waterfall`. Histogram and Pareto are semantic
forms of `clusteredColumn`/`paretoLine`; no additional layout identifier is
inferred from cached values. A future identifier outside the current schema is
retained verbatim and reaches the explicit unsupported-chart placeholder.

| ID | Layout | Status | Notes |
| --- | --- | --- |
| X-LAYOUT-001 | Waterfall, clustered column (including histogram), Pareto/`paretoLine`, funnel | Supported | Includes bounded layout data and shared Chart Style paint. |
| X-LAYOUT-002 | Box-and-whisker | Supported | Includes mean/outlier/non-outlier roles and visibility controls. |
| X-LAYOUT-003 | Treemap and sunburst | Supported | Hierarchy depth/slot budgets apply before tree construction. |
| X-LAYOUT-004 | Region Map | Partial | Deterministic offline country geometry only; external geocoding is out of scope. |
| X-LAYOUT-005 | Unknown or future `layoutId` values | Supported | Preserved verbatim and failed closed with a placeholder; no layout is guessed from cached values or a visually similar known family. |

## Maintenance rules

- Update #1276 and this summary in the same pull request that changes support
  status.
- Do not mark a row Supported from a single screenshot or sample-specific
  adjustment.
- Keep self-VRT and Office-fidelity validation separate: self-VRT detects
  regressions, while Office exports adjudicate compatibility.
- Add a new row when valid authored markup is intentionally deferred. Silently
  discarding a newly discovered rendering property is not an acceptable steady
  state.
- Private workbooks and Office exports remain local. Public tests must use small,
  synthetic fixtures that isolate the relevant OOXML contract.
