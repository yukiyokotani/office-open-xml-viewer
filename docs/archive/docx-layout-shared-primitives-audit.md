# DOCX Layout Shared-Primitives Audit

## Purpose

This audit fixes the ownership boundary for the DOCX layout-engine migration.
It does not introduce a new OOXML interpretation. It records which existing
format-neutral primitives must be reused and which WordprocessingML layout
rules must remain in `packages/docx`.

The normative map uses ECMA-376 Part 1 (5th edition) and Part 4. Microsoft
implementation notes remain compatibility evidence, not a replacement for the
standard. In particular, Part 4 owns transitional VML syntax; it does not make
VML page anchoring or WordprocessingML flow a generic DrawingML concern.

## Decision rule

A primitive belongs in a shared package only when all of the following hold:

1. The OOXML grammar or operation is defined independently of its host format.
2. DOCX, XLSX, and PPTX can consume the same semantic result without adding
   host-specific inheritance, pagination, or coordinate-origin behavior.
3. Sharing removes duplicated parsing or rendering logic rather than merely
   giving differently-behaving concepts the same TypeScript shape.

Host code may locate a shared XML leaf, supply already-resolved defaults, and
map the shared result into its own wire model. It must not duplicate the leaf
grammar. Conversely, a common-looking rectangle is not automatically a shared
layout abstraction when its coordinate space and flow semantics differ.

## Inventory and ownership

| Existing primitive | Normative owner | Current consumers | A1 decision |
|---|---|---|---|
| Drawing fills, strokes, effects, image-fill descriptors, DrawingML text-body data in `core/src/types/common.ts` | Part 1 §§20.1 and 21.1; chart relationships are described in §14.2 and chart markup in §21.2 | DOCX, XLSX, PPTX TypeScript renderers | Reuse unchanged. Layout nodes carry resolved values or stable resource keys; DOCX must not define parallel DrawingML fill/stroke/effect types. |
| `DrawingGroupSpec`, `DrawingGroupTransform`, `DrawingRect`, and non-visual `hidden` parsing in `ooxml-common/src/drawing.rs` | Part 1 §§20.1.2.2.8, 20.1.7.5 and Annex L.4.7; Word drawing host element §20.4.2.5 | DOCX, XLSX, PPTX Rust parsers | Reuse unchanged in the parser. Nested group-transform composition is DrawingML grammar, not DOCX page layout. |
| `SpaceLine`, `parse_lnspc`, `BodyPr`, `parse_body_pr` in `ooxml-common/src/text.rs` | Part 1 §§21.1.2.1.1-.4 and 21.1.2.2.5 | PPTX and XLSX parsers; applicable DrawingML text in DOCX uses the same leaf grammar | Reuse the leaf parser and keep host inheritance/default selection outside it. WordprocessingML paragraph spacing (`w:spacing`) is a different grammar and remains DOCX-local. |
| Color source extraction, transforms, theme resolution contract, and color-space conversion in `ooxml-common/src/color.rs` | Part 1 §§20.1.2.3 and 20.1.6 | DOCX, XLSX, PPTX Rust parsers | Reuse unchanged. Each host supplies its theme/clrMap resolver and the required tint compatibility mode; layout receives resolved color facts and never reruns the cascade. |
| `EMU_PER_PX_96DPI` in `ooxml-common/src/units.rs` and unit constants in core | Part 1 §20.1.10 coordinate simple types; the 96-DPI conversion is an application conversion, not page flow | DOCX, XLSX, PPTX | Reuse conversions at parser/resource or paint boundaries. `DocumentLayout` itself uses points at scale 1 and contains no display scale or DPR. |
| Ref-counted `FontFace` registry in `core/src/fonts/font-registry.ts` | Browser resource lifetime; OOXML font references originate in host markup | DOCX, XLSX, PPTX font loaders | Reuse unchanged. A2's font service may snapshot resolved metrics but must not duplicate global face registration, ref-counting, or release behavior. |
| Local font-metric loading in `core/src/fonts/local-metrics.ts` | Browser font-resource capability; requested metric data is format-neutral | DOCX and other host loaders through core | Reuse behind A2's immutable text-service snapshot. Layout must not retain `FontFace`, `FontFaceSet`, callbacks, or other platform objects. |
| Raster header inspection and pixel-budget checks in `core/src/image/raster-dimensions.ts` | Image container formats and the shared decoder safety policy, not an OOXML host rule | Shared image pipeline used by the three renderers | Reuse behind `ImageMetadataService`. DOCX owns only placement and wrap geometry; it must not add another PNG/JPEG/GIF/WebP/TIFF/BMP/DIB dimension parser. |
| Canonical chart model and `core/src/chart/renderer.ts` | Part 1 §14.2 and §21.2, with DrawingML styling from §§20-21 | DOCX, XLSX, PPTX | Keep the format-neutral chart subsystem in core. DOCX layout stores a stable chart resource key and authored/placed bounds; DOCX paint delegates the atomic chart content to core and does not adapt chart fields or recompute page placement. Chart-internal label layout remains owned and tested by the shared chart subsystem, not by DOCX pagination. |

## New A1 data contracts

`LayoutRect`, `PointPt`, `Matrix2DData`, and `ClipPathData` remain internal to
DOCX for A1. This is intentional rather than duplication:

- `core/src/types/common.ts` models authored cross-format OOXML values, mostly
  in EMUs, and some types intentionally mention Canvas contracts.
- A1 models a retained, structured-clone-safe page result in points at scale 1.
- XLSX uses sheet/cell coordinates and PPTX uses a fixed slide coordinate
  system; neither currently exposes the same flow ownership or page invariant.
- `Matrix2DData` is a serialized affine result for retained paint. It does not
  replace the normative DrawingML group-transform parser in
  `ooxml-common/src/drawing.rs`, nor does it expose `DOMMatrix`.

Moving these types to core now would create a name-level abstraction without a
shared lifecycle or invariant. If XLSX or PPTX later adopt the same immutable
retained-scene contract, the plain-data geometry types can move to core in a
separate API-reviewed change.

## Why WordprocessingML layout stays in DOCX

The following responsibilities are defined by WordprocessingML, not by shared
DrawingML:

- paragraph/run flow and section properties (`w:p`, `w:r`, and `w:sectPr`;
  Part 1 §§17.3 and 17.6);
- table grid, rows, cells, conditional table styles, and pagination
  (`w:tbl`, `w:tblGrid`, `w:tr`, and `w:tc`; Part 1 §17.4);
- headers, footers, footnotes, endnotes, and their section/page selection
  (Part 1 §§17.10 and 17.11);
- columns, page/column breaks, keep/widow behavior, floating-object wrap
  interaction, and the ordered body state machine;
- transitional VML hosted by WordprocessingML (Part 4 §§9, 14, and 19.3),
  after shared VML/DrawingML shape facts have been parsed.

These operations depend on WordprocessingML adjacency, story identity, section
state, and pagination. Sharing them with XLSX or PPTX would either leak DOCX
semantics into the other renderers or hide host-specific branches inside core.
They therefore belong under `packages/docx/src/layout`.

## Layout-to-paint boundary

The parser and shared resource layers may preserve authored or resolved OOXML
facts. DOCX layout turns those facts into point-space geometry, flow ownership,
layer order, clips, transforms, and stable resource keys. DOCX paint consumes
that retained result and may invoke an explicitly shared atomic painter (for
example the core chart painter); it must not call DOCX measurement, pagination,
style resolution, or parser-model helpers.

The boundary also applies in worker mode: `DocumentLayout` contains only plain
structured-clone data. Browser resources and callbacks remain behind main- and
worker-capable service/resource registries and are addressed by stable keys.

## Consequences for later series PRs

- A2 wraps core font and raster facilities; it does not fork them.
- A3 stores chart/image/math keys and placement bounds, while shared content
  renderers stay in core.
- Parser changes extend `ooxml-common` when the XML leaf is genuinely shared;
  WordprocessingML inheritance and flow remain in the DOCX parser/layout.
- A proposed new shared type must identify its common normative owner and at
  least two format consumers. Similar field names alone are insufficient.
- Paint importing a DOCX layout algorithm, or layout importing a Canvas/display
  contract, is an architecture violation enforced by the A1 static gates.
