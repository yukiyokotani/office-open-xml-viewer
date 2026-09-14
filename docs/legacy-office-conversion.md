# Opt-in legacy Office conversion

The byte-conversion API can normalize legacy binary Office bytes before an
existing OOXML parser runs:

- `.doc` to macro-free `.docx`
- `.xls` to macro-free `.xlsx`
- `.ppt` to macro-free `.pptx`

The opt-in `@silurus/ooxml/legacy-conversion` entry contains both a purpose-built
local WASM converter and the implementation-neutral adapter API. Ordinary DOCX,
XLSX, and PPTX entry points do not import, fetch, initialize, or retain the
converter Worker or its WASM. If neither a converter nor the matching direct
source is supplied, legacy input continues to reject with
`OoxmlError.code === 'legacy-binary-format'`.

Experimental direct sources are available for DOC, PPT and XLS. They read
supported binary subsets into the existing document, presentation or workbook
models without generating an OOXML ZIP or reparsing generated XML. These paths
use the ordinary layout and Canvas renderers.

## Experimental direct DOC source (browser)

Import the separate `@silurus/ooxml/legacy-doc` entry and enable it only for
`doc`. The factory returns a validated descriptor without fetching or
initializing WASM. The browser document loader opens the dedicated
`legacy_doc_direct_bg.wasm` asset only when a DOC input selects this source.

```typescript
import { DocxDocument } from '@silurus/ooxml/docx';
import { createLegacyDocSource } from '@silurus/ooxml/legacy-doc';

const document = await DocxDocument.load(legacyDocArrayBuffer, {
  legacyConversion: {
    doc: { source: createLegacyDocSource() },
  },
});
const canvas = window.document.querySelector('canvas') as HTMLCanvasElement;

try {
  await document.renderPage(canvas, 0, { width: 960 });
} finally {
  document.destroy();
}
```

Both browser rendering modes and progressive layout use the same retained
layout pipeline. A custom asset pipeline can supply an absolute `wasmUrl` to
`createLegacyDocSource`. The native source remains alive for image reads until
the document is destroyed; finishing a model cursor does not dispose it.
An optional `legacyConversion.doc.signal` can cancel loading and remains
attached to the loaded document until destruction.

This is a narrow experimental reader, not full DOC support. Unsupported
formatting, fields, numbering, notes and other unimplemented structures may
reject the entire document. The byte-converter support matrix below does not
describe native-reader coverage. No fallback to byte conversion occurs after
a native failure. ZIP resource metrics and Markdown export are unsupported for
this source. Node document APIs do not yet support the direct DOC source;
continue using byte conversion there. Neither route executes macros.

## Experimental direct PPT source

Import `createLegacyPptSource` from the separate
`@silurus/ooxml/legacy-ppt` entry and enable it only for `ppt`. Importing the
entry makes the dedicated `legacy_ppt_direct_bg.wasm` asset available, but the
factory only returns a validated source descriptor: it does not fetch or
initialize WASM. The presentation loader initializes that asset only when a
legacy PPT input selects the source.

Browser:

```typescript
import { PptxPresentation } from '@silurus/ooxml/pptx';
import { createLegacyPptSource } from '@silurus/ooxml/legacy-ppt';

const presentation = await PptxPresentation.load(legacyPptArrayBuffer, {
  legacyConversion: {
    ppt: { source: createLegacyPptSource() },
  },
});
const canvas = document.querySelector('canvas') as HTMLCanvasElement;

try {
  await presentation.renderSlide(canvas, 0, { width: 960 });
} finally {
  presentation.destroy();
}
```

Node:

```typescript
import { openPptxPresentation } from '@silurus/ooxml/node';
import { createLegacyPptSource } from '@silurus/ooxml/legacy-ppt';

const session = await openPptxPresentation(legacyPptBytes, {
  legacyConversion: {
    ppt: { source: createLegacyPptSource() },
  },
});

try {
  for await (const slide of session.slides()) {
    // Consume each ordinary shared slide model.
  }
} finally {
  await session.close();
}
```

Applications with a custom asset pipeline can override the emitted asset URL:

```typescript
const source = createLegacyPptSource({
  wasmUrl: new URL('/assets/legacy_ppt_direct_bg.wasm', location.href).href,
});
```

The direct source is an experimental, bounded subset, not a full-fidelity
PowerPoint implementation. Its current omissions include audio/video media,
embedded fonts, Markdown production, and ZIP-based resource-usage metrics.
The support matrix below describes the byte-conversion engine, not the direct
reader's admission contract. The direct reader is currently narrower and can
reject otherwise convertible constructs, including unresolved paragraph
margin or indentation and positive paragraph before/after percentages. Its
explicit unsupported diagnostics are authoritative; conversion support does
not imply direct-reader support.

## Experimental direct XLS source

Like DOC and PPT, XLS is a per-format opt-in. XLS applications can import
`createLegacyXlsSource` from `@silurus/ooxml/legacy-xls` and select it only in
`legacyConversion.xls`:

```typescript
import { XlsxWorkbook } from '@silurus/ooxml/xlsx';
import { createLegacyXlsSource } from '@silurus/ooxml/legacy-xls';

const workbook = await XlsxWorkbook.load(legacyXlsArrayBuffer, {
  legacyConversion: {
    xls: { source: createLegacyXlsSource() },
  },
});
```

The direct XLS source is likewise experimental and bounded. It preserves its
supported BIFF workbook subset in the shared worksheet model; unsupported
features still reject or remain omitted as documented by diagnostics. Its
dedicated asset is `legacy_xls_direct_bg.wasm`, and the factory accepts the same
kind of optional absolute `wasmUrl` override as the PPT factory.

Existing converter options and `convert()` behavior are unchanged, so no
migration is required unless an application chooses a native source.
Each `doc`, `ppt` or `xls` configuration selects
either `source` or `converter`, never both.
Importing the ordinary DOCX, XLSX, or PPTX entries alone does not select this
functionality or fetch a dedicated WASM asset.

## Built-in browser converter

Use one shared converter instance so all viewers share its bounded queue. Each
active conversion receives a new Worker; the Worker and converter WASM memory
are released before the converted package enters the existing parser Worker.

```typescript
import { DocxViewer } from '@silurus/ooxml/docx';
import {
  createLegacyOfficeWasmWorkerConverter,
} from '@silurus/ooxml/legacy-conversion';

const legacyConverter = createLegacyOfficeWasmWorkerConverter({
  maxConcurrency: 1,
  maxQueuedConversions: 4,
});

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new DocxViewer(canvas, {
  legacyConversion: {
    doc: {
      converter: legacyConverter,
      timeoutMs: 120_000,
    },
  },
});
await viewer.load(legacyDocBytes);
```

`XlsxViewer` and `PptxViewer` use the matching `xls` and `ppt` fields. Each
field is an independent opt-in: configuring `doc` does not enable legacy input
for either other viewer. Importing the opt-in entry emits a separate
`legacy_office_converter_bg.wasm` asset. Applications must serve that asset with
the other package assets; it is not fetched or initialized until an enabled
legacy input actually reaches the converter.

For Node, use `createLegacyOfficeWasmConverter()`. Its default loader reads the
emitted WASM asset locally; `wasm` can be supplied explicitly as bytes or a
compiled module when an application has its own asset pipeline. Direct browser
use is also possible, but conversion is synchronous after WASM initialization
and should therefore remain inside a Worker.

## Initial support matrix

This first engine version is suitable for feasibility testing and text/value
ingestion experiments. It is not a general-fidelity replacement for opening a
legacy document in Microsoft Office.

| Input | Accepted subset | Preserved | Deliberately omitted / rejected |
|---|---|---|---|
| DOC | CFB Word 97-2003 documents with a readable main-story CLX piece table | main-story text, paragraphs, tabs, custom tab stops and document-wide default tab interval, line/page/column breaks, displayed field results, font names and explicit sizes, paragraph-style character defaults, character styles and direct bold/italic/underline/strike/caps/color/spacing properties, paragraph alignment/indentation/line spacing/before-after spacing/keep options, ordinary single-level and multilevel list definitions, list starts/restarts and marker formatting, literal list-bullet glyph references with resolved marker fonts, nested table structure, explicit cell widths/margins/borders/merges and row heights, section boundaries, page size/orientation, explicit body margins and gutter, columns, vertical alignment, document grid, inline JPEG/PNG/EMF/WMF picture frames with display size, cropping, rotation and flips, explicitly positioned main-story floating JPEG/PNG/EMF/WMF frames with basic wrapping, formatted header/footer variants and supported passive page-number fields, paragraph borders, formatted footnote/endnote content and references | frames, unsupported paragraph/list-style interactions and conditional table styles, legacy automatic-numbering fields and unrepresentable list templates, advanced table/character/section properties, header/footer and note floating drawings, note numbering/positioning/custom separators and custom-mark rendering, advanced floating drawings, non-raster images other than EMF/WMF, picture borders/effects and nonrectangular geometry, revisions, OLE; non-Western compressed code-page pieces are not decoded yet |
| XLS | CFB BIFF8 workbooks, including shared-string character data split across `CONTINUE` records | worksheet names and visibility (including very hidden), scalar values, cached formula results, merged ranges, date system, BIFF8 number formats, fonts, palette colors and supported checksum-bound extension colors, fills, borders, alignment, shared-string rich-text runs, styled blank cells, row heights and column widths, row/column hiding and outlines, print setup/margins/options, basic header/footer commands, manual page breaks, and measured passive embedded PNG/JPEG/EMF/WMF picture frames with supported cell anchors, cropping, rotation and flips | formula programs, phonetic string data, unsupported extended styles/theme colors/gradients, conditional formatting, print areas/titles, extended headers/footers, saved custom views, charts, non-picture drawings, grouped or active/linked picture objects, picture effects, external links, pre-BIFF8 sheets |
| PPT | CFB PowerPoint 97-2003 files with a resolvable current edit chain and persist directory | live slide order and dimensions, UTF-16/compressed Unicode text and outline references, individual shape anchors, nested group coordinates, basic rotation/flips, direct text margins/wrapping/vertical anchoring, direct font names/sizes/bold/italic/underline, literal and slide/master-scheme colors, paragraph alignment/spacing and explicit local ruler custom tabs, character bullets, explicit shape-local automatic numbering and paragraph-style offsets, verified-placeholder and explicit master-shape text-style inheritance, manual line breaks, unmodified basic presets with direct or explicitly linked master solid fill/line colors, supported classic linear gradients, line widths and opacity, line caps/joins, arrow ends and standard dash patterns, embedded/delayed JPEG, PNG, EMF and WMF picture frames with signed cropping, local/inherited solid, linear-gradient and stretched-image backgrounds, eligible foreground picture fills on supported preset and uniform custom paths, enabled non-placeholder master objects using the same supported drawing subset, static slide-number metacharacters, explicit full-coordinate line/cubic paths with uniform path paint; superseded slides, deleted and explicitly hidden shapes are not emitted | unlinked placeholder and nonuniform master text overrides, unlinked/drawing-default paint, master placeholder content and header/footer fields, system/palette color indices, inherited/outline automatic numbering, picture bullets, text-ruler offsets/default intervals and inherited ruler tabs, advanced character formatting, embedded fonts, guide-dependent or compact custom geometry, arc/editing escapes, mixed per-path paint, some rotated/grouped geometry, rotated gradient shapes or gradients inside rotated/reflected groups, non-linear gradients, patterns, custom dash arrays/compound lines, effects, custom fill rectangles and origins, charts, notes, PICT/DIB/TIFF/other image formats, picture effects and advanced foreground image-fill sizing, audio/video, transitions, animations, actions, OLE |

XLS literal, untinted RGBA colors in checksum-matched XF extensions are preserved
for text, pattern fills, and all five border edges. Cell-specific font colors do
not change other cells sharing the original font. Missing or stale XF checksums
retain the base palette formatting. Untinted accent and hyperlink theme colors
are resolved through the embedded theme package's internal relationships and named color slots,
including saved system-color `lastClr` values. The theme is read only when an
owned extension needs it; no theme package, links, or active content are copied
to the output. The first four light/dark theme indices retain their BIFF palette
fallback: the documented index ordering conflicts with observed Office output,
and no compatibility remapping is inferred. Version-only default themes,
unsupported theme color forms or transforms, tinted extension colors, gradient fills, and font-scheme extensions
still use the base-format fallback; extended-style warnings remain. No migration
is required.

Embedded XLS theme parsing is a bounded metadata subset, not a general OPC
validator: it accepts UTF-8 XML and internal, unescaped part names; rejects
ambiguous packages, external theme relationships, DTDs, and malformed XML; and
caps ZIP input at 4 MiB, entries at 64, each expanded part at 256 KiB, and declared
aggregate expansion at 2 MiB. XML depth, events, attributes, and retained strings
are also bounded. These are converter resource policies, not Office format
limits. No host system-color lookup or default-theme guess is performed.

XLS gridline visibility, zero-value display, and right-to-left sheet direction
are preserved through ordinary OOXML sheet views. Row/column header visibility
is retained as metadata but is not yet applied by the viewer. Multiple window
associations are retained; the converter does not reconstruct window placement,
pane selections, scrolling, or zoom. No migration is required; XLS remains a
separate opt-in.

XLS worksheet visibility is preserved as OOXML metadata. Display follows the
existing `XlsxViewer` `hiddenSheetMode` option; its default remains `'show'`.
Hidden sheets and their cell data are retained, not removed.

PPT slide visibility is also preserved as OOXML metadata. The slide's own
`SlideShowSlideInfoAtom.fHidden` becomes `p:sld/@show="0"`; hidden slides and
their content remain in the presentation. Display follows the existing
`PptxViewer` `hiddenSlideMode` option, whose default remains `'show'`.
Master visibility is not inherited. Transitions, sounds, and actions remain
omitted. No migration or additional opt-in is required for this metadata fix;
PPT conversion still requires its existing per-format opt-in.

PPT local text-body ruler custom tabs are preserved as explicit DrawingML
paragraph tab lists, including signed positions, all four alignment values, and
an explicitly empty list. This applies to owned inline text and outline text
references, including master objects that are themselves rendered. No migration
is required; PPT conversion retains its existing independent opt-in.

The converter reads the local `TextRulerAtom` (MS-PPT 2.9.23-24, 2.9.29-30)
without inventing a paragraph-margin adjustment. It charges both decoding and
each paragraph's tab emission against the existing work and XML budgets. The
ordinary PPTX parser and renderer consume the output. Ruler margin/indent fields,
ruler default intervals, document default rulers, linked-master ruler inheritance,
and conflicting direct paragraph tab arrays remain unsupported.
Malformed local tab records fail rather than producing a partially decoded list.
Multiple local ruler records are rejected as unsupported ambiguity; the inline
record grammar permits them, but precedence is not inferred by this subset.
This is not a claim of full binary/Office visual equality.

OfficeArt `metroBlob` alternative shape XML is currently ignored. A modern
Office-saved PPT can retain paragraph properties there rather than in its
classic text ruler; see the [controlled probe protocol](../scripts/legacy-ppt-ruler-probes.md)
before attributing those differences to an implicit ruler rule.

Local Office-reference checks confirm that restoring these stops improves
tab-separated text, but residual RTL anchoring and ruler-indent differences
remain. Fidelity checks must load the intended fonts: fallback metrics can cause
extra wrapping even when the tab position matches Office. These checks are
separate from byte-exact previous-converter/unchanged-renderer comparisons.

PPT paragraph text direction is retained as DrawingML `a:pPr/@rtl`
(MS-PPT 2.9.20/21 and 2.13.30; ECMA-376 21.1.2.2.7). Supported master and
direct formatting paths inherit an absent direction, while explicit left-to-right
clears an inherited right-to-left value. Alignment remains independent and
logical Unicode text order is unchanged. Reserved direction values are rejected.
The ordinary PPTX parser and renderer handle the resulting metadata; no
renderer changes, migration or additional opt-in are required. This does not
claim complete Office-equivalent bidirectional shaping or punctuation placement.

PPT text-body direction preserves an owned primary `txflTextFlow=1` as
DrawingML `bodyPr/@vert="eaVert"` when `cdirFont` is absent or zero. This is a
PowerPoint compatibility mapping, not a literal interpretation of the generic
MS-ODRAW enumeration. Controlled local Office down-save/reopen/PDF/roundtrip
tests found that both OOXML `vert` and `eaVert` become the same binary value,
and that Office reopens that value with East Asian vertical behavior. The
converter targets the binary presentation, not recovery of a source OOXML
distinction lost by Office during down-save.

The controls used Latin, CJK, mixed text and punctuation with two font families;
additional controls varied unrotated frames, positive/negative 45- and 90-degree
rotation, 180-degree rotation, and horizontal/vertical/both flips. The mapping
does not branch on characters, fonts, sample names, or rotation angles. A group
control lost its geometry during Office down-save and is not evidence of group
visual fidelity. Combined transforms, other producers and arbitrary grouped
layouts are not certified by these controls. Other text-flow values, nonzero
font direction, tertiary direction properties and inherited direction remain
unsupported; the existing advanced-text warning still applies. Invalid primary
direction enums, scalar flags and duplicate direction properties are rejected.
Geometry and the ordinary OOXML renderer are unchanged. Restoring body direction
does not implement Office's automatic upright digit grouping or resolve existing
font and geometry differences. No migration is required.

XLS extended indentation (`XFExt` property `0x000F`, MS-XLS 2.5.108) is preserved
as ordinary SpreadsheetML `alignment/@indent`, including values above the base
four-bit limit through 250. Extensions require the existing matching XF checksum
and owning-XF checks. Cell and style XFs retain their own indentation and reading
order; style XF fields are not treated as reserved (MS-XLS 2.5.249). Invalid
extension sizes/ranges and duplicate properties fail through the existing error
contract. This is a direct specification mapping, not an Office-layout heuristic;
the existing XLSX parser and renderer consume it without a binary-specific path.
No migration or opt-in change is required.

Version-3 and version-4 CFB containers are admitted. Password-protected legacy
binaries and pre-CFB Office formats are rejected. These limits are structural,
not filename-based. Unsupported binary structures fail with
`reason === 'unsupported-input'`. Accepted documents can still lose the features
listed above: their warning identifiers are not a fidelity certificate.

DOC inline/floating images, measured XLS picture frames, and PPT
picture/background images can retain passive
EMF and WMF BLIPs (MS-ODRAW 2.2.24-25/31). Both UID layouts and uncompressed or RFC 1950
zlib-compressed data are supported. Validated metafile bytes become ordinary image
parts without rasterization, geometry rewriting or execution of metafile
commands by the converter. Existing OOXML image handling renders the supported
EMF/WMF drawing subset; this is not full metafile or Office visual parity. PICT
and unsupported binary drawing containers remain omitted. XLS support is limited
to the eligible passive picture subset described below; it does not reconstruct
charts, arbitrary drawing shapes, grouped pictures or active/linked objects.
The shared OOXML image renderer supports retained line, polygon, rectangle and
cubic-Bezier paths, including fill, stroke, stroke-and-fill, abort and saved-DC
path state (MS-EMF 2.3.10, 2.3.5.9, 2.3.5.38-39 and 3.1.1.2.4). This can restore
outline-based content in retained EMFs without a binary-format renderer.
Glyph-to-path, ellipse/arc path construction, flattening and widening remain
unsupported; affected paths are omitted rather than painted as fragments.
Path clipping retains intersection-only support. Other clip combination modes
and full GDI pen/brush semantics are not implemented. Preserving EMF bytes is
still not proof of complete visible output or original-binary layout fidelity.
Restored outlines also do not establish color or opacity fidelity. Visual
evaluation must compare every changed page with the previous renderer and an
Office reference, and record remaining differences separately from restored
content. A reference exported from OOXML does not certify binary-input layout.

As renderer resource policy (not format limits), one path retains at most
65,536 commands; one EMF playback allocates at most 262,144 path commands and
replays at most 1,048,576 stored commands. Saved DCs share immutable geometry
and these budgets. Malformed or over-budget path geometry is discarded; a
replay-budget rejection issues no partial path drawing. The existing image
failure/omission behavior and cache ownership remain unchanged.

Metafile extraction checks the declared compressed/expanded lengths, stream end,
header and record envelope. As resource policy, each metafile is limited to 32 MiB
stored and expanded, with the existing 128 MiB per-media-store retention cap
checked before decompression. DOC inline and floating stores have separate
caps. Repeated image references share one retained buffer; cache lifetime ends
with conversion. Existing work and generated-package limits still apply.

WMF admission accepts standard and placeable headers and validates the bounded
record envelope through its terminal record (MS-WMF 2.3 and 2.3.2.1-3). It
preserves the original drawing bytes, including unknown drawing operations;
it does not evaluate them. This is structural validation, not a claim that all
WMF operations are supported by the existing Canvas image player. Text-heavy
metafiles can retain their image dimensions while their text remains unpainted
by that player; preserving the image is not evidence of equation-text fidelity.
When a size-declared WMF contains data after a valid terminal record, the entire
image remains unsupported and is omitted through the existing media-omission
path. The converter neither removes nor forwards those trailing bytes. This is
an admission policy, not a rule that Office padding is valid or safely ignorable.
Malformed records or over-budget expansion fail the existing media validation.
DOC and PPT propagate those errors as unsupported input; XLS retains its
existing picture-set omission and warning behavior. No public API, per-format
opt-in or application migration change
is required. DOC, XLS and PPT conversion remain independently opt-in.

PPT master object inheritance follows `SlideFlags.fMasterObjects` independently
of color-scheme and background inheritance (MS-PPT 2.5.10-11). Live main/title
master chains are resolved through the persist directory and emitted below
slide-local objects in ordinary PresentationML shape order (ECMA-376 19.3.1.45).
The destination slide's resolved color scheme applies to inherited objects.
Master placeholder exemplars are not copied as visible content; header/footer
field synthesis remains unsupported. Explicit hidden flags are respected and
script anchors are omitted before following text or image references.
One writer shares IDs, image relationships and work/XML budgets across layers.
Borrowed master chains are cached per conversion with cycle and depth checks;
expanded output is still charged for every destination slide. Missing local
drawings retain the warned unpositioned-text fallback without duplicating IDs.
No renderer, worker protocol or per-format opt-in migration is required.
Geometry support remains partial: an omitted foreground shape can expose a
master object that Office would cover. Object inheritance alone is not a
guarantee of visual fidelity.

PPT vector shapes can also carry explicit custom paths: full 32-bit coordinate
pairs, straight lines, cubic Bezier curves, moves, closes and path ends are
converted into ordinary DrawingML custom geometry (MS-ODRAW 2.2.51/53-55,
2.3.6.1-9, 2.4.9/30-31; ECMA-376 20.1.9). Geometry-space origins and reversed
axes are normalized algebraically; explicitly linked master geometry inherits
individual properties without copying source arrays. Path-level no-fill/no-line
flags remain separate from shape paint. Point and segment expansion consumes
the shared work budget, and generated path XML consumes the output budget.
Compact coordinate encodings, guide-dependent points, arc/editing escapes,
mixed per-path paint and picture-frame clipping geometry remain unsupported.
Rotated shapes and nonuniformly transformed groups can still differ in aspect
ratio, orientation or placement from Office; explicit path support does not
resolve those existing transform limitations. Effects remain unsupported.
DOC/XLS drawing reconstruction is not enabled by this PPT integration; the
OfficeArt decoder is shared so later format-specific wiring need not duplicate it.

PPT classic linear gradients retain ordered shade colors, signed focus and the
16.16 angle through the compatibility and direct-model projections. Explicit
shade-array stops at either endpoint take precedence over scalar front or back
colors; a missing final endpoint uses the scalar back color. The current subset
requires a nonempty shade array beginning at position zero and linear RGB
interpolation. Admission excludes rotated leaves, rotated or reflected ancestor
groups, nonopaque fills,
`fillUseRect`, non-shape fill modes and other gradient types. Leaf flips and
unrotated scaled groups remain supported, and `rotateFillWithShape` is
preserved. Twenty controlled Office comparisons covered focus, angle, endpoint
conflicts and leaf flips. Native rational positions were within one DrawingML
position unit of Office's serialized integers. Compatibility gradient XML
matched 18 cases semantically; two focus cases differed by one serialized
position unit. These
checks define the tested subset and do not claim complete gradient fidelity.

Plain eligible `msofillPicture` foreground fills resolve through the converter's
existing validated passive media store and become ordinary DrawingML image fills.
Supported presets and uniform supported custom paths clip the image through the
same OOXML shape geometry used for solid paint. A text-bearing source remains one
ordinary OOXML `ShapeElement`, retaining its image fill, outline and text rather
than being converted into a picture with a separate text overlay. This uses the
existing PPT conversion opt-in; no migration or additional option is required.

This is a bounded mapping, not full picture-fill fidelity. Custom binary fill
rectangles, fill origins and other unsupported sizing controls are not
reconstructed. Image-frame placement for rotated or flipped shapes that
explicitly disable rotation with the shape remains a fidelity limit. These
limits are separate from the supported geometry clip and passive media lookup,
and no inferred transform or sample-specific sizing is applied.

Supported PPT solid outlines preserve flat, round and square caps; bevel, round
and miter joins; and triangular, stealth, diamond, oval and open-arrow ends with
independent widths and lengths (MS-ODRAW 2.3.8.15/20-27, 2.4.16-20;
ECMA-376 20.1.8.38/43/57 and CT_LineProperties). Explicitly linked master
properties inherit independently, including explicit no-arrow overrides.
The binary defaults are flat caps and round joins; these are emitted explicitly
instead of relying on renderer defaults. Arrow editability does not suppress
authored ends. Unrepresentable miter limits are rejected rather than clamped.
These properties use the ordinary PPTX parser and renderer, with no opt-in API
change or migration. They do not restore unsupported connector geometry or
guarantee Office-identical arrow sizing and shaft trimming.

All ten standard OfficeArt dash/dot patterns map to their DrawingML preset
counterparts (MS-ODRAW 2.3.8.17/2.4.15; ECMA-376 20.1.8.48/20.1.10.49).
An explicit solid style clears an inherited preset; absent styles still inherit.
Dash patterns do not suppress the line, its cap or its arrow ends. Invalid
preset enums are rejected. Custom `lineDashStyle` arrays are not reconstructed.
The ordinary OOXML renderer's existing preset-cadence approximations remain,
so this mapping does not promise Office-identical dash spacing. No migration
or renderer change is required.

Unadjusted straight connectors (`msosptStraightConnector1`, MS-ODRAW 2.4.24)
also retain their static DrawingML `straightConnector1` path, including zero
width/height, line styling and arrow ends. They do not acquire a fill. Conversion
preserves the saved geometry; it does not recreate editable endpoint bindings
or run a routing algorithm. Bent/curved connector presets and adjusted geometry
remain outside this preset mapping, and the existing rotation/group-transform
limitations still apply to connector placement.

Slide-number metacharacters in positioned text, including inherited ordinary
master objects and outline-referenced text, become static decimal text using
the document's starting number and live slide order (MS-PPT 2.4.2, 2.9.47).
Only declared character positions are replaced; literal asterisks remain text.
Original UTF-16 style boundaries are retained even for multi-digit numbers.
This does not synthesize missing master placeholders, evaluate arbitrary fields,
or add dynamic numbering to the generated presentation.

PPT paragraph default tab intervals are retained through direct and supported
master text-style inheritance (MS-PPT 2.9.20/2.2.29). Signed master-unit values
become ordinary DrawingML `defTabSz` coordinates (ECMA-376 21.1.2.2.7), including
explicit zero instead of accidentally inheriting another interval. The existing
OOXML parser/renderer uses positive intervals; nonpositive values remain in the
package but currently use the viewer's fallback interval. Custom tab-stop lists
and TextRuler properties remain unsupported. No migration, legacy-specific
renderer change, or opt-in API change is required.

XLS shared-string formatting uses `FormatRun` UTF-16 character offsets and
`FontIndex` references, including the reserved index-4 gap and ignored terminal
run (MS-XLS 2.5.129, 2.5.132 and 2.5.293). Run fonts become ordinary
SpreadsheetML `r/rPr/rFont` properties (ECMA-376 18.4.4-7); an unformatted prefix
retains the cell font, while explicit normal formatting resets bold/italic and
other run properties. Continued character fragments are joined before decoding
UTF-16, so surrogate pairs spanning record boundaries stay intact. Phonetic
extensions are skipped, not promoted into visible text. Invalid live font
references, unordered/out-of-range run starts and surrogate-splitting boundaries
reject the input rather than attaching formatting to the wrong characters.
The converter shares immutable encoded string fragments between cells and
caches run font properties for one workbook, encoding each entry before reading
the next. Resource policies cap SST entries and total format runs at one million
each, retained encoded strings at 256 MiB and aggregate
worksheet XML at 256 MiB, independently of the compressed output limit.
Retained run properties still depend on existing OOXML parser/renderer support;
automatic font-color resets and advanced font effects do not have verified
visual parity. This does not change the per-format opt-in API.

DOC character properties follow physical FKP ranges through the logical CLX
piece table, including UTF-16 positions and displayed-field gaps. Supported
style properties are resolved into ordinary OOXML run properties; fonts are
referenced by name, not embedded or downloaded. Missing formatting tables use
explicitly warned defaults. Style depth, formatting pages/runs and property
application work have converter resource limits; these are implementation
policies, not limits of the Office file format. The generated DOC main XML part
also has a 256 MiB resource ceiling, separate from the output ZIP byte limit.
The main story is limited to 64 Mi UTF-16 units before decoding repeated pieces
and one million control characters before constructing paragraphs/tokens.
Supported paragraph properties resolve through styles, direct PAPX and piece
properties, including bounded references into the binary Data stream. Fixed,
minimum and proportional line spacing retain their original units. Table rows
use the definitions on their terminating marks; nested cells remain nested and
row marks do not become visible paragraphs. Shared grids retain explicit edges,
including zero-width cells, and horizontal merges become ordinary OOXML spans.
Table style inheritance, preferred percentage widths, shading, text rotation,
floating/frame placement and protection-bookmark table separation are incomplete.
Unknown optional border-side flags are omitted with a warning, not reinterpreted.
Nesting (32), rows per section (100,000) and grid boundaries (65,536) have resource
ceilings. Paragraph text and pending tables remain bounded by the XML budget.
Floating-frame support remains partial; line wrapping and pagination can differ
significantly. No existing renderer changes are required.

DOC list tables and list-format overrides are emitted as ordinary
WordprocessingML numbering. Supported output includes single-level and
multilevel lists, supported literal bullets and level templates, number formats,
suffixes and justification, starts and restart boundaries, and marker-only
character formatting. Lists with the same binary LSID share a counter even when
paragraphs use different LFOs; a start override applies once for its original
LFO and level. The main story, each header/footer story and each note use
independent numbering scopes. Legal-number formatting is retained without
collapsing a zero-padded decimal format to ordinary decimal.

List paragraph and marker properties are resolved without replaying document
defaults as direct formatting. Marker CHPX can vary without creating a new
counter sequence. Explicit direct paragraph bidi, common alignment values and
absolute twip indentation (including zero) are retained after list formatting;
style-only, default and relative-indent values are not promoted by this
compatibility rule. This precedence is based on bounded Word-produced controls,
not the literal list-last ordering described by MS-DOC. The controls covered
ordinary common alignments and both paragraph directions, but did not establish
universal behavior for piece-level direct formatting or exotic alignment enum
values.

The list decoding and counter mapping follow MS-DOC 2.4.6.3-4 and 2.9.150,
and ECMA-376 17.9.10 and 17.9.26. The direct-format precedence exception above
is an observed Office behavior, not an amendment to those specifications.

Local checkpoints exercised 36 numbered paragraphs across nine list cases, 352
paragraphs for indentation and bidi compatibility, and 18 common-alignment
cases. These measurements bound the compatibility claim; they are not a general
Office conformance certificate.

The converter emits OOXML consumed by the existing DOCX parser and renderer; no
legacy-specific renderer path was added. This support does not guarantee exact
Office display, pagination, marker placement or bidirectional word order.
Literal UTF-16 bullet glyph references and their resolved marker fonts are
preserved without converting code points or branching on font names or font
charsets. This compatibility behavior is based on 250 used body-list references
across 16 Office-produced pairs: one companion DOCX was produced from an
original DOC, while the other pairs began as OOXML and were saved as DOC.
Direct DOC display confirmed representative Symbol, Wingdings, Arial,
and Japanese-font markers; two less common marker shapes remain unadjudicated.
Font availability and font-axis selection in the unchanged renderer can still
affect display. Unsupported automatic-number fields and list templates that
cannot be expressed safely report
`legacy-doc:unsupported-numbering-text-or-autonum-omitted`.

DOC tables retain the compatibility cell-shading arrays (`Shd80` and `Shd`),
explicit cell ranges (including alternating cells), and table-wide shading
(MS-DOC 2.6.3, 2.9.52-53, 2.9.247-249, 2.9.308). Foreground and background
colors, automatic colors, no-shading sentinels and the 38 documented OOXML
pattern mappings become ordinary `w:shd` properties (ECMA-376 17.3.5,
17.4.30-32, 17.18.78). Row-level exceptions preserve a shared table grid;
omitted trailing entries in modern shading arrays clear stale segment values.
Unmappable binary patterns remain warned and omitted,
not approximated by a percentage tint. Array/range work consumes the formatting
budget, and output consumes the existing XML budget.

This converter does not interpret conditional table styles, so it uses the
legacy compatibility shading specified for readers without table-style support.
The separate `ShdRaw` style-inheritance arrays remain unsupported; in particular,
their `ShdNil` is not treated as an explicit clear override of the compatibility
array. These limitations remain visible in conversion warnings. The existing
DOCX viewer currently renders background fills but does not reproduce every
shading pattern or automatic-color/inheritance case. Preserving that metadata
does not claim pattern-level visual parity. No renderer change, migration, or
additional opt-in is required.

DOC table-level floating positions now become ordinary `w:tblpPr` and
`w:tblOverlap` properties (MS-DOC 2.4.3, 2.6.3, 2.7.13, 2.9.208/351/357;
ECMA-376 17.4.57). The converter preserves page/margin/column/text anchors,
symbolic alignment, signed coordinates, physical text clearances and explicit
overlap prevention. Encoded absolute distances are decremented by one as
specified; reserved alignment values are mapped separately. Non-positioned
anchor codes remain inline. No renderer-specific offset or wrapping correction
is applied. Paragraph-frame-derived table positioning remains unsupported, and
the existing DOCX renderer's floating-table layout limitations still apply.
No migration or per-format opt-in change is required.

DOC header/footer stories follow MS-DOC 2.3.3 and 2.8.22: the six separator
stories are not page headers, and each section has even/default/first header
and footer slots. Zero-length ranges inherit the previous section's matching
variant; an explicit blank paragraph creates an empty part instead. Guard marks
are removed, while paragraph/table formatting and supported inline pictures
use the same physical piece/FKP resolution as the main story. Image relationships
are scoped to their containing part. Document-facing-page and section-title-page
flags become ordinary OOXML settings (ECMA-376 17.10); no legacy-specific
renderer path is added.

Unnested, unlocked PAGE and NUMPAGES fields with supported general formatting
switches retain their dynamic meaning in headers/footers. The field table's
lock flag keeps cached text, and private field results are suppressed.
Other field instructions are discarded while their cached display is retained;
they are not evaluated and cannot open links, files, macros or external services.
Section page-number formats, continuation and explicit restarts become ordinary
`w:pgNumType` (MS-DOC 2.6.4; MS-OSHARED 2.2.1.3; ECMA-376 17.6.12). A stored
start is ignored unless restart is enabled; an enabled restart without an
explicit start retains the binary format's default zero. Both unsigned 16-bit
and 32-bit starts are supported. Formats reset independently per section; the
non-counting bullet format uses the decimal fallback allowed by MS-DOC.
`none` suppresses the number. Language-dependent/unsupported number formats
retain their OOXML token, but the shared renderer may still display decimal.
The shared field-number formatter bounds each expanded ordinal to 4,096 UTF-16
units before allocating repeated glyphs. Exceeding this resource budget fails
rendering instead of changing the format; large decimal starts remain supported.
This does not restore active main-story fields: their cached display remains.
Header floating drawings, advanced field switches, chapter-number prefixes,
and exact Office pagination remain incomplete.
Each aggregate main/header story has a 64 Mi UTF-16-unit decoding ceiling and
one million controls; headers additionally allow at most 4,096 nonempty parts.
The aggregate generated XML has a 256 MiB ceiling. These are resource policies,
not format limits. No migration or opt-in API change is required.

Footnote and endnote text now retains its paragraphs, character formatting,
supported tables and inline pictures in ordinary `footnotes.xml`/`endnotes.xml`
parts. The converter joins the main-story reference PLC with the corresponding
note-text PLC (MS-DOC 2.3.2/5, 2.8.16/17/19/20), preserving UTF-16 positions
across the main, footnote, header and comment documents. Automatic reference
characters require the special-character property. An empty reference PLC is
distinct from a malformed or missing text range. Fields retain cached text;
their instructions are not emitted or executed. The aggregate decoded note
text has a 64 Mi UTF-16-unit budget and each kind allows at most 65,536 notes,
in addition to the existing XML/structure budgets. These are resource policies.

Note numbering formats/restarts/offsets, positioning, custom separators and
floating drawings remain incomplete. Literal custom marks are retained with
the standard `customMarkFollows` attribute, but the existing OOXML viewer does
not yet honor that attribute when numbering/painting references; an extra
number can appear. No legacy-only renderer is added to compensate for this
shared limitation. Page-bottom notes now use the page's terminal continuous
section region, avoiding placement inside an earlier region's body text.
The shared DOCX paginator now keeps at most 128 recent paragraph acquisition
candidates, releasing older measurement copies that previously caused heap
exhaustion in long note-bearing documents. This is a cache-retention policy,
not an OOXML limit or an overall heap-byte guarantee: retained document geometry
and other resources still require substantial memory for large documents.
Preserving note content does not imply exact Office pagination or successful
rendering for every input. No migration or opt-in API change is required.

Paragraph borders retain top, bottom, logical left/right and between edges from
both Brc80 and Brc operands (MS-DOC 2.6.2, 2.9.16/17/21). The converter resolves
style inheritance and direct/piece overrides per edge, then projects logical
sides after the final paragraph direction is known. Width, spacing, color and
applicable shadow/frame flags become ordinary `w:pBdr` (ECMA-376 17.3.1.24).
This also restores paragraph rules in supported headers, footers and table cells.
Explicit `none` clears an edge. As input recovery, the converter also recognizes
the documented NilBrc/Brc80MayBeNil no-border sentinels in paragraph operands,
where Office can store them, and preserves them as `nil`; this is distinct from
the ordinary Brc value constraints. Adjacency/group painting remains owned by
the existing OOXML renderer. Binary PGP grouping metadata, paragraph shading,
frames and exact Office border-effect appearance remain incomplete.

Custom paragraph tabs resolve `sprmPChgTabsPapx` and `sprmPChgTabs` through the
same style/PAPX/PRM cascade (MS-DOC 2.9.179-183). Deletions remove inherited stops
within the specified range, including the normative 25-twip minimum tolerance;
`XAS_plusOne` deletion distances are decoded before use. The resulting sorted
stops preserve signed positions, alignment and leaders as ordinary `w:tabs`
(ECMA-376 17.3.1.37-38). Binary heavy leaders mean underscores, not OOXML heavy
lines; bar-tab leaders and unused descriptor bits are ignored as specified.
Variable-length edits consume the formatting-work budget, and the resolved set
has a 256-stop resource cap independent of the per-record 64-entry format limit.
The document-wide default interval is read from `DopBase.dxaTab` (MS-DOC
2.7.2) and written into a related `word/settings.xml` part as `w:defaultTabStop`
(ECMA-376 17.15.1.25). The ordinary DOCX parser and layout retain precedence of
custom paragraph stops over automatic stops. Missing document properties use
the OOXML default with a warning as an explicit recovery policy; a present but
truncated DOP or zero interval is rejected, not silently assigned new spacing.
Only the shared DOP prefix is interpreted; this does not claim preservation of
other document settings or version-specific compatibility flags. Preserving tabs
or list metadata does not imply Word pagination equivalence.

Section text flow preserves the basic top-to-bottom, right-to-left-column mode
(`sprmSTextFlow` / `msotxflTtoBA`, MS-DOC 2.6.4 and MS-ODRAW 2.4.5) as ordinary
`w:sectPr/w:textDirection w:val="tbRl"` (ECMA-376 17.6.20, 17.18.93 and Part 4
14.11.7). Each section resolves its own properties; an explicit horizontal reset
does not retain a previous vertical direction. Existing DOCX layout and Canvas
painting handle the orientation, with no binary-only renderer path. Other
rotation variants and version-dependent column-direction modes remain omitted
under the advanced-section-property warning; unknown enumeration values reject.
This does not yet preserve frame/cell text directions, all drawings, every list
marker form, all East Asian character formatting, or exact Word line wrapping.

DOC inline pictures follow `sprmCFSpec` and `sprmCPicLocation` through the same
style/CHPX/CLX cascade as other character properties. Only passive picture-frame
JPEG/PNG BLIPs are retained; binary-data and OLE markers are not dereferenced.
`PICMID` supplies the scaled display extent (MS-DOC 2.9.190-193). Inline BLIPs
are matched by property encounter order, not their ignored index or complex flag
(MS-ODRAW 2.2.15). Cropping and transforms become ordinary DrawingML pictures;
the existing DOCX parser and renderer remain unchanged. Unsupported inline
pictures emit a loss warning. Restoring image extents can change line heights
and pagination; this is not a claim of complete Word layout fidelity.

DOC floating JPEG/PNG picture frames use main-story `PlcfSpa` anchors and
`OfficeArtClientAnchor` indices (MS-DOC 2.8.27, 2.9.168, 2.9.253). The drawing
store's delayed BLIPs refer to `WordDocument`, not the inline picture `Data`
stream (2.9.171). Explicit signed positions relative to the page, margin,
column or paragraph, rectangular extents, cropping, flips, top/bottom and
square wrapping, front/behind placement, and overlap/anchor settings become
ordinary DrawingML anchors (ECMA-376 20.4.2.3). No DOC-specific layout or paint
path is introduced. Header drawings and nested groups are not reassigned to
the body. Rotated or alignment-based floating positions, tight/through wrap
contours, non-picture shapes and non-raster media remain omitted with a loss
warning. Alignment-based positions require further reconciliation of producer
values with the published OfficeArt position-origin enumeration; the converter
does not guess that mapping. Restored floating pictures can alter wrapping;
preserving an anchor does not establish Word-compatible pagination.

The inline and floating picture caches are document-owned and each limited to
100,000 source locations/anchors, one million record/property/marker operations
and 128 MiB of retained media. Floating occurrences also have a 100,000 limit.
Raster dimension validation is shared with PPT. Repeated references reuse the
same borrowed image bytes and package part, with unique drawing occurrence IDs.
These are resource policies, not binary-format limits. No source filename or
external picture URL is followed or copied into the output package.

Every output package is created from scratch and contains no source macro,
VBA/Excel 4.0 program, ActiveX control, OLE object, hyperlink action, or external
relationship. The converter never executes formulas, field programs, actions,
links, or macros; passive slide-number substitution is described above.
Fixed, content-free warning identifiers report the intentional loss
class in the conversion provenance record.

PPT picture frames resolve their one-based BLIP references through the current
document's image store (MS-PPT 2.1.3/2.4.3 and MS-ODRAW 2.2.20–32). Only referenced
JPEG/PNG bytes are packaged; repeated references share a part, and each slide
gets only its own internal image relationships. Unsupported encodings, including
payloads whose signature disagrees with the BLIP type, are omitted rather than
relabelled or fetched elsewhere. Malformed supported headers/ranges fail closed.
The converter checks image headers without decoding pixels: supported JPEG
frames use 8-bit Huffman baseline, sequential or progressive encoding. Resource
policy caps each image at 32,768 pixels per side and 40 million pixels total,
and retained media parts at 128 MiB independently of the output ZIP limit.
These are implementation limits, not Office format limits. Normal OOXML image
decoding and rendering still apply; the header checks are not full codec validation.
Signed crop fractions (MS-ODRAW 2.3.23) become ordinary DrawingML `a:srcRect`,
with existing picture/group transforms preserving positions, rotations and flips.
No additional opt-in or renderer-specific legacy path is introduced.

PPT backgrounds follow `SlideFlags.fMasterBackground` independently of scheme
and foreground-object inheritance (MS-PPT 2.5.10–11). Current main/title masters
are resolved with cycle/depth/work checks and a per-conversion cache. The
ungrouped live OfficeArt background shape supplies fill properties; it is not
rendered as a foreground rectangle. Supported solid colors, opacity and picture
fills become PresentationML `p:bgPr` before the shape tree. Master scheme-color
references resolve against the destination slide's active scheme. Image bytes
and relationships use the same bounded store as picture frames. Gradient,
pattern, texture and custom-rectangle background fills remain unsupported.
Background fidelity alone does not imply complete foreground-text fidelity.
Explicit `fHaveMaster` / `hspMaster` links (MS-ODRAW 2.2.40 / 2.3.2.1) now
resolve against live master shapes, independently of placeholder-position metadata.
Uniform character/paragraph formatting at each master indent level overrides the
containing master's text-type defaults; direct slide-run properties still win,
including explicit black text and false bold/italic values. Referenced master
chains are checked for missing IDs, cycles and excessive depth. Only immutable
resolved levels and paint remain after parsing; master-shape metadata has a 100,000-node
resource cap. Exemplar text itself, actions and links are never copied.
Nonuniform exemplar formatting at a level is omitted with a warning rather than
selecting an arbitrary run. Unsupported font indices are also omitted, leaving
normal font fallback without guessing a replacement index. Unlinked placeholder
formatting and other unsupported text properties remain fidelity limitations.
No contrast-based recoloring or sample-specific background suppression is applied.

The same explicit master-shape links also supply solid fill/line properties,
including color, opacity and line width. Local properties override inherited
values independently; Boolean use bits preserve explicit no-fill/no-line and
geometry paint vetoes. Scheme colors resolve only at the destination slide.
Unsupported inherited fill types and dashed lines remain omitted rather than
being replaced by solid defaults. Inherited custom geometry still suppresses
unadjusted preset reconstruction. Unlinked masters, drawing defaults, master
foreground objects and advanced paint remain unsupported. This does not add
legacy-specific behavior to any OOXML parser or renderer.

Character bullets preserve `TextPFException` flags and values independently
(MS-PPT 2.9.20-22): direct no-bullet and follow-text flags override inherited
choices without discarding still-inherited glyph, font, size or color values.
Valid UTF-16 BMP glyphs become DrawingML `buChar`; color resolves through the
destination slide's scheme. Bullet size preserves percentages or absolute
points (MS-PPT 2.2.3). Unsupported glyphs suppress the marker, while unsupported
optional size, font and color values are omitted with a warning. No substitute
glyph, guessed font index or size clamping is applied.

Explicit shape-local automatic numbering uses `PP9ShapeBinaryTagExtension` and
`StyleTextProp9Atom` (MS-PPT 2.7.18, 2.9.26-27, 2.9.67-68). The converter follows
the owning shape's passive `___PPT9` tag and binds entries to consecutive
character-run groups using the specified `pp9rt` modulo-16 matching rule.
An enabled bullet with an explicit numbering flag, scheme and valid start number
becomes DrawingML `buAutoNum`; all 41 numbering scheme identities and starts
1-32767 are retained. Numbering does not replace the text or bypass the ordinary
PPTX parser/renderer. Bullet color, size and typeface remain independent.
No migration is required; PPT conversion remains separately opt-in.

This subset requires a uniform explicit choice across the whole paragraph,
including its terminator. Missing flags/schemes, conflicting paragraph choices,
picture bullets, outline PP9 bindings and PP9 master/default inheritance remain
unsupported, preserving the available base character bullet. No default scheme,
restart or picture-vs-number precedence is guessed. Retaining a valid scheme is
not a guarantee that every script/font is faithfully rendered. Existing offset
limitations can also leave multi-digit numbers too close to or overlapping text;
the converter does not silently expand margins to hide that fidelity gap.

Paragraph-style text and bullet offsets are converted from master units into
DrawingML `marL` and relative `indent`, after per-level inheritance. Negative or
out-of-range left margins and offsets outside the DrawingML schema bounds are
omitted, not clamped. A first-line offset without a resolved text offset is also
omitted. `TextRulerAtom` overrides and tab stops remain unsupported. The ordinary
PPTX renderer preserves signed non-bullet first-line indents consistently in
measurement, wrapping and painting: a negative indent extends the first line
left of `marL`, while continuation lines keep that margin. This is general
DrawingML support, not a binary-specific rendering exception. Font metrics,
tabs and unsupported geometry can still cause visible differences from Office;
preserving offsets does not guarantee layout equivalence.

## Measured XLS picture conversion

No migration is required. DOC, XLS and PPT remain separate format opt-ins, and
the converter still produces ordinary macro-free OOXML for the existing parsers
and renderers. Passive XLS pictures additionally need the workbook's actual
Normal-font metrics. Enable them with `measureXlsNormalFont` on either
`createLegacyOfficeWasmConverter` or `createLegacyOfficeWasmWorkerConverter`:

```typescript
const converter = createLegacyOfficeWasmWorkerConverter({
  measureXlsNormalFont: async (font, signal) => {
    // Application-owned: load this font from your trusted font collection,
    // measure digits 0–9 at font.sizePoints / 72 * 96 CSS pixels (including
    // font.bold / font.italic), and round the maximum measured advance.
    // Return undefined if the intended font is unavailable. Do not substitute
    // a guessed width or treat an untrusted font.family as a resource URL.
    return measureInstalledNormalFont(font, signal);
  },
});
```

The callback returns an integer width from 1 to 4096 pixels, or `undefined`.
The upper limit is resource policy, not an Office font-layout rule. Load the
same font used by the viewer before measuring. Callback rejection fails the
conversion; unavailable metrics omit pictures with an explicit warning.
Direct converters retain at most one prepared XLS model while measuring;
another measured XLS request reports `capacity-exceeded`. The worker adapter
keeps its existing bounded queue and concurrency settings. Cancellation frees
the prepared model or terminates its worker, aborts the host measurement signal,
and discards late replies. Apply a conversion timeout if a font loader can hang.

The worker sends only a bounded font descriptor to the main thread and receives
a numeric width or failure. Functions and WASM pointers never cross realms.
CFB/BIFF parsing is performed once; the owned cell/style/picture data survives
measurement without retaining the parser's source slices. DOC and PPT never
invoke this XLS hook, and omitting the hook preserves the previous output path.

The current subset emits owned, explicitly sized, embedded PNG/JPEG/validated
EMF/WMF picture frames outside nested groups. It preserves cell-relative anchor
offsets, movement/resize behavior, crop and local flip/rotation attributes in
ordinary SpreadsheetDrawing XML. Coordinate conversion uses MS-XLS 2.5.193 and
ECMA-376 18.3.1.13/81, not a fixed assumed digit width. Normal font selection
follows XF zero and its FontIndex (MS-XLS 2.2.6.1.2.2/2.5.129).

Nested group transforms, active/linked objects, unsupported font variants,
picture effects and unresolved geometry remain best-effort omissions. A rejected
optional drawing/media stage drops pictures with a warning while preserving
otherwise valid cells; it never repairs or copies rejected payloads. The shared
image validators are unchanged. Geometry prefix construction has a cumulative
two-million-operation budget, in addition to the existing drawing/media limits.
Output ZIP bytes are bounded and repeated image references reuse one media part.
This is not a claim of complete Excel display fidelity: the ordinary XLSX
renderer also has its own DrawingML capability limits. In particular, the
current XLSX image model does not expose the saved flip/rotation attributes,
and its fixed-size handling currently specializes `editAs="oneCell"`; retaining
those properties in OOXML does not establish their full display fidelity yet.

## XLS drawing inspection for development

Separately, a native-only inspection helper can
extract the supported passive PNG, JPEG, EMF and WMF entries from a BIFF8 global
image store without requiring font metrics or generating an OOXML package:

```sh
cargo run -p legacy-office-converter --features inspection \
  --example inspect_xls_images -- sample.xls fresh-output-directory
```

Omit the output directory to print catalog indices, formats and byte counts
without saving images. The optional directory must not already exist. Extracted
images can contain private document content and must remain local. A catalog
entry is not proof that an image is displayed on a worksheet.

The helper follows MS-XLS 2.4.58/171 and the documented first-continuation
exception in 2.1.7.20.3 implementation note 6. It bounds the stream to workbook
globals and requires one MS-ODRAW 2.2.12/20/22 drawing-group/image-store owner.
Shared passive-image validation remains unchanged; malformed supported images
fail inspection rather than being silently repaired. Unsupported encodings,
unreferenced slots and unresolved delayed entries are not exposed. No external
resource, macro, OLE object or drawing action is evaluated.

Inspection caps the source at 256 MiB, assembled drawing data and retained media
at 128 MiB each, and the drawing/decoding walk at two million records. Shared
image limits also apply. These are independent resource ceilings, not a combined
process-memory guarantee: source, workbook, assembled data and extracted images
can coexist. The helper is excluded from production WASM even when its Cargo
feature is enabled. It does not change the converter contract or renderer.

The companion `inspect_xls_anchors` native example prints raw worksheet anchor
metadata without extracting images:

```sh
cargo run -p legacy-office-converter --features inspection \
  --example inspect_xls_anchors -- sample.xls
```

It uses BoundSheet tab order and disjoint worksheet substreams, excludes nested
chart streams, and joins only owned drawing fragments. MS-XLS 2.5.194/195 client
markers must end at the exact fragment boundary immediately preceding the
matching Obj/TxO record. The inspector retains the shape identity, FtCmo object
identity/type/flags, enclosing group depth and signed MS-XLS 2.5.193 endpoint
fractions. It does not flatten groups or interpret client formulas and actions.
The fractions are not pixels or EMUs; negative and beyond-cell fractions remain
unchanged. Reserved anchor bits are ignored as specified. Invalid movement flags,
duplicate identities, ambiguous clients and truncated streams fail inspection.
The current subset does not reclassify Continue records following Obj/TxO as
drawing data; producer output requiring that interleaving remains unsupported.

For explicitly owned plain picture objects, anchor metadata also retains a
one-based BStore reference, raw signed crop/rotation values, clipboard format and
the aspect-preservation flag. The reference must be a scalar `pib` with `fBid`
set in the shape's own FOPT; complex BLIPs and drawing-wide properties are not
substituted. FtCf and FtPioGrbit must occupy their specified Obj fields. DDE,
ActiveX, camera, icon, dynamic/default-sized, controls-stream and auto-load forms,
additional client fields, linked BLIPs, explicit hidden/script anchors and
deleted/OLE/group/background shape flags do not produce passive references.
Unknown property content is never decoded as a script, URL or nested object.

Use `inspect_xls_images --used sample.xls fresh-output-directory` with the same
Cargo invocation to extract only supported images referenced by those objects.
The native `inspect_xls_pictures` helper parses the workbook stream once, binds
anchors to the global catalog by index, and decodes each requested image at most
once. It returns only anchors with a corresponding supported image. Unused
catalog entries are not inflated; an invalid referenced image or out-of-range
index still fails inspection. There is no fallback to another image, file or URL.
The anchor and media stages each retain their separate two-million-work budget.
This raw inspection is not an assertion of complete inherited visibility or
geometry. The measured production path applies additional picture eligibility.

Anchor inspection limits cumulative drawing bytes to 128 MiB, record work to
two million, substream/group nesting to 32, retained anchors to 65,536, and
per-sheet shape/client identities to 65,536. The disjoint ranges prevent repeated
scanning through overlapping worksheet references. Native metadata does not prove
visibility, image eligibility or complete object validity: non-picture objects,
deleted shapes and OLE-marked shapes can have anchors too. These development
helpers remain separate from the measured runtime conversion described above.

## Custom converter contract

```typescript
import { DocxDocument, type LegacyOfficeConverter } from '@silurus/ooxml/docx';

const converter: LegacyOfficeConverter = {
  async convert({ bytes, from, to, maxOutputBytes, signal }) {
    // Run an application-owned local engine or explicitly configured service.
    // The library never supplies a remote endpoint or uploads these bytes.
    const result = await convertLegacyOffice(bytes, {
      from,
      to,
      maxOutputBytes,
      signal,
    });
    return {
      bytes: result.bytes,
      engine: 'example-engine',
      engineVersion: '1.0.0',
      outputSha256: result.outputSha256,
      warnings: result.warnings,
    };
  },
};

const document = await DocxDocument.load(input, {
  legacyConversion: {
    doc: {
      converter,
      timeoutMs: 120_000,
      maxInputBytes: 256 * 1024 * 1024,
      maxOutputBytes: 512 * 1024 * 1024,
      onResult(record) {
        // Content-free provenance: formats, sizes, engine, version, digest, warnings.
        conversionAuditLog.push(record);
      },
    },
  },
});
```

The matching format must be opted in independently. `DocxDocument` reads only
`legacyConversion.doc` and requests `doc -> docx`; `XlsxWorkbook` reads only
`legacyConversion.xls` and requests `xls -> xlsx`; `PptxPresentation` reads only
`legacyConversion.ppt` and requests `ppt -> pptx`. Supplying one field never
enables the other two. A converter must still verify the binary structures it
receives and return `unsupported-input` for an unsupported version or feature
set.

The same `legacyConversion` option is available on the browser viewers and the
Node `open*` / `materialize*` APIs. Node resolves conversion before it lazily
initializes parser WASM.

## Custom disposable Worker adapter

CPU-heavy browser conversion should run in a dedicated Worker. The shared
adapter transfers the source `ArrayBuffer` into one disposable Worker, transfers
the generated package back, and terminates that Worker on success, failure,
cancellation, or timeout:

```typescript
import {
  createDisposableWorkerLegacyOfficeConverter,
} from '@silurus/ooxml/legacy-conversion';

const converter = createDisposableWorkerLegacyOfficeConverter(
  () => new Worker(new URL('./legacy-office.worker.js', import.meta.url), {
    type: 'module',
  }),
  {
    maxConcurrency: 1,
    maxQueuedConversions: 4,
  },
);
```

The Worker installs the matching one-shot host around the application-owned
converter or WASM wrapper:

```typescript
import {
  installLegacyOfficeConversionWorkerHandler,
  type LegacyOfficeConverter,
} from '@silurus/ooxml/legacy-conversion';

const wasmConverter: LegacyOfficeConverter = {
  async convert(input) {
    await initializeConverterWasm();
    return convertWithWasm(input);
  },
};

installLegacyOfficeConversionWorkerHandler(self, wasmConverter);
```

Converter WASM and parser WASM own separate linear memories. The generated
OOXML package must therefore materialize as a standalone buffer once. Transfer
lists prevent additional JavaScript-realm clones, but they cannot eliminate the
copy from converter memory into that buffer or the parser's later copy into its
own memory.

The converter owns its request bytes and may detach their backing buffer. After
resolution, ownership of the returned bytes belongs to the host. Neither side
may retain or mutate bytes after ownership has moved.

## Validation and failure behavior

Converter output is rejected before parser handoff unless it has a bounded,
consistent ZIP central directory, the requested main document part, and a
readable `[Content_Types].xml` that declares that part as DOCX, XLSX, or PPTX.
The preflight also rejects ZIP encryption, unsupported ZIP compression,
duplicate entries and content-type declarations, macro-capable content types,
known VBA and ActiveX part names, malformed content-types XML, and output beyond
the configured limit. These checks are defense in depth; the converter remains
responsible for removing macros and embedded executable content, and the viewer
never executes those features. The ordinary OOXML path is intentionally
unchanged by this converter-only preflight.

Conversion failures are `LegacyOfficeConversionError` instances with stable
`code === 'legacy-office-conversion'`, `stage === 'conversion'`, formats, and one
of these reasons:

- `aborted`
- `timeout`
- `source-too-large`
- `output-too-large`
- `unsupported-input`
- `failed`
- `invalid-output`

Free-form converter exception messages are not propagated. Converter identity,
version, optional lowercase SHA-256, and warnings are bounded metadata supplied
by the converter and must never include document text, filenames, source URLs,
or passwords. The host checks the digest syntax but deliberately does not make a
second full pass over potentially large output. Verify it independently when the
converter is outside the ingestion system's trust boundary.

The defaults admit a 256 MiB source, a 512 MiB output, and two minutes of
conversion. Both byte limits have a non-configurable 1 GiB hard ceiling. A
custom in-process converter must honor the supplied `AbortSignal`; the disposable
Worker adapter enforces cancellation by terminating the Worker even if its WASM
code cannot cooperatively yield. The converter also receives `maxOutputBytes`
so it can stop before materializing an inadmissible package. Viewer reload,
supersession, and destruction
are combined with an application-supplied conversion signal, so an in-flight
disposable converter Worker is also terminated when its owning view no longer
needs the result. One disposable-adapter instance defaults to one live Worker
and four queued conversions; share that instance wherever an application needs
one common concurrency boundary. A full queue rejects with
`capacity-exceeded`.

Encrypted legacy binaries are outside this initial contract. They continue to
fail through the existing encryption path. Macros, external-link updates, and
embedded code must never be executed by a converter.

## Current implementation boundary

The repository now contains the first purpose-built WASM engine in addition to
the opt-in contract, browser/Node normalization, converter-output preflight,
typed errors, and disposable Worker transport. An opt-in local regression run
checks every installed Office-produced legacy counterpart and passes each
generated package to the existing OOXML parser:

```bash
pnpm build:wasm
pnpm test:legacy-converter-private
```

The corpus is deliberately not redistributed. Broader binary-record coverage,
visual fidelity evaluation against Office, fuzzing, and resource measurements
remain part of
[issue #1472](https://github.com/yukiyokotani/office-open-xml-viewer/issues/1472).

## Best-effort fidelity evaluation

Parser acceptance is a smoke test, **not converter completion**. The target is
useful best-effort preservation of the binary input's content and display, with
missing content and visual differences explicitly reported. Pixel equality and
byte-identical ZIP files are not required for each incremental improvement.
Pairing a legacy file with its original
OOXML is useful for investigation, but does not prove fidelity: saving to an old
format can itself change or remove features. Use Office opening the actual legacy
file as the visual reference. Office's upgraded OOXML is useful for mapping
binary records to XML, but conversion itself can change layout and is not an
absolute visual oracle. In particular, rebuilt/down-saved corpus members
must not silently be treated as lossless copies of their original OOXML.

The local macOS oracle opens disposable copies using installed Microsoft Office,
with macros disabled and Word/Excel external-link updates disabled, and exports
both the legacy file and this converter's OOXML to PDF. It compares page counts,
page sizes, and every pixel at 96 DPI. Missing pages, export errors, and any pixel
difference make the exact-comparison run nonzero; this is a diagnostic finding,
not a requirement to add sample-specific adjustments. No blur, registration, resized comparison, or relaxed
threshold hides a discrepancy. PDFs, page images, difference images, source/output
hashes, converter WASM hash, and a report stay in a newly created local temporary
directory; no private artifact is committed or uploaded.

```bash
pnpm --filter @silurus/ooxml-legacy-converter wasm
node scripts/legacy-office-fidelity.mjs --format=xls --limit=10 --python=python3
python3 scripts/legacy-office-compare.test.py
```

To reuse an already exported binary-input reference, explicitly supply
`--format=doc --input=PATH --reference-pdf=PATH` (or the matching XLS/PPT format).
Both paths are required together. The tool hashes the supplied PDF and exports
only the candidate OOXML through Office; it never infers PDF provenance from a
filename. The corpus smoke test also accepts `OOXML_LEGACY_CORPUS_ROOT` for a
separate local checkout and discovers nested files without following symlinks.

The oracle requires Office for macOS, macOS automation permission for each Office
application, Poppler (`pdftoppm`), and Python with Pillow and pypdf. Omit `--format`
and `--limit` to select the full locally installed corpus. Runs are sequential;
an Office failure stops the batch rather than accumulating open documents or
dialogs. Original corpus files and existing visual references are never changed.
Temporary Office-container copies are intentionally retained for diagnosis.

The exporter refuses to open a document unless Office reports its automation
security setting and confirms that macros are disabled. PowerPoint builds that
return no value for this property are currently blocked, even after macOS
automation permission is granted. Word and Excel PDF export have been exercised;
the PowerPoint export path is not yet validated end to end. Do not weaken this
guard to obtain a passing report.

Office-versus-Office PDF comparison helps isolate conversion loss. It is not
sufficient for the viewer's evaluation: compare the converted OOXML's
Canvas output to the same Office oracle separately. Keep renderer self-regression
tests against the previous renderer separate from both fidelity comparisons.
Neither whole-corpus Office equality nor Canvas display equality has been reached.

DOC section decoding is bounded to 16,384 sections and one million property
operations per input (resource policy, not format limits). Intermediate section
properties remain attached to the section-ending paragraph; manual page breaks
are distinct from section breaks. Missing header/footer distances use the
MS-DOC §2.6.4 defaults for the stored producer installation LCID when that LCID
is listed by the specification. Explicit values, including zero, win. Unlisted
languages retain the unresolved-margin warning and zero-distance recovery;
known body margins are retained. The host locale and document text language are
not used to guess the producer's installation settings.
XLS saved custom-view print records cannot override the active worksheet's print
settings, and undefined printer fields are not emitted. These changes do not add
legacy-specific renderer behavior or change per-format opt-in defaults.

Resource policy for PPT reconstruction additionally limits each of retained
outline text and emitted slide text to 128 MiB, and charges persist-directory
entries against the record-work budget. Repeated references cannot bypass the
text limit. Expanded slide XML is capped at 256 MiB across the presentation;
escaping and paragraph markup are charged before appending. Shape property
entries share the record-work budget, group nesting is bounded to 64, and each
slide can emit at most 100,000 shapes/groups. These are implementation resource
policies, not file-format limits. XLS style tables are bounded, repeated fills/borders are interned,
and the BIFF column-256 default-format sentinel never creates an extra column.

PPT text-frame reconstruction follows [MS-PPT] `OfficeArtClientAnchor` and
`OfficeArtClientTextbox`, and [MS-ODRAW] group/child anchors and shape properties.
Only a shape's own text container supplies its text; action data is not traversed
for display content. Child coordinate systems are preserved as ordinary PPTX
groups. Basic unmodified rectangles, ellipses, diamonds, isosceles/right triangles,
straight lines and text boxes map from [MS-ODRAW] `MSOSPT` to ordinary DrawingML
presets. Shapes without text retain their place in the drawing order. Explicit
solid fill/line properties retain literal RGB, opacity and line width; geometry
and style Boolean use bits can independently suppress paint. Entirely absent
paint layers stay transparent pending master/drawing-default resolution. Within
an explicit layer, unspecified properties use the documented MS-ODRAW defaults;
this is not a claim of complete inheritance support. Unresolved colors, nonsolid
paint and custom fill rectangles are omitted, not substituted with guessed colors.
Unknown or customized geometry keeps its text as a transparent frame, without
painting a replacement rectangle. Direct `StyleTextPropAtom`
character and paragraph runs are retained for both inline and outline-referenced
text. Run boundaries count UTF-16 units and include the implicit final paragraph
mark; invalid counts and surrogate-splitting runs fail closed. Font names are
escaped and referenced, not embedded or fetched. Literal RGB and scheme colors are
retained. For verified placeholders, missing supported character/paragraph
properties inherit by text type and indentation level from the main master's
`TextMasterStyleAtom`, then the document's master-style defaults. Direct properties,
including explicit bold/italic/underline resets, take precedence. Ordinary text
boxes do not inherit placeholder styles. Missing unresolved font sizes still use
the warned 18-point fallback. Negative paragraph spacing converts from
master units, while nonnegative spacing retains its percentage semantics.
VT, LF and Unicode line separators become DrawingML line breaks within the same
paragraph; CR remains the PPT paragraph boundary. This follows Unicode UAX #14
BK/LF semantics and does not add a binary-specific renderer path.

Color schemes follow [MS-PPT] `SlideFlags.fMasterScheme`: either the slide's active
eight-color scheme or the scheme of its referenced main/title master is used.
Master IDs resolve through the current persist directory, not stream order;
available-scheme lists are not mistaken for the active scheme. Main masters are
roots, while title masters may inherit. Cycles, dangling references and malformed
active schemes fail closed. A master scheme is cached per conversion, with a
64-level traversal limit and every reference charged to the parsing-work budget.
These bounds are resource policy, not normative file-format limits. Text's
`ColorIndexStruct` and OfficeArt's `fSchemeIndex` have separate index encodings;
both resolve to literal DrawingML RGB for the existing renderer. `fSystemRGB`
also contains literal RGB; system/palette indices remain unresolved. Color-scheme
inheritance does not imply support for master objects, backgrounds or paint-property
inheritance.

Placeholder text inheritance follows [MS-PPT] 2.7.8 and 2.9.35-36/41/44:
only a direct `PlaceholderAtom` whose position is not `0xFFFFFFFF` enables it.
The corresponding inline or outline-referenced `TextHeaderAtom` selects the text
type. Non-placeholder title-like text does not automatically inherit a title style;
compatibility for detached placeholder metadata remains unsupported. Title-master
references are followed to the main master, but title-master shape-specific text
overrides remain omitted. Style tables are parsed once per main master and shared
across slides; at most 10,000 master references, eight types per master and five
levels per type can be retained. The master-count limit is resource policy; type
and level bounds follow the format. Character and paragraph properties beyond the
supported direct subset, including inherited custom rulers, remain omitted.

Missing drawing records retain the earlier unpositioned-text fallback with a
separate warning. Invalid or missing anchors in emitted drawing-backed shapes, ambiguous
coordinate spaces and zero-scale groups fail closed instead of guessing positions.
No migration is required and the independent per-format opt-ins are unchanged.

Text style runs and tab entries share the bounded parsing-work budget. Outline
style records borrow their input bytes until their owning frame is emitted;
unreferenced outline text is not treated as a visible slide object. Intermediate
run tables are released after each text body, and expanded XML is charged against
the presentation-wide limit, including manual breaks and escaped font/text data.

Treat converted OOXML as a derived search/view representation, preserve the
original binary as the authoritative source, and gate production use on a
corpus representative of the documents being ingested. A custom local or remote
adapter remains supported for applications that need a broader conversion
engine; the library still supplies no remote endpoint and never silently uploads
document bytes.

## Direct DOC implementation backlog

Work proceeds in explicitly selected batches; completing one item does not
start the next automatically. The direct binary-to-model reader is the target
architecture. Established behavior is represented by implementation and focused
tests, with specification or bounded Office-observation comments beside them.
Temporary experiment inputs, exports, diagnostics and builds belong under
`outputs/` and are removed after their verification work is finished.
Comparison worktrees must use separate `CARGO_TARGET_DIR` directories as well
as separate WASM output directories; sharing Cargo output can reuse the other
worktree's test executable.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-00 | Recover or repeat the table-style inheritance controls | Repeated for unconditional color and LTR alignment |
| DOC-01 | Focused PAPX alignment regression tests | Implemented |
| DOC-02 | Row-owned TIstd selection and paragraph context | Connected; table admission remains gated |
| DOC-03 | Observed unconditional CHPX color in direct projection | Connected; table admission remains gated |
| DOC-04 | Observed unconditional LTR PAPX alignment in direct projection | Connected; table admission remains gated |
| DOC-05 | Checked logical-column context for irregular rows | Implemented: checked source-cell order, including horizontal/vertical merge slots and RTL; varied direct-DOC Office controls verified |
| DOC-06 | Conditional formatting in the production cascade | Bounded character color/absolute size and logical paragraph alignment are connected through story projection; conditional fonts, physical table alignment and unsupported TAPX remain gated; bounded inherited color/size/PJc are covered by DOC-51 |
| DOC-07 | Further TAPX/PAPX/CHPX precedence | Selected PAPX and CHPX work split into DOC-17 through DOC-22; remaining properties stay gated |
| DOC-08 | Evaluate removal of specific admission gates | DOC-23 evaluates the bounded table-style subset; broader admission still requires the remaining table-property/shading work and end-to-end Office evidence |
| DOC-09 | Validate PrcData property-array limits | Implemented: enforce the MS-DOC cbGrpprl maximum with exact boundary tests; strict FKP framing tracked separately |
| DOC-10 | Exercise FKP-to-Data table-property acquisition | Implemented: actual FKP acquisition tests cover PTableProps, mixed chains, ignored tails, later PCD overrides, physical ranges and failures |
| DOC-11 | Validate the effective mutation target in Office probes | Implemented: reusable read-only inspector and exact-edit validator follow FKP, Data and later PCD properties; bounded TIstd/TTlp assertions require an actual top-level TTP |
| DOC-12 | Reject reserved TC80 vertical-merge values | Implemented: reject used TC80 value 2; retain valid 0/1/3, wholly omitted defaults and ignored excess descriptors |
| DOC-13 | Resolve conditional-region presence across property families | Bounded color presence connected: empty/absent edges, singleton priority, eligible corners and disabled corner flags verified with varied Word controls; noncolor CHPX, PAPX and TAPX still require separate property support |
| DOC-14 | Validate PHugePapx ownership constraints in FKP | Implemented: validate the sole PHugePapx PRL and zero istd at the bounded FKP acquisition boundary; generic Data traversal is unchanged |
| DOC-15 | Validate partial TC80 array records | Implemented: reject incomplete 20-byte array records while preserving whole omitted defaults and ignored complete excess entries |
| DOC-16 | Enforce isolated build caches in the comparison harness | Implemented: bounded exact-model comparison with independent build targets, source/input fingerprints and ownership-checked cleanup; live comparison of 59 DOC inputs preserved admission, errors and streamed models |
| DOC-17 | Share bounded conditional-property operand parsing | Implemented: shared borrowed CNF framing and condition validation for character and paragraph properties; no recursive expansion or newly admitted paragraph properties |
| DOC-18 | Cache the supported table paragraph-style profile | Implemented: the bounded shared profile retains one validated alignment patch and unsupported-property state; different paragraph contexts reuse it without rescanning the table style |
| DOC-19 | Project conditional table paragraph alignment | Implemented within the verified subset: conditional LTR PJc alignment and cross-family presence; table-style PJc80 stays gated after contrary native Word controls (DOC-28) |
| DOC-20 | Project table-style character size | Implemented and tested: absolute CHps inheritance, conditional size/presence and direct overrides. Relative and complex-script size operations remain gated |
| DOC-21 | Project table-style font selection | Implemented within the verified subset: unconditional ASCII/high-ANSI selection, inheritance, direct overrides and existing font-index validation. Conditional font references remain gated under DOC-27 |
| DOC-22 | Verify combined paragraph/character conditional precedence | Implemented: actual 3x3 story projection reproduces the observed color/size/PJc ordering with independent unconditional ASCII/high-ANSI fonts; focused tests cover direct overrides and paragraph marks |
| DOC-23 | Evaluate bounded table-style admission | Evaluation complete for this batch: retain admission gates. DOC-25 and DOC-26 remain prerequisites for general table styles; DOC-27/28 cover the new Office counterexamples. Verified internal projection does not establish full-document admission |
| DOC-24 | Make table-style Office probes reproducible after cleanup | Implemented: deterministic generator and exact mutation validation passed native Word end-to-end controls; source-cell U+0007 ownership is verified separately from the TTP. Generated files remain disposable under outputs |
| DOC-25 | Resolve table-style-aware direct cell shading | Bounded style-aware shading, including both clear-RGB compatibility/RawNil orders, reaches the actual story. General table-style admission still depends on DOC-26 and unsupported style cases |
| DOC-26 | Apply TIstd table properties with specified preservation | Partial: native margin/reset acquisition is covered; cell-edit and preferred-width composition remains unresolved. TIstd/TTlp admission stays gated |
| DOC-27 | Resolve conditional font-table references | Investigated further: native conditional marker fonts stay unchanged when all seven FFN records are reversed or rotated, while unconditional font references follow the reordered table. Preserve the gate; neither a direct index nor a fixed remapping is established |
| DOC-28 | Resolve table-style physical alignment compatibility | Eight additional native controls cover both property orders and opposing left/right values in unconditional and first-row PAPX. PJc determines the display in both orders; PJc80 remains gated pending a documented compatibility policy |

### Selected batch: DOC-24 through DOC-53

This batch contains 30 tracked items, including the existing prerequisite
items and their separately reviewable implementation slices. Parent items
DOC-25/26 are complete only after their runtime and Office acceptance work;
completing a parser helper does not complete those parent items. New findings
outside this selection are recorded without starting another automatic batch.

| Item | Scope | Acceptance / current state |
| --- | --- | --- |
| DOC-29 | Acquire the effective FIB version | Implemented: Read counted arrays and nFibNew using MS-DOC 2.5.14-15; reject truncated/unknown effective versions in native acquisition|
| DOC-30 | Validate the TAPX property category | Implemented: Only table SPRMs are permitted; valid unsupported table properties retain an explicit gate|
| DOC-31 | Validate prohibited and TIstd-preserved TAPX properties | Implemented: Enforce the explicit exclusion list and preservation rules without treating unsupported valid properties as malformed|
| DOC-32 | Validate the default table-style width-before exception | Implemented: Require the specified zero dxa default and reject it in other styles|
| DOC-33 | Validate TCnf framing and bounded nesting | Implemented: Use the shared borrowed CNF parser and work budget; reject recursive TCnf|
| DOC-34 | Validate conditional TAPX property scope | Implemented: Admit the documented conditional border exceptions only in the proper scope; ignore embedded TIstd|
| DOC-35 | Integrate TAPX validation with the cached profile | Implemented: Validate and collect supported facts in one bounded pass; exercise production lookup and malformed inputs|
| DOC-36 | Prepare Raw shading segment boundaries | Preparation implemented; runtime remains under DOC-41: Validate the 22/22/19 cell segments, incomplete operands and empty short rows|
| DOC-37 | Distinguish style-deferred shading from explicit no-fill | Preparation implemented; runtime remains under DOC-41: Preserve Raw ShdNil, ShdAuto, concrete values and unsupported patterns without a guessed background fallback|
| DOC-38 | Apply Raw shading replacement order | Preparation implemented; runtime remains under DOC-41: Later nil or omitted entries remove stale earlier direct shading in the addressed segment|
| DOC-39 | Preserve shading ownership across cell edits | Preparation implemented; runtime remains under DOC-41: Insert/delete/redefine operations move or discard only the corresponding source-cell facts|
| DOC-40 | Select compatibility shading by version and capability | Preparation implemented; runtime remains under DOC-41: Preserve the legacy path; ignore compatibility arrays/ranges only under the documented style-capable rule|
| DOC-41 | Connect style-aware shading to native acquisition | Connected for native projection: effective FIB policy is configured before profile caching; actual-story tests resolve source-cell Raw values against the selected style. XML conversion keeps its previous path and unresolved compatibility arrays remain gated |
| DOC-42 | Preserve table positioning at TIstd application | Connected bounded reset: anchors, positive absolute positions and four wrapping distances survive TIstd; only the independently authored no-overlap flag resets. Native controls and actual acquisition tests cover both property orders |
| DOC-43 | Preserve dimensions and style options at TIstd application | Connected preservation of height, preferred width, autofit, gap and style options. Independent bidi properties now retain their own last boolean values and combine by the documented OR; native 5664 compatibility remains gated under DOC-62. Gap controls establish only before/after equality; revision semantics remain unsupported |
| DOC-44 | Reset remaining row properties at TIstd application | Connected bounded reset of row alignment, header, modern cantSplit and no-overlap, with later direct overrides and repeated TIstd acquisition tests. Cell geometry and TDxaLeft survive. This does not resolve the full TAPX replacement cascade or legacy cantSplit compatibility |
| DOC-45 | Project unconditional TAPX cell shading | Connected for bounded D687 shading: native controls verify child/empty/grandchild precedence, authored Nil versus omission, Auto and direct Raw ordering. Ordinary ipatNil remains no-fill; unresolved compatibility behavior stays gated |
| DOC-46 | Project default and style cell margins | Connected bounded unconditional D634/D63E and direct margin precedence. Clean native controls confirm style D634 after removing masking direct Nil; noninherited same-side D63E/D634 composition is now supported |
| DOC-47 | Project unconditional TAPX table borders | Connected modern LTR D613 inheritance and post-TIstd D613/D62F source-cell overrides, with actual-story tests and budgeted late border payloads. Unstyled compatibility behavior is preserved. TC80/style interaction, old/Nil borders with styles, repeated reset interactions, RTL and broader conditional borders remain separately gated |
| DOC-48 | Connect conditional table shading and presence | Connected for single-style D687 conditions and cross-family presence. Native absent/empty/Nil/Auto/concrete controls verify singleton precedence; inherited TCnf and unsupported shading patterns remain gated |
| DOC-49 | Project conditional cell borders | Connected for one non-inherited FIRST_ROW or FIRST_COLUMN border condition on rectangular unmerged LTR tables. Native controls establish exterior/interior region edges and singleton behavior. Actual-story tests cover six sides and unconditional-border coexistence; multiple conditions, later headers, irregular geometry and conditional/direct-border interactions remain gated |
| DOC-50 | Establish TAPX inheritance precedence | Connected bounded unconditional D687, D613 and independent D63E/D634 inheritance with native child, empty-child, grandchild, reverse-value and direct-override controls. Nil rules are property-specific; inherited TCnf, mixed margin-family chains and other TAPX properties remain gated |
| DOC-51 | Establish conditional PAPX/CHPX inheritance | Implemented for supported color, absolute CHps and logical PJc: varied native controls verify matching/nonmatching conditions, empty and partial children, parent conditional versus child unconditional values and direct overrides. An actual-story child-style test covers cache/context selection; conditional fonts, PJc80 and unsupported properties remain gated |
| DOC-52 | Exercise combined table properties through story projection | Actual-story regression covers source paragraph IDs, ragged/merged cells, conditional color/size/alignment, shading, margins and direct overrides. Live-resource finalization removes merged-continuation orphans without changing projection order. Numbering through discarded continuations remains under DOC-59; no generalized merged-style fidelity claim |
| DOC-53 | Re-evaluate bounded full-document admission | Evaluation retains the unresolved gates: fresh isolated release WASMs agree on 143 inputs (59 private and 84 controls), with four private inputs admitted. Native controls and actual-story tests support only the documented subsets; DOC-25/26 and conditional-font/physical-alignment compatibility remain incomplete |

The DOC-29 through DOC-40 checkpoint passed 785 Rust unit tests and nine
integration tests, with six existing ignores. Fresh, separately built WASM
versions were compared on 59 DOC inputs: admission remained four inputs, and
58 complete results matched. One previously rejected input changed its error
from unsupported notes to an undocumented FIB version (DOC-54); this is an
intentional specification-based rejection, not a visual difference. The
comparison is not an Office-fidelity or renderer-regression result. Prepared
Raw shading was not yet called by native row acquisition at that checkpoint.
The subsequent DOC-41/45 implementation connects that path while retaining
unresolved admission gates.

The margin, border, row-reset and combined-resource checkpoint passed 837 Rust
unit tests and nine integration tests, with six existing ignores; the
no-default-features build also passed. Typechecking and package-build checks
passed after rebuilding all required WASM assets. The final static layout
boundary check and its 95 tests passed, along with 17 compatibility-checker
tests and 26 public-API-checker tests. Fresh isolated release WASMs agree on
all 143 compared inputs, including admission, errors and exact streamed models.
Four of 59 private inputs remain admitted; all 84 controls remain gated.
Equality of rejected results is not a display-fidelity result. Additional
native first-column boundary controls independently confirm the lower and
vertical exterior edges used by the bounded conditional-border projection.

Scoped adversarial review checked exact nested border framing, effective row
origins, source-cell ownership, unchanged unstyled border order, fixed profile
storage, late payload accounting and live-resource finalization. Unsupported
conditional cascades do not expose a partial border patch. No renderer or
shared-model change was needed for this checkpoint; XLS/PPT conversion paths
remain separate. Full-branch architecture and browser/visual acceptance remain
outstanding before integration.

DOC-24 native acceptance subsequently verified 16 marker locations in the
negative and style-connected controls. Removing the selected flattened direct
properties restored the negative controls to default formatting, while the
style-connected document matched the original native-DOC Word PDF exactly at
150 dpi. The validator now distinguishes depth-one cell marks (U+0007) from
separate row marks carrying PFTtp, as required by MS-DOC 2.4.3. Nineteen focused
probe tests passed. This establishes the probe workflow and its selected
unconditional character/paragraph controls; it does not establish TAPX or
conditional inheritance compatibility.

The subsequent native shading and conditional-inheritance checkpoint passed
796 Rust unit tests and nine integration tests, with six existing ignores.
Fresh isolated WASM builds compared 59 private inputs and 43 controls with no
result differences. Four private inputs remain admitted; all 43 controls remain
admission-gated. Actual-story tests separately verify Raw/style shading,
ordinary ipatNil no-fill and inherited conditional color/size/alignment; equal
rejection results do not establish display fidelity.

Adversarial review found and fixed a newly reachable projection defect that
used an ordinary ipatNil background color as a fill. A focused actual-story
test failed before the fix and passed afterward. Review also confirmed bounded
profile caches and condition sets, source-cell ownership without row clones,
unchanged XML compatibility handling and explicit gates for unresolved Native
Word behavior. No renderer or cross-format public API was changed. Whole-branch
architecture and browser/visual acceptance remain outstanding.

### Additional findings outside the selected batch

| Item | Scope | Status |
| --- | --- | --- |
| DOC-54 | Investigate the undocumented effective FIB version 0x00C3 | Reviewed against the official version/count tables and product notes. The documented C0/C2 exceptions do not establish C3 support; corpus occurrence and another reader's numeric clustering are insufficient. Explicit rejection remains |
| DOC-55 | Resolve native Word Raw-Nil compatibility shading | Implemented bounded clear-RGB modern compatibility fallback in both orders, including empty/short replacement tails. Patterns, range shading and broader native style/version interactions remain incomplete |
| DOC-56 | Resolve authored Nil cell-margin precedence | Bounded prerequisite resolved with DOC-46: native D632 and direct D634 Nil controls match explicit Dxa zero, while omission exposes the applicable style value; authored style D634 Nil masks inherited Dxa. Retain authored state until resolution. This does not establish conditional CSSA or mixed D63E/D634 style-chain composition |
| DOC-57 | Recover a usable native oracle for ordinary ipatNil | A new ordinary-ipatNil control with valid Auto COLORREFs still causes a native Word display-update error. No PDF oracle was obtained. This trial family is paused to avoid repeated Word errors; specification-based no-fill remains separately tested |
| DOC-58 | Avoid resources orphaned by merged-cell projection | Fixed within DOC-52: horizontal/vertical continuation regressions fail before live-resource finalization and pass afterward. Retained nested/body/header/footer/section references remain deduplicated; traversal scratch is budgeted and image-free documents bypass the pass |
| DOC-59 | Establish numbering behavior in merged continuations | Bounded behavior verified: two horizontal merge ranges and a vertical continuation suppress visible paragraphs while still advancing the shared list. Actual-CFB regression tests preserve the existing counter order; no counter implementation change |
| DOC-60 | Validate nested operand framing in positive Office probes | Implemented bounded replacement-operand checks, explicit negative placement mode and validation-coverage reporting. Historical recipe plans replay exact provenance and receive newly computed checks; malformed nested lengths cannot pass as positive evidence |
| DOC-61 | Audit border-origin projection beyond the bounded subset | Additional native controls retain white single and thin double direct borders: black wins equal-width white ties, while the thin double wins its equal-weight style tie. Current bounded weighting is supported; opposing direct layers, spacing, Nil and RTL remain unverified |
| DOC-62 | Resolve native TFBiDi90 compatibility | Read-only revalidation confirms independent last-value Bool16 state and the documented OR in acquisition. Current native PDF versus save-time DOCX disagreement remains unresolved; TIstd/TTlp gates still reject all eight controls regardless of the final direction |
| DOC-63 | Distinguish legacy and modern cantSplit compatibility | Implemented an explicit current-Word native-reader policy: validate but ignore legacy 3403 and apply modern 3466. Selection does not depend on document nFib. Generic/XML behavior remains unchanged; direct acquisition and reset ordering are tested |
| DOC-64 | Audit table identity for explicit default values | Implemented semantic false/omitted identity for no-overlap, with nonmutating invalid-value checks and actual acquisition/source-row grouping tests. Native false/omitted controls match; native inline true also stays grouped and is tracked separately as DOC-84 |

The current formatting tests exercise the typed projection methods; they do not
establish full-document admission or visual compatibility. The complete branch
still requires its outstanding architecture and visual-regression review before
merge or release. XLS and PPT follow separate backlogs and are outside this batch.

### Selected follow-up batch: 30 items

The next user-authorized batch selects the eleven unresolved items DOC-27,
DOC-28, DOC-54, DOC-55, DOC-57, DOC-59, DOC-60, DOC-61, DOC-62, DOC-63 and
DOC-64, plus the nineteen bounded slices below. Prior test counts remain
historical. Investigations can establish a narrower supported behavior or
identify a concrete remaining blocker; a retained gate is not implemented
compatibility. DOC-25/26 remain parent acceptance items. Findings outside this
selection are recorded without automatically extending the batch.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-65 | Validate nested conditional-border operand lengths | Implemented exact nested BrcOperand length checks; non-normative NilBrc recovery cannot certify a positive BrcOperand |
| DOC-66 | Validate probe CNF conditions and nesting | Implemented all twelve CNFC values, exact wrapper ownership, family checks and bounded nonrecursive nesting |
| DOC-67 | Validate probe cell-margin operand framing | Implemented exact CSSA size, cell range, side mask, units and documented width constraints for the recognized operands |
| DOC-68 | Validate probe Raw shading segment framing | Implemented exact RawShd segment framing and 22/22/19 cell limits; individual shading values are not claimed as fully validated semantics |
| DOC-69 | Validate positive-probe property families and PAPX prefixes | Implemented replacement UPX family and PAPX-owner checks, recognized TAPX placement restrictions and required default-style width checks; untouched property sets are not recertified semantically |
| DOC-70 | Bound positive-probe validation work and negative-control handling | Implemented shared 1 MiB/4096-SPRM validation budgets and an explicit specification-invalid placement mode that cannot bypass malformed operand framing |
| DOC-71 | Separate native row compatibility policy from XML conversion | Implemented explicit native-reader cantSplit policy while preserving generic/XML behavior |
| DOC-72 | Verify ordered legacy and modern row-split properties | Implemented and tested ordered 3403/3466 handling, independent of nFib, with native-current-Word behavior distinguished from generic/XML acquisition |
| DOC-73 | Establish semantic defaults used for row identity | Implemented only the documented false no-overlap default normalization. Other raw identity defaults remain outside this slice |
| DOC-74 | Verify repeated TIstd row-reset interactions | Verified repeated TIstd resets and later modern cantSplit overrides through actual native acquisition; ignored legacy values cannot resurrect cleared state |
| DOC-75 | Exercise native acquisition and source-row grouping | Added native/generic acquisition and source-row grouping coverage, preserving source TTP IDs and separating row-local cantSplit from table identity |
| DOC-76 | Establish last-row conditional border regions | Connected bounded LAST_ROW border regions through cached profile and actual story projection; six sides, singleton, disabled flag and existing shape gates are covered |
| DOC-77 | Establish last-column conditional border regions | Connected bounded LAST_COLUMN border regions through cached profile and actual story projection; six sides, singleton, disabled flag and existing shape gates are covered |
| DOC-78 | Establish overlapping conditional border precedence | Further native different-side and interior-edge controls support the bounded FIRST_COLUMN plus FIRST_ROW cascade in either serialized order. Its production projection is connected in the next batch; other condition pairs remain gated |
| DOC-79 | Establish inherited conditional TAPX border behavior | Further partial-side, reverse-valued and repeated-record controls support bounded per-side conditional borders. FIRST_ROW and FIRST_COLUMN logical-left inheritance are connected; other inherited border sides/regions and shading remain gated |
| DOC-80 | Bound merged and irregular conditional border geometry | Clean sole-TDef horizontal source-span geometry is verified in both merge positions. Cell-edit/preferred-width combinations still differ; the conditional merged-geometry gate remains |
| DOC-81 | Establish conditional margin composition | Implemented single non-inherited FIRST_ROW D63E on all four physical sides over a matching unconditional D63E baseline. Conditional D634 and broader/inherited composition remain gated |
| DOC-82 | Run isolated model comparison and adjudicate changes | Fresh isolated release WASMs agree on all 187 inputs (59 private and 128 controls), including errors and streamed models; four private inputs admitted. The final comparison includes the reviewed border-presence correction |
| DOC-83 | Review the selected batch and re-evaluate admission | Scoped adversarial review found and fixed a competing singleton condition gap: active cross-family edge conditions now gate the border patch, with actual-story regressions. Final model revalidation passed after this correction. Parent admission and full-branch architecture/browser/visual acceptance remain outstanding |

### Additional finding from the follow-up batch

| Item | Scope | Status |
| --- | --- | --- |
| DOC-84 | Resolve native inline no-overlap grouping | A matched current-Word control keeps an inline middle row in the same table even with no-overlap true; its PDF matches omission and its saved DOCX keeps the three-row table. The existing normative true identity distinction is not thereby disproved for positioned tables. Isolate inline versus positioned behavior before changing that rule |
| DOC-85 | Integrate cross-family border condition matching | Bounded row/column border pairs now cover all four combinations in native controls and actual story tests. Competing nonborder edge conditions, eligible corners, wider cascades and merged/irregular shapes remain conservatively gated |

The follow-up batch implements the bounded probe-validation, current-Word
row-split policy, false no-overlap identity, numbering regressions and final-edge
border slices. Native investigations for the remaining compatibility items
record narrower evidence and concrete next controls; they are not claims that
all thirty selected items now provide complete format support.

Final verification for this follow-up checkpoint: 846 Rust unit tests and nine
integration tests passed, with six existing ignores; the no-default-features
check passed. The table-style probe suite passed 34 tests, its PAPX dependency
passed nine, and the unchanged renderer border suites passed 33. The static
layout-boundary check and diff checks passed. Fresh isolated baseline and
candidate release WASMs agree on 187 inputs; admission remains four of 59
private DOCs, and all 128 controls are gated. This comparison is not visual
fidelity or full-branch architecture acceptance.

Adversarial review corrected positive-probe NilBrc certification and table-SPRM
placement drift, preserved exact historical-plan replay with current coverage
reporting, and added the competing-edge border gate. Runtime additions reuse
bounded profile state and existing source-row ownership; no renderer, shared
model or XLS/PPT behavior was added. Generic DOC-to-XML no-overlap identity now
uses the same documented false default, while its cantSplit behavior is retained.
The remaining investigations and DOC-84/85 prevent any general table-style
admission or completion claim.

### Selected next batch: 30 items

This user-authorized batch selects DOC-27, DOC-28, DOC-54, DOC-55, DOC-57,
DOC-61, DOC-62, DOC-78, DOC-79, DOC-80, DOC-81, DOC-84 and DOC-85,
together with the seventeen bounded slices below. Parent investigations and
implementation prerequisites are tracked separately; no helper or control alone
establishes general admission. The ordinary-ipatNil trial family remains paused
after native Word display-update errors. New findings do not extend the batch.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-86 | Verify different-side row/column border overlap | Native different-side controls preserve both row-top and column-left edges at overlap in either record order; bounded projection is connected with final review and regression tests |
| DOC-87 | Verify conditional interior-border overlap | Native row insideV and column insideH coexist with matching PDFs in either record order; broader edge combinations remain outside this control |
| DOC-88 | Verify repeated conditional border records | Native repeated disjoint sides compose; a repeated same-side property uses the later serialized value in both value orders |
| DOC-89 | Verify partial inherited conditional borders | Native partial child overrides preserve omitted parent exterior sides; a separately connected three-by-three child table verifies the inherited/replaced interior vertical edges |
| DOC-90 | Verify reversed conditional-border inheritance | Native reverse-valued child replaces the top side while retaining the parent left side; broader inherited conditions remain gated unless independently covered |
| DOC-91 | Bound condition-indexed border profile state | Implemented fixed four-condition/six-side raw storage within the existing bounded profile cache; unsupported inheritance and Nil clear the border projection |
| DOC-92 | Apply accepted multi-condition border precedence | Connected FIRST_COLUMN then FIRST_ROW through shared selection; raw winners are resolved before decoding and charging retained strings. Rectangular unmerged LTR and direct-layer restrictions remain |
| DOC-93 | Isolate conditional margin property families | Native isolation confirms unconditional and FIRST_ROW D63E margins; conditional D634 does not create a row-specific PDF offset despite its saved representation |
| DOC-94 | Verify margin serialization versus application order | Both conditional/unconditional D634 serialization orders have identical native geometry; conditional D634 stays gated |
| DOC-95 | Verify partial conditional margins and direct overrides | Native direct D632 Dxa, authored Nil and omission remain distinct; table indentation and mixed-family baselines limit general composition acceptance |
| DOC-96 | Compare positioned no-overlap identity | Native positioned omission/false keep one three-row table; true splits it into three one-row tables. Source-row regression preserves the corresponding TTP ownership; inline compatibility remains DOC-84 |
| DOC-97 | Isolate compatibility shading cell ranges | Native D609 one/three-entry arrays are ignored; post-TIstd D612 affects exactly its one/three-cell prefix with later Raw Nil. Other segments and replacements remain unverified |
| DOC-98 | Verify compatibility shading reset interactions | Native D612 before TIstd is discarded while the same values after TIstd survive later Raw Nil. Retain the specification-conflicting compatibility gate pending broader controls |
| DOC-99 | Verify complementary merged/irregular border geometry | Native opposite horizontal and partial vertical merges suppress swallowed boundaries and retain the remaining boundary; compensating widths preserve the source exterior edge. Combined irregular merges remain gated |
| DOC-100 | Verify opposing direct border origins | Native swapped opposing red/blue direct borders select red under the documented brightness tie-break; exterior thin-double control agrees with the existing conflict result. Spacing, Nil and full reading-order ties remain unverified |
| DOC-101 | Compare the final isolated native models | Fresh independent release WASMs agree on 221 cases: 59 private inputs and 162 controls. Admission remains four private inputs, with controls gated; equal errors are not display-fidelity evidence |
| DOC-102 | Review final changes and reassess admission | Scoped adversarial review corrected eligible-corner selection and retained-string accounting. Final tests passed; full-document admission and full-branch architecture/browser/visual acceptance remain outstanding |

The current batch's native trials also distinguish a positioned no-overlap
identity change from the unresolved inline case. They do not justify applying
that distinction as a new undocumented compatibility rule. Conditional margins
and compatibility shading retain their gates where the observed result has not
yet isolated the complete cascade.

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-103 | Respect eligible corner presence in border admission | Adversarial review identified that a character corner can activate an edge absent from ordinary border presence, suppressing an opposing singleton edge. A conservative eligibility gate and source-story regression prevent partial border projection; each mapped side also uses shared selection. Generalized corner borders remain separate work |

This checkpoint adds bounded multi-condition and inherited border projection,
plus regressions for positioned row identity and retained conditional-margin
gates. Native Word completed 34 new control exports without a display-update
error; the previously failing ordinary-ipatNil family was not retried. The
compatibility investigations are not complete merely because their individual
controls now have an observed result.

Fresh verification: 855 Rust unit tests and nine integration tests passed, with
six existing ignores. The no-default-features check, 34 probe tests, nine PAPX
inspector tests, 33 unchanged renderer border tests, static layout-boundary
check, targeted formatting and diff checks passed. The final isolated release
WASM comparison covers 221 cases with unchanged admission, errors and exact
streamed models. Four of 59 private inputs remain admitted; this checkpoint
does not establish full table-style admission or visual compatibility.

The reviewed runtime delta is confined to native DOC table-style caching and
border projection. Its four inline condition slots reuse the existing bounded
cache; condition evaluation allocates no collection, source rows remain owned
by the prepared story, and only final raw side winners are decoded and charged.
Unconditional row fallback and cell-side borders retain their existing model
roles. Renderer, worker protocol, shared models, public opt-in options and the
independent XLS/PPT paths are unchanged. This scoped acceptance does not replace
the outstanding review of the full feature branch.

### Selected admission-prerequisite batch: 30 items

This user-authorized batch selects parent items DOC-25, DOC-26, DOC-55,
DOC-80, DOC-81 and DOC-85, plus the twenty-four bounded slices below.
Previously blocked font, FIB-version and ordinary-ipatNil investigations are not
repeated without new evidence. Scope is finite; newly found issues are recorded
without automatically starting another batch. Parent admission requires actual
runtime and native Office acceptance, not only helper tests.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-104 | Verify empty compatibility-shading replacement | Verified: empty replacement removes the modern compatibility prefix before a later Raw array |
| DOC-105 | Verify shading representation ordering | Verified: D612 colors survive either D609 serialization order |
| DOC-106 | Verify Raw concrete/Auto shading precedence | Verified: later Raw Auto clears and concrete Raw colors replace the compatibility fallback |
| DOC-107 | Verify the second compatibility-shading segment | Verified: the second segment starts at source cell 23 with adjacent-cell countercontrols |
| DOC-108 | Verify the third compatibility-shading segment | Verified: the third segment starts at source cell 45 with adjacent-cell countercontrols |
| DOC-109 | Define the bounded native shading policy | Implemented within a bounded clear-RGB fallback policy; unsupported orders and patterns remain gated |
| DOC-110 | Preserve compatibility shading through cell edits | Implemented: compatibility state follows source-cell insertion, deletion and redefinition |
| DOC-111 | Exercise shading through actual container acquisition | Verified actual CFB-to-story fallback; full-document TIstd/TTlp admission remains blocked |
| DOC-112 | Anchor margin controls to explicit cell edges | Corrected: explicit native borders now anchor each pilot independently; initial eight borderless controls excluded |
| DOC-113 | Verify conditional style-margin sides | Implemented all four physical FIRST_ROW D63E sides after corrected direct-right, active/disabled and reversed-magnitude native controls |
| DOC-114 | Verify disabled conditional-margin flags | Observed: disabling FIRST_ROW moves the grid and removes its relative left offset; general cascade unresolved |
| DOC-115 | Verify direct margins with paired geometry | Verified direct left and right D632 Dxa/Nil/omission with visible cell borders and actual CFB story precedence. Broader composition remains bounded |
| DOC-116 | Cache accepted conditional-margin patches | Implemented fixed optional FIRST_ROW margin patch within the existing bounded style cache, with final order-independent baseline checks |
| DOC-117 | Project accepted conditional margins in the story | Connected bounded per-source-cell margin projection through shared condition selection before direct formatting resolution |
| DOC-118 | Bound conditional-margin retention and work | Verified fixed 63-cell scratch storage, count validation and common precedence helper; no per-row heap allocation |
| DOC-119 | Reassess conditional-margin admission | Reviewed: bounded FIRST_ROW D63E is projected on all physical sides; full TIstd/TTlp admission and other conditional margins remain gated |
| DOC-120 | Verify additional row/column border pairs | Implemented: all four one-row/one-column pairs; native exterior/interior serialization controls agree |
| DOC-121 | Verify inherited first-column borders | Implemented: inherited FIRST_COLUMN six-side profile with a partial child override and visible insideH |
| DOC-122 | Verify mixed-family edge-condition selection | Verified in actual story models: column-before-row order and source-owned exterior/interior edges |
| DOC-123 | Resolve conditional boundaries from explicit merges | Implemented tested FIRST_ROW source-end right ownership and swallowed insideV suppression as groundwork; explicit geometry gate remains pending DOC-150 |
| DOC-124 | Bound conditional-border geometry work | Reviewed: current unmerged geometry remains bounded; merged source-owner work is still required |
| DOC-125 | Reassess the bounded border projection gates | Reviewed: exactly one row plus one column supported; wider cascades, merges, irregular grids and other gates retained |
| DOC-126 | Compare final isolated model results | Verified: 257/259 exact results; two expected D660 rejection-message changes, no admission or streamed-model change |
| DOC-127 | Review final changes and parent admission | Scoped review complete; required bounded checks pass and unresolved parent admission remains explicit |


| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-128 | Reset row-level D660 shading at TIstd | Normative preceding-D660 reset implemented from the TIstd preservation list; native trio is inconclusive because none visibly paints the expected fill. Reverse order remains gated |
| DOC-129 | Resolve preferred cell width at TIstd | Before/after TIstd retains D635 in diagnostics and matched native display. Interactions with cell-edit geometry are not isolated; no general preferred-width layout policy is established |
| DOC-130 | Verify visible anchors in margin controls | Corrected controls contain visible borders. Review also corrected use of one PDF's border coordinate for another: measure each grid independently. Margin origin and D634/D63E composition remain unresolved |
| DOC-131 | Resolve compatibility shading after authored Raw Nil | Implemented bounded clear-RGB reverse replacement after native full/empty/short countercontrols. Authored Raw Nil is updated across the entire segment; unsupported patterns/ranges/style cases remain gated |


### Admission-prerequisite batch acceptance

The thirty selected items were reviewed to the bounded scope above. Conditional
border projection now composes every one-row/one-column edge pair and retains
all supported inherited FIRST_COLUMN sides. Modern compatibility shading now
supplies clear RGB fills to a later explicit Raw ShdNil in all three source-cell
segments. These facts reach the actual CFB-to-story path; independent TIstd and
TTlp admission gates remain.

Native Word produced 38 binary-input PDF controls and diagnostic DOCX exports.
Only controls with independently verified visible geometry or fills support the
implementation. Eight initial borderless margin controls do not prove relative
margins, and the D660 trio does not establish native shading precedence. The
corrected margin pilots expose grid-origin movement and leave composition
unresolved. Additional findings are recorded as DOC-128 through DOC-131.

Fresh verification passed 861 Rust unit tests and nine integration tests, with
six existing ignores; feature-off compilation, 34 probe tests, nine PAPX tests,
33 unchanged renderer border tests, static layout boundaries, scoped formatting
and diff checks also passed. Independent release WASM builds over 259 identical
inputs produced 257 exact matches. The remaining two controls stay rejected:
resetting preceding D660 or gating following D660 changes only which unsupported
formatting diagnostic is reported. No admitted stream or admission status
changed. Four of 59 private inputs remain admitted; this is not a fidelity or
full-table-style completion claim.

Adversarial review corrected empty-segment slicing and gated reverse-order
compatibility replacements across their entire source segment. Per-cell state
uses existing size-based accounting and fixed record/cell bounds. Border
composition retains at most two inline patches, chooses source-owned sides in
specified order, and introduces no geometry inference. Renderer, shared model,
worker and independent XLS/PPT paths are unchanged. Merged boundary ownership,
conditional-margin composition and whole-branch architecture/browser/visual
acceptance remain outstanding. This finite batch does not start further work.

### Selected source-ownership and reset batch: 30 items

This finite user-authorized batch selects DOC-25, DOC-26, DOC-55, DOC-80,
DOC-81, DOC-113, DOC-115, DOC-116, DOC-117, DOC-118, DOC-123,
DOC-129 and DOC-131, plus the seventeen slices below. Existing blocked native
trials are not repeated without a changed, informative control. New findings
are added to the backlog without extending this batch.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-132 | Map explicit horizontal source spans | Implemented fixed source-span mapping; native width redistribution remains DOC-150 |
| DOC-133 | Transfer the source-end right border to the retained cell | Implemented conditional source-end selection onto the retained owner; generic direct merging unchanged |
| DOC-134 | Suppress swallowed conditional insideV borders | Verified swallowed insideV suppression in both merge positions and actual story models |
| DOC-135 | Preserve direct-border and unsupported merge gates | Retained direct-layer, vertical, bidi, malformed and unsupported-condition gates |
| DOC-136 | Test merged boundaries through actual CFB-to-story projection | Verified actual CFB owner projection with an independent merged-geometry rejection gate |
| DOC-137 | Bound source-span geometry and retained memory | Verified fixed 63-slot span bound and malformed/max-count cases; no retained heap map |
| DOC-138 | Measure each native table grid independently | Verified each PDF against its own visible borders; rejected intent-only alignment assumptions |
| DOC-139 | Verify four physical conditional D63E sides | Implemented evidenced top/left/bottom FIRST_ROW D63E; right remains DOC-151 |
| DOC-140 | Isolate D634 and D63E baselines | Observed D63E/D634 distinction; unresolved style D634 behavior recorded as DOC-149 |
| DOC-141 | Verify anchored direct D632 Dxa/Nil/omission | Verified left D632 Dxa/Nil/omission offsets and actual per-cell model precedence |
| DOC-142 | Verify reversed Raw Nil and compatibility arrays | Implemented clear-RGB fallback in either RawNil/compatibility order; other forms remain gated |
| DOC-143 | Verify empty and shorter reversed replacements | Implemented empty/short replacement clearing the full segment tail, with native countercontrols |
| DOC-144 | Distinguish D635 preferred width from cell-definition geometry | Verified retained D635 serialization across TIstd; identical PDFs do not establish live width |
| DOC-145 | Exercise table reset through actual acquisition | Verified actual acquisition for reverse shading, repeated reset and distinct D635/physical width |
| DOC-146 | Audit remaining TIstd and TTlp admission prerequisites | Reviewed concrete TIstd/TTlp prerequisites; blanket admission remains gated |
| DOC-147 | Compare final isolated models | Verified: all 280 results match exactly in independently built release WASMs; admission and successful models unchanged |
| DOC-148 | Adversarially review and close the finite batch | Scoped review and proportionate checks complete; full parent/architecture acceptance remains incomplete |

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-149 | Reconcile unconditional table-style D634 with native output | Resolved the observed discrepancy: later direct D634 Nil masked the style default. Clean D634=288 controls show the margin. Noninherited same-side D63E-over-D634 composition is now supported; inherited overlap remains gated |
| DOC-150 | Resolve native grid redistribution after horizontal merging | Refined: sole TDef source spans match native 6000/3000 and 1500/7500 merges. Competing TInsert/TDxaCol/D635 families change the result; no generic equalization rule is valid. Keep the geometry gate pending DOC-172 |
| DOC-151 | Establish a visible conditional-right margin oracle | Verified with correctly selected, visibly bordered T07 and effective direct PJc=2. Conditional-right active/disabled/reversed and direct D632 cases support the bounded implementation; ineffective pilots are excluded |

### Source-ownership and reset batch acceptance

The selected thirty items are closed at their explicit bounded statuses above;
unresolved parent and evidence-dependent work remains in the backlog. This
checkpoint connects FIRST_ROW top/left/bottom margin patches before direct cell
formatting, updates reverse compatibility shading across complete replacement
segments, and prepares source-owned horizontal conditional borders behind an
independent geometry gate. Existing plain/direct merge behavior is preserved.

Native Word exported 21 controls. Root review rejected the attempted right-margin
oracle because paragraph formatting kept the text left aligned. Merged controls
validate visible boundary ownership but expose a different native grid, so they
do not establish full geometry. D635 diagnostic serialization is retained across
TIstd, while identical PDFs do not establish a live width effect. These limits
are recorded in DOC-149 through DOC-151 without speculative compatibility rules.

Fresh checks passed 874 unit tests and nine integration tests, with six existing
ignores; feature-off compilation, 34 probe tests, nine PAPX tests, 33 unchanged
renderer border tests, static layout boundaries, scoped formatting and diff
checks also passed. Independent release WASMs match exactly for 280 identical
inputs. Four of 59 private inputs remain admitted. Equal rejections are not a
visual-fidelity result, and no full table-style admission is claimed.

Adversarial review enforced final order-independent margin baseline validation,
kept shape gates specific to actual conditional-margin presence, removed a
proposed generic merge change, and eliminated duplicate retained border winners.
Both margin and span scratch storage are fixed to 63 source cells. Renderer,
shared models, workers and independent XLS/PPT paths are unchanged. Whole-branch
architecture/browser/visual acceptance remains outstanding. No further batch is
started by this checkpoint.


### Selected margin and geometry follow-up: 30 items

This finite batch selects unresolved DOC-26, DOC-46, DOC-80, DOC-81,
DOC-113, DOC-115, DOC-119, DOC-129 and DOC-149 through DOC-151, plus the
nineteen bounded slices below. A measured counterexample or retained gate does
not complete the parent implementation. New findings are recorded separately.
Native exports run serially, with exact source mutation provenance. Historical
results do not substitute for fresh acceptance of the final source.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-152 | Establish direct marker-paragraph alignment ownership | Verified effective PJc operand and selected table style, not merely opcode presence; rejected center-aligned and wrong-style pilots |
| DOC-153 | Measure conditional-right active and disabled controls | Verified three first-row cells against their own borders and the second row, with matching disabled controls |
| DOC-154 | Verify reversed and varied right-margin magnitudes | Verified reversed 72/288-twip right-margin values; the selected conditional value wins in either magnitude direction |
| DOC-155 | Verify direct right D632 Dxa, Nil and omission | Verified right D632 Dxa432, Nil and omission in a bordered three-cell row; actual CFB story regression added |
| DOC-156 | Extend the bounded conditional-margin profile if evidenced | Implemented physical right in the existing bounded four-side patch; no new allocation or cache key |
| DOC-157 | Exercise all accepted physical sides in actual story projection | Verified actual CFB projection for enabled/disabled all-side margins and direct-right precedence |
| DOC-158 | Isolate default-margin omission from authored D634 | Verified clean no-direct controls; later direct Nil had masked the style default in earlier pilots |
| DOC-159 | Distinguish default-margin zero, Nil and nonzero values | Verified clean style D634 Dxa0 and Nil controls show no inset, while Dxa288 is visible; acquisition preserves the distinction from omission |
| DOC-160 | Verify D634/D63E serialized ordering with explicit borders | Verified D63E wins over same-side D634 in both serialized orders and reversed magnitudes; inherited cross-family cases remain gated |
| DOC-161 | Reconcile cached default margins with evidenced native scope | Implemented bounded noninherited same-side composition; retained the separate row-default and style-cell patches |
| DOC-162 | Exercise default-margin precedence through native acquisition | Verified native acquisition for unmasked D634, masking direct Nil, and D63E above final direct D634; independent TIstd admission gate asserted |
| DOC-163 | Isolate TC80 preferred width from horizontal source spans | Investigated matched/equal/nil TC80 preferences; competing cell-edit families prevent a general preference conclusion |
| DOC-164 | Isolate autofit during merged-width calculation | Matched fixed/autofit controls display identically within the tested competing-property fixture; this does not establish general autofit behavior |
| DOC-165 | Verify unequal source geometry in both merge positions | Verified clean unequal sole-TDef source spans in left and right merge positions, with no competing geometry records |
| DOC-166 | Compare D635 order with an explicit physical-width oracle | Verified matched D635-before/after-TIstd display and retention; a general live-width rule remains unresolved because cell-edit properties coexist |
| DOC-167 | Apply only supported width behavior and preserve explicit gates | Preserved source-span planning and the explicit conditional merge gate; no width equalization heuristic added |
| DOC-168 | Check margin and geometry bounds and retained ownership | Reviewed unchanged fixed four-side/63-cell storage, bounded source-span planning and ownership; added raw-record-to-planner regression |
| DOC-169 | Compare final independently built release models | Verified all 317 inputs match exactly, including admission, errors and streamed models |
| DOC-170 | Adversarially review final scope and admission prerequisites | Scoped review and final checks complete; full TIstd/TTlp admission and whole-branch architecture/browser/visual acceptance remain outstanding |


| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-171 | Validate the effective experimental target across property families | Add reusable probe assertions for operand values, selected style membership, and masking/competing properties across complete FKP/Data/PCD acquisition. Exact byte edits alone do not establish an effective native control; current corrected controls retain local assertions |
| DOC-172 | Resolve cell-edit and preferred-width family precedence | Isolate TInsert/TDxaCol/D635 against a sole TDefTable and then implement only an evidenced rule. Earlier cell-edit properties can alter native merged widths despite later byte-order TDefTable; keep DOC-150 geometry admission gated |


### Margin and geometry follow-up acceptance

The selected thirty items have the explicit implementation, verification or
unresolved statuses above. This checkpoint supersedes earlier right-margin and
D634 discrepancy statuses; it does not complete the parent table-style or
preferred-width admission tasks, and does not start another automatic batch.

The bounded FIRST_ROW D63E path now projects all four physical sides, including
direct right D632 Dxa/Nil/omission precedence. Noninherited unconditional D63E
and D634 can coexist on the same side, with D63E supplying the style-cell value
above the row default. Inherited cross-family and conditional D634 overlap
remain gated. Native controls established both serialized orders and reversed
magnitudes. Removing residual direct Nil showed that the earlier apparent
missing style D634 was masking, not evidence that Word ignores the property.

Clean sole-TDef controls establish horizontal source-span sums in both merge
positions. Earlier cell edits and preferred-width properties can change native
results; those combinations remain unresolved. The planner is unchanged, and
the independent conditional merged-geometry gate remains. Ineffective alignment,
wrong-style and competing-property pilots were excluded from the behavioral
argument. Exact mutation provenance must include effective operands and selected
style membership, not merely successful file export or opcode presence.

Fresh checks passed 878 unit tests and nine integration tests, with six existing
ignores; feature-off compilation, 34 probe tests, nine PAPX tests, 33 unchanged
renderer-border tests, static layout boundaries, scoped formatting and diff
checks also passed. Independently built release WASMs match exactly for all
317 identical inputs. Four of 59 private inputs remain admitted. Equality of
rejections is not visual fidelity, and full table-style admission is not claimed.

Scoped adversarial review rejected the confounded native interpretations,
verified the remaining admission gates, and confirmed unchanged fixed four-side
and 63-cell storage. No renderer, shared model, worker or XLS/PPT behavior changed.
Whole-branch architecture/browser/visual acceptance remains outstanding before
integration. New DOC-171 and DOC-172 retain the concrete follow-up work.
