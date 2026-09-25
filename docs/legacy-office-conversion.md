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

Fields in the direct DOC source are checked against each story's own field
table: the main document, headers and footers, footnotes, endnotes and
textboxes (MS-DOC 2.8.25). They map onto the fields the DOCX reader already
supports:

- PAGE, NUMPAGES, DATE and TIME are computed when the document is laid out,
  as for DOCX. Word does the same: a DOC header DATE field prints the export
  date in Word's PDF, not the date stored in the file. Only switches the
  renderer interprets exactly are accepted: numeric page formats,
  MERGEFORMAT or CHARFORMAT, and date pictures that need no language data.
- Form check boxes show their stored state and size.
- Every other field shows its stored result. Fields are never executed.
  HYPERLINK fields and REF or PAGEREF fields with `\h` become links. Inside a
  table of contents, link text keeps the paragraph's color and underline,
  which matches Word's PDF. Stored results are not recomputed, so a
  PAGEREF number can differ from a PDF that Word produced after updating it.
- Some fields reject the document. These include equations (EQ, often used
  for phonetic guides), macro buttons, drop-down form fields and SYMBOL
  fields without a stored result. Others are fields shown as codes, locked
  or edited page and date fields, and nested hyperlinks. A private result
  with content is rejected unless it is an INCLUDEPICTURE picture, which
  Word's PDF shows.

Footnotes and endnotes become DOCX notes. Automatic reference marks and the
numbers inside each note use the document-wide note format and starting
value. Arabic, Roman and letter formats are supported, with Word's
lowercase-Roman default for endnotes in later Word versions. Footnotes must be
at the page bottom and endnotes at the end of the document. Documents are
rejected if they use custom reference marks, number restarts per section or
per page, sections with different note numbering, or custom separator
stories.

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

The direct reader shows an embedded OLE object, such as an Excel or Graph
chart, as the presentation picture the file stores for it: an OLE shape is a
picture frame whose `pib` names the BLIP to display (MS-ODRAW 2.2.40 and
2.3.23.5), resolved through the shape's `ExObjRefAtom` to the document's
external object list (MS-PPT 2.7.7 and 2.10.1). The object storage is never
read or activated. PowerPoint's own PDF exports show that stored picture
unchanged for embedded objects drawn as content. Icon or thumbnail aspects,
linked objects and ActiveX controls, pictures without a supported BLIP, and
pattern, texture or non-stretched picture fills are rejected instead of being
drawn without them.

PowerPoint displays GIF data that a producer stored in a PNG picture slot,
so the direct PPT reader identifies such a slot by its GIF87a/GIF89a
signature and emits it as `image/gif`; other mismatched content stays
rejected.

Pattern fills on unrotated shapes become tiled picture fills that follow
PowerPoint's own output: the 8x8 area of the stored 10x10 pattern bitmap,
one pattern pixel per point, white pixels in the fill colour and black
pixels in the background colour. Pattern fills on rotated or flipped shapes,
other pattern bitmap sizes, translucent pattern colours, background
patterns and texture fills stay rejected.

Picture colour settings follow how PowerPoint itself reads the binary
properties when it saves a binary deck as PPTX: "Black and White" becomes
DrawingML `grayscl` plus `biLevel` at 50%, and a transparent colour becomes a
`clrChange` to the same colour with zero alpha. The presentation renderer
applies these blip effects in document order for PPTX files as well.
Brightness/contrast (washout), grayscale or black-and-white alone, recolouring
and adjustments on picture fills stay rejected until Office output confirms
their rendering.

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

A modern Office-saved PPT can retain paragraph properties in the OfficeArt
`metroBlob` alternative shape XML rather than in its classic text ruler. The
direct PPT source adopts that XML under the rule described with the release
gap inventory below; see also the [controlled probe protocol](../scripts/legacy-ppt-ruler-probes.md)
before attributing differences to an implicit ruler rule.

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

Checkpoint entries below are historical. Settled behavior and its evidence
boundaries are maintained beside the production implementation; the latest
checkpoint records current task status, not a separate specification.

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
| DOC-129 | Resolve preferred cell width at TIstd | Clean sole-TDef controls now show a live D635 width effect in both TIstd orders and at varied magnitudes. This establishes preservation, not a general preferred-width/merge layout algorithm; see DOC-194 |
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
| DOC-171 | Validate the effective experimental target across property families | Implemented source-bound acquired-trace and fixed-fixture target assertions, including full FKP/Data/complex-PCD acquisition, exact style names, PJc opcodes/operands, ordering, absence and ownership. Nonzero simple PCD PRMs fail explicitly; checks do not infer Word family precedence |
| DOC-172 | Resolve cell-edit and preferred-width family precedence | Refined with clean D635 controls and unmerged repeated-TDef/TDxaCol countercontrols. Native TDxaCol effects survive the later TDef in the tested pair, contradicting naive raw-order geometry. Keep DOC-150 gated; resolve DOC-193 and DOC-194 before generalizing |


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


### Selected probe assertions and width precedence batch: 30 items

This finite batch selects DOC-26, DOC-80, DOC-129, DOC-150, DOC-163,
DOC-164, DOC-166, DOC-167, DOC-171 and DOC-172, plus the twenty slices
below. Probe checks establish acquired records and declared ownership, not an
inferred Word layout algorithm. Native width controls must remove or explicitly
account for competing cell definitions and width properties.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-173 | Bind read-only trace assertions to source hashes | Implemented exact serialized-source SHA-256 binding and read-only check-trace CLI |
| DOC-174 | Share complete FKP/Data/PCD trace acquisition | Implemented one shared acquisition session for FKP/Data and later complex PCD; unsupported nonzero simple PRMs fail explicitly |
| DOC-175 | Assert exact operand lists and absent competitors | Implemented exact framed operand arrays and empty arrays for absent competitors |
| DOC-176 | Assert ordered acquired property sequences | Implemented exact filtered acquired order, including repeated property records |
| DOC-177 | Validate paragraph versus row-mark ownership | Implemented character-boundary and top-level row ownership checks; fixed-probe marker bodies and terminators must retain table depth one |
| DOC-178 | Check direct alignment operands rather than opcode presence | Implemented exact serialized PJc opcode and operand checks, including later PCD overrides; no bidi-aware alignment claim |
| DOC-179 | Check the selected table-style identity at the measured target | Implemented exact authored style-name/istd matching, including cache-isolated names; rejects duplicate names and wrong kinds |
| DOC-180 | Expose fixed-probe marker and row assertion commands | Implemented bounded check-target CLI for the fixed eight-table marker and row inventory |
| DOC-181 | Bound assertion input, lookup, trace and output resources | Verified interval overlap/reordering, shared-cache work, trace/payload/output budgets, strict JSON and schema rejection |
| DOC-182 | Preserve historical mutation-plan validation | Verified historical mutation-plan tests remain passing; duplicate JSON keys now fail explicitly |
| DOC-183 | Certify new native controls with reusable assertions | Certified ten frozen controls with source hashes, exact edits, full ordered rows, actual TDelete absence and exact style/alignment assertions |
| DOC-184 | Isolate clean D635 before/after TIstd | Verified clean D635 before/after TIstd controls preserve the same native width effect |
| DOC-185 | Verify opposing preferred-width magnitudes | Verified varied D635 values change the visible split: left merged controls show 4500/4500 and 3000/6000 in either TIstd order |
| DOC-186 | Isolate valid TDxaCol and TDefTable redefinition order | Verified unmerged repeated-TDef/TDxaCol controls show 3000/3000/3000 in both orders. The merged pair alone was nondiscriminating; generic precedence remains unresolved |
| DOC-187 | Verify width effects across both merge positions | Verified clean D635 effects in left and right merge positions; no arbitrary width equalization rule added |
| DOC-188 | Determine supported width facts without a guessed layout rule | Established bounded native facts and rejected the proposed raw-order width expectation; general geometry remains gated |
| DOC-189 | Exercise supported acquisition facts and preserve admission gates | Verified a focused native-acquisition regression retains the TIstd gate for both unresolved repeated-TDef/TDxaCol orders |
| DOC-190 | Review ownership and resource costs of the final changes | Reviewed bounded physical interval indices, one document-bound session, sequential marker checks and aggregate cached-trace work |
| DOC-191 | Compare final independently built release models | Verified all 327 results match exactly in independently built release WASMs; production WASM bytes are identical |
| DOC-192 | Adversarially review source and native evidence | Scoped adversarial review and proportionate final checks complete; parent geometry and whole-branch architecture/browser/visual acceptance remain outstanding |

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-193 | Resolve TDxaCol versus repeated TDefTable in native acquisition | Refined by twelve varied native controls, including omitted preferred table width: explicit TDxaCol survives later same-count TDef in the tested fixed unmerged scope. DOC-216 now implements bounded native acquisition; DOC-240 and DOC-241 track remaining source and structural limits, and TIstd admission stays gated |
| DOC-194 | Resolve D635 preferred width with merges and table constraints | Refined by eight native controls separating merge ownership and preferred table width with nil TC80 and fixed layout. Primary-only model projection is regression-tested; general allocation, other TC80 and autofit remain open under DOC-217 |
| DOC-195 | Extend probe acquisition to nonzero simple PCD PRMs when required | Implemented the closed MS-DOC Prm0 mapping for acquired paragraph traces, with exact two-byte origin, recognized character-only provenance, reserved-index rejection and native alignment controls |

### Probe assertion contracts

The read-only `check-trace INPUT ASSERTIONS` command in
`scripts/legacy-doc-papx-probes.py` accepts `legacy-doc-property-trace/v1`.
Each source-hash-bound target declares a physical `fc`, `owner` (`ttp` or
`paragraph`), and nonempty `properties` mapping lowercase four-digit SPRM codes
to exact framed operand arrays. An empty array asserts absence. Optional `order`
asserts the complete acquired sequence filtered to those property codes.

The fixed-fixture `check-target INPUT ASSERTIONS` command in
`scripts/legacy-doc-table-style-probes.py` accepts
`legacy-doc-table-style-target-assertions/v1`. It requires an exact marker or
table/row selector, an exact authored `style` name, `direct_pjc` as null or an
exact `{code, operand}` object, and `acquired` property assertions. Despite the
field name, PJc checks include later complex PCD records and identify serialized
opcodes/operands only. They do not resolve bidi alignment, style-family
precedence or displayed geometry. Both commands decode documented paragraph Prm0 values,
retain recognized character Prm0 records as unapplied paragraph-trace provenance, and reject
reserved indices and ambiguous ownership. A synthesized Prl identifies its
actual two-byte Pcd.Prm source range explicitly.

### Probe assertions and width precedence acceptance

This thirty-item checkpoint delivers read-only probe assertions and a focused
admission-gate regression. The eight fresh merged controls and two discriminating
unmerged controls were exported by local Word, with exact source edits/hashes
and 130 target checks. Independent PDF border measurements and visual inspection
support the bounded width facts above. The saved DOCX output is corroborating
diagnostic evidence. No production width algorithm or admission gate changed.

The merged TDxaCol pair initially hid the distinction being tested. Unmerged
countercontrols exposed a mismatch with the current raw-order expectation.
Adversarial review rejected the proposed width assertion and retained an explicit
gate test instead. Clean varied D635 controls establish a live width effect
across TIstd, but do not establish a general merge/preferred-width allocation
rule. Existing DOC-26, DOC-80, DOC-129, DOC-150, DOC-163, DOC-164, DOC-166 and
DOC-167 remain partial within these limits; DOC-171 is implemented in the stated
probe scope and DOC-172 has new counterevidence. DOC-193 through DOC-195 record
the additional unresolved work.

Fresh checks passed 879 unit tests and nine integration tests, with six existing
ignores; feature-off compilation, 44 target-probe tests, 18 PAPX-probe tests,
33 unchanged renderer-border tests, static layout boundaries and diff checks
also passed. All 327 identical inputs match in independently built release
WASMs, including errors and streamed models. The release WASM bytes themselves
are identical. Four of 59 private inputs remain admitted; equal rejections are
not visual fidelity.

Final review corrected incomplete later-PCD checks, physical-range ambiguity,
marker/terminator ownership and cached-trace work amplification. No renderer,
shared-model, worker, XLS or PPT behavior changed. Whole-branch architecture,
browser and visual acceptance remains outstanding. This checkpoint does not
complete table-style support or start another batch automatically.

### Selected cell-width and simple-PRM follow-up: about 30 items

This finite batch selects DOC-26, DOC-80, DOC-129, DOC-150, DOC-163,
DOC-164, DOC-172 and DOC-193 through DOC-195, plus the twenty bounded
slices below. Existing parent items remain partial until their implementation
and native evidence support admission; inspecting an unresolved dependency is
not completion of that feature. Native controls keep one intended variable per
comparison and are hash-bound before export.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-196 | Decode the documented Prm0 index table | Implemented the closed MS-DOC 2.9.215 mapping; PJc uses 0x2461 |
| DOC-197 | Separate paragraph Prm0 from character/no-op provenance | Implemented paragraph application, character-only unapplied provenance and zero no-op |
| DOC-198 | Apply simple PCD overrides to marker and row ownership checks | Verified later PJc and structural Prm0 flags participate in ownership and target assertions |
| DOC-199 | Retain truthful two-byte PCD provenance for expanded Prls | Implemented explicit synthesized-Prl framing with the actual two-byte Pcd.Prm origin |
| DOC-200 | Bound simple-PRM decoding and preserve historical probe contracts | Verified reserved-index, work/payload bounds and historical probe tests; inspection and mutation contracts remain unchanged |
| DOC-201 | Check simple-PRM assertions against native controls | Verified native left/right Prm0 alignment and unchanged paragraph alignment for character-only bold in three controls |
| DOC-202 | Vary TDxaCol cell ranges around repeated TDefTable | Verified middle-cell TDxaCol before/after repeated TDef in matched unmerged controls |
| DOC-203 | Reverse TDxaCol magnitudes in matched native controls | Verified 1500/6000 width variations; no single-sample scale factor inferred |
| DOC-204 | Use distinct first and second TDefTable definitions | Verified unequal definitions with second-definition outer widths and surviving TDxaCol middle width in both orders |
| DOC-205 | Test repeated width operations within the same property family | Verified 6000 then 1500 matches single final 1500 within the tested TDxaCol family |
| DOC-206 | Certify unmerged geometry and absence of competing width records | Certified all width controls with complete row traces, exact style identity and absent competing insert/delete/merge/width properties as applicable |
| DOC-207 | Review native TDxaCol evidence independently | Verified twelve native TDxaCol controls independently; four omit preferred table width to isolate physical acquisition |
| DOC-208 | Compare D635 with and without merging | Verified D635 differs between unmerged and merged layouts while source cell definitions remain fixed |
| DOC-209 | Isolate primary versus continuation-cell preferred widths | Verified continuation-only D635 1500/6000 has no visible effect on the merged region; primary width ownership is covered by focused projection tests |
| DOC-210 | Vary preferred table width with fixed cell preferences | Verified omitted/6000 preferred table width yields 3000/3000 while 9000 yields 4500/4500 for the same primary preference |
| DOC-211 | Hold TC80 and autofit constant across width controls | Verified controls hold nil TC80 and fixed layout constant; broader TC80/autofit behavior remains unresolved |
| DOC-212 | Apply only supported behavior and retain unresolved admission gates | Retained existing geometry/admission limits and added a primary-source preference projection regression; no inferred width allocation rule |
| DOC-213 | Compare final independent release models | Verified all 350 inputs match in independently built release WASMs, including errors and streamed models; release WASM bytes are identical |
| DOC-214 | Review final source and native evidence adversarially | Accepted the diagnostic and focused-test diff after specification, native-evidence, resource, ownership and regression review; broader architecture and visual acceptance remains open |
| DOC-215 | Record bounded conclusions and new unresolved dependencies | Recorded distinct acquisition and normalized-layout dependencies as DOC-216 and DOC-217 |

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-216 | Model persistent TDxaCol overrides across cell redefinition | Partially implemented: call-local native acquisition preserves explicit ranges across same-count definitions in the fixed LTR/default-TC80 profile. Full-CFB and native controls cover the admitted rule. Changed count, insert/delete, preferred-width and cross-source competition remain gated; parent style admission remains open |
| DOC-217 | Resolve normalized preferred-width allocation independently | Open: ten additional native controls isolate primary/continuation TC80, preferred totals and AutoFit. The model preserves these independent facts; native allocation still differs in constraint-dependent ways. Reconcile the existing layout algorithm without a guessed scale or rounding constant |

### Cell-width and simple-PRM evidence checkpoint

The native set contains 23 new inputs: twelve TDxaCol controls, eight D635
controls and three Prm0 controls. Exact mutation and acquired-target assertions
preceded export. One initial Word automation timeout produced a delayed PDF;
a successful separate re-export matched its pixels. All accepted native
results have complete exports, independent PDF measurements and corroborating
saved-DOCX diagnostics.

With preferred table width omitted, middle-cell TDxaCol 1500 survives either
side of a later identical TDef, giving native widths 1500/1500/3000. With a
different second definition, its outer widths survive while TDxaCol supplies
the middle 3000, giving 2500/3000/5000 in both orders. This separates the
acquisition mismatch from normalization. The tested scope remains fixed,
unmerged rows without insertion/deletion or changing cell counts.

For D635, only the primary source cell supplies the merged region's projected
preference, consistent with MS-DOC 2.6.3 TMerge. Changing only a continuation
cell preference between 1500 and 6000 does not change its native merged
geometry. Preferred table width remains a separate constraint. The added
projection test protects this ownership distinction without treating raw
physical grid edges as the final native layout.

Prm0 acquisition follows MS-DOC 2.4.6.1 and 2.9.215. Native left/right controls
confirm the mapped PJc override; a character-only bold control preserves
paragraph alignment. This extends diagnostic paragraph acquisition and does
not claim complete direct-character acquisition.

Fresh final checks passed 880 unit and nine integration tests, with six existing
ignores; feature-off compilation, 22 PAPX-probe and 44 target-probe tests,
33 unchanged renderer-border tests, static layout boundaries and diff checks
also passed. Independent release builds match all 350 inputs exactly and
produce identical WASM bytes. Four of 59 private inputs remain admitted; equal
rejections do not establish display fidelity.

Adversarial review accepted the closed Prm0 mapping, explicit synthetic
provenance and bounded trace accounting. The Rust change is a projection test
only. No production width algorithm, admission gate, renderer, shared model,
worker, XLS or PPT behavior changed. Whole-branch architecture, browser and
visual acceptance remains outstanding. DOC-216 and DOC-217 separate the two
remaining width problems; this finite checkpoint does not complete their parent
features or automatically start another batch.

### Selected persistent-width acquisition checkpoint: about 30 items

This finite batch selects existing DOC-26, DOC-80, DOC-129, DOC-150,
DOC-172, DOC-193, DOC-216 and DOC-217, together with the twenty-two bounded
slices below. Parent items remain partial until their actual admission and
native acceptance requirements are satisfied. Each slice reports its own
implementation or evidence boundary; a completed experiment does not complete
the corresponding format feature.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-218 | Define explicit TDxaCol acquisition separately from TDef defaults | Reviewed MS-DOC 2.6.3 creation defaults and TDxaCol overrides against native order controls; no general raw-Prl-order assumption |
| DOC-219 | Keep native geometry acquisition separate from XML compatibility | Implemented call-local native state; shared Row application and XML compatibility remain unchanged |
| DOC-220 | Preserve explicit widths across same-count cell redefinition | Implemented bounded same-count persistence in the native production acquisition path |
| DOC-221 | Compose full valid cell ranges with last-write ownership | Implemented per-cell last-write composition for every valid range, with first/last/disjoint/overlap regressions |
| DOC-222 | Contain unresolved merge interactions with repeated definitions | Repeated-definition merge competition gates, including late changes; simple pre-existing TDxa-plus-merge remains unchanged |
| DOC-223 | Contain unresolved TC80 width interactions | Repeated-definition TC80 and direct preferred-width competition gates; unrelated D635-only sequences retain existing behavior |
| DOC-224 | Gate unresolved insert/delete combinations | Insert/delete competition requiring persistence remains gated |
| DOC-225 | Gate unresolved cell-count changes | Changed-count competition remains gated rather than remapping cell ownership by inference |
| DOC-226 | Preserve source boundaries and gate unresolved cross-source changes | Direct-PAPX and later PCD competition remains gated; no global state or source carry by inference |
| DOC-227 | Protect the existing compatibility acquisition path | Verified a discriminating XML raw-order regression, separate from native acquisition |
| DOC-228 | Bound temporary width state and repeated acquisition work | Verified fixed call-local 63-slot storage, full-range boundary tests and strict range validation; no retained Row/cache growth |
| DOC-229 | Verify first/last TDxaCol ranges in native Word | Verified native first/last pairs across unequal same-count definitions |
| DOC-230 | Verify disjoint/overlapping TDxaCol ranges in native Word | Verified native disjoint and both overlapping orders; later TDxaCol wins within the shared range |
| DOC-231 | Isolate primary/continuation TC80 width preferences | Verified raw TC80 primary/continuation ownership in native controls and focused Writer projection tests |
| DOC-232 | Isolate fixed versus automatic table layout | Verified fixed/AutoFit native controls; content-dependent differences remain a separate layout problem |
| DOC-233 | Isolate preferred table totals from cell preferences | Verified live preferred totals with nil/primary/continuation cell preferences; no normalization formula inferred |
| DOC-234 | Audit preferred-width facts through the existing model and layout | Reviewed parser-to-model-to-layout ownership and preserved independent physical-grid, cell-preference and table-constraint facts |
| DOC-235 | Check shared model and main/worker boundaries | Verified full-CFB direct-model wiring, cell text and trailing paragraph preservation, plus unchanged source-owner/runtime tests |
| DOC-236 | Compare independently built final release models | Verified all 376 inputs match across separate release builds, including errors and streamed models; four of 59 private inputs admitted |
| DOC-237 | Recheck architecture gates and unresolved branch-level acceptance | Rechecked structural, compatibility and API gates; final package/type checks and branch-level scope recorded below |
| DOC-238 | Review the final diff and native evidence adversarially | Reviewed state-machine ordering, native evidence, bounds, module separation and actual admission; corrected bidi and false-rejection findings |
| DOC-239 | Record bounded results and additional dependencies | Recorded bounded implementation and remaining source/structural/layout dependencies |

### Persistent-width acquisition acceptance

Native DOC acquisition now retains explicit TDxaCol changes separately from
TDefTable creation defaults. A later definition with the same cell count
preserves those changes within one direct-property source. The implementation
uses a fixed 63-slot array local to acquisition; it does not enlarge retained
rows or caches. XML compatibility acquisition keeps its previous behavior.

The supported persistence profile is fixed, left-to-right, unmerged cells with
default TC80 flags and no competing cell preference. Nonzero TC80 flags are a
coverage boundary, not a claim of malformed input. Cell-count changes,
insertion/deletion, preferred-width and cross-source interactions remain gated
when persistence would require an unsupported inference. A later incompatible
change also gates an already-used persistence result. Existing simple
noncompeting sequences and the independent TIstd admission gate are preserved.

Twenty-six new native controls cover ten varied range/order inputs, ten
TC80/AutoFit/total-width inputs and six controls with TIstd removed. Exact
source edits, complete acquired row chains and source hashes preceded export.
The five range pairs match in native PDF geometry and pixels; reversing
overlapping TDxaCol operations changes the shared range as expected.

Removing TIstd also removes the test table's inherited borders. In those six
controls, native PDF marker starts independently confirm equal first, middle
and overlapping-range advances before and after the later definition. Saved
DOCX grids corroborate acquisition; they do not independently prove the absent
rightmost visible edge. These controls close the omission-equivalence question
within this bounded scope without claiming whole-document fidelity.

Raw TC80 controls confirm that the primary source preference affects a merged
region while continuation preferences are not promoted. Preferred table total
and AutoFit remain independent facts. AutoFit changes the observed allocation
for one nil-preference constraint case; this does not justify a general scaling
or rounding rule. A focused Writer test protects the projection, and a
full-CFB regression verifies corrected grid widths, cell text, trailing content
and late unsupported-state rejection through the actual direct-model path.

Adversarial review corrected an early bidi setting that reopened eligibility,
a late incompatible-state admission gap, and unnecessary D635-only rejection.
It also replaced a nondiscriminating compatibility test and required native
TIstd-omission controls. No renderer formula, sample-specific count/range
branch, empirical constant or shared-model extension was introduced.

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-240 | Establish native precedence across direct PAPX and complex PCD | Open: build valid PTableProps-based cross-source controls with certified CLX/PCD ownership; raw table records in PCD are not sufficient evidence |
| DOC-241 | Establish width ownership across structural cell changes | Open: vary changed counts and insertion/deletion around repeated definitions with valid cell markers; do not infer index remapping from fixed-count controls |
| DOC-242 | Extend persistence beyond default TC80 and horizontal LTR cells | Open: isolate relevant TC80 flags and preferred-width interactions before widening the current coverage boundary; coordinate final allocation with DOC-217 |

Fresh final checks passed 891 unit and nine integration tests, with six existing
ignores; feature-off compilation, 22 PAPX-probe and 44 target-probe tests, and
59 source-owner/runtime/renderer tests also passed. Independently rebuilt
release WASMs match all 376 identical inputs, including streamed models and
errors. Four of 59 private inputs remain admitted; unchanged rejections do not
establish display fidelity.

All deterministic architecture commands passed: final layout boundaries,
boundary tests, compatibility evidence tests, public API tests, package builds
and type checking. Missing generated WASM was rebuilt before accepting these
results. Whole-branch semantic architecture, browser and visual acceptance
remains open. This finite checkpoint fixes native acquisition within its stated
scope; it does not complete table-style support or begin another batch.

### Selected source and structural acquisition checkpoint: about 30 items

This finite batch advances DOC-193, DOC-216, DOC-217 and DOC-240 through
DOC-242, together with the twenty-four bounded slices below. Parent features
remain partial. An experiment or admission check does not complete the
corresponding table-style or layout feature.

| Item | Scope | Status |
| --- | --- | --- |
| DOC-243 | Select paragraph properties from complex PCD sources | Implemented the MS-DOC 2.4.6.1 step 5 filter in native acquisition and the diagnostic probe |
| DOC-244 | Preserve referenced Data traversal after source selection | Verified paragraph indirection preserves its referenced property records; unresolved native table effects remain gated |
| DOC-245 | Evaluate PHugePapx after selecting eligible source records | Verified ignored non-paragraph records do not consume the first eligible position; preceding paragraph records do |
| DOC-246 | Separate raw and interpreted diagnostic caches | Implemented source-tagged cache keys; raw inspection remains serialized provenance |
| DOC-247 | Apply source selection to mutation-target validation | Verified raw table records do not shadow paragraph sources, while interpreted indirect records remain in the acquisition trace |
| DOC-248 | Keep native paragraph and table acquisition consistent | Connected both paths to the same bounded traversal; XML compatibility retains its existing policy |
| DOC-249 | Preserve unsupported vertical-merge results | Fixed intercepted TVertMerge records whose unsupported result was previously discarded |
| DOC-250 | Bound vertical-alignment interactions with persistent widths | Added admission checks for nondefault TVertAlign before, between and after width persistence |
| DOC-251 | Preserve supported alignment no-ops | Verified explicit default alignment and empty ranges; malformed and reserved operands remain rejected |
| DOC-252 | Verify actual binary model admission | Added full-CFB regressions for malformed vertical merge, unsupported alignment combinations, and raw versus redirected complex-PCD properties |
| DOC-253 | Test insertion followed by removal of the inserted definition | Native paired controls retain the earlier width on the surviving original cell |
| DOC-254 | Test deletion and reinsertion of the width-bearing definition | Native paired controls distinguish the deleted override from a later width assignment |
| DOC-255 | Test movement to a higher cell index | Native leading-insertion controls retain the override on the surviving original definition |
| DOC-256 | Test movement to a lower cell index | Native leading-deletion controls retain the override on the surviving original definition |
| DOC-257 | Test repeated definitions with increasing and decreasing counts | Native three-to-four-to-three controls retain the tested override; general count-change admission remains closed |
| DOC-258 | Test a shorter definition followed by insertion | Native three-to-two-plus-insert controls retain the tested surviving override; removed-slot behavior remains a dependency |
| DOC-259 | Test direct vertical alignment around cell redefinition | Native center/bottom controls preserve alignment in all three tested positions; no alignment cascade implemented |
| DOC-260 | Compare explicit top alignment with omission | Native pixels and glyph positions match with the same minimum-height constraint |
| DOC-261 | Construct valid cross-source property references | Certified single-character PCD ownership, CLX relocation, Data references and exact binary changes |
| DOC-262 | Check native activation before inferring precedence | Direct complex-PCD paragraph alignment is active; wrapped table-width effects remain unresolved |
| DOC-263 | Verify shared traversal failure containment | Covered filtered-record work, cycles, ignored tails, malformed framing and feature-off compilation |
| DOC-264 | Compare separate release builds on the extended corpus | Final comparison and admission counts are recorded below |
| DOC-265 | Recheck package and architecture boundaries | Deterministic checks are recorded below; whole-branch semantic and visual acceptance remains open |
| DOC-266 | Review bounded claims and record additional dependencies | Final diff and native evidence reviewed separately from broad feature completion |

### Source selection and unsupported-state acceptance

Native complex-PCD acquisition selects top-level paragraph SPRMs before
processing paragraph indirection. Non-paragraph records remain visible as
unapplied diagnostic provenance and still consume framing/work budgets.
Filtering stops at referenced Data: the shared traversal preserves those
records, its cycle/depth limits, and the rule that indirection discards the
remaining array tail. Native paragraph and table readers use the same policy;
style, direct-PAPX and XML compatibility paths retain their existing behavior.
The diagnostic probe applies source selection consistently to acquired traces
and mutation-target validation, without rewriting raw inspection.

This selection rule does not prove how indirect table properties from a
complex PCD affect Word's table model. Four native controls using valid
PTableProps references show only the direct-PAPX widths in the target row.
Two independent paragraph-alignment controls confirm that complex-PCD
activation works. These results isolate an unresolved table-acquisition
question; they do not establish a universal precedence rule. Native table
properties reached through that complex-PCD source are retained for diagnosis
but close admission. Ordinary complex-PCD paragraph properties remain enabled.

Intercepted unsupported TVertMerge and TVertAlign records can no longer bypass
the native table-format gate. Direct nondefault vertical alignment also closes
admission when it competes with the bounded persistent-width profile. Eight
native controls, with identical minimum row height and no table-style
selection, show that center/bottom alignment survives the tested later
same-count definition. Widths alone would therefore be an incomplete result.
Explicit top and empty-range no-ops retain existing supported behavior.

Twelve structural controls independently vary operations around a width
assignment. Their native borders show that insertion/deletion can move a
surviving definition's override to another index, while deletion and reinsertion
of the target loses that override. Two changed-count sequences retain the
tested surviving override. These controls retain table-style selection for
visible borders. They do not justify admitting arbitrary changed counts,
removed/reintroduced slots, multiple-cell operations, merges, preferred-width
constraints or other source combinations.

No renderer, shared model, public option or worker protocol changed. No
empirical width formula, sample-specific branch or extra retained row/cache
state was introduced. Broader table-style support, DOC-217 allocation and
whole-branch architecture/visual acceptance remain incomplete.

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-267 | Model explicit vertical-alignment persistence | Implemented for the bounded same-count, fixed horizontal LTR profile; explicit resets, overlapping ranges and TC80 descriptor conflicts are covered |
| DOC-268 | Generalize surviving-cell width ownership | Open: test removed/reintroduced slots, range boundaries, multiple-cell edits and style omission before extending DOC-241 |
| DOC-269 | Resolve indirect table-property application from complex PCD | Direct replacement suppression is resolved and implemented; the remaining reachable-PCD question is tracked by DOC-272 |
| DOC-270 | Revisit conservative structural no-op admission | Implemented for unmerged TSplit, clearing an already-unmerged vertical cell, and empty TMerge ranges; other merge interactions remain gated |

Fresh final checks passed 897 unit and nine integration tests, with six existing
ignores; feature-off compilation, 26 PAPX-probe and 44 target-probe tests, and
59 source-owner/runtime/renderer tests also passed. Separate release builds
match all 402 identical inputs, including streamed models and errors. A final
test-only full-CFB addition was followed by another all-features test run and
release build; its WASM is byte-identical to the compared candidate. Four of
59 private inputs remain admitted; unchanged rejections do not establish
display fidelity.

All deterministic architecture checks, package builds and type checking passed
with freshly generated WASM. Independent final review accepted the scoped
source-selection and admission fixes after the full-CFB coverage gap was
closed. The 26 native controls provide bounded evidence for the open alignment,
structural and source-acquisition questions. Whole-branch semantic, browser
and visual acceptance remains open. This finite checkpoint does not complete
table-style support or start another batch.

### Concatenated property arrays and cell alignment checkpoint

Native acquisition now follows the single logical PAPX-plus-Prm array required
by MS-DOC 2.4.6.1. A direct replacement discards the appended tail, and
PHugePapx first-position eligibility spans that boundary. The diagnostic probe
uses the same rule for acquisition and mutation validation. This resolves the
earlier direct-PAPX/PCD observation without inventing a precedence exception.
The implementation and normative rationale are in `doc/sprm.rs` and
`doc/formatting.rs`.

Bounded direct alignment persistence and proven structural no-ops are
implemented in `doc/table/geometry.rs`. Its adjacent comments distinguish the
specification, native observations and intentionally unsupported combinations.
Twenty-three native controls covered resets, ranges, descriptor activation,
width-independent alignment, no-ops and structural counterexamples. Parent
table-style and general structural support remain incomplete.

The duplicated Formatting geometry matrix and its remaining malformed-merge
case were removed. Geometry boundary tests and
`full_cfb_variable_vertical_merge_length_keeps_the_native_gate_closed` retain
that protection; full-CFB alignment and property-array tests cover the actual
model path. Exploratory native matrices are not permanent regression fixtures.

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-271 | Distinguish TDef count changes from explicit deletion | Open: temporarily absent slots retained widths in shrink/regrow controls, while explicit deletion/reinsertion cleared replacement widths; general structural acquisition remains gated |
| DOC-272 | Determine reachable complex-PCD table indirection | Open: validate table properties reached from an appended PCD when no earlier direct replacement suppresses it; keep the existing admission gate |

Fresh verification passed 901 unit and nine integration tests (six existing
ignores), 72 probe tests, feature-off compilation and the layout boundary check.
Separate release WASM builds produced identical outcomes for 96 selected inputs,
including all 59 private inputs; private admission remains four of 59. Rejected
inputs do not establish rendering fidelity. Final adversarial review accepted
this bounded change; whole-branch semantic, browser and visual acceptance remain
open.

### Complex-PCD property acquisition checkpoint

Nineteen native Word exports (Word 16.113, macOS; one unmodified baseline) replaced the earlier
single-array reading of direct PAPX plus Pcd.Prm. Each control changed one
Word-saved row mark or table-cell paragraph. It was compared through Word's
PDF and saved DOCX, and carried a piece alignment witness proving the split
pieces were active. `doc/sprm.rs` now records and implements the observed
rules:

- a direct sprmPTableProps is followed only as the first Prl of its array;
- piece properties, both Prm0 and Prm1, apply after the whole direct chain,
  including PrcData reached through a direct first-position redirect;
- piece sprmPTableProps and sprmPHugePapx are ignored rather than followed,
  even as the paragraph's first Prl, and later piece Prls still apply;
- top-level piece table SPRMs apply to the row mark together with paragraph
  SPRMs.

DOC-269's suppression rule is therefore withdrawn, and DOC-272 is resolved:
reachable complex-PCD table indirection does not occur. The diagnostic PAPX
probe follows the same rules. Piece table SPRMs use the existing per-code
admission gates.

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-273 | Cross-source row definition after explicit widths | Open: one control showed a piece TDefTable replacing earlier direct-chain TDxaCol widths; the geometry source boundary stays gated |
| TOOL-1 | Office export restoration | Open: `scripts/legacy-office-export.applescript` cannot restore Word/Excel settings when `open` returns no value; it also adopts an unrestored ForceDisable baseline |
| LEGACY-OOXML | Remove the OOXML-generation path | Open: legacy support is unreleased, so delete the byte converter, its WASM/TS entry points and XML writers instead of deprecating them; direct paths must not depend on them |

### Direct-model table-style admission checkpoint

The direct DOC model now admits table-style selection instead of rejecting
every table that carries sprmTIstd or sprmTTlp. The table-style profile
(TAPX/PAPX/CHPX, conditional selection, borders, margins, shading) keeps its
own per-property gates; the XML conversion path is unchanged. The decisions
and their evidence are recorded next to the code in
`doc/table/native_admission.rs`, `doc/table/position.rs`,
`doc/direct_model/tables.rs` and `doc/direct_model/story/borders.rs` and
`preferences.rs`.

- sprmTIstd is admitted when every table property authored before it in the
  row chain is on the MS-DOC 2.6.3 preserved list, is reset by an implemented
  applier, or is cell geometry that the earlier Word controls (DOC-44,
  DOC-129, DOC-184, DOC-193) show surviving the selection. Any other earlier
  table property keeps the row gated.
- sprmTTlp feeds conditional selection; its itl is historical metadata.
  sprmTRsid has no presentation semantics.
- sprmTWidthIndent (direct or inherited from the style) is validated but not
  projected: TDxaLeft/TDxaGapHalf/TDefTable define the physical origin. In
  two Word PDF exports of left-to-right documents whose preference differs
  from that origin, the borders sit at the physical origin; the paired OOXML
  documents carry the preference as `w:tblInd`. RTL rows are admitted only
  when the effective preference equals the origin.
- sprmTWidthBefore/sprmTWidthAfter (direct, or the default style's required
  zero width-before) are admitted only when ftsNil or equal to the physical
  leading/trailing grid width that the projection emits.
- sprmTFCellNoWrap is admitted only for cells with an ftsDxa preferred
  width, where MS-DOC 2.9.28 says it is ignored.
- Direct NilBrc borders are projected as explicit no-border edges under a
  table style, and Nil diagonals as absent diagonals (MS-DOC 2.9.20,
  2.9.157). Style borders on RTL rows, TC80/Brc80 borders with styles,
  repeated TIstd border resets and drawn diagonals remain gated.
- Main-story tables with nondefault position or wrapping properties
  (MS-DOC 2.6.3, 2.7.13) leave the ordinary flow as floating tables.
- Table paragraph frame properties that mirror the enclosing depth-1 row's
  table position are dropped: MS-DOC 2.4.3 consults them only when no row
  carries table positioning, and the paired OOXML documents keep only the
  table's tblpPr (nested tables inside the positioned table repeat the same
  frame values). A mirror has the same anchors, X/Y, no-overlap flag,
  automatic size and around-wrapping, and frame distances equal to the
  table's left/top distances. Framed table paragraphs whose outer row is not
  positioned remain gated (legacy framed tables have no specified layout).
- Dop2000 Copts.fDontAdjustLineHeightInTable (MS-DOC 2.7.13, Dop offset
  512 bit 3) is projected as the inverse `adjustLineHeightInTable`, so the
  section line grid applies inside table cells as in Word. The decoded bit
  is the inverse of the paired OOXML element in all 57 comparable private
  inputs.

Separate release-candidate builds were not compared for this checkpoint.
The local private census admits ten of 59 DOC inputs (previously four); the
positioned-table and border-interaction gates no longer fire for any input.
Admission is not visual fidelity.

| Additional item | Scope | Status |
| --- | --- | --- |
| DOC-TBL-1 | Replacement of other pre-TIstd table properties | Open: needs Word controls that author e.g. TVertAlign, TSetBrc80, TMerge or TCellFHideMark before TIstd and compare with the same record after it |
| DOC-TBL-2 | RTL preferred indent | Open: vary sprmTWidthIndent against TDxaLeft/TDxaGapHalf in right-to-left tables (styled and unstyled) and measure the border position in Word's PDF |
| DOC-TBL-3 | Width-before/after disagreement | Open: vary sprmTWidthBefore/After against the physical leading/trailing grid slot (fixed and AutoFit) and measure the row edges |
| DOC-TBL-4 | Positioned-table exceptions | Open: controls for sprmTDyaAbs 0 (inline) and for left/zero X with zero Y and column/margin anchors (the MS-OI29500 2.1.162 counterpart), plus positioned tables in headers, footers and notes |
| DOC-TBL-5 | Hide-mark, cell text flow, no-wrap | Open: the shared cell model has no hideMark, cell text direction or no-wrap facts; MS-DOC 2.6.3 (all cells empty) and ECMA-376 17.4.21 (per-cell end mark) describe hideMark differently, so Word controls are needed before a model capability is designed |
| DOC-TBL-7 | Legacy framed tables | Open: a table whose cell paragraphs carry frame properties but whose rows carry no table positioning; build Word controls varying the frame anchors/offsets on the first-cell paragraph only versus every cell paragraph, and compare the table placement |
| DOC-TBL-6 | Compatibility shading without a table style | Open: current Word's use of sprmTDefTableShd in rows without sprmTIstd, which the specification says style-capable readers ignore |

### Local direct-render survey

`packages/{docx,pptx,xlsx}/tests/visual/legacy-corpus.spec.ts` render each
local private legacy sample through its direct source. Each sample is written
beside its same-named Office PDF export as paired PNGs and a summary. The
survey reports only: it never gates, updates references, or generates OOXML.
Run it with an output directory outside the checkout:

```bash
LEGACY_CORPUS=1 LEGACY_CORPUS_OUT=/tmp/legacy-survey VRT_PRIVATE_CORPUS=1 \
  pnpm --filter @silurus/ooxml-pptx exec playwright test \
  --config playwright.config.ts --project=chrome legacy-corpus.spec.ts
```

Pixel percentages are only a triage signal. For example, a slide can score
above 95% while a chart or autoshape is missing, so review the pairs visually.

### Release gap inventory

The direct-render survey was reviewed visually against the Office PDF
exports, and the findings are grouped here. Counts are local private samples
affected. They record open work, not supported behavior. Omission is
acceptable only where the caller did not enable an opt-in module such as
chartex. Every other gap below is unimplemented behavior or a bug that must
be closed before an experimental release.

| Area | Gap | Samples |
| --- | --- | --- |
| XLS | ~~BIFF8 embedded charts are not projected into `ChartModel`~~ Projected (89f3db02, f7dcb0fa, 06da6a7a); chart sheets and the items below remain | 127 of 139 |
| XLS | EMF pictures written by GDI+ (EMF+ comment records, short EMR_EOF and a record count off by one) are rejected by the passive validator. Core rendering also treats EMF+ as out of scope, so shared EMF+ drawing is needed for all formats | 3 |
| XLS | Chart and picture anchors need the Normal font's digit width. The browser default measures only an installed face, so they are omitted when Office fonts such as Calibri are not installed; shared reference font metrics are needed | all with drawings |
| XLS | Chart text omits TextPropsStream (its checksum is not implemented), Fbi font autoscaling, the outline Excel draws around inverted negative points, plot-area layout, drop/high-low lines and 3-D walls | most chart samples |
| XLS | ~~Extended colors (XFExt theme/tint) fall back to palette approximations~~ Resolved: tints (b09eae6a) and theme 0-3 in Excel's lt1/dk1/lt2/dk2 order (0adbc794) | about 6 |
| XLS | Table (ListObject) styles, conditional-format data bars/icons and pivot styling are absent | about 5 |
| XLS | Formula text is not decompiled from Ptg tokens, so volatile functions are not recalculated as Excel does at export | 2 |
| XLS | Clip-art pictures, text boxes, strikethrough and one vertical merge are missing | 1 to 3 each |
| XLS | The direct reader rejects, instead of omitting, drawn objects it does not project: EMF+-only pictures, shapes and text boxes, grouped shapes, chart and macro sheets, and drawings whose OfficeArt data continues after an Obj record | 19 of 139 |
| PPT | ~~Only seven MS-ODRAW shape types map to presets~~ 100+ shape types map as PowerPoint converts them, with evidenced adjust formulas (officeart::preset); adjusted callout2/3 families, arrow callouts, curved arrows, ribbons and tall cubes/hexagons/parallelograms still fail closed | several |
| PPT | ~~Native/OLE charts are missing~~ Resolved: embedded OLE objects show their stored presentation picture (bfc835d3) | 3 |
| PPT | ~~Rotation by multiples of 90 degrees and combined flips use the wrong bounds or order~~ Resolved from the 120-case PowerPoint control (aa9dc5c1) | 1 |
| PPT | ~~Slide gradient backgrounds~~ linear/scaled/two-colour/translucent shades resolved (95b74d19); path (5, 6) and title (8) shades now fail closed. Bullets resolved through master levels; letter spacing, shrink-to-fit and per-paragraph indents come from adopted alternative shape XML where it agrees with the binary | several |
| PPT | ~~Gradients on rotated shapes (or inside rotated/flipped groups) are replaced by the solid fill colour~~ Resolved (ef41f03a) | several |
| PPT | Custom geometry with per-path fill/stroke flags is rejected; the PPTX model has no per-path `fill`/`stroke` (ECMA-376 §20.1.9.15), a generic PPTX gap | 1 |
| PPT | ~~Unmapped shape types are dropped silently~~ Now rejected | several |
| PPT | Picture brightness/contrast (washout), pattern/texture fills, and OLE icons, links and controls are rejected | several |
| PPT | Implicit paragraph margin/indent and percentage spacing are rejected | 12 of 34 load failures |
| DOC | 55 of 59 samples are rejected (formatting, notes, fields, positioned tables, drawings, header pictures, non-PNG/JPEG images, list ancestry, FIB version, language ID) | 55 |
| DOC | Picture washout/brightness and space-before after a page break differ from Word | 2 |

PowerPoint 2007+ also stores a `metroBlob` (MS-ODRAW 2.3.4.41, an OPC
package with the shape's DrawingML, `drs/shapexml.xml`) on most shapes. The
specification says it SHOULD be ignored; implementation note 32 says Office
2007 and 2010 use it. PowerPoint 16 controls settle how it is used:

- With only the alternative XML edited (binary byte-identical, the package
  rebuilt at the same length), PowerPoint's PDF renders the edited XML: a card
  fill changed in the XML only is drawn in the new color, and character
  spacing raised in the XML only is drawn raised.
- With a shape's binary adjust and fill edited but its alternative XML kept,
  PowerPoint's PDF follows the binary, identically to a copy whose XML was
  removed.
- A corpus deck whose alternative states 10.5 pt text over a binary 10 pt run
  renders without the alternative's other run properties, while decks whose
  sizes agree render them.

The direct PPT source therefore adopts the alternative XML for a shape only
when it agrees with the binary shape on what both record and the direct model
can compare: the same preset (and adjust values within master-unit rounding)
or custom geometry on both sides, the same untransformed position, size,
rotation and flips (group children in their unscaled child space), the same
solid fill color when both have one, the same run font size, bold and italic
where both state them, and the same paragraph, run and line-break structure
at equal UTF-16 lengths. The XML's text characters are masked, so the adopted
shape takes its characters, transform and identifier from the binary. The
XML is parsed by the ordinary PPTX shape parser, resolving theme references
against the master's round-trip theme and color map; no OOXML is generated.
Placeholders (no layout context), shapes whose XML references relationships,
math runs, non-shape parts (pictures, groups, connectors, SmartArt), and any
parse, resource or budget failure keep the binary projection. The downrev
checksums beside the XML are not recomputable (they are not checksums of the
binary records), so structural agreement stands in for them; a binary edit
that preserves every compared property would still adopt a stale XML.
Glyph shadow and emboss, which the binary cannot express, reject a shape only
when its alternative XML is not adopted.
