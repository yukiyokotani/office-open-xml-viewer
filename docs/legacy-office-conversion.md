# Opt-in legacy Office conversion

The viewer can normalize legacy binary Office bytes before its existing OOXML
parser runs:

- `.doc` to macro-free `.docx`
- `.xls` to macro-free `.xlsx`
- `.ppt` to macro-free `.pptx`

The opt-in `@silurus/ooxml/legacy-conversion` entry contains both a purpose-built
local WASM converter and the implementation-neutral adapter API. Ordinary DOCX,
XLSX, and PPTX entry points do not import, fetch, initialize, or retain the
converter Worker or its WASM. If no converter is supplied, legacy input
continues to reject with `OoxmlError.code === 'legacy-binary-format'`.

The renderer remains OOXML-only. A successful conversion enters exactly the
same parser, model, layout, and Canvas renderer as a native OOXML package.

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
| DOC | CFB Word 97-2003 documents with a readable main-story CLX piece table | main-story text, paragraphs, tabs, custom tab stops and document-wide default tab interval, line/page/column breaks, displayed field results, font names and explicit sizes, paragraph-style character defaults, character styles and direct bold/italic/underline/strike/caps/color/spacing properties, paragraph alignment/indentation/line spacing/before-after spacing/keep options, nested table structure, explicit cell widths/margins/borders/merges and row heights, section boundaries, page size/orientation, explicit body margins and gutter, columns, vertical alignment, document grid, inline JPEG/PNG/EMF picture frames with display size, cropping, rotation and flips, explicitly positioned main-story floating JPEG/PNG/EMF frames with basic wrapping, formatted header/footer variants and supported passive page-number fields, paragraph borders, formatted footnote/endnote content and references | frames, list/style-remapping and conditional table styles, advanced table/character/section properties, header/footer and note floating drawings, note numbering/positioning/custom separators and custom-mark rendering, lists, advanced floating drawings, non-raster images other than EMF, picture borders/effects and nonrectangular geometry, revisions, OLE; non-Western compressed code-page pieces are not decoded yet |
| XLS | CFB BIFF8 workbooks, including shared-string character data split across `CONTINUE` records | worksheet names and visibility (including very hidden), scalar values, cached formula results, merged ranges, date system, BIFF8 number formats, fonts, palette colors and supported checksum-bound extension colors, fills, borders, alignment, shared-string rich-text runs, styled blank cells, row heights and column widths, row/column hiding and outlines, print setup/margins/options, basic header/footer commands, manual page breaks, and measured passive embedded PNG/JPEG/EMF picture frames with supported cell anchors, cropping, rotation and flips | formula programs, phonetic string data, unsupported extended styles/theme colors/gradients, conditional formatting, print areas/titles, extended headers/footers, saved custom views, charts, non-picture drawings, grouped or active/linked picture objects, picture effects, external links, pre-BIFF8 sheets |
| PPT | CFB PowerPoint 97-2003 files with a resolvable current edit chain and persist directory | live slide order and dimensions, UTF-16/compressed Unicode text and outline references, individual shape anchors, nested group coordinates, basic rotation/flips, direct text margins/wrapping/vertical anchoring, direct font names/sizes/bold/italic/underline, literal and slide/master-scheme colors, paragraph alignment/spacing and explicit local ruler custom tabs, character bullets, explicit shape-local automatic numbering and paragraph-style offsets, verified-placeholder and explicit master-shape text-style inheritance, manual line breaks, unmodified basic presets with direct or explicitly linked master solid fill/line colors, line widths and opacity, line caps/joins, arrow ends and standard dash patterns, embedded/delayed JPEG, PNG and EMF picture frames with signed cropping, local/inherited solid and stretched-image backgrounds, eligible foreground picture fills on supported preset and uniform custom paths, enabled non-placeholder master objects using the same supported drawing subset, static slide-number metacharacters, explicit full-coordinate line/cubic paths with uniform path paint; superseded slides, deleted and explicitly hidden shapes are not emitted | unlinked placeholder and nonuniform master text overrides, unlinked/drawing-default paint, master placeholder content and header/footer fields, system/palette color indices, inherited/outline automatic numbering, picture bullets, text-ruler offsets/default intervals and inherited ruler tabs, advanced character formatting, embedded fonts, guide-dependent or compact custom geometry, arc/editing escapes, mixed per-path paint, some rotated/grouped geometry, gradients/patterns, custom dash arrays/compound lines, effects, custom fill rectangles and origins, charts, notes, WMF/PICT/DIB/TIFF/other image formats, picture effects and advanced foreground image-fill sizing, audio/video, transitions, animations, actions, OLE |

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

Version-3 and version-4 CFB containers are admitted. Password-protected legacy
binaries and pre-CFB Office formats are rejected. These limits are structural,
not filename-based. Unsupported binary structures fail with
`reason === 'unsupported-input'`. Accepted documents can still lose the features
listed above: their warning identifiers are not a fidelity certificate.

DOC inline/floating images, measured XLS picture frames, and PPT
picture/background images can retain passive
EMF BLIPs (MS-ODRAW 2.2.24/31). Both UID layouts and uncompressed or RFC 1950
zlib-compressed data are supported. Validated EMF bytes become ordinary image
parts without rasterization, geometry rewriting or execution of metafile
commands by the converter. Existing OOXML image handling renders the supported
EMF drawing subset; this is not full EMF/EMF+ or Office visual parity. WMF, PICT
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

EMF extraction checks the declared compressed/expanded lengths, stream end,
header and record envelope. As resource policy, each EMF is limited to 32 MiB
stored and expanded, with the existing 128 MiB per-media-store retention cap
checked before decompression. DOC inline and floating stores have separate
caps. Repeated image references share one retained buffer; cache lifetime ends
with conversion. Existing work and generated-package limits still apply.
Malformed records or over-budget expansion reject the input through the existing
error contract. No public API, per-format opt-in or application migration change
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
Floating frames and numbering are still absent; line wrapping and pagination
can differ significantly. No existing renderer changes are required.

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
other document settings or version-specific compatibility flags. List generation
remains unsupported. Preserving tabs does not imply Word pagination equivalence.

Section text flow preserves the basic top-to-bottom, right-to-left-column mode
(`sprmSTextFlow` / `msotxflTtoBA`, MS-DOC 2.6.4 and MS-ODRAW 2.4.5) as ordinary
`w:sectPr/w:textDirection w:val="tbRl"` (ECMA-376 17.6.20, 17.18.93 and Part 4
14.11.7). Each section resolves its own properties; an explicit horizontal reset
does not retain a previous vertical direction. Existing DOCX layout and Canvas
painting handle the orientation, with no binary-only renderer path. Other
rotation variants and version-dependent column-direction modes remain omitted
under the advanced-section-property warning; unknown enumeration values reject.
This does not yet preserve frame/cell text directions, drawings, list markers,
all East Asian character formatting, or exact Word line wrapping.

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
EMF picture frames outside nested groups. It preserves cell-relative anchor
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
extract the supported passive PNG, JPEG and EMF entries from a BIFF8 global
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
