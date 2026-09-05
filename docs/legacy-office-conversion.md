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
| DOC | CFB Word 97-2003 documents with a readable main-story CLX piece table | main-story text, paragraphs, tabs, custom tab stops and document-wide default tab interval, line/page/column breaks, displayed field results, font names and explicit sizes, paragraph-style character defaults, character styles and direct bold/italic/underline/strike/caps/color/spacing properties, paragraph alignment/indentation/line spacing/before-after spacing/keep options, nested table structure, explicit cell widths/margins/borders/merges and row heights, section boundaries, page size/orientation, explicit body margins and gutter, columns, vertical alignment, document grid, inline JPEG/PNG picture frames with display size, cropping, rotation and flips, explicitly positioned main-story floating JPEG/PNG frames with basic wrapping, formatted header/footer variants and supported passive page-number fields | frames, list/style-remapping and conditional table styles, advanced table/character/section properties, header/footer floating drawings, notes, lists, advanced floating drawings, non-raster images, picture borders/effects and nonrectangular geometry, revisions, OLE; non-Western compressed code-page pieces are not decoded yet |
| XLS | CFB BIFF8 workbooks, including shared-string character data split across `CONTINUE` records | worksheet names, scalar values, cached formula results, merged ranges, date system, BIFF8 number formats, fonts, palette colors, fills, borders, alignment, shared-string rich-text runs, styled blank cells, row heights and column widths, row/column hiding and outlines, print setup/margins/options, basic header/footer commands and manual page breaks | formula programs, phonetic string data, extended styles/themes/gradients, conditional formatting, print areas/titles, extended headers/footers, saved custom views, charts, drawings, external links, pre-BIFF8 sheets |
| PPT | CFB PowerPoint 97-2003 files with a resolvable current edit chain and persist directory | live slide order and dimensions, UTF-16/compressed Unicode text and outline references, individual shape anchors, nested group coordinates, basic rotation/flips, direct text margins/wrapping/vertical anchoring, direct font names/sizes/bold/italic/underline, literal and slide/master-scheme colors, paragraph alignment/spacing, character bullets and paragraph-style offsets, verified-placeholder and explicit master-shape text-style inheritance, manual line breaks, unmodified basic presets with direct or explicitly linked master solid fill/line colors, line widths and opacity, line caps/joins, arrow ends and standard dash patterns, embedded/delayed JPEG and PNG picture frames with signed cropping, local/inherited solid and stretched-image backgrounds, enabled non-placeholder master objects using the same supported drawing subset, static slide-number metacharacters, explicit full-coordinate line/cubic paths with uniform path paint; superseded slides, deleted and explicitly hidden shapes are not emitted | unlinked placeholder and nonuniform master text overrides, unlinked/drawing-default paint, master placeholder content and header/footer fields, system/palette color indices, automatic numbering, picture bullets, text-ruler offsets and tabs, advanced character formatting, embedded fonts, guide-dependent or compact custom geometry, arc/editing escapes, mixed per-path paint, some rotated/grouped geometry, gradients/patterns, custom dash arrays/compound lines, effects, custom fill rectangles, charts, notes, vector/DIB/TIFF/other image formats, picture effects and foreground image fills, audio/video, transitions, animations, actions, OLE |

Version-3 and version-4 CFB containers are admitted. Password-protected legacy
binaries and pre-CFB Office formats are rejected. These limits are structural,
not filename-based. Unsupported binary structures fail with
`reason === 'unsupported-input'`. Accepted documents can still lose the features
listed above: their warning identifiers are not a fidelity certificate.

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
Header floating drawings, advanced field switches, numbering restarts/formats
from section properties, and exact Office pagination remain incomplete.
Each aggregate main/header story has a 64 Mi UTF-16-unit decoding ceiling and
one million controls; headers additionally allow at most 4,096 nonempty parts.
The aggregate generated XML has a 256 MiB ceiling. These are resource policies,
not format limits. No migration or opt-in API change is required.

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
glyph, guessed font index or size clamping is applied. Automatic numbering and
picture-bullet extensions remain unsupported; only the available base character
bullet is preserved when such extensions are present.

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
supported direct subset, including bullets and custom rulers, remain omitted.

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
