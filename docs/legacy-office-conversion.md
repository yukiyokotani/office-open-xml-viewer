# Experimental legacy Office sources

Legacy binary Office files can be opened with the ordinary viewers and
loaders through optional model sources:

- `.doc` with the DOCX loaders and `DocxViewer`
- `.xls` with the XLSX loaders and `XlsxViewer`
- `.ppt` with the PPTX loaders and `PptxViewer`

Each reader is a separate opt-in entry. It reads a supported binary subset
directly into the existing document, workbook or presentation model, without
generating an OOXML package, and the ordinary layout and Canvas renderers
draw the result. Importing the ordinary DOCX, XLSX or PPTX entries alone does
not import, fetch or initialize any legacy reader. Without a matching source,
legacy input continues to reject with
`OoxmlError.code === 'legacy-binary-format'`, so no migration is required for
applications that do not opt in.

These readers are experimental and deliberately narrow. They reject input
they cannot represent instead of showing a partial document, and none of them
executes macros, fields, formulas, actions or embedded objects.

## Enabling a source

Pass the source in the format-generic `modelSources` option. It is available
on the DOCX, XLSX and PPTX `load()` options, on the viewers, and on the Node
session options (`openDocxDocument`, `materializeDocxDocument`,
`openXlsxWorkbook`, `openPptxPresentation` and their siblings).

```typescript
import { DocxViewer } from '@silurus/ooxml/docx';
import { legacyDocSource } from '@silurus/ooxml/legacy-doc';

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new DocxViewer(canvas, { modelSources: [legacyDocSource()] });
await viewer.load(docOrDocxBytes);
```

| Input | Entry | Factory | Loaders and viewers |
| --- | --- | --- | --- |
| DOC | `@silurus/ooxml/legacy-doc` | `legacyDocSource()` | `DocxViewer`, `DocxDocument.load`, `openDocxDocument`, `materializeDocxDocument` |
| XLS | `@silurus/ooxml/legacy-xls` | `legacyXlsSource()` | `XlsxViewer`, `XlsxWorkbook.load`, `openXlsxWorkbook` |
| PPT | `@silurus/ooxml/legacy-ppt` | `legacyPptSource()` | `PptxViewer`, `PptxPresentation.load`, `openPptxPresentation` |

Each factory accepts the same optional settings:

```typescript
interface LegacySourceOptions {
  wasmUrl?: string; // absolute URL of the reader's WASM
  moduleUrl?: string; // absolute URL of the reader's source module
  maxInputBytes?: number; // defaults to, and may not exceed, 256 MiB
}
```

Creating a source fetches nothing. When a claimed file loads, the parser
Worker (or Node) imports the source's self-contained ES module, emitted as
`legacy-<format>-source-module*.js` next to the package files, and initializes
the reader's dedicated WASM: `legacy_doc_direct_bg.wasm`,
`legacy_xls_direct_bg.wasm` or `legacy_ppt_direct_bg.wasm`. Serve these files
with the other package assets, and allow the module URL wherever a Content
Security Policy restricts `script-src` or Worker imports. Applications with a
custom asset pipeline pass absolute `moduleUrl` and `wasmUrl` values:

```typescript
const source = legacyPptSource({
  moduleUrl: new URL('/assets/legacy-ppt-source-module.js', location.href).href,
  wasmUrl: new URL('/assets/legacy_ppt_direct_bg.wasm', location.href).href,
});
```

Module URLs come only from application options; document content never
selects or rewrites them.

A source can report the document's own view preferences. Precedence is: an
explicit caller option, then the source's view default, then the renderer
default. Capabilities that a source lacks degrade the same way for every
format: resource metrics are reported without a ZIP usage snapshot, and
`toMarkdown()` rejects with an "... is unsupported for this source" error.
All three legacy readers lack Markdown export and ZIP accounting.

To cancel a load, destroy the document or viewer, or start another load in
its place. Node sessions keep their `signal` option.

## Admission and failure behavior

A source's synchronous `claim(bytes)` accepts only its own [MS-CFB] family:
a compound file whose directory names `WordDocument` (DOC), `Workbook` or
`Book` (XLS), or `PowerPoint Document` (PPT). A container that names more than
one family, or that carries `EncryptionInfo`, is not claimed. Every input a
source does not claim takes the unchanged OOXML path, so:

- a legacy file without a matching source rejects with the typed
  `legacy-binary-format` error;
- encrypted OOXML packages keep their existing encryption errors;
- configuring `legacyDocSource()` does not enable XLS or PPT input.

Claimed input larger than `maxInputBytes` throws a `RangeError`. The limit is
resource policy, not an Office format limit. After a source claims a file, a
reader failure rejects the load: there is no fallback to another source or to
the OOXML path. Password-protected legacy binaries, pre-CFB Office formats and
unsupported binary structures are rejected by the readers. These checks are
structural, never filename-based.

## Experimental direct DOC source

```typescript
import { DocxDocument } from '@silurus/ooxml/docx';
import { legacyDocSource } from '@silurus/ooxml/legacy-doc';

const document = await DocxDocument.load(legacyDocArrayBuffer, {
  modelSources: [legacyDocSource()],
});
const canvas = window.document.querySelector('canvas') as HTMLCanvasElement;

try {
  await document.renderPage(canvas, 0, { width: 960 });
} finally {
  document.destroy();
}
```

Node supports DOC through the same reader:

```typescript
import { openDocxDocument } from '@silurus/ooxml/node';
import { legacyDocSource } from '@silurus/ooxml/legacy-doc';

const session = await openDocxDocument(legacyDocBytes, {
  factory,
  modelSources: [legacyDocSource()],
});
```

Both browser rendering modes and progressive layout use the same retained
layout pipeline. The native source remains alive for image reads until the
document is destroyed; finishing a model cursor does not dispose it.

This is a narrow experimental reader, not full DOC support. Unsupported
formatting, fields, numbering, notes and other unimplemented structures may
reject the entire document.

When a DOC prints its revision markup (MS-DOC `DopBase.fRMPrint`) and carries
revision marks, the reader reports `showTrackedChanges: true` as its view
default. An explicit `showTrackedChanges` option, including `false`, always
wins.

Missing header and footer distances use the MS-DOC 2.6.4 defaults for the
stored producer installation LCID when the specification lists that LCID.
Explicit values, including zero, win. Unlisted languages keep the
unresolved-margin recovery. The host locale and the document's text language
are not used to guess the producer's settings.

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

Browser:

```typescript
import { PptxPresentation } from '@silurus/ooxml/pptx';
import { legacyPptSource } from '@silurus/ooxml/legacy-ppt';

const presentation = await PptxPresentation.load(legacyPptArrayBuffer, {
  modelSources: [legacyPptSource()],
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
import { legacyPptSource } from '@silurus/ooxml/legacy-ppt';

const session = await openPptxPresentation(legacyPptBytes, {
  modelSources: [legacyPptSource()],
});

try {
  for await (const slide of session.slides()) {
    // Consume each ordinary shared slide model.
  }
} finally {
  await session.close();
}
```

The direct source is an experimental, bounded subset, not a full-fidelity
PowerPoint implementation. Besides Markdown and ZIP accounting, it currently
lacks audio/video media and embedded fonts. It rejects constructs it cannot
represent, including unresolved paragraph margin or indentation and positive
paragraph before/after percentages. Its explicit unsupported diagnostics are
authoritative.

A slide's own `SlideShowSlideInfoAtom.fHidden` marks it hidden; hidden slides
and their content remain in the presentation. Display follows the existing
`PptxViewer` `hiddenSlideMode` option, whose default remains `'show'`.

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

Classic linear gradients (MS-ODRAW `msofillShade` and `msofillShadeScale`)
keep their ordered shade colours, signed focus and 16.16 angle. Explicit
shade-array stops at either endpoint take precedence over scalar front or
back colours; a missing final endpoint uses the scalar back colour. Linear,
scaled, two-colour and translucent shades are projected, including on rotated
shapes and inside rotated or flipped groups. Path and title shades, other
shade types, custom fill rectangles and opacity combined with a shade-colour
array fail closed. Twenty controlled Office comparisons covered focus, angle,
endpoint conflicts and leaf flips; the projected positions were within one
DrawingML position unit of Office's serialized integers.

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
Brightness/contrast (washout) on a picture becomes DrawingML `lum`: bright is
the brightness over 0x8000, and contrast is k - 1 or 1 - 1/k for the stored
16.16 slope k. PowerPoint writes exactly these values when it saves the
binary deck as PPTX. The renderer's `lum` formula reproduces PowerPoint's PDF
of a gray ramp under a bright × contrast grid, from both PPTX and .ppt, to
within 0.7 of 255. Brightness/contrast combined with a transparent colour or
black-and-white (whose order has no evidence), grayscale or black-and-white
alone, recolouring and adjustments on picture fills stay rejected until
Office output confirms their rendering.

A modern Office-saved PPT can retain paragraph properties in the OfficeArt
`metroBlob` alternative shape XML rather than in its classic text ruler. The
direct PPT source adopts that XML under the rule described with the release
gap inventory below; see also the [controlled probe protocol](../packages/legacy-converter/tools/legacy-ppt-ruler-probes.md)
before attributing differences to an implicit ruler rule.

## Experimental direct XLS source

```typescript
import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { legacyXlsSource } from '@silurus/ooxml/legacy-xls';

const container = document.querySelector('#workbook') as HTMLElement;
const viewer = new XlsxViewer(container, { modelSources: [legacyXlsSource()] });
await viewer.load(legacyXlsBytes);
```

`XlsxWorkbook.load(bytes, { modelSources: [legacyXlsSource()] })` and the Node
`openXlsxWorkbook(bytes, { modelSources: [legacyXlsSource()], factory })`
accept the same source.

The direct XLS source is experimental and bounded. It reads a BIFF8 workbook
subset into the shared worksheet model: sheet names, cell values and cached
formula results, merged ranges, the date system, number formats, fonts,
palette and extension colors, fills, borders, alignment, rich-text runs, row
heights and column widths, and row/column hiding and outlines. It also
projects conditional formatting (classic CF with its CFEx extensions, CF12
comparison, formula, color-scale, data-bar and icon-set rules, their
differential formats and decompiled formulas), Excel tables with custom table
styles, frozen panes, tab colors, XFExt gradient fills, hyperlinks, worksheet
AutoFilter ranges (whose application-inserted drop-down objects become the
renderer's filter buttons; active filter criteria reject), data validations
and defined names, plus embedded charts, chart sheets, pictures, and
rectangles, text boxes and their groups as shape anchors. Variants the shared
model cannot show (for example inactive rules, suppressed list drop-downs,
displayed phonetic guides, or table-style elements outside the model) reject
the workbook instead of being dropped. Print areas and titles travel as
defined names; page setup, headers and footers affect printing only and are
not part of the model.

Worksheet visibility, including very hidden sheets, is kept as model
metadata; display follows the existing `XlsxViewer` `hiddenSheetMode` option,
whose default remains `'show'`. Gridline visibility, zero-value display and
right-to-left direction come from the sheet's window settings. Pane
selections, scrolling, zoom and window placement are not reconstructed.

Column widths and drawing anchors depend on the Normal font's maximum digit
width (ECMA-376 §18.3.1.13). The XLS source asks the host for it through a
generic host-layout capability, and the XLSX renderer answers with its own
`computeMdw`, the same measurement that sizes the painted grid. It measures
in the rendering realm: the render Worker in `mode: 'worker'`, the page in
main mode, and the session canvas `factory` in Node. There is no measurement
callback option. Without a measurable font (for example in Node without a
`factory`), charts, pictures and shapes are omitted. On macOS the renderer
quantizes the advance to whole points like Excel for Mac, so XLS drawings
follow that platform's rule.

## XLS drawing inspection for development

A native-only inspection helper can
extract the supported passive PNG, JPEG, EMF and WMF entries from a BIFF8 global
image store without requiring font metrics:

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
feature is enabled. It does not change the reader or renderer.

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
geometry. The direct XLS reader applies additional picture eligibility.

Anchor inspection limits cumulative drawing bytes to 128 MiB, record work to
two million, substream/group nesting to 32, retained anchors to 65,536, and
per-sheet shape/client identities to 65,536. The disjoint ranges prevent repeated
scanning through overlapping worksheet references. Native metadata does not prove
visibility, image eligibility or complete object validity: non-picture objects,
deleted shapes and OLE-marked shapes can have anchors too. These development
helpers remain separate from the direct XLS reader.

## Current implementation boundary

The repository contains only the direct readers; legacy input is never
converted to OOXML. The Rust crate `legacy-office-converter` builds one reader
per feature: `direct-doc`, `direct-xls` and `direct-ppt`. Building it for
`wasm32` with none of them enabled is a compile error. The `inspection`
feature adds the native-only examples above, and `fuzzing` exposes the fuzz
entry points.

The local direct-render survey described below renders every installed
legacy sample through its source and pairs it with the Office-exported PDF.
The corpus is deliberately not redistributed. Broader binary-record coverage,
visual fidelity against Office, fuzzing and resource measurements remain part
of [issue #1472](https://github.com/yukiyokotani/office-open-xml-viewer/issues/1472).

## Best-effort fidelity evaluation

Loading without an error is a smoke test, **not reader completion**. The
target is useful best-effort preservation of the binary input's content and
display, with missing content and visual differences explicitly reported.
Pixel equality is not required for each incremental improvement. Pairing a
legacy file with its original OOXML is useful for investigation, but does not
prove fidelity: saving to an old format can itself change or remove features.
Use Office opening the actual legacy file as the visual reference. Office's
upgraded OOXML is useful for mapping binary records to XML, but Office's own
conversion can change layout and is not an absolute visual oracle. In
particular, rebuilt or down-saved corpus members must not silently be treated
as lossless copies of their original OOXML.

The local macOS exporter opens a disposable copy with installed Microsoft
Office, with macros disabled and Word/Excel external-link updates disabled,
and writes a PDF to a fresh output path:

```bash
osascript packages/legacy-converter/tools/legacy-office-export.applescript \
  doc disposable-copy.doc fresh-output.pdf
```

The first argument is `doc`, `xls` or `ppt`. The exporter requires Office for
macOS and macOS automation permission for each Office application. Runs are
sequential; an Office failure stops rather than accumulating open documents
or dialogs. Original corpus files and existing visual references are never
changed, and PDFs and page images stay in a local temporary directory; no
private artifact is committed or uploaded.

The exporter refuses to open a document unless Office reports its automation
security setting and confirms that macros are disabled. PowerPoint builds that
return no value for this property are currently blocked, even after macOS
automation permission is granted. Word and Excel PDF export have been exercised;
the PowerPoint export path is not yet validated end to end. Do not weaken this
guard to obtain a passing report.

Compare the direct reader's Canvas output with those Office PDFs through the
local direct-render survey. Keep renderer self-regression tests against the
previous renderer separate from this fidelity comparison. Neither whole-corpus
Office equality nor Canvas display equality has been reached.

Treat the rendered model as a derived view of the original binary, keep the
binary as the authoritative source, and gate production use on a corpus
representative of the documents being ingested.

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
| DOC-41 | Connect style-aware shading to native acquisition | Connected for native projection: effective FIB policy is configured before profile caching; actual-story tests resolve source-cell Raw values against the selected style. Unresolved compatibility arrays remain gated |
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
shared-model change was needed for this checkpoint. Full-branch architecture and browser/visual acceptance remain
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
| DOC-71 | Separate native row compatibility policy from the generic path | Implemented explicit native-reader cantSplit policy while preserving generic behavior |
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
`packages/legacy-converter/tools/legacy-doc-papx-probes.py` accepts `legacy-doc-property-trace/v1`.
Each source-hash-bound target declares a physical `fc`, `owner` (`ttp` or
`paragraph`), and nonempty `properties` mapping lowercase four-digit SPRM codes
to exact framed operand arrays. An empty array asserts absence. Optional `order`
asserts the complete acquired sequence filtered to those property codes.

The fixed-fixture `check-target INPUT ASSERTIONS` command in
`packages/legacy-converter/tools/legacy-doc-table-style-probes.py` accepts
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
style and direct-PAPX paths retain their existing behavior.
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
| TOOL-1 | Office export restoration | Open: `packages/legacy-converter/tools/legacy-office-export.applescript` cannot restore Word/Excel settings when `open` returns no value; it also adopts an unrestored ForceDisable baseline |
| LEGACY-OOXML | Remove the OOXML-generation path | Done: legacy support was unreleased, so the OOXML generator, its WASM/TS entry points and XML writers were deleted rather than deprecated; the direct readers do not depend on them |

### Direct-model table-style admission checkpoint

The direct DOC model now admits table-style selection instead of rejecting
every table that carries sprmTIstd or sprmTTlp. The table-style profile
(TAPX/PAPX/CHPX, conditional selection, borders, margins, shading) keeps its
own per-property gates. The decisions
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

`packages/legacy-converter/tests/survey/{doc,ppt,xls}.spec.ts` render each
local private legacy sample through its direct source, on the matching viewer
package's own VRT fixture and dev server. Each sample is written beside its
same-named Office PDF export as paired PNGs and a summary. The survey reports
only: it never gates, updates references, or generates OOXML. Run it with an
output directory outside the checkout (`VRT_PORT` serves DOC, `+1` PPT and
`+2` XLS; `LEGACY_CORPUS_FORMATS` and `LEGACY_CORPUS_FILTER` narrow the run):

```bash
LEGACY_CORPUS=1 LEGACY_CORPUS_OUT=/tmp/legacy-survey LEGACY_CORPUS_FORMATS=ppt \
  pnpm --filter @silurus/ooxml-legacy-converter survey
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
| XLS | Chart and picture anchors need the Normal font's digit width. The renderer measures it with the same `computeMdw` that sizes the painted grid, so anchors follow the grid, but they match Excel only when Excel's fonts (such as Calibri) are available; without a measurable font they are omitted, and shared reference font metrics are needed | all with drawings |
| XLS | Chart text omits TextPropsStream (its checksum is not implemented), Fbi font autoscaling, the outline Excel draws around inverted negative points, plot-area layout, drop/high-low lines and 3-D walls | most chart samples |
| XLS | ~~Extended colors (XFExt theme/tint) fall back to palette approximations~~ Resolved: tints (b09eae6a) and theme 0-3 in Excel's lt1/dk1/lt2/dk2 order (0adbc794) | about 6 |
| XLS | Table (ListObject) styles, conditional-format data bars/icons and pivot styling are absent | about 5 |
| XLS | Formula text is not decompiled from Ptg tokens, so volatile functions are not recalculated as Excel does at export | 2 |
| XLS | Clip-art pictures, strikethrough and one vertical merge are missing | 1 to 3 each |
| XLS | The direct reader rejects, instead of omitting, drawn objects it does not project: lines, ovals and other shape types, grouped charts, macro sheets and AutoFilter criteria. Chart sheets are projected as chart-sheet worksheets with the chart at its Chart record rectangle; rectangles, text boxes, freeform polygons, pictures and their (rotated, flipped or nested) sheet groups are projected as XLSX-model shape anchors with solid paint and TxO text, following Excel's own XLSX of the same workbooks; OfficeArt data that Excel continues in Continue records after a complete Obj, chart substream or TxO is assembled by native record length | 0 of 139 |
| XLS | ~~The chart area's automatic border is not projected~~ Resolved: an automatic chart area takes the BIFF outline Excel writes for it | ~~1~~ |
| PPT | ~~Only seven MS-ODRAW shape types map to presets~~ 100+ shape types map as PowerPoint converts them, with evidenced adjust formulas (officeart::preset); adjusted callout2/3 families, arrow callouts, curved arrows, ribbons and tall cubes/hexagons/parallelograms still fail closed | several |
| PPT | ~~Native/OLE charts are missing~~ Resolved: embedded OLE objects show their stored presentation picture (bfc835d3) | 3 |
| PPT | ~~Rotation by multiples of 90 degrees and combined flips use the wrong bounds or order~~ Resolved from the 120-case PowerPoint control (aa9dc5c1) | 1 |
| PPT | ~~Slide gradient backgrounds~~ linear/scaled/two-colour/translucent shades resolved (95b74d19); path (5, 6) and title (8) shades now fail closed. Bullets resolved through master levels; letter spacing, shrink-to-fit and per-paragraph indents come from adopted alternative shape XML where it agrees with the binary | several |
| PPT | ~~Gradients on rotated shapes (or inside rotated/flipped groups) are replaced by the solid fill colour~~ Resolved (ef41f03a) | several |
| PPT | Custom geometry with per-path fill/stroke flags is rejected; the PPTX model has no per-path `fill`/`stroke` (ECMA-376 §20.1.9.15), a generic PPTX gap | 1 |
| PPT | ~~Unmapped shape types are dropped silently~~ Now rejected | several |
| PPT | ~~Picture brightness/contrast (washout)~~ projected as `lum` from the gray-ramp control; pattern fills (including on rotated shapes) are projected; pattern fills on flipped shapes, texture fills, and OLE icons, links and controls are rejected | several |
| PPT | Implicit paragraph margin/indent and percentage spacing are rejected | 12 of 34 load failures |
| PPT | Alternative shape XML (metroBlob) that cannot be verified against the binary shape fails closed: placeholders (no slide layout), a preset against freeform geometry, and fills stated in non-comparable forms | first error of 14 of 34 (10 previously loaded) |
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

The alternative therefore carries display information the binary lacks, so
the direct PPT source resolves each shape's alternative to one of three
outcomes:

- The package names no alternative part (only the `downRev` checksums), or
  the alternative verifiably disagrees with the binary on a compared
  attribute: the binary projection is used, as PowerPoint does.
- The alternative agrees on every compared attribute: it is adopted. Its
  text characters are masked, so the adopted shape takes its characters,
  transform and identifier from the binary. The XML is parsed by the ordinary
  PPTX shape parser, resolving theme references against the master's
  round-trip theme and color map; no OOXML is generated. Its serialized size
  is charged to the session's model budget.
- Otherwise agreement cannot be established and the presentation fails
  closed as unsupported input: an oversized, over-budget, duplicated or
  unreadable blob, package or round-trip theme; an alternative part that is
  not a shape or connector (pictures, groups, SmartArt, ink); a placeholder,
  whose alternative inherits from a slide layout the binary file does not
  have; relationship references; and any compared attribute the two forms
  state in ways the reader cannot equate.

The compared attributes are those the controls cover, each in one unit
system. Presets compare by name and by adjust values within the range the
binary inputs' rounding allows (a whole 21600-based value on an anchor that
rounds PowerPoint's extent to one master unit), an omitted value standing
for the preset default of ECMA-376 `presetShapeDefinitions.xml`; a preset
against custom geometry is not comparable, because PowerPoint saves a preset
without an MS-ODRAW shape type as a freeform. Custom geometry compares path
by path in normalized path coordinates, including the authored per-path fill
and stroke flags, after dropping a closing line to the subpath start that the
binary stores explicitly. Position, size, rotation and flips compare in the
binary anchor's unit (master units, or a group's unscaled child units) within
one unit. The recorded fill compares before PowerPoint's open-path display
rule: solid colors within one unit per channel (independent rounding of a
theme color transform), gradients and patterns only when identical within
their fixed-point rounding. Run font size, bold and italic compare where both
state them, and paragraphs, runs and line breaks must line up at equal UTF-16
lengths. XML 1.0 cannot carry the vertical tab that breaks a line in a binary
paragraph; the masked alternative states it as one masked character of the
enclosing run, which becomes a line break in the adopted shape. The downrev
checksums beside the XML are not recomputable (they are not checksums of the
binary records), so structural agreement stands in for them; a binary edit
that preserves every compared property would still adopt a stale XML.
Glyph shadow and emboss, which the binary cannot express, reject a shape only
when its alternative XML is not adopted.
