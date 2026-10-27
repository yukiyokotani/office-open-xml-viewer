# Experimental legacy Office sources

Install the viewer and its separate optional companion at matching versions:

```sh
npm i @silurus/ooxml @silurus/ooxml-legacy-office
```

The companion pins its `@silurus/ooxml` peer to the exact same version.
A stale runtime or model-source protocol fails before input loading with
`CompanionRuntimeMismatchError` (`code: 'companion-runtime-mismatch'`), exported
from the internal `@silurus/ooxml/internal/companion` ABI. That entry is reserved
for companion packages; applications use the ordinary viewer/loader entries.

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

Import the reader and pass it as its own option, like the other opt-in
modules (`{ math }`, `{ chartEx }`, `{ tiff }`). The option is `legacyDoc` on
the DOCX surfaces, `legacyXls` on the XLSX surfaces and `legacyPpt` on the
PPTX surfaces: the `load()` options, the viewers and scroll viewers, and the
Node session options (`openDocxDocument`, `materializeDocxDocument`,
`openXlsxWorkbook`, `openPptxPresentation` and their siblings). Each option
accepts only a source for its own format.

```typescript
import { DocxViewer } from '@silurus/ooxml/docx';
import { legacyDoc } from '@silurus/ooxml-legacy-office/doc';

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new DocxViewer(canvas, { legacyDoc });
await viewer.load(docOrDocxBytes);
```

| Input | Entry | Reader / factory | Option | Loaders and viewers |
| --- | --- | --- | --- | --- |
| DOC | `@silurus/ooxml-legacy-office/doc` | `legacyDoc` / `createLegacyDoc()` | `legacyDoc` | `DocxViewer`, `DocxScrollViewer`, `DocxDocument.load`, `openDocxDocument`, `materializeDocxDocument` |
| XLS | `@silurus/ooxml-legacy-office/xls` | `legacyXls` / `createLegacyXls()` | `legacyXls` | `XlsxViewer`, `XlsxSheetViewer`, `XlsxWorkbook.load`, `openXlsxWorkbook` |
| PPT | `@silurus/ooxml-legacy-office/ppt` | `legacyPpt` / `createLegacyPpt()` | `legacyPpt` | `PptxViewer`, `PptxScrollViewer`, `PptxPresentation.load`, `openPptxPresentation` |

`legacyDoc`, `legacyXls` and `legacyPpt` use the bundled asset URLs. To change
them, or to set an input limit, create a reader with the matching factory; each
factory accepts the same optional settings:

```typescript
interface LegacySourceOptions {
  wasmUrl?: string; // absolute URL of the reader's WASM
  moduleUrl?: string; // absolute URL of the reader's source module
  maxInputBytes?: number; // optional caller limit; omitted means no whole-file gate
}
```

Importing a reader or creating one with a factory fetches nothing. When a claimed file loads, the parser
Worker (or Node) imports the source's self-contained ES module, emitted as
`legacy-<format>-source-module*.mjs` next to the package files, and initializes
the reader's dedicated WASM: `legacy_doc_direct_bg.wasm`, `legacy_xls_direct_bg.wasm` or
`legacy_ppt_direct_bg.wasm`. Serve these files
with the other package assets, and allow the module URL wherever a Content
Security Policy restricts `script-src` or Worker imports. Applications with a
custom asset pipeline pass absolute `moduleUrl` and `wasmUrl` values:

```typescript
import { createLegacyPpt } from '@silurus/ooxml-legacy-office/ppt';

const legacyPpt = createLegacyPpt({
  moduleUrl: new URL('/assets/legacy-ppt-source-module.js', location.href).href,
  wasmUrl: new URL('/assets/legacy_ppt_direct_bg.wasm', location.href).href,
});
const viewer = new PptxViewer(canvas, { legacyPpt });
```

Module URLs come only from application options; document content never
selects or rewrites them.

A source can report the document's own view preferences. Precedence is: an
explicit caller option, then the source's view default, then the renderer
default. Capabilities that a source lacks degrade the same way for every
format: resource metrics are reported without a ZIP usage snapshot, and
`toMarkdown()` rejects with an "... is unsupported for this source" error.
All three legacy readers lack ZIP accounting, and the legacy XLS and PPT
readers lack Markdown export.

The legacy DOC reader does offer Markdown export: the direct reader projects
into the same document model the DOCX parser produces, so `toMarkdown()`
runs the shared DOCX markdown projection (headings from outline level,
lists, tables, notes, character formatting; geometry discarded) on that
model. The projection re-reads the retained source bytes, so it works
independently of the streaming load cursor, and it shares the DOCX markdown
byte ceiling. Legacy XLS and PPT keep the rejection: the XLSX markdown
projection walks OOXML worksheet JSON against the package shared-strings
table, which the typed BIFF projection does not carry, and the PPTX
projection consumes a whole presentation including notes and comments,
which the streaming legacy PPT session does not assemble.

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
- configuring `legacyDoc` does not enable XLS or PPT input.

When the caller sets `maxInputBytes`, claimed input larger than that positive
safe integer throws a `RangeError`. There is no default or library-imposed
whole-file size ceiling. Applications that require the former 256 MiB input
limit must pass `maxInputBytes: 256 * 1024 * 1024` explicitly. Source copies remain
fallible: allocation failure is a classified resource error. Internal budgets
apply to actual inflated or retained work, rather than compressed or declared
sizes. These budgets are library policy, not Office format limits.

After a source claims a file, a reader failure rejects the load: there is no fallback to another source or to
the OOXML path. Password-protected legacy binaries, pre-CFB Office formats and
unsupported binary structures are rejected by the readers. These checks are
structural, never filename-based.

After admission, representational rejections use the viewer-owned `OoxmlError`
with `code: 'unsupported-content'` and frozen, content-free details:
`{ format: 'doc' | 'xls' | 'ppt', reason: 'source-subset' }`. Offer conversion
for this code; do not match diagnostic text. The category includes malformed
and unknown structures and does not certify file validity. It applies to open
and lazy worksheet, slide and image operations, in browser main/worker rendering
modes and Node. Background failures use the same error through `onError`.

Encryption uses `unsupported-encryption`. `OoxmlResourceLimitError`,
`parser-crashed`, `companion-runtime-mismatch`, network and worker failures keep
their own meaning. Native budgets without measured resource records remain
ordinary errors, as does cursor misuse. Input size guards remain `RangeError`.
The source module protocol is v2; rebuild custom modules using the documented
model-source rejection signal and install matching viewer/companion versions.
No migration is required for applications that do not enable model sources.

## Experimental direct DOC source

```typescript
import { DocxDocument } from '@silurus/ooxml/docx';
import { legacyDoc } from '@silurus/ooxml-legacy-office/doc';

const document = await DocxDocument.load(legacyDocArrayBuffer, {
  legacyDoc,
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
import { legacyDoc } from '@silurus/ooxml-legacy-office/doc';

const session = await openDocxDocument(legacyDocBytes, {
  factory,
  legacyDoc,
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

A default character language of 0x007F (reserved in MS-LCID 2.2; Windows
names it LOCALE_INVARIANT), stamped by some third-party producers, projects
no language: Word opens such documents and writes an empty `w:lang
w:val=""` on DOCX save, which the DOCX reader treats as no language. The
other language axes have no Office evidence for this value and stay
rejected, as do unlisted LCIDs.

A table whose style selection resolves to istd 0x000B (MS-DOC 2.6.3
sprmTIstd) in a stylesheet that does not define even that slot gets no
table style: Word opens such documents and saves the tables without
`w:tblStyle`. Unsupported content inside a defined table style stays
rejected.

A table-styled cell whose modern compatibility shading array
(sprmTDefTableShd and its later segments, MS-DOC 2.6.3) carries the clear
pattern over automatic colors (ShdAuto, MS-DOC 2.9.247) projects as an
explicit no-fill value, the same projection the Raw shading arrays already
give it. Patterned entries remain unsupported.

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
at the page bottom and endnotes at the end of the document. A custom
reference mark is shown as its own literal text and does not use up a number,
as in Word. A separator story that is only an empty paragraph keeps that
paragraph's space with no line. Documents are rejected if they use number
restarts per section or per page, sections with different note numbering, or
other custom separator stories.

Section text flow (MS-DOC sprmSTextFlow, an MS-ODRAW MSOTXFL value) is shown
per section, following the basic rules of the specification:

- Values 0 and 4 are horizontal. No different glyph treatment is inferred
  for value 4 from its name.
- Value 1 is vertical with upright East Asian characters, as before.
- Values 3 and 5 run downward with later lines to the left and every
  character rotated a quarter turn clockwise.
- Value 2 runs upward with later lines to the right and every character
  rotated a quarter turn counter-clockwise.

MS-ODRAW notes that Word 2007 and 2010 place later lines of value 5 to the
right. The file does not record which application displays it, so the
documented basic rule is used. Word's placement of headers, footers, notes,
anchored drawings and upright tables in value-2 sections has not been
established. They follow the library's general placement for vertical
sections, turned with the section's own direction, without a claim of
matching Word. Horizontal-in-vertical runs remain supported only in value-1
sections.

## Experimental direct PPT source

Browser:

```typescript
import { PptxPresentation } from '@silurus/ooxml/pptx';
import { legacyPpt } from '@silurus/ooxml-legacy-office/ppt';

const presentation = await PptxPresentation.load(legacyPptArrayBuffer, {
  legacyPpt,
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
import { legacyPpt } from '@silurus/ooxml-legacy-office/ppt';

const session = await openPptxPresentation(legacyPptBytes, {
  legacyPpt,
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
represent, including paragraph origins that cannot be resolved from the
stored text and master records. Its explicit unsupported diagnostics are
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
unsupported fill placements are rejected instead of being drawn without them.

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
pixels in the background colour. Supported placements retain their shape transforms. Unsupported bitmap
sizes, translucent pattern colours and unprojected texture placements
remain rejected.

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
direct PPT source adopts that XML under the rule described below.

## Experimental direct XLS source

```typescript
import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { legacyXls } from '@silurus/ooxml-legacy-office/xls';

const container = document.querySelector('#workbook') as HTMLElement;
const viewer = new XlsxViewer(container, { legacyXls });
await viewer.load(legacyXlsBytes);
```

`XlsxWorkbook.load(bytes, { legacyXls })` and the Node
`openXlsxWorkbook(bytes, { legacyXls, factory })`
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
cargo run -p legacy-office-parser --features inspection \
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

Native inspection examples read the complete source with fallible allocation and
no default input-length ceiling. Inspection caps assembled drawing data and retained
media at 128 MiB each, and the drawing/decoding walk at two million records. Shared
image limits also apply. These are independent resource ceilings, not a combined
process-memory guarantee: source, workbook, assembled data and extracted images
can coexist. The helper is excluded from production WASM even when its Cargo
feature is enabled. It does not change the reader or renderer.

The companion `inspect_xls_anchors` native example prints raw worksheet anchor
metadata without extracting images:

```sh
cargo run -p legacy-office-parser --features inspection \
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
converted to OOXML. The Rust crate `legacy-office-parser` builds one reader
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
osascript packages/legacy-office/tools/legacy-office-export.applescript \
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

## Alternative shape XML

Modern Office can retain DrawingML in an OfficeArt `metroBlob`
([MS-ODRAW] §2.3.4.41). The source uses three outcomes: a package with no
alternative or a verified disagreement uses the binary projection; a verified
consistent alternative is adopted; an unreadable or unverifiable alternative
rejects the load. Exhausting an optional alternative's verification budgets
keeps that shape's binary projection and does not poison the document. Input
work uses the standard OOXML package allowance (256 MiB of actual consumed
input per alternative, including rereads); inflated output remains bounded to
1 MiB per part and 4 MiB per alternative package. Round-trip theme verification
also has its own input-work allowance. Failed attempts retain no model, text,
or image charges. A checksum-only package does not require a theme.

Adopted alternatives retain their formatting and resource references, while
characters, identifiers and transforms come from the binary. Theme style
references must resolve to existing entries. Compatibility comparisons cover
only the properties documented beside the implementation; edits to other
properties can leave a stale alternative undetected. This is an experimental
compatibility boundary, not a claim of complete PowerPoint equivalence.

Per-entry and cumulative budgets bound input, records, decoded themes, text,
model storage and emitted JSON. Repeated hyperlink targets and split text
styles consume the model budget before they are copied. These ceilings are
library resource policy, not Office format restrictions.

The private Office corpus and its visual survey remain local development
material. Current completion requires fresh parser-backed tests, renderer
self-regression checks and separately adjudicated Office fidelity evidence.
Historical survey counts do not establish the current release's behavior.

### XLS worksheet ownership

XLS worksheets use the ordinary XLSX cache admission, leases, LRU and render
worker eviction. Native projections belong to one cursor and are released on
terminal acknowledgement, cancellation or close. Reopening projects only the
selected sheet from retained, bounded neutral BIFF facts.

The cursor emits an explicit `metadata-unavailable` preview before row units.
XLS viewers therefore wait for the complete worksheet. Early viewport rendering
requires a separately verified BIFF preview eligibility rule.
