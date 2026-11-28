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

These readers are experimental and deliberately narrow. Many structures they
cannot map still reject the whole document. Some details without a model
equivalent are left out of an otherwise complete document instead; the
sections below give examples. None of the readers executes macros, fields,
formulas, actions or embedded objects.

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

The legacy DOC reader offers Markdown export: it builds the same document
model as the DOCX parser, and `toMarkdown()` runs the shared DOCX Markdown
projection on that model (headings from outline level, lists, tables, notes,
character formatting; geometry discarded). Document text and hyperlink
targets are escaped exactly as for a DOCX. The reader keeps the source bytes
for the life of the document and projects from them on each call, so
`toMarkdown()` does not depend on the state of the streaming load. The
output shares the DOCX Markdown byte ceiling. The source is kept only while
it is no longer than the reader's model byte budget (256 MiB by default); a longer
source still loads and renders, and its `toMarkdown()` rejects with a
resource-limit error. Legacy XLS and PPT keep the
rejection: the XLSX projection walks OOXML worksheet JSON against the package
shared-strings table, which the typed BIFF projection does not carry, and the
PPTX projection consumes a whole presentation including notes and comments,
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

### Difference from the OpenDocument readers

The two optional reader packages currently follow different policies. The
legacy readers reject the whole document with `unsupported-content` when they
meet many structures they cannot map. The OpenDocument readers
(`@silurus/ooxml-opendocument`, see [opendocument.md](./opendocument.md))
render every feature they read and reject only corrupt, encrypted or unsafe
input, or input over a resource limit. Keep a conversion fallback for legacy
input. The legacy behaviour may change in a later release: aligning it with
the OpenDocument policy is tracked in [issue #2263](https://github.com/yukiyokotani/office-open-xml-viewer/issues/2263).

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

A picture's `imagePath` in the DOC model is a reader-owned key such as
`legacy-doc/image/0`, not a ZIP part. Read its bytes through the open session:

```typescript
const picture = await session.getImage(run.imagePath, run.mimeType); // Blob
```

`openXlsxWorkbook(bytes, { legacyXls })` sessions offer the same `getImage`
for worksheet pictures, and `openPptxPresentation` sessions already do.

To read pictures without rendering, and so without a canvas `factory`, open
a picture session instead. It parses the document but creates no layout
state, and offers only `getImage`, `resourceUsage` and `close`:

```typescript
import { materializeDocxDocument, openDocxPictures } from '@silurus/ooxml/node';

const document = await materializeDocxDocument(legacyDocBytes, { legacyDoc });
const pictures = await openDocxPictures(legacyDocBytes, { legacyDoc });
try {
  const picture = await pictures.getImage(run.imagePath, run.mimeType);
} finally {
  await pictures.close();
}
```

`openDocxDocument` still requires a `factory`, because pagination measures
text.

Both browser rendering modes and progressive layout use the same retained
layout pipeline. The native source remains alive for image reads until the
document is destroyed; finishing a model cursor does not dispose it.

This is a narrow experimental reader, not full DOC support. Unsupported
formatting, fields, numbering, notes and other unimplemented structures may
reject the entire document.

When a DOC prints its revision markup (MS-DOC `DopBase.fRMPrint`) and carries
revision marks, the reader reports `showTrackedChanges: true` as its view
default. An explicit `showTrackedChanges` option, including `false`, takes
precedence.

Missing header and footer distances use the MS-DOC 2.6.4 defaults for the
stored producer installation language when the specification lists that
language; explicit values, including zero, take priority. The host locale and
the document's text language are not used. A missing distance that no listed
language resolves rejects the document.

Fields are never executed. Visible fields map onto the fields the DOCX reader
supports:

- PAGE and NUMPAGES are computed when the document is laid out.
- DATE and TIME show the current date or time when the field has a date-time
  picture the renderer can format; otherwise they show the result stored in
  the file.
- A locked PAGE, NUMPAGES, DATE or TIME field is not recomputed; it shows its
  stored result unless that result is marked private.
- EQ fields become phonetic guides or inline equations.
- HYPERLINK fields whose instruction the reader can parse, and REF or PAGEREF
  fields with a bookmark and `\h`, become links. Stored results are not
  recomputed, so a PAGEREF number can differ from a PDF that Word produced
  after updating it.
- A visible MACROBUTTON, GOTOBUTTON, ADVANCE or FORMDROPDOWN field rejects the
  document. A field nested inside another field's instruction is hidden and
  is not checked against this list.

Footnotes and endnotes become DOCX notes. Word numbers notes per section;
when that numbering is not what the document-wide model produces, the reader
writes each automatic mark as text, and in those marks number formats other
than Arabic, Roman, letters and Chicago symbols are written as Arabic
numbers. A per-page restart counts continuously across pages. A custom
reference mark shows its own text and uses no number.

Section text flow is read per section. Horizontal-in-vertical runs are
rendered only in body paragraphs of sections with upright East Asian vertical
text (`tbRl`); in other sections, in table cells and in headers or footers
they reject the document. For upward-flowing sections, the
placement of headers, footers, notes, anchored drawings and upright tables
follows the library's general vertical placement, without a claim of
matching Word.

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
PowerPoint implementation. Besides Markdown and ZIP accounting, it does not
load the presentation's embedded fonts.

A slide's own `SlideShowSlideInfoAtom.fHidden` marks it hidden; hidden slides
and their content remain in the presentation. Display follows the existing
`PptxViewer` `hiddenSlideMode` option, whose default remains `'show'`.

An embedded or linked OLE object is drawn from the presentation picture the
file stores for it; the object itself is never read or activated. Some OLE
structures, for example an OLE group shape, reject the document.

Some modern Office-saved PPT files also store DrawingML for a shape in an
alternative shape XML part. Depending on how it compares with the binary
shape, the reader draws the alternative or the binary projection, and some
disagreements reject the load. The comparison rules are documented in the
reader source (`packages/legacy-office/parser/src/ppt/metro.rs`).

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

The direct XLS source is experimental and bounded. It projects a BIFF8
workbook into the shared worksheet model: sheet names and visibility, cell
values with their cached formula results, merged ranges, cell styles,
shared-string character formatting and phonetic guides, row and column
geometry, sheet views, conditional formatting, tables, hyperlinks, filters,
data validations, defined names, cell comments (shown as comment indicators),
PivotTable layout, charts, chart sheets, pictures and shapes. Formula token
programs, external links and macros are never evaluated.

Some content without a model equivalent is left out while the workbook still
loads, for example AutoFilter criteria, list validations whose drop-down is
suppressed, the comment box Excel keeps visible on the sheet, and drawing
shadows. Other structures the model cannot state reject the workbook, for
example inactive conditional-formatting rules. Page setup, headers, footers
and page breaks are not part of the model; print areas and titles are
defined names.

Worksheet visibility, including very hidden sheets, is kept as model
metadata; display follows the existing `XlsxViewer` `hiddenSheetMode` option,
whose default remains `'show'`. Of the sheet views, only display flags and
frozen panes are projected.

A cell-anchored drawing stores each corner as a row and column plus signed
offsets in 1/1024 of the column width and 1/256 of the row height. Resolving
column widths needs the Normal font's maximum digit width (ECMA-376
§18.3.1.13). The XLS source asks the host for it through a generic
host-layout capability, and the XLSX renderer answers with its own
`computeMdw`, the same measurement that sizes the painted grid, in the
rendering realm: the render Worker in `mode: 'worker'`, the page in main
mode, and the session canvas `factory` in Node. There is no measurement
callback option. In Node without a `factory` nothing can measure the font
and the session renders nothing, so anchors resolve with the grid's fallback
width of 8 pixels, the width the grid itself uses when it cannot measure.
That width is recorded as the workbook's `layoutMetrics`, so a later render
sizes the grid to match; the drawings and their `getImage` bytes stay
available, but column positions are approximate when the Normal font's
real digit width differs. When a `factory` is given but its measurement is
unusable, drawings that need the worksheet grid to resolve their anchors are
omitted; a chart sheet's charts use absolute anchors and remain.

XLS worksheets are delivered complete, without an early row preview.

## Current implementation boundary

The repository contains only the direct readers; legacy input is never
converted to OOXML. The Rust crate `legacy-office-parser` builds one reader
per feature: `direct-doc`, `direct-xls` and `direct-ppt`. Building it for
`wasm32` with none of them enabled is a compile error. The `inspection`
feature adds native-only development examples, and `fuzzing` exposes the fuzz
entry points.

Treat the rendered model as a derived view of the original binary, keep the
binary as the authoritative source, and gate production use on a corpus
representative of the documents being ingested. The readers' internal budgets
are library resource policy, not Office format restrictions.
