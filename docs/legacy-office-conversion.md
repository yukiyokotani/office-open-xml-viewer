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
| DOC | CFB Word 97-2003 documents with a readable main-story CLX piece table | main-story text, paragraphs, tabs, line/page/column breaks, displayed field results, section boundaries, page size/orientation, explicit body margins and gutter, columns, vertical alignment, document grid | character/paragraph formatting, advanced section properties, headers/footers, notes, lists, tables, drawings, revisions, OLE; non-Western compressed code-page pieces are not decoded yet |
| XLS | CFB BIFF8 workbooks, including shared-string character data split across `CONTINUE` records | worksheet names, scalar values, cached formula results, merged ranges, date system, BIFF8 number formats, fonts, palette colors, fills, borders, alignment, styled blank cells, row heights and column widths, row/column hiding and outlines, print setup/margins/options, basic header/footer commands and manual page breaks | formula programs, rich-text runs, extended styles/themes/gradients, conditional formatting, print areas/titles, extended headers/footers, saved custom views, charts, drawings, external links, pre-BIFF8 sheets |
| PPT | CFB PowerPoint 97-2003 files with a resolvable current edit chain and persist directory | live slide order, slide dimensions, Unicode/Windows-1252 text, outline text references, slide boundaries; superseded and deleted slides are not emitted | masters/layout fidelity, formatting, shapes, charts, notes, media, transitions, animations, actions, OLE |

Version-3 and version-4 CFB containers are admitted. Password-protected legacy
binaries and pre-CFB Office formats are rejected. These limits are structural,
not filename-based. Unsupported binary structures fail with
`reason === 'unsupported-input'`. Accepted documents can still lose the features
listed above: their warning identifiers are not a fidelity certificate.

Every output package is created from scratch and contains no source macro,
VBA/Excel 4.0 program, ActiveX control, OLE object, hyperlink action, or external
relationship. The converter never evaluates formulas, fields, actions, links,
or macros. Fixed, content-free warning identifiers report the intentional loss
class in the conversion provenance record.

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
are distinct from section breaks. Missing header/footer distances use zero only
for the currently omitted header/footer stories, with an explicit warning;
known body margins are retained. The producer's locale defaults are not guessed.
XLS saved custom-view print records cannot override the active worksheet's print
settings, and undefined printer fields are not emitted. These changes do not add
legacy-specific renderer behavior or change per-format opt-in defaults.

Resource policy for PPT reconstruction additionally limits each of retained
outline text and emitted slide text to 128 MiB, and charges persist-directory
entries against the record-work budget. Repeated references cannot bypass the
text limit. XLS style tables are bounded, repeated fills/borders are interned,
and the BIFF column-256 default-format sentinel never creates an extra column.

Treat converted OOXML as a derived search/view representation, preserve the
original binary as the authoritative source, and gate production use on a
corpus representative of the documents being ingested. A custom local or remote
adapter remains supported for applications that need a broader conversion
engine; the library still supplies no remote endpoint and never silently uploads
document bytes.
