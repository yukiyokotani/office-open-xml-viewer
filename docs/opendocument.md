# Experimental OpenDocument sources

Install the viewer and its optional companion at matching versions:

```sh
npm i @silurus/ooxml @silurus/ooxml-opendocument
```

The companion pins its `@silurus/ooxml` peer to the exact same version. A
stale runtime or model-source protocol fails before input loading with
`CompanionRuntimeMismatchError` (`code: 'companion-runtime-mismatch'`).

OpenDocument text can be opened with the ordinary DOCX viewers and loaders
through an optional model source:

| Input | Entry | Reader / factory | Option | Loaders and viewers |
| --- | --- | --- | --- | --- |
| ODT, OTT | `@silurus/ooxml-opendocument/odt` | `odt` / `createOdt()` | `odt` | `DocxViewer`, `DocxScrollViewer`, `DocxDocument.load`, `openDocxDocument`, `materializeDocxDocument` |

```typescript
import { DocxViewer } from '@silurus/ooxml/docx';
import { odt } from '@silurus/ooxml-opendocument/odt';

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new DocxViewer(canvas, { odt });
await viewer.load(odtOrDocxBytes);
```

The reader maps the document directly into the DOCX document model, without
generating an OOXML package; the ordinary layout and Canvas renderer draw it.
Importing the entry or calling `createOdt()` fetches nothing. When an ODT file
is opened, the parser Worker (or Node) imports the reader's self-contained ES
module (`odt-source-module.mjs`) and its WASM (`opendocument_odt_bg.wasm`).
Serve these files with the other package assets and allow the module URL in a
Content Security Policy. Applications with a custom asset pipeline pass
absolute URLs:

```typescript
import { createOdt } from '@silurus/ooxml-opendocument/odt';

const odt = createOdt({
  moduleUrl: new URL('/assets/odt-source-module.mjs', location.href).href,
  wasmUrl: new URL('/assets/opendocument_odt_bg.wasm', location.href).href,
});
```

## Admission

The source claims a package only by its leading `mimetype` file, which a
packaged OpenDocument text document must carry (ODF 1.3 Part 2 §3.3, Part 3
§2.2.3). Other inputs, including other OpenDocument families, continue on the
OOXML path unchanged. The ZIP is read with the same structural checks and
inflation limits as OOXML packages, including the caller's `resourceLimits`;
item names are compared exactly. Encrypted packages reject with
`unsupported-encryption`.

## Supported content

The reader is fail-closed: a document that uses anything outside the admitted
subset rejects as a whole with `OoxmlError.code === 'unsupported-content'`
instead of being shown partially. It never executes scripts or fetches
external resources.

Currently admitted:

- paragraphs, headings (without heading numbering), plain sections, spans,
  spaces, tabs (in paragraphs without an indent, at the document's default
  tab interval) and line breaks between visible text set in the same font;
- page size and margins of a single page style without displayed headers or
  footers;
- paragraph indents, alignment, page breaks, keep-with-next /
  keep-together, widow/orphan control of two lines, and paragraph spacing
  including contextual spacing;
- character fonts, sizes (including a relative size without raise or
  lower), bold, italic, single/double underline and strike-through, colours,
  background colour, letter spacing and scaling;
- text in the Latin script class of ODF Part 3 Table 22.

Lists, tables, frames and images, notes, fields, comments, tracked changes,
headers and footers, Asian and complex scripts, and other formatting reject
for now. The subset grows release by release.
