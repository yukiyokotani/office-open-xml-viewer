# @silurus/ooxml-opendocument

Experimental OpenDocument model sources for `@silurus/ooxml`. Both packages
must have the same version.

```sh
npm i @silurus/ooxml @silurus/ooxml-opendocument
```

```ts
import { DocxViewer } from '@silurus/ooxml/docx';
import { odt } from '@silurus/ooxml-opendocument/odt';

const canvas = document.getElementById('viewer') as HTMLCanvasElement;
const viewer = new DocxViewer(canvas, { odt });
```

The reader maps OpenDocument text (`.odt`, `.ott`) directly into the DOCX
document model; the ordinary DOCX layout and Canvas renderer draw it. Importing
the reader loads nothing: its source module and WASM load only when an ODT file
is opened. When an asset pipeline relocates the bundled files, create the
reader with `createOdt({ wasmUrl, moduleUrl })` and pass it as the same option.

For presentations (`.odp`, `.otp`) pass the `odp` reader to the PPTX viewers
and loaders:

```ts
import { PptxViewer } from '@silurus/ooxml/pptx';
import { odp } from '@silurus/ooxml-opendocument/odp';

const viewer = new PptxViewer(canvas, { odp });
```

It maps the presentation directly into the PPTX slide model and renders every
feature it reads, rejecting only corrupt, encrypted or over-limit packages
(see `docs/opendocument.md`).

The ODT reader is fail-closed: a document that uses anything outside the admitted
subset is rejected as a whole with `OoxmlError.code === 'unsupported-content'`
instead of being shown partially. The admitted subset is currently narrow
(paragraphs, headings and character formatting of Latin-script text on one page
style) and grows release by release.
