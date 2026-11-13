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

The reader renders every OpenDocument text feature, following the ODF 1.3
specification and, where it leaves a choice, LibreOffice. It rejects only
corrupt, encrypted or unsafe input and input beyond the resource limits;
features the DOCX model cannot express exactly are drawn with their nearest
DOCX equivalent, so small differences from LibreOffice remain. See
[docs/opendocument.md](../../docs/opendocument.md) for what is rendered and
the known differences.
