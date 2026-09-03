# @silurus/ooxml-markdown

Opt-in GitHub-flavoured Markdown projection for OOXML documents. Each format
entry point loads only its matching parser and dedicated WebAssembly binary;
the main `@silurus/ooxml` viewer package does not include this code.

```typescript
import { initFromBytes, toMarkdown } from '@silurus/ooxml-markdown/pptx';
import wasmUrl from '@silurus/ooxml-markdown/pptx/wasm-binary?url';

initFromBytes(await fetch(wasmUrl).then((response) => response.arrayBuffer()));
const markdown = toMarkdown(await file.arrayBuffer());
```

Replace `/pptx` with `/docx` or `/xlsx` for the other formats. In Node.js, the
`ooxml-md input.pptx` CLI resolves and initializes the matching binary for you.
