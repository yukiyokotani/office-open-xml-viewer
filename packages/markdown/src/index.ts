/**
 * Convert .pptx / .docx / .xlsx to GitHub-flavoured markdown via the
 * workspace WASM parsers. The underlying `pptx_to_markdown` /
 * `docx_to_markdown` / `xlsx_to_markdown` functions are wasm_bindgen exports
 * that call the Rust markdown projection directly (no JSON round-trip).
 *
 * Usage modes:
 *
 *   - **Node / workspace CLI**: call the matching `init*FromBytes()` once at
 *     startup, then pass file bytes to the matching `*ToMarkdown` function.
 *
 *   - **Browser bundlers (Vite / webpack)**: import the `.wasm` URL via the
 *     bundler-specific syntax (`?url` in Vite), fetch it, and call
 *     `init*FromBytes()`. Then call `*ToMarkdown` on a `File.arrayBuffer()`.
 */

// Resolve each parser's wasm-bindgen glue through its package `exports` map
// (`./wasm` → `src/wasm/<fmt>_parser.js`) rather than a monorepo-relative
// sibling path. The package remains an internal workspace adapter, but using
// package boundaries keeps its development CLI and tests independent of the
// monorepo's physical directory layout. This mirrors the CLI, which reads the
// raw binary through the sibling `./wasm-binary` export.
// @ts-ignore — wasm-pack JS shim, exports typed individually below
import * as pptxWasm from '@silurus/ooxml-pptx/wasm';
// @ts-ignore
import * as docxWasm from '@silurus/ooxml-docx/wasm';
// @ts-ignore
import * as xlsxWasm from '@silurus/ooxml-xlsx/wasm';

type WasmModule = {
  initSync: (init: { module: WebAssembly.Module }) => unknown;
};

const pptxState = { initialized: false };
const docxState = { initialized: false };
const xlsxState = { initialized: false };

function syncInit(mod: WasmModule, bytes: Uint8Array, state: { initialized: boolean }): void {
  if (state.initialized) return;
  // The DOM TS lib types `WebAssembly.Module(...)` as `BufferSource =
  // ArrayBufferView<ArrayBuffer> | ArrayBuffer`, which excludes
  // `Uint8Array<SharedArrayBuffer>`. We always pass non-shared bytes here
  // (read from `readFileSync` or `fetch().arrayBuffer()`), so the cast is
  // safe.
  const module = new WebAssembly.Module(bytes as unknown as ArrayBuffer);
  mod.initSync({ module });
  state.initialized = true;
}

/** Initialise from raw WASM bytes. Works in both Node and browser. */
export function initPptxFromBytes(bytes: ArrayBuffer | Uint8Array): void {
  syncInit(pptxWasm as WasmModule, toUint8(bytes), pptxState);
}
export function initDocxFromBytes(bytes: ArrayBuffer | Uint8Array): void {
  syncInit(docxWasm as WasmModule, toUint8(bytes), docxState);
}
export function initXlsxFromBytes(bytes: ArrayBuffer | Uint8Array): void {
  syncInit(xlsxWasm as WasmModule, toUint8(bytes), xlsxState);
}

// The public buffer type is `ArrayBuffer | Uint8Array` ONLY. A Node `Buffer`
// IS a `Uint8Array`, so `readFileSync` output is accepted as-is — but naming
// `Buffer` in an exported signature would bake a `@types/node` global into the
// published `.d.ts` and break browser consumers that (correctly) do not
// install Node types.
function toUint8(buffer: ArrayBuffer | Uint8Array): Uint8Array {
  if (ArrayBuffer.isView(buffer)) {
    return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  }
  return new Uint8Array(buffer);
}

/** Convert a `.pptx` archive's bytes to GitHub-flavoured markdown. Title
 *  slides become `# heading`s, body shapes become nested bullets at the
 *  paragraph's `lvl`, tables become pipe tables, charts become summarised
 *  bullets, speaker notes and comments are collated. */
export function pptxToMarkdown(buffer: ArrayBuffer | Uint8Array): string {
  if (!pptxState.initialized) throw new Error('pptx wasm not initialized — call initPptxFromBytes() first');
  return (pptxWasm as unknown as { pptx_to_markdown: (b: Uint8Array) => string }).pptx_to_markdown(toUint8(buffer));
}

/** Convert a `.docx` archive's bytes to GitHub-flavoured markdown. Headings
 *  come from `<w:outlineLvl>`, lists honour the abstractNum format, tables
 *  preserve vMerge continuation, footnotes/endnotes/comments are collated. */
export function docxToMarkdown(buffer: ArrayBuffer | Uint8Array): string {
  if (!docxState.initialized) throw new Error('docx wasm not initialized — call initDocxFromBytes() first');
  return (docxWasm as unknown as { docx_to_markdown: (b: Uint8Array) => string }).docx_to_markdown(toUint8(buffer));
}

/** Convert a `.xlsx` archive's bytes to GitHub-flavoured markdown. Each
 *  sheet becomes a `## SheetName` section followed by a pipe table of its
 *  populated bbox; fully-empty middle rows are trimmed, ULP noise is
 *  masked. */
export function xlsxToMarkdown(buffer: ArrayBuffer | Uint8Array): string {
  if (!xlsxState.initialized) throw new Error('xlsx wasm not initialized — call initXlsxFromBytes() first');
  return (xlsxWasm as unknown as { xlsx_to_markdown: (b: Uint8Array) => string }).xlsx_to_markdown(toUint8(buffer));
}
