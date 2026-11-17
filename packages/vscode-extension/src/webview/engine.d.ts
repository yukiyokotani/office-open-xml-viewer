// The prebuilt MathJax + STIX Two Math engine is a self-contained IIFE asset
// (no exports) that sets `globalThis.__ooxmlStix2` when evaluated. We import it
// for its side effect so esbuild bundles it into the webview (the library's
// normal lazy <script> injection is blocked by the webview's nonce CSP).
declare module '@silurus/ooxml-core/mathjax-stix2';

interface Stix2Engine {
  mathml2svg(mathml: string): string;
}
// eslint-disable-next-line no-var
declare var __ooxmlStix2: Stix2Engine | undefined;

// Parser WASM binaries are bundled with esbuild's `file` loader: the default
// export is the emitted asset's URL, relative to the webview's <base>.
declare module '@silurus/ooxml-docx/wasm-binary' {
  const url: string;
  export default url;
}
declare module '@silurus/ooxml-xlsx/wasm-binary' {
  const url: string;
  export default url;
}
declare module '@silurus/ooxml-pptx/wasm-binary' {
  const url: string;
  export default url;
}
