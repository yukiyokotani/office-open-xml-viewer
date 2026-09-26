import { defineConfig } from 'rolldown';

// `@silurus/ooxml-core` ships TypeScript source only, so the compiled entry
// inlines the typed-error helpers it uses; the parser glue stays external and
// resolves through each format package's `./wasm` export.
export default defineConfig({
  input: 'src/index.ts',
  external: /^@silurus\/ooxml-(docx|pptx|xlsx)\/wasm$/,
  output: { file: 'dist/index.js', format: 'esm' },
});
