/// <reference types="vite/client" />

// Brings Vite's ambient module declarations (e.g. `declare module '*?url'`) into
// scope so the `?url` asset import in `math/engine.ts` type-checks. The bundle is
// emitted as a real asset (not a base64 data URL) by the root `wasmAssetUrl`
// build plugin (see `vite.config.ts`); this file only supplies the
// `import … from '…?url'` typing.

// Vite's built-in `*?url` declaration does not match the explicit inline
// modifier used by the optional bundled Office substitute chunks.
declare module '*.ttf?url&inline' {
  const url: string;
  export default url;
}
