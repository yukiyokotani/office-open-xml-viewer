# Vendored Formualizer engine glue (static assets, no bundler WASM import)

`formualizer_wasm_bg.js` + `formualizer_wasm_bg.wasm` are byte-identical
copies of the corresponding files from the `formualizer@0.13.0` npm package
(`pkg/`). They are vendored here instead of imported from the package
because the package's entry (`pkg/formualizer_wasm.js`) loads its WASM via
wasm-pack `--target bundler` output — `import * as wasm from
"./formualizer_wasm_bg.wasm"`, an ES-module WASM import that bundlers
without a WASM plugin (notably the Angular application builder / esbuild,
which accept no custom plugins) reject at build time:

> "ESM integration proposal for Wasm" is not supported currently.

The wrapper itself (`formualizer_wasm_bg.js`) is plain JavaScript with no
module imports, so it bundles anywhere. `../formualizer-loader.ts` fetches
the sibling `.wasm` bytes (emitted as a real asset file by the repo's
`?url` handling, staged by consumers exactly like the parser WASMs,
optionally overridden via `LoadOptions.formualizerWasm`) and
instantiates them by hand:

- imports are derived at runtime from the binary (`WebAssembly.Module.imports`),
  all from the wrapper's own shim namespace;
- then `__wbg_set_wasm(instance.exports)` + `__wbindgen_start()`, mirroring
  the package's own glue line-for-line.

When upgrading Formualizer, re-copy both files from the new package and
re-verify with the recalc suite (`src/recalc.test.ts`); the runtime-derived
imports tolerate added/renamed shims as long as the wrapper still exports
them. License: MIT OR Apache-2.0, see `THIRD_PARTY_NOTICES.md`.
