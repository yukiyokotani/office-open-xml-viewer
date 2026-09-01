import { defineConfig, type Plugin } from 'vite';
import wasm from 'vite-plugin-wasm';
import { dts } from 'rolldown-plugin-dts';
import { resolve, dirname, basename } from 'path';
import { fileURLToPath } from 'url';
import { readFile } from 'fs/promises';

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Emit `?url` asset imports as real asset files instead of base64 `data:` URLs.
 * (Named hashlessly here — `rollupOptions.output.assetFileNames` is `[name]`.)
 *
 * Vite **library mode** force-inlines every `?url` asset as a
 * `data:<mime>;base64,…` string regardless of `assetsInlineLimit` (a number or a
 * `() => false` function does NOT override it — `build.lib` unconditionally
 * returns `true` from Vite's internal `shouldInline`). Two heavy asset kinds ride
 * on this path:
 *   - the three parser WASM modules (`*_parser_bg.wasm?url`, ~0.6–0.7 MB each) —
 *     base64 inflates them +33 % and blocks `WebAssembly.compileStreaming`
 *     (a data URL cannot be fetch-streamed; the worker must `atob` by hand);
 *   - the MathJax + STIX Two Math engine (`assets/mathjax-stix2.js?url`, ~3 MB) —
 *     inlined it turned the opt-in `math.mjs` chunk into a 4.1 MB base64 blob,
 *     even though consumers only import it when a document actually has equations.
 *
 * All of these are `?url` imports in a single owner module each (the format
 * main-thread handles — `document.ts` / `presentation.ts` / `workbook.ts` — and
 * `math/engine.ts`). We intercept the `?url` variant here, `emitFile` the bytes
 * as an asset next to the chunk, and hand back the standard ESM asset reference
 * `new URL('<name>', import.meta.url)` — the form Vite / webpack 5 / Rollup /
 * Turbopack rewrite when they re-bundle our `.mjs`, and which resolves
 * correctly for a plain `<script type=module>` too. (esbuild — and therefore
 * the Angular CLI — does NOT process it; those consumers use the `wasmUrl`
 * load option, see the README bundler note.) wasm-bindgen's `--target web`
 * glue then `fetch()`es its URL and hits `instantiateStreaming`; the math engine
 * is lazy-loaded via a `<script src>` pointed at the emitted asset.
 *
 * Runs with `enforce: 'pre'` and claims every `?url` import; a bare-`.wasm`
 * import (owned by `vite-plugin-wasm`) is untouched. Nothing in the tree imports
 * bare `.wasm`, and every `?url` here is a real on-disk asset we want emitted —
 * exactly what Vite's non-lib mode would do anyway.
 */
export function wasmAssetUrl(): Plugin {
  const SUFFIX = '?url';
  return {
    name: 'wasm-asset-url',
    enforce: 'pre',
    // Build-only: emitFile/ROLLUP_FILE_URL are Rollup build machinery and do
    // not exist on the dev server. In dev, Vite's stock `?url` handling serves
    // the file directly (the pre-E4 behavior) — intercepting there returned an
    // unresolvable reference and broke every WASM load (caught by CI smoke).
    apply: 'build',
    async load(id) {
      if (!id.endsWith(SUFFIX)) return null;
      const filePath = id.slice(0, -SUFFIX.length);
      const source = await readFile(filePath);
      const referenceId = this.emitFile({
        type: 'asset',
        name: basename(filePath),
        source,
      });
      // `import.meta.ROLLUP_FILE_URL_<id>` expands at render time to Rollup's
      // ES default resolution — `new URL('<name>', import.meta.url).href` — an
      // absolute href the worker (or the math engine's `<script>` loader) can
      // fetch from any realm. It must be emitted BARE: wrapping it in another
      // `new URL(…, import.meta.url)` ships a nested pattern that webpack 5
      // compiles into a critical-dependency ContextModule (the outer first arg
      // is now an expression, not a literal), throwing MODULE_NOT_FOUND at
      // module evaluation — `import '@silurus/ooxml/docx'` alone crashed every
      // webpack consumer. The single-level form is the exact shape webpack 5 /
      // Turbopack / Vite statically rewrite when re-bundling our `.mjs`.
      return `export default import.meta.ROLLUP_FILE_URL_${referenceId};`;
    },
  };
}

export default defineConfig(({ command, mode }) => ({
  // Published library assets must resolve from the imported module URL, not
  // from the hosting page's origin root. This is especially important for the
  // standalone module workers and sibling assets when consumers serve the
  // package below a subpath or from a CDN.
  base: './',
  plugins: [
    wasmAssetUrl(),
    wasm(),
    // Storybook loads the root Vite config in serve mode. The declaration
    // plugins are build-only: their Rolldown buildStart hooks expect library
    // inputs and fail against Storybook's dev-server graph.
    ...(command === 'build' && mode !== 'runtime'
      ? dts({
          // TypeScript 7 is the repository's sole compiler. Its native tsgo
          // declaration generator avoids the removed JavaScript Compiler API.
          generator: 'tsgo',
          tsconfig: './tsconfig.lib.json',
        })
      : []),
  ],
  // rolldown-plugin-dts emits declarations into the same graph. Oxc should
  // leave those declaration modules (and normal JavaScript inputs) untouched.
  oxc: {
    exclude: [/\.js$/, /\.d\.[cm]?ts$/],
  },
  build: {
    lib: {
      entry: {
        index: resolve(__dirname, 'src/index.ts'),
        pptx:  resolve(__dirname, 'src/pptx.ts'),
        xlsx:  resolve(__dirname, 'src/xlsx.ts'),
        docx:  resolve(__dirname, 'src/docx.ts'),
        // Opt-in math engine (MathJax + STIX Two Math). Separate entry so the
        // ~3 MB asset stays out of the docx/pptx bundles unless imported.
        math:  resolve(__dirname, 'src/math.ts'),
        // Opt-in model-space 3-D chart mesh/camera painter. Ordinary format
        // bundles keep the lightweight contract and 2-D fallback only.
        'three-d': resolve(__dirname, 'src/three-d.ts'),
        // Opt-in offline ChartEx Region Map geometry/projector.
        'region-map': resolve(__dirname, 'src/region-map.ts'),
        // Opt-in Microsoft ChartEx family renderer. Classic 2-D charts remain
        // part of every format entry.
        'chart-ex': resolve(__dirname, 'src/chart-ex.ts'),
        // Opt-in TIFF 6.0 software decoder. Native raster users retain only the
        // lightweight codec contract and header guard.
        tiff: resolve(__dirname, 'src/tiff.ts'),
        // Node-only bounded sessions and server render helpers. Kept as a
        // separate entry so browser consumers never load Node built-ins.
        node:  resolve(__dirname, 'src/node.ts'),
      },
      // ESM-only: the published bundle inlines a large math engine; emitting a
      // duplicate CJS copy of every chunk roughly doubled the package size.
      // Every modern bundler (Vite / webpack / Rollup / esbuild / Next) and
      // Node ≥ 20 consume ESM, so we ship `.mjs` only.
      formats: ['es'],
      fileName: (_format, name) =>
        name.endsWith('.d')
          ? `.types-work/${name.slice(0, -2)}.d.ts`
          : `${name}.mjs`,
    },
    rollupOptions: {
      external: [/^node:/, 'skia-canvas'],
      output: {
        assetFileNames: '[name][extname]',
        chunkFileNames: (chunk) =>
          chunk.name.endsWith('.d')
            ? '.types-work/[name]-[hash].d.ts'
            : '[name]-[hash].js',
      },
    },
    target: 'esnext',
  },
  worker: {
    format: 'es',
    // Built-in worker renderers lazy-import the same optional math engine. Keep
    // its ~3 MB `?url` asset external in nested worker builds too; otherwise
    // library mode base64-inlines one copy into every format worker chunk.
    plugins: () => [wasmAssetUrl(), wasm()],
    rollupOptions: {
      output: {
        assetFileNames: '[name][extname]',
        // The published format entry is commonly re-bundled by a consumer.
        // A prebuilt worker URL is an opaque asset to that second bundler, so
        // sibling JS chunks imported by the worker would not be copied. Keep
        // each render worker as one self-contained module asset; optional
        // renderer code is fetched only when mode:'worker' loads this asset.
        codeSplitting: false,
      },
    },
  },
}));
