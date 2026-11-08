import { defineConfig } from 'vite';
import wasm from 'vite-plugin-wasm';
import { resolve } from 'path';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { legacyBundleBoundary, wasmAssetUrl } from '../../vite.config';

const dirname =
  typeof __dirname !== 'undefined'
    ? __dirname
    : path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  define: { __OOXML_MODEL_SOURCES__: 'true' },
  // Library assets (parser WASM, workers) must resolve from the importing
  // module's URL, never from the hosting page's origin root, as in the root
  // build. Without it Vite 8.3 emits root-absolute "/<asset>" references.
  base: './',
  plugins: [wasmAssetUrl(), wasm(), legacyBundleBoundary()],
  root: dirname,
  resolve: {
    alias: {
      '@ooxml-test-three-d-renderer': resolve(dirname, '../../src/three-d.ts'),
      '@ooxml-test-region-map-renderer': resolve(dirname, '../../src/region-map.ts'),
      '@ooxml-test-math-renderer': resolve(dirname, '../../src/math.ts'),
      '@ooxml-test-tiff-renderer': resolve(dirname, '../../src/tiff.ts'),
      '@ooxml-test-chart-ex-renderer': resolve(dirname, '../../src/chart-ex.ts'),
    },
  },
  server: { port: 5175, strictPort: true },
  build: {
    // Serve public/ (sample fixtures) from the dev server for VRT, but don't
    // copy it into the published dist/.
    copyPublicDir: false,
    lib: {
      entry: resolve(dirname, 'src/index.ts'),
      name: 'XlsxViewer',
      formats: ['es', 'cjs'],
      fileName: (format) => `index.${format === 'es' ? 'mjs' : 'cjs'}`,
    },
    target: 'esnext',
    rollupOptions: {
      output: {
        assetFileNames: '[name][extname]',
      },
    },
  },
  worker: {
    format: 'es',
    plugins: () => [wasmAssetUrl(), wasm(), legacyBundleBoundary()],
    rollupOptions: {
      output: { assetFileNames: '[name][extname]' },
    },
  },
});
