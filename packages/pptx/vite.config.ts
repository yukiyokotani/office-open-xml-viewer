import { defineConfig } from 'vite';
import wasm from 'vite-plugin-wasm';
import { resolve } from 'path';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { wasmAssetUrl } from '../../vite.config';

const dirname =
  typeof __dirname !== 'undefined'
    ? __dirname
    : path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  plugins: [wasm()],
  resolve: {
    alias: {
      '@ooxml-test-three-d-renderer': resolve(dirname, '../../src/three-d.ts'),
      '@ooxml-test-region-map-renderer': resolve(dirname, '../../src/region-map.ts'),
      '@ooxml-test-math-renderer': resolve(dirname, '../../src/math.ts'),
      '@ooxml-test-tiff-renderer': resolve(dirname, '../../src/tiff.ts'),
    },
  },
  server: {
    fs: {
      // Include monorepo root so node_modules/.pnpm/ fontsource files can be served
      // (pnpm symlinks to ../../node_modules/.pnpm/...).
      allow: [dirname, path.resolve(dirname, '../..')],
    },
  },
  build: {
    // Serve public/ (sample fixtures) from the dev server for VRT, but don't
    // copy it into the published dist/.
    copyPublicDir: false,
    lib: {
      entry: resolve(dirname, 'src/index.ts'),
      name: 'OoxmlViewer',
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
    plugins: () => [wasmAssetUrl(), wasm()],
    rollupOptions: {
      output: { assetFileNames: '[name][extname]' },
    },
  },
});
