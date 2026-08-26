import { defineConfig } from 'astro/config';
import wasm from 'vite-plugin-wasm';
import topLevelAwait from 'vite-plugin-top-level-await';
import { fileURLToPath } from 'node:url';

const pkgSrc = (p) => fileURLToPath(new URL(`../packages/${p}/src/index.ts`, import.meta.url));

// GitHub Pages base path. Custom domain => '/', project pages => '/office-open-xml-viewer/'.
const SITE_BASE = process.env.SITE_BASE ?? '/';

// https://astro.build
export default defineConfig({
  base: SITE_BASE,
  trailingSlash: 'ignore',
  vite: {
    plugins: [wasm(), topLevelAwait()],
    worker: {
      format: 'es',
      plugins: () => [wasm(), topLevelAwait()],
    },
    resolve: {
      // Pull the workspace packages from source so Vite processes their
      // `?worker&inline` / `?url` imports (same flow as Storybook).
      alias: {
        // Public package entry used by executable documentation examples.
        // Resolve it to the same source entry as the site's internal package
        // import so the displayed code is also the code that runs.
        '@silurus/ooxml/docx': pkgSrc('docx'),
        '@silurus/ooxml-pptx': pkgSrc('pptx'),
        '@silurus/ooxml-xlsx': pkgSrc('xlsx'),
        '@silurus/ooxml-docx': pkgSrc('docx'),
        // Keep core subpaths ahead of the package-root prefix alias; otherwise
        // Vite appends them to the root entry file (for example,
        // `src/index.ts/internal/resource-measurement`).
        '@silurus/ooxml-core/internal/resource-measurement': fileURLToPath(
          new URL('../packages/core/src/internal/resource-measurement.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/bounded-async-lru-cache': fileURLToPath(
          new URL('../packages/core/src/internal/bounded-async-lru-cache.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/bounded-raw-part-cache': fileURLToPath(
          new URL('../packages/core/src/internal/bounded-raw-part-cache.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/canvas-viewer-mechanics': fileURLToPath(
          new URL('../packages/core/src/internal/canvas-viewer-mechanics.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/chart-context': fileURLToPath(
          new URL('../packages/core/src/internal/chart-context.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/comment-context': fileURLToPath(
          new URL('../packages/core/src/internal/comment-context.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/read-only-comment-contract': fileURLToPath(
          new URL('../packages/core/src/internal/read-only-comment-contract.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/read-only-comment-margin': fileURLToPath(
          new URL('../packages/core/src/internal/read-only-comment-margin.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/read-only-comment-decoration': fileURLToPath(
          new URL('../packages/core/src/internal/read-only-comment-decoration.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/virtual-scroll': fileURLToPath(
          new URL('../packages/core/src/internal/virtual-scroll.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/dom-interaction-boundary': fileURLToPath(
          new URL('../packages/core/src/internal/dom-interaction-boundary.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/dom-geometry': fileURLToPath(
          new URL('../packages/core/src/internal/dom-geometry.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/internal/script-preload-accumulator': fileURLToPath(
          new URL('../packages/core/src/internal/script-preload-accumulator.ts', import.meta.url),
        ),
        '@silurus/ooxml-core/worker': fileURLToPath(
          new URL('../packages/core/src/worker/index.ts', import.meta.url),
        ),
        '@silurus/ooxml-core': pkgSrc('core'),
      },
    },
  },
});
