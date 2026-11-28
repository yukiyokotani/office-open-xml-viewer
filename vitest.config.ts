import { defineConfig } from 'vitest/config';

// Unit tests only. `*.spec.ts` is reserved for the Playwright visual-regression
// suites (run via `pnpm vrt`), so restrict vitest to `*.test.ts` under src and
// keep it out of the tests/visual directories.
export default defineConfig({
  resolve: { alias: { '@silurus/ooxml/internal/companion': new URL('./src/internal-companion.ts', import.meta.url).pathname } },
  define: { __OOXML_MODEL_SOURCES__: 'true' },
  test: {
    include: ['packages/**/src/**/*.test.ts', 'site/src/**/*.test.ts', 'tests/asset-sidecar-build.test.ts'],
    exclude: ['**/node_modules/**', '**/dist/**', '**/tests/visual/**'],
    setupFiles: ['tests/vitest-locale-setup.ts'],
  },
});
