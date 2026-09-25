import { defineConfig } from 'vitest/config';

// Unit tests only. `*.spec.ts` is reserved for the Playwright visual-regression
// suites (run via `pnpm vrt`), so restrict vitest to `*.test.ts` under src and
// keep it out of the tests/visual directories.
export default defineConfig({
  test: {
    include: ['packages/**/src/**/*.test.ts', 'site/src/**/*.test.ts', 'tests/asset-sidecar-build.test.ts'],
    exclude: ['**/node_modules/**', '**/dist/**', '**/tests/visual/**'],
  },
});
