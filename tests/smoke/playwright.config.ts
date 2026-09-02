import { defineConfig, devices } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const REPOSITORY_ROOT = fileURLToPath(new URL('../..', import.meta.url));
const SMOKE_PORT = Number(process.env.OOXML_SMOKE_PORT ?? 6007);
if (!Number.isInteger(SMOKE_PORT) || SMOKE_PORT < 1 || SMOKE_PORT > 65_535) {
  throw new Error(`invalid OOXML_SMOKE_PORT: ${process.env.OOXML_SMOKE_PORT}`);
}

export default defineConfig({
  testDir: '.',
  testMatch: '**/*.spec.ts',
  fullyParallel: false,
  reporter: [['list']],
  use: {
    baseURL: `http://localhost:${SMOKE_PORT}`,
    actionTimeout: 30_000,
  },
  // The smoke assertion is canvasHasInk (a count of non-white pixels), which is
  // font-independent, so it stays stable across engines. Running webkit and
  // firefox alongside chromium catches engine-specific breakage in the
  // parse -> render pipeline (OffscreenCanvas quirks, worker transfer, canvas
  // API gaps) that a Chrome-only smoke would miss. (VRT match-% comparisons,
  // which ARE font-sensitive, stay Chrome-only and local — see visual.spec.ts.)
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
        deviceScaleFactor: 1,
        viewport: { width: 1400, height: 900 },
      },
    },
    {
      name: 'webkit',
      use: {
        ...devices['Desktop Safari'],
        deviceScaleFactor: 1,
        viewport: { width: 1400, height: 900 },
      },
    },
    {
      name: 'firefox',
      use: {
        ...devices['Desktop Firefox'],
        deviceScaleFactor: 1,
        viewport: { width: 1400, height: 900 },
      },
    },
  ],
  webServer: {
    command: `pnpm exec storybook dev --port ${SMOKE_PORT} --no-open`,
    cwd: REPOSITORY_ROOT,
    url: `http://localhost:${SMOKE_PORT}/iframe.html`,
    // Locally, reuse a Storybook already serving on 6007; in CI always boot a
    // fresh one so the run never binds to a stale/unrelated server.
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
