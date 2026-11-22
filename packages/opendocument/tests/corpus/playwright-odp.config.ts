// Playwright configuration for the local ODP corpus harness. Run it through
// `node packages/opendocument/tests/corpus/run-odp.mjs`, which prepares the
// run directory this configuration and the spec read from the environment.
import { defineConfig } from '@playwright/test';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.ODF_HARNESS_PORT ?? 5192);
const runDir = process.env.ODF_HARNESS_RUN_DIR;
if (!runDir) throw new Error('ODF_HARNESS_RUN_DIR is required; start the harness with run-odp.mjs');

export default defineConfig({
  testDir: here,
  testMatch: 'odp-corpus.spec.ts',
  fullyParallel: true,
  workers: Number(process.env.ODF_HARNESS_JOBS ?? 4),
  timeout: Number(process.env.ODF_HARNESS_TIMEOUT ?? 240) * 1000,
  reporter: [['dot']],
  outputDir: resolve(runDir, 'playwright'),
  use: {
    baseURL: `http://127.0.0.1:${port}`,
    channel: 'chrome',
    launchOptions: { args: ['--disable-gpu'] },
    deviceScaleFactor: 1,
    viewport: { width: 1280, height: 960 },
  },
  // The PPTX package's dev server renders ODP through the `odp` model source.
  webServer: {
    command: `pnpm exec vite --config ${resolve(here, 'vite-odp.config.ts')} --host 127.0.0.1 --port ${port} --strictPort`,
    cwd: resolve(here, '../../../pptx'),
    url: `http://127.0.0.1:${port}/tests/visual/fixture.html`,
    reuseExistingServer: false,
    timeout: 120_000,
  },
});
