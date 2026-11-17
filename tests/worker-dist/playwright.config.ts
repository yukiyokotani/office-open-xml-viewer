import { defineConfig } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const port = Number(process.env.WORKER_DIST_PORT ?? 6012);
const repositoryRoot = fileURLToPath(new URL('../..', import.meta.url));

export default defineConfig({
  testDir: '.',
  testMatch: ['production-workers.spec.ts', 'base-href.spec.ts'],
  reporter: [['list']],
  use: {
    baseURL: `http://127.0.0.1:${port}`,
    channel: 'chrome',
    viewport: { width: 1200, height: 900 },
  },
  webServer: {
    command: 'node scripts/build-worker-consumer-fixture.mjs && node tests/worker-dist/server.mjs',
    cwd: repositoryRoot,
    url: `http://127.0.0.1:${port}`,
    reuseExistingServer: false,
    timeout: 30_000,
  },
});
