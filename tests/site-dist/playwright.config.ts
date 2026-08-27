import { defineConfig } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const port = Number(process.env.SITE_DIST_PORT ?? 4322);
const siteRoot = fileURLToPath(new URL('../../site', import.meta.url));

export default defineConfig({
  testDir: '.',
  testMatch: 'viewer-pages.spec.ts',
  reporter: [['list']],
  use: {
    baseURL: `http://127.0.0.1:${port}`,
    channel: 'chrome',
    viewport: { width: 1440, height: 1000 },
    deviceScaleFactor: 2,
  },
  webServer: {
    command: `./node_modules/.bin/astro preview --host 127.0.0.1 --port ${port}`,
    cwd: siteRoot,
    url: `http://127.0.0.1:${port}`,
    reuseExistingServer: false,
    timeout: 30_000,
  },
});
