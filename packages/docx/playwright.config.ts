import { defineConfig, devices } from '@playwright/test';

const vrtPort = Number(process.env.VRT_PORT ?? 5180);
if (!Number.isInteger(vrtPort) || vrtPort < 1 || vrtPort > 65_535) {
  throw new Error(`invalid VRT_PORT: ${process.env.VRT_PORT}`);
}
const privateCorpus = process.env.VRT_PRIVATE_CORPUS === '1';

export default defineConfig({
  testDir: './tests/visual',
  testMatch: '**/*.spec.ts',
  fullyParallel: false,
  // The private corpus already persists exact PNG artifacts. Avoid the HTML
  // reporter there: large main/worker runs can keep its teardown alive after
  // every test has completed.
  reporter: privateCorpus
    ? [['list']]
    : [
        ['list'],
        ['html', { outputFolder: 'tests/visual/report', open: 'never' }],
      ],
  use: {
    baseURL: `http://127.0.0.1:${vrtPort}`,
    actionTimeout: 30_000,
  },
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
        // Keep exact-pixel VRT independent of GPU driver scheduling and
        // antialiasing differences between otherwise identical Chrome runs.
        launchOptions: {
          args: ['--disable-gpu'],
        },
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 960 },
      },
    },
    {
      name: 'firefox',
      testMatch: '**/conformance.spec.ts',
      use: {
        ...devices['Desktop Firefox'],
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 960 },
      },
    },
    {
      name: 'webkit',
      testMatch: '**/conformance.spec.ts',
      use: {
        ...devices['Desktop Safari'],
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 960 },
      },
    },
  ],
  // Start the Vite dev server separately before running tests:
  //   pnpm exec vite --port 5180
  webServer: {
    command: `pnpm exec vite --host 127.0.0.1 --port ${vrtPort} --strictPort`,
    url: `http://127.0.0.1:${vrtPort}/tests/visual/fixture.html`,
    reuseExistingServer: privateCorpus ? false : !process.env.CI,
    timeout: 120_000,
  },
});
