import { defineConfig, devices } from '@playwright/test';

const vrtPort = Number(process.env.VRT_PORT ?? 5173);
if (!Number.isInteger(vrtPort) || vrtPort < 1 || vrtPort > 65_535) {
  throw new Error(`invalid VRT_PORT: ${process.env.VRT_PORT}`);
}
const privateCorpus = process.env.VRT_PRIVATE_CORPUS === '1';

export default defineConfig({
  testDir: './tests/visual',
  testMatch: '**/*.spec.ts',
  // Run slides sequentially for stable output
  fullyParallel: false,
  // The private run exercises both main-thread and worker fixtures over the
  // complete corpus. Keep those files in one Playwright worker: parallel test
  // worker teardown can otherwise leave the shared Vite server waiting after
  // every pixel comparison has completed.
  workers: privateCorpus ? 1 : undefined,
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
        // Use the system-installed Google Chrome so host fonts (Hiragino etc.)
        // match what the user sees in the browser.
        channel: 'chrome',
        // Keep exact-pixel VRT independent of GPU driver scheduling and
        // antialiasing differences between otherwise identical Chrome runs.
        launchOptions: {
          args: ['--disable-gpu'],
        },
        // Force DPR=1 so canvas physical size matches the 1280×720
        // PowerPoint reference images (toDataURL returns canvas.width × canvas.height).
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 720 },
      },
    },
  ],
  webServer: {
    command: `pnpm exec vite --host 127.0.0.1 --port ${vrtPort} --strictPort`,
    url: `http://127.0.0.1:${vrtPort}/tests/visual/fixture.html`,
    reuseExistingServer: privateCorpus ? false : !process.env.CI,
    timeout: 60_000,
  },
});
