import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests/visual',
  testMatch: '**/*.spec.ts',
  fullyParallel: false,
  reporter: [
    ['list'],
    ['html', { outputFolder: 'tests/visual/report', open: 'never' }],
  ],
  use: {
    baseURL: 'http://localhost:5180',
    actionTimeout: 30_000,
  },
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 960 },
      },
    },
  ],
  // Start the Vite dev server separately before running tests:
  //   npx vite --port 5180
  webServer: {
    command: 'npx vite --port 5180 --strictPort',
    url: 'http://localhost:5180/tests/visual/fixture.html',
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
