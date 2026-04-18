import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './packages/pptx/tests/visual',
  testMatch: '**/*.spec.ts',
  fullyParallel: false,
  reporter: [
    ['list'],
    ['html', { outputFolder: 'packages/pptx/tests/visual/report', open: 'never' }],
  ],
  use: {
    baseURL: 'http://localhost:5173',
    actionTimeout: 30_000,
  },
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 720 },
      },
    },
  ],
  webServer: {
    command: 'pnpm --filter @ooxml/pptx dev --port 5173',
    url: 'http://localhost:5173/sample-1.pptx',
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  },
});
