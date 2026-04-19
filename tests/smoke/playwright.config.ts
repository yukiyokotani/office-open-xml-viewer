import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: '.',
  testMatch: '**/*.spec.ts',
  fullyParallel: false,
  reporter: [['list']],
  use: {
    baseURL: 'http://localhost:6007',
    actionTimeout: 30_000,
  },
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
        deviceScaleFactor: 1,
        viewport: { width: 1400, height: 900 },
      },
    },
  ],
  webServer: {
    command: 'pnpm storybook --port 6007 --no-open',
    url: 'http://localhost:6007/iframe.html',
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
