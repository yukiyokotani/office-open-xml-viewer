import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './tests/visual',
  testMatch: '**/*.spec.ts',
  // Run slides sequentially for stable output
  fullyParallel: false,
  reporter: [
    ['list'],
    ['html', { outputFolder: 'tests/visual/report', open: 'never' }],
  ],
  use: {
    baseURL: 'http://localhost:5173',
    actionTimeout: 30_000,
  },
  projects: [
    {
      name: 'chrome',
      use: {
        // Use the system-installed Google Chrome so fonts (Hiragino etc.)
        // and rendering exactly match what the user sees in the browser.
        channel: 'chrome',
        // Force DPR=1 so canvas physical size matches the 1280×720
        // PowerPoint reference images (toDataURL returns canvas.width × canvas.height).
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 720 },
      },
    },
  ],
  webServer: {
    command: 'npx vite dev --port 5173',
    url: 'http://localhost:5173/demo/sample-1.pptx',
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  },
});
