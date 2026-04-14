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
      name: 'chromium',
      use: {
        ...devices['Desktop Chrome'],
        // Viewport must be at least as wide as the canvas (1920px); height ≥ 1080px.
        // Force dpr=1 so canvas physical size == CSS size == 1920×1080.
        deviceScaleFactor: 1,
        viewport: { width: 1920, height: 1080 },
      },
    },
  ],
  webServer: {
    command: 'npx vite dev --port 5173',
    // /sample.pptx is a static file that always returns 200 — use it as the ready check
    url: 'http://localhost:5173/sample.pptx',
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  },
});
