import { defineConfig } from '@playwright/test';
import { resolve } from 'node:path';
import { packagesDir, surveyPorts } from './tests/survey/survey.js';

// Local-only legacy corpus surveys. Each format renders on its viewer
// package's dev server; LEGACY_CORPUS_FORMATS (default doc,ppt,xls) limits
// which servers start.
const ports = surveyPorts();
const formats = (process.env.LEGACY_CORPUS_FORMATS ?? 'doc,ppt,xls')
  .split(',')
  .filter((format): format is 'doc' | 'ppt' | 'xls' => format in ports);
const packages = { doc: 'docx', ppt: 'pptx', xls: 'xlsx' } as const;

export default defineConfig({
  testDir: './tests/survey',
  testMatch: formats.map((format) => `${format}.spec.ts`),
  fullyParallel: false,
  reporter: [['list']],
  use: {
    actionTimeout: 30_000,
    // Number separators and built-in Excel date formats follow
    // navigator.language (charts included). Pin the previous-renderer and
    // candidate contexts to the same locale so a host-language change cannot
    // masquerade as a regression.
    locale: process.env.VRT_LOCALE || 'en-US',
  },
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
        launchOptions: { args: ['--disable-gpu'] },
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 720 },
      },
    },
  ],
  webServer: formats.map((format) => ({
    command: `pnpm exec vite --host 127.0.0.1 --port ${ports[format]} --strictPort`,
    cwd: resolve(packagesDir, packages[format]),
    url: `http://127.0.0.1:${ports[format]}/tests/visual/fixture.html`,
    reuseExistingServer: false,
    timeout: 60_000,
  })),
});
