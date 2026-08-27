import { describe, expect, it } from 'vitest';
import { build } from 'esbuild';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { mainThreadOnlyWorkerStubs } from '../esbuild-worker-stub.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const EXTENSION_ROOT = resolve(HERE, '..');

describe('VS Code webview math engine bundle', () => {
  it('retains the MathJax STIX2 side-effect entry used by DOCX, XLSX, and PPTX', async () => {
    const result = await build({
      stdin: {
        contents: "import '@silurus/ooxml-core/mathjax-stix2';",
        resolveDir: HERE,
        loader: 'js',
      },
      bundle: true,
      write: false,
      format: 'iife',
      platform: 'browser',
      target: 'es2020',
      logLevel: 'silent',
    });

    expect(result.warnings).toEqual([]);
    expect(result.outputFiles[0]?.text).toContain('globalThis.__ooxmlStix2');
  });

  it('excludes worker-only Vite asset imports from the main-thread webview bundle', async () => {
    const result = await build({
      entryPoints: [resolve(EXTENSION_ROOT, 'src/webview/bootstrap.ts')],
      bundle: true,
      write: false,
      outdir: 'dist-test',
      format: 'iife',
      platform: 'browser',
      target: 'es2020',
      logLevel: 'silent',
      alias: {
        '@silurus/ooxml-docx': resolve(EXTENSION_ROOT, '../docx/src/index.ts'),
        '@silurus/ooxml-xlsx': resolve(EXTENSION_ROOT, '../xlsx/src/index.ts'),
        '@silurus/ooxml-pptx': resolve(EXTENSION_ROOT, '../pptx/src/index.ts'),
      },
      external: ['*.wasm'],
      loader: { '.wasm': 'file' },
      plugins: [mainThreadOnlyWorkerStubs],
    });

    expect(result.warnings).toEqual([]);
    const bundle = result.outputFiles.find((file) => file.path.endsWith('/bootstrap.js'))?.text
      ?? result.outputFiles.find((file) => file.path.endsWith('.js'))?.text
      ?? '';
    expect(bundle).toContain('globalThis.__ooxmlStix2');
    expect(bundle).toContain('renderChartExChart');
    expect(bundle).toContain('renderRegionMapChart');
    expect(bundle).toContain('renderSimpleThreeDChart');
    expect(bundle).not.toContain('Failed to load math engine from');
    expect(bundle).not.toContain('ooxml-worker-renderer-module');
  });
});
