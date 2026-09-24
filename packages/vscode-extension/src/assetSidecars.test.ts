import { afterEach, describe, expect, it } from 'vitest';
import { build } from 'esbuild';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { bundledAssetSidecars } from '../esbuild-asset-sidecars.mjs';

const fixtures: string[] = [];
afterEach(() => {
  for (const fixture of fixtures.splice(0)) rmSync(fixture, { recursive: true, force: true });
});

describe('VS Code webview package sidecars', () => {
  it('emits Vite font and parser assets as files referenced by the IIFE', async () => {
    const fixture = mkdtempSync(resolve(tmpdir(), 'ooxml-sidecars-'));
    fixtures.push(fixture);
    const dist = resolve(fixture, 'packages/docx/dist');
    mkdirSync(dist, { recursive: true });
    writeFileSync(resolve(dist, 'index.mjs'), `
      export const font = new URL("Carlito-Regular.ttf", import.meta.url).href;
      export const wasm = new URL("docx_parser_bg.wasm", import.meta.url).href;
    `);
    writeFileSync(resolve(dist, 'Carlito-Regular.ttf'), 'font bytes');
    writeFileSync(resolve(dist, 'docx_parser_bg.wasm'), 'wasm bytes');
    const result = await build({
      entryPoints: [resolve(dist, 'index.mjs')], bundle: true, write: false,
      outdir: resolve(fixture, 'out'), assetNames: 'assets/[name]-[hash]',
      format: 'iife', platform: 'browser', logLevel: 'silent',
      loader: { '.ttf': 'file', '.wasm': 'file' },
      plugins: [bundledAssetSidecars],
    });

    expect(result.warnings).toEqual([]);
    const js = result.outputFiles.find((file) => file.path.endsWith('.js'))?.text ?? '';
    for (const ext of ['ttf', 'wasm']) {
      const asset = result.outputFiles.find((file) => file.path.endsWith(`.${ext}`));
      expect(asset).toBeDefined();
      expect(js).toContain(`./assets/${asset?.path.split('/').at(-1)}`);
    }
    expect(js).not.toContain('import.meta');
  });
});
