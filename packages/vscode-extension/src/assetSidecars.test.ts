import { afterEach, describe, expect, it } from 'vitest';
import { build } from 'esbuild';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { bundledAssetSidecars, parserWasmManifest, PARSER_WASM_MANIFEST } from '../esbuild-asset-sidecars.mjs';

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

describe('VS Code parser WASM manifest', () => {
  it('names each emitted parser file for the extension host', async () => {
    const fixture = mkdtempSync(resolve(tmpdir(), 'ooxml-parser-manifest-'));
    fixtures.push(fixture);
    const entry = resolve(fixture, 'entry.mjs');
    for (const format of ['docx', 'xlsx', 'pptx']) {
      writeFileSync(resolve(fixture, `${format}_parser_bg.wasm`), `${format} wasm bytes`);
    }
    writeFileSync(entry, ['docx', 'xlsx', 'pptx']
      .map((format) => `export { default as ${format} } from './${format}_parser_bg.wasm';`)
      .join('\n'));
    const outdir = resolve(fixture, 'dist');
    await build({
      entryPoints: [entry], bundle: true, outdir, assetNames: 'assets/[name]-[hash]',
      format: 'iife', platform: 'browser', logLevel: 'silent', absWorkingDir: fixture,
      loader: { '.wasm': 'file' },
      plugins: [parserWasmManifest],
    });

    const manifest = JSON.parse(readFileSync(resolve(outdir, PARSER_WASM_MANIFEST), 'utf8'));
    expect(Object.keys(manifest).sort()).toEqual(['docx', 'pptx', 'xlsx']);
    for (const format of ['docx', 'xlsx', 'pptx']) {
      expect(manifest[format]).toMatch(new RegExp(`^assets/${format}_parser_bg-[A-Z0-9]+\\.wasm$`));
      expect(readFileSync(resolve(outdir, manifest[format]), 'utf8')).toBe(`${format} wasm bytes`);
    }
  });

  it('fails the build when a format has no parser output', async () => {
    const fixture = mkdtempSync(resolve(tmpdir(), 'ooxml-parser-manifest-'));
    fixtures.push(fixture);
    writeFileSync(resolve(fixture, 'docx_parser_bg.wasm'), 'docx');
    writeFileSync(resolve(fixture, 'entry.mjs'), "export { default } from './docx_parser_bg.wasm';");
    await expect(build({
      entryPoints: [resolve(fixture, 'entry.mjs')], bundle: true, outdir: resolve(fixture, 'dist'),
      assetNames: 'assets/[name]-[hash]', format: 'iife', platform: 'browser', logLevel: 'silent',
      absWorkingDir: fixture, loader: { '.wasm': 'file' }, plugins: [parserWasmManifest],
    })).rejects.toThrow(/no parser WASM output for xlsx, pptx/);
  });
});
