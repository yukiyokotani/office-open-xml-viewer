import { afterEach, describe, expect, it } from 'vitest';
import { build } from 'esbuild';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import {
  bundledAssetSidecars,
  parserWasmOutputProblems,
  singleParserWasmPerFormat,
} from '../esbuild-asset-sidecars.mjs';

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

describe('single parser WASM per format', () => {
  const valid = [
    'dist/webview.js',
    'dist/assets/docx_parser_bg-AAAA1111.wasm',
    'dist/assets/xlsx_parser_bg-BBBB2222.wasm',
    'dist/assets/pptx_parser_bg-CCCC3333.wasm',
  ];

  it('accepts exactly one parser per format', () => {
    expect(parserWasmOutputProblems(valid)).toEqual([]);
  });

  it('rejects a second parser for a format, a missing parser, and any other WASM', () => {
    expect(parserWasmOutputProblems([...valid, 'dist/assets/docx_parser_bg-DDDD4444.wasm'])).toEqual([
      expect.stringContaining('expected one docx parser WASM, found 2'),
    ]);
    expect(parserWasmOutputProblems(valid.filter((output) => !output.includes('pptx')))).toEqual([
      expect.stringContaining('expected one pptx parser WASM, found 0'),
    ]);
    expect(parserWasmOutputProblems([...valid, 'dist/assets/other-EEEE5555.wasm'])).toEqual([
      expect.stringContaining('unexpected WASM output'),
    ]);
  });

  it('fails an esbuild build whose two imports of one parser differ in content', async () => {
    const fixture = mkdtempSync(resolve(tmpdir(), 'ooxml-single-wasm-'));
    fixtures.push(fixture);
    for (const format of ['docx', 'xlsx', 'pptx']) {
      mkdirSync(resolve(fixture, 'src'), { recursive: true });
      mkdirSync(resolve(fixture, 'dist'), { recursive: true });
      writeFileSync(resolve(fixture, `src/${format}_parser_bg.wasm`), `${format} parser`);
      writeFileSync(resolve(fixture, `dist/${format}_parser_bg.wasm`), `${format} parser`);
    }
    const entry = resolve(fixture, 'entry.js');
    writeFileSync(entry, ['docx', 'xlsx', 'pptx'].flatMap((format) => [
      `import ${format}Src from './src/${format}_parser_bg.wasm';`,
      `import ${format}Dist from './dist/${format}_parser_bg.wasm';`,
      `console.log(${format}Src, ${format}Dist);`,
    ]).join('\n'));
    const options = {
      entryPoints: [entry], bundle: true, write: false, outdir: resolve(fixture, 'out'),
      assetNames: 'assets/[name]-[hash]', format: 'iife' as const, platform: 'browser' as const,
      logLevel: 'silent' as const, loader: { '.wasm': 'file' as const },
      plugins: [singleParserWasmPerFormat],
    };

    // Identical bytes behind both imports: esbuild emits one file per format.
    const same = await build(options);
    expect((same.outputFiles ?? []).filter((file) => file.path.endsWith('.wasm'))).toHaveLength(3);

    writeFileSync(resolve(fixture, 'dist/docx_parser_bg.wasm'), 'stale docx parser');
    await expect(build(options)).rejects.toThrow('expected one docx parser WASM, found 2');
  });
});
