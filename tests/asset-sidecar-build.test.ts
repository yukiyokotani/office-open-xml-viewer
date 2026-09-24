import { afterEach, describe, expect, it } from 'vitest';
import { build } from 'vite';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, readdirSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { wasmAssetUrl } from '../vite.config';

const fixtures: string[] = [];
afterEach(() => {
  for (const fixture of fixtures.splice(0)) rmSync(fixture, { recursive: true, force: true });
});

describe('library URL sidecar build', () => {
  it('keeps a font outside ESM and CJS bundles with usable file URLs', async () => {
    const fixture = mkdtempSync(resolve(tmpdir(), 'ooxml-asset-build-'));
    fixtures.push(fixture);
    mkdirSync(resolve(fixture, 'src'));
    writeFileSync(resolve(fixture, 'src/font.ttf'), 'font bytes');
    writeFileSync(resolve(fixture, 'src/index.ts'), `import url from './font.ttf?url'; export { url };`);
    await build({
      configFile: false, root: fixture, logLevel: 'silent',
      plugins: [wasmAssetUrl()],
      build: {
        lib: { entry: resolve(fixture, 'src/index.ts'), formats: ['es', 'cjs'], fileName: (format) => `index.${format === 'es' ? 'mjs' : 'cjs'}` },
        rollupOptions: { output: { assetFileNames: '[name][extname]' } },
      },
    });

    expect(readFileSync(resolve(fixture, 'dist/font.ttf'), 'utf8')).toBe('font bytes');
    const esm = readFileSync(resolve(fixture, 'dist/index.mjs'), 'utf8');
    const cjs = readFileSync(resolve(fixture, 'dist/index.cjs'), 'utf8');
    expect(esm).not.toContain('data:font');
    expect(cjs).not.toContain('data:font');
    const cjsUrl = execFileSync(process.execPath, ['-e', `console.log(require(${JSON.stringify(resolve(fixture, 'dist/index.cjs'))}).url)`], { encoding: 'utf8' }).trim();
    expect(realpathSync(new URL(cjsUrl))).toBe(realpathSync(resolve(fixture, 'dist/font.ttf')));
  });

  it('does not carry font bytes into a single-file consumer rebundle', async () => {
    const fixture = mkdtempSync(resolve(tmpdir(), 'ooxml-portable-font-build-'));
    fixtures.push(fixture);
    mkdirSync(resolve(fixture, 'src'));
    const source = resolve('packages/core/src/fonts/office-fallback.ts');
    writeFileSync(resolve(fixture, 'src/index.ts'),
      `export { loadOfficeFontFallbacks, unloadOfficeFontFallbacks } from ${JSON.stringify(source)};`);
    await build({
      configFile: false, root: fixture, logLevel: 'silent',
      build: { lib: { entry: resolve(fixture, 'src/index.ts'), formats: ['es'], fileName: () => 'index.mjs' } },
    });

    const libraryDir = resolve(fixture, 'dist');
    const libraryFiles = readdirSync(libraryDir);
    expect(libraryFiles.some((name) => /\.(?:ttf|otf|woff2?)$/.test(name))).toBe(false);
    expect(libraryFiles.every((name) => !readFileSync(resolve(libraryDir, name), 'utf8')
      .includes('data:font/ttf;base64,'))).toBe(true);

    const single = resolve(fixture, 'consumer.mjs');
    const esbuildBin = resolve('packages/core/node_modules/esbuild/bin/esbuild');
    execFileSync(esbuildBin, [resolve(libraryDir, 'index.mjs'), '--bundle', '--format=esm',
      '--platform=node', `--outfile=${single}`, '--log-level=silent']);
    expect(readFileSync(single, 'utf8')).not.toContain('data:font/ttf;base64,');

    const originalFontFace = globalThis.FontFace;
    const faces: Array<{ source: string | ArrayBuffer; status: FontFaceLoadStatus }> = [];
    class FakeFace {
      status: FontFaceLoadStatus = 'unloaded';
      constructor(readonly family: string, readonly source: string | ArrayBuffer) {}
      async load() {
        if (typeof this.source === 'string') throw new Error('no local Calibri');
        this.status = 'loaded';
        return this;
      }
    }
    Object.assign(globalThis, { FontFace: FakeFace });
    try {
      const bundled = await import(pathToFileURL(single).href);
      const set = {
        add(face: FakeFace) { faces.push(face); },
        delete(face: FakeFace) {
          const index = faces.indexOf(face);
          if (index < 0) return false;
          faces.splice(index, 1);
          return true;
        },
      } as unknown as FontFaceSet;
      const loaded = await bundled.loadOfficeFontFallbacks([{ family: 'Calibri' }], set);
      expect(loaded.routes).toEqual({});
      expect(faces).toEqual([]);
      bundled.unloadOfficeFontFallbacks(loaded.faces);
    } finally {
      Object.assign(globalThis, { FontFace: originalFontFace });
    }
  });
});
