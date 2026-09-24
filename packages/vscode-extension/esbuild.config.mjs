import * as esbuild from 'esbuild';
import { mkdir, readdir, rm } from 'node:fs/promises';
import { mainThreadOnlyWorkerStubs } from './esbuild-worker-stub.mjs';
import { bundledAssetSidecars } from './esbuild-asset-sidecars.mjs';

const production = process.argv.includes('--production');
const watch = process.argv.includes('--watch');

/** @type {esbuild.BuildOptions} */
const extensionConfig = {
  entryPoints: ['src/extension.ts'],
  bundle: true,
  format: 'cjs',
  platform: 'node',
  target: 'node18',
  external: ['vscode'],
  outfile: 'dist/extension.js',
  sourcemap: !production,
  minify: production,
};

/** @type {esbuild.BuildOptions} */
const webviewConfig = {
  entryPoints: ['src/webview/bootstrap.ts'],
  bundle: true,
  format: 'iife',
  platform: 'browser',
  target: 'es2020',
  outfile: 'dist/webview.js',
  // Vite's plain ?url imports (notably parser WASM) are real files in this
  // bundle. The viewer does not bundle font binaries.
  assetNames: 'assets/[name]-[hash]',
  sourcemap: !production,
  minify: production,
  // Static Vite sidecars become file imports through bundledAssetSidecars.
  loader: {
    '.wasm': 'file',
    '.ttf': 'file',
  },
  plugins: [mainThreadOnlyWorkerStubs, bundledAssetSidecars],
};

async function build() {
  if (watch) {
    const [extCtx, wvCtx] = await Promise.all([
      esbuild.context(extensionConfig),
      esbuild.context(webviewConfig),
    ]);
    await Promise.all([extCtx.watch(), wvCtx.watch()]);
    console.log('[esbuild] watching...');
  } else {
    await Promise.all([
      esbuild.build(extensionConfig),
      esbuild.build(webviewConfig),
    ]);
    // Older builds emitted Carlito files. Remove those stale outputs so an
    // incremental VSIX does not redistribute bytes that are no longer used.
    await mkdir('dist/assets', { recursive: true });
    for (const name of await readdir('dist/assets')) {
      if (/^Carlito-(?:Regular|Bold|Italic|BoldItalic)-[A-Z0-9]+\.ttf$/.test(name)
        || name === 'Carlito-OFL.txt') {
        await rm(`dist/assets/${name}`);
      }
    }
    console.log('[esbuild] build complete');
  }
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
