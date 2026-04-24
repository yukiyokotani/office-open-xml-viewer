import * as esbuild from 'esbuild';
import { readFileSync } from 'fs';

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
  sourcemap: !production,
  minify: production,
  // WASM files are loaded at runtime via fetch — exclude from bundle
  external: ['*.wasm'],
  loader: {
    '.wasm': 'file',
  },
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
    console.log('[esbuild] build complete');
  }
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
