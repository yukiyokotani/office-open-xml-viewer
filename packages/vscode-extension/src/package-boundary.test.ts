import { execFileSync } from 'node:child_process';
import {
  copyFileSync,
  mkdirSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

const extensionRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const vscePackage = require.resolve('@vscode/vsce/package.json');
const vsce = resolve(dirname(vscePackage), 'vsce');
const fixtures: string[] = [];

afterEach(() => {
  for (const fixture of fixtures.splice(0)) rmSync(fixture, { recursive: true, force: true });
});

describe('VS Code extension package boundary', () => {
  it('never includes development source maps left by a previous build', () => {
    const fixture = mkdtempSync(resolve(tmpdir(), 'ooxml-vscode-package-'));
    fixtures.push(fixture);
    mkdirSync(resolve(fixture, 'dist'));
    for (const file of ['package.json', 'README.md', 'icon.png', 'LICENSE', '.vscodeignore']) {
      copyFileSync(resolve(extensionRoot, file), resolve(fixture, file));
    }
    writeFileSync(resolve(fixture, 'dist/extension.js'), 'module.exports = {};');
    writeFileSync(resolve(fixture, 'dist/webview.js'), '(() => {})();');
    writeFileSync(resolve(fixture, 'dist/extension.js.map'), '{"version":3}');
    writeFileSync(resolve(fixture, 'dist/webview.js.map'), '{"version":3}');
    writeFileSync(resolve(fixture, 'esbuild-worker-stub.mjs'), 'export const plugin = {};');
    writeFileSync(resolve(fixture, 'esbuild-worker-stub.d.mts'), 'export const plugin: unknown;');
    writeFileSync(resolve(fixture, 'esbuild-asset-sidecars.mjs'), 'export const plugin = {};');

    const files = execFileSync(process.execPath, [vsce, 'ls', '--no-dependencies'], {
      cwd: fixture,
      encoding: 'utf8',
    }).trim().split(/\r?\n/);

    expect(files).toContain('dist/extension.js');
    expect(files).toContain('dist/webview.js');
    expect(files).toContain('LICENSE');
    expect(files.filter((file) => file.endsWith('.map'))).toEqual([]);
    expect(files.filter((file) => file.startsWith('esbuild-worker-stub.'))).toEqual([]);
    expect(files).not.toContain('esbuild-asset-sidecars.mjs');
  });
});
