import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

const distDir = resolve(process.argv[2] ?? 'dist');
const files = readdirSync(distDir);
const hosts = files.filter((name) => /^render-worker-host-[\w-]+\.js$/.test(name));

if (hosts.length !== 3) {
  throw new Error(`Expected 3 production render-worker hosts, found ${hosts.length}`);
}

const workerAssets = new Set();
for (const hostName of hosts) {
  const source = readFileSync(join(distDir, hostName), 'utf8');
  if (/\b(?:blob|data):/i.test(source)) {
    throw new Error(`${hostName} inlines a worker instead of emitting a module asset`);
  }
  if (/new URL\(\s*["']\/assets\//.test(source)) {
    throw new Error(`${hostName} resolves its worker from the origin root`);
  }
  const match = /new URL\(\s*["'](assets\/render-worker-[^"']+\.js)["']\s*,\s*import\.meta\.url\s*\)/.exec(source);
  if (!match) {
    throw new Error(`${hostName} does not resolve a module worker from import.meta.url`);
  }
  const assetPath = join(distDir, match[1]);
  if (!existsSync(assetPath)) {
    throw new Error(`${hostName} references missing worker asset ${match[1]}`);
  }
  workerAssets.add(assetPath);
}

for (const assetPath of workerAssets) {
  const source = readFileSync(assetPath, 'utf8');
  if (/\bimport\.meta\b/.test(source)) {
    throw new Error(
      `${assetPath} contains import.meta; opaque worker assets must remain parseable as classic scripts`,
    );
  }
  const specifiers = [
    ...source.matchAll(/\bfrom\s*["']([^"']+)["']/g),
    ...source.matchAll(/\bimport\(\s*["']([^"']+)["']\s*\)/g),
  ].map((match) => match[1]);
  for (const specifier of specifiers) {
    if (/^(?:blob|data):/i.test(specifier)) {
      throw new Error(`${assetPath} imports a Blob/data module dependency`);
    }
    if (specifier.startsWith('./')) {
      throw new Error(
        `${assetPath} imports split chunk ${specifier}; consumer bundlers copy the worker as an opaque asset`,
      );
    }
  }
}

console.log(
  `Production render workers: ${hosts.length} import.meta.url hosts, ${workerAssets.size} self-contained module assets`,
);
