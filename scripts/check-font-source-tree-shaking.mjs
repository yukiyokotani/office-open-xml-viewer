#!/usr/bin/env node
// A consumer that imports only the font source factories (the import the
// `useGoogleFonts` deprecation message recommends) must not pay for a viewer.
// Each published browser entry is re-bundled the way an application bundler
// would (Rolldown, the bundler behind Vite), and the result must be one small
// chunk: no parser, renderer or worker code, and no lazily imported chunk.
//
// This guards the library build's chunk graph. Rolldown keeps a chunk that
// two or more of its dynamically imported chunks import back, even when the
// consumer uses nothing from it, so a cycle between a format chunk and its lazy
// chunks makes every import of that chunk retain the whole format.
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { rolldown } from 'rolldown';

const distDir = resolve(process.argv[2] ?? 'dist');
const FACTORIES = ['fontCssApi', 'fontCss', 'fontFiles', 'fontResolver', 'googleFonts', 'googleFontsSubstitutions', 'officeFontSubstitutions'];
// The factories, their validation and the Google preset data minify to about
// 12 KiB. The budget leaves room for growth, not for any viewer code.
const BUDGET_BYTES = 32 * 1024;
const VIEWER_MARKERS = ['_parser_bg.wasm', 'OffscreenCanvas', 'getContext('];

// The other direction: a document loaded without `LoadOptions.fonts` must
// not pay for the font-source runtime, the worker bridge host with its
// source cache, failure record and face downloads included (only its gate,
// resolver-bridge-gate.ts, is static; it loads the host for a document with
// a resolver). These modules are reachable from a
// format entry only through dynamic `import()`, so none may be retained by a
// chunk of the entry's static graph (its static `import` edges, as
// scripts/measure-bundle-size.mjs counts them). One static import of any
// of them that is used, or of one with module-level state (the realm-shared
// caches and registries), moves it and what it imports into a static chunk,
// so the check reads the retained modules of a build with the published
// configuration (scripts/module-stats-plugin.mjs) instead of output text.
export const LAZY_FONT_SOURCE_MODULES = Object.freeze([
  ...[
    'source-session', 'css-source', 'files-source',
    'resolver-resolution', 'resolver-admission', 'resolver-bridge-responder',
    'source-metric-tables', 'web-font-sfnt', 'open-type-variations', 'open-type-instances',
    'installed-families',
    'resolver-bridge-host', 'source-cache', 'resolver-failures', 'face-download',
  ].map((name) => `packages/core/src/fonts/${name}.ts`),
  ...['docx', 'xlsx', 'pptx'].map((format) => `packages/${format}/src/font-source-load.ts`),
  'packages/docx/src/font-demand-plan.ts',
  'packages/docx/src/served-font-metrics.ts',
]);

/** `[entry, module, chunk]` for every lazy font-source module that a chunk
 * in the static graph of a format entry retains. */
export function staticLazyFontSourceModules(stats, entries = ['docx.mjs', 'xlsx.mjs', 'pptx.mjs']) {
  const lazy = new Set(LAZY_FONT_SOURCE_MODULES);
  const violations = [];
  for (const entry of entries) {
    if (!stats.chunks[entry]) throw new Error(`module stats lack the entry ${entry}`);
    const seen = new Set();
    const pending = [entry];
    while (pending.length > 0) {
      const name = pending.pop();
      const chunk = stats.chunks[name];
      if (seen.has(name) || !chunk) continue;
      seen.add(name);
      pending.push(...chunk.imports);
      for (const module of Object.keys(chunk.modules)) {
        if (lazy.has(module)) violations.push([entry, module, name]);
      }
    }
  }
  return violations;
}

async function checkFactoriesOnlyImport() {
  const work = mkdtempSync(join(tmpdir(), 'ooxml-font-shake-'));
  try {
    for (const entry of ['index.mjs', 'docx.mjs', 'xlsx.mjs', 'pptx.mjs']) {
      const consumer = join(work, `consumer-${entry}`);
      writeFileSync(consumer, [
        `import { ${FACTORIES.join(', ')} } from ${JSON.stringify(join(distDir, entry))};`,
        `console.log(googleFonts(), fontCssApi, fontCss, fontFiles, fontResolver, googleFontsSubstitutions, officeFontSubstitutions);`,
        '',
      ].join('\n'));
      const bundle = await rolldown({ input: consumer, logLevel: 'silent' });
      try {
        const { output } = await bundle.generate({ format: 'es', minify: true });
        const chunks = output.filter((item) => item.type === 'chunk');
        if (chunks.length !== 1) {
          throw new Error(`${entry}: a factories-only import emitted ${chunks.length} chunks (${chunks.map((chunk) => chunk.fileName).join(', ')}); expected 1`);
        }
        const bytes = Buffer.byteLength(chunks[0].code);
        if (bytes > BUDGET_BYTES) {
          throw new Error(`${entry}: a factories-only import is ${bytes} bytes, over the ${BUDGET_BYTES}-byte budget`);
        }
        const marker = VIEWER_MARKERS.find((text) => chunks[0].code.includes(text));
        if (marker) throw new Error(`${entry}: a factories-only import contains viewer code (${marker})`);
        console.log(`${entry}: factories-only import ${bytes}/${BUDGET_BYTES} bytes, 1 chunk`);
      } finally {
        await bundle.close();
      }
    }
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

function checkStaticGraphs() {
  const root = mkdtempSync(join(process.platform === 'darwin' ? '/private/tmp' : tmpdir(), 'ooxml-font-graph-'));
  try {
    const statsPath = join(root, 'stats.json');
    execFileSync(process.execPath, [
      'node_modules/vite/bin/vite.js', 'build', '--outDir', join(root, 'dist'), '--emptyOutDir', '--logLevel', 'error',
    ], { encoding: 'utf8', timeout: 300_000, maxBuffer: 10 * 1024 * 1024, env: { ...process.env, OOXML_MODULE_STATS: statsPath } });
    const stats = JSON.parse(readFileSync(statsPath, 'utf8'));
    const retained = new Set(Object.values(stats.chunks).flatMap((chunk) => Object.keys(chunk.modules)));
    // A module list that no longer names real modules would pass vacuously.
    const missing = LAZY_FONT_SOURCE_MODULES.filter((module) => !retained.has(module));
    if (missing.length > 0) throw new Error(`lazy font-source modules missing from the build: ${missing.join(', ')}`);
    const violations = staticLazyFontSourceModules(stats);
    if (violations.length > 0) {
      throw new Error(`font-source code in a format entry's static graph (load it only through import()):\n${
        violations.map(([entry, module, chunk]) => `  ${entry}: ${module} (chunk ${chunk})`).join('\n')}`);
    }
    console.log(`docx.mjs, xlsx.mjs, pptx.mjs: static graphs retain none of ${LAZY_FONT_SOURCE_MODULES.length} font-source modules`);
  } finally {
    // A unique directory created by this run; rmSync does not follow symlinks.
    rmSync(root, { recursive: true, force: true });
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await checkFactoriesOnlyImport();
  checkStaticGraphs();
}
