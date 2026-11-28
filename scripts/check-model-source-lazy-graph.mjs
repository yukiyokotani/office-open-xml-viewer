#!/usr/bin/env node
// Runtime smoke proves ordinary loads do not request optional source code.
// This companion guard checks eager source and emitted graphs, inline workers,
// sidecars, markers, and the differential build cost of selected-source code.
// It guards accidental coupling; it is not a security boundary.
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { resolve, dirname, relative, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parse } from '@babel/parser';
import { execFileSync } from 'node:child_process';
import { performance } from 'node:perf_hooks';
import { tmpdir } from 'node:os';
import { gzipSync } from 'node:zlib';
import { createTypeScriptResolver } from './check-core-legacy-boundary.mjs';

const resolveModule = createTypeScriptResolver();

// The only entry-to-source transitions are these opt-in dispatch functions.
// Keep this list exact: a new caller must explain why it can run only after
// a format's opt-in source option (or a worker source descriptor) is present.
export const SOURCE_DISPATCH = new Map([
  ['packages/core/src/source/model-source.ts', ['openModelSourceModule']],
  ['packages/docx/src/document.ts', ['load']],
  ['packages/xlsx/src/workbook.ts', ['load']],
  ['packages/pptx/src/presentation.ts', ['load']],
  ['packages/node/src/docx.ts', ['openDocxDocument', 'materializeDocxDocument', 'openDocxPictures']],
  ['packages/node/src/xlsx.ts', ['openXlsxWorkbook']],
  ['packages/node/src/pptx.ts', ['openPptxPresentationImpl']],
  ['packages/node/src/docx-model-source.ts', ['acquireDocxInput']],
  ['packages/node/src/xlsx-model-source.ts', ['acquireXlsxInput']],
  ['packages/node/src/pptx-model-source.ts', ['acquirePptxInput']],
  ...['docx', 'xlsx', 'pptx'].flatMap((format) => [
    [`packages/${format}/src/worker-source.ts`, ['self.onmessage']],
    [`packages/${format}/src/render-worker.ts`,
      format === 'pptx' ? ['executeArchiveFromNew'] : ['self.onmessage']],
  ]),
]);

// A feature budget, not a historical whole-entry baseline. It counts every
// eagerly retained rendered byte that disappears when selected-source dispatch
// is compiled out, per module, even if preparation was moved into another
// static module (assertDispatchCost, normalised token bytes). Measured at
// introduction: docx 217, xlsx 81, pptx 92, node 494; each budget leaves room
// for a small dispatch change, not for another copy of shared geometry or
// layout.
const DISPATCH_COST_BUDGET = Object.freeze({ docx: 800, xlsx: 500, pptx: 500, node: 1_200 });

// Independent sidecar compression plus registration/owner bindings adds a
// small overhead to the reading code moved out of the eager worker. This
// budget allows that code, not another copy of shared geometry or layout.
const DOCX_SOURCE_SIDECAR_GZIP_BUDGET = 16 * 1024;

function assertSidecarBudget(file) {
  const bytes = gzipSync(readFileSync(file), { level: 9 }).length;
  if (bytes > DOCX_SOURCE_SIDECAR_GZIP_BUDGET) {
    throw new Error(`${file} exceeds ${DOCX_SOURCE_SIDECAR_GZIP_BUDGET}-byte gzip sidecar budget: ${bytes}`);
  }
  console.log(`${file}: ${bytes}/${DOCX_SOURCE_SIDECAR_GZIP_BUDGET} gzip bytes`);
}

export function checkSourceDispatchImports(files) {
  for (const { path, text } of files) {
    if (/\.(?:test|spec|stories|probe)\.[cm]?[jt]sx?$/.test(path)) continue;
    const ast = parse(text, { sourceType: 'module', plugins: ['typescript', 'jsx'] });
    const allowed = SOURCE_DISPATCH.get(path) ?? [];
    function walk(node, functionName) {
      if (!node || typeof node !== 'object' || !node.type) return;
      let current = functionName;
      if (node.type === 'FunctionDeclaration') current = node.id?.name;
      if (['ClassMethod', 'ObjectMethod'].includes(node.type)) current = node.key?.name;
      if (['ArrowFunctionExpression', 'FunctionExpression'].includes(node.type)) {
        // Preserve the containing named dispatch for its local callback/IIFE.
        current = functionName;
      }
      if ((node.type === 'CallExpression' && node.callee.type === 'Import')
        || node.type === 'ImportExpression') {
        const target = node.arguments?.[0] ?? node.source;
        const specifier = target?.value;
        const sourceImport = typeof specifier === 'string'
          ? /model-source/.test(specifier)
          : text.slice(target?.start ?? 0, target?.end ?? 0).includes('sourceOwnerUrl')
            || text.slice(target?.start ?? 0, target?.end ?? 0).includes('sourceModule.moduleUrl');
        if (sourceImport && !allowed.includes(current)) {
          throw new Error(`${path}:${node.loc?.start.line} model-source import outside an allowed dispatch function (${current ?? 'top level'})`);
        }
      }
      // Assignment to self.onmessage is a named worker dispatch boundary.
      if (node.type === 'AssignmentExpression' && node.left.type === 'MemberExpression'
        && node.left.object.name === 'self' && node.left.property.name === 'onmessage') {
        walk(node.right, 'self.onmessage');
        return;
      }
      for (const value of Object.values(node)) {
        if (Array.isArray(value)) value.forEach((child) => walk(child, current));
        else if (value && typeof value === 'object' && value.type) walk(value, current);
      }
    }
    walk(ast, undefined);
  }
}

function trackedSourceFiles() {
  return execFileSync('git', ['ls-files', 'packages/core/src', 'packages/docx/src',
    'packages/xlsx/src', 'packages/pptx/src', 'packages/node/src'], { encoding: 'utf8' })
    .split('\n').filter((path) => /\.[cm]?[jt]sx?$/.test(path) && existsSync(path))
    .map((path) => ({ path, text: readFileSync(path, 'utf8') }));
}

function staticSpecifiers(ast) {
  return ast.program.body.flatMap((node) =>
    node.source && ['ImportDeclaration', 'ExportNamedDeclaration', 'ExportAllDeclaration'].includes(node.type)
      && node.importKind !== 'type' && node.exportKind !== 'type'
      && !(node.type === 'ImportDeclaration' && node.specifiers.length > 0
        && node.specifiers.every((specifier) => specifier.importKind === 'type'))
      ? [node.source.value] : []);
}

export function assertNoTopLevelModelImport(code, name) {
  const ast = parse(code, { sourceType: 'module' });
  function walk(node, functionDepth) {
    if (!node || typeof node !== 'object' || !node.type) return;
    const target = node.type === 'ImportExpression' ? node.source
      : node.type === 'CallExpression' && node.callee.type === 'Import' ? node.arguments[0]
        : undefined;
    if (functionDepth === 0 && typeof target?.value === 'string'
      && target.value.includes('model-source')) {
      throw new Error(`${name} imports a model-source chunk at top level`);
    }
    const nested = functionDepth + (['FunctionDeclaration', 'FunctionExpression',
      'ArrowFunctionExpression', 'ClassMethod', 'ObjectMethod'].includes(node.type) ? 1 : 0);
    for (const value of Object.values(node)) {
      if (Array.isArray(value)) value.forEach((child) => walk(child, nested));
      else if (value && typeof value === 'object' && value.type) walk(value, nested);
    }
  }
  walk(ast, 0);
}

/** The real static source graph, including workspace package exports. */
export function eagerModules(entry) {
  const visited = new Set();
  const pending = [resolve(entry)];
  while (pending.length) {
    const file = pending.pop();
    if (visited.has(file)) continue;
    visited.add(file);
    if (file.includes('/src/wasm/') || !/\.[cm]?[jt]sx?$/.test(file)) continue;
    // Reparse on each audit. A test or editor can change a file between two
    // calls in one process; a path-only cache would silently preserve its old
    // import graph.
    const ast = parse(readFileSync(file, 'utf8'), { sourceType: 'module', plugins: ['typescript', 'jsx'] });
    const imports = staticSpecifiers(ast);
    const importer = relative(process.cwd(), file).replaceAll('\\', '/');
    for (const specifier of imports) {
      if (specifier.includes('?')) continue;
      let target;
      if (specifier.startsWith('.')) {
        const stem = resolve(dirname(file), specifier).replace(/\.js$/, '');
        target = [stem + '.ts', stem + '.tsx', stem + '.js', stem + '.mjs']
          .find((candidate) => existsSync(candidate));
      } else if (file.startsWith(process.cwd())) {
        target = resolveModule(importer, specifier);
      }
      if (target) pending.push(resolve(target));
    }
  }
  return visited;
}

export function assertLazySourceOwner(entry, owner) {
  const modules = eagerModules(entry);
  if (modules.has(resolve(owner))) throw new Error(`${entry} statically reaches ${owner}`);
}

function bundleGraph(entry) {
  const visited = new Set();
  const pending = [resolve(entry)];
  let bytes = 0;
  let joined = '';
  const codes = [];
  while (pending.length) {
    const file = pending.pop();
    if (visited.has(file)) continue;
    visited.add(file);
    const code = readFileSync(file, 'utf8');
    assertNoTopLevelModelImport(code, file);
    bytes += Buffer.byteLength(code);
    joined += '\n' + code;
    codes.push(code);
    const ast = parse(code, { sourceType: 'module' });
    for (const specifier of staticSpecifiers(ast)) {
      if (specifier.startsWith('.')) pending.push(resolve(dirname(file), specifier));
    }
  }
  return { bytes, files: visited.size, joined, codes };
}

function inlineWorkers(code) {
  const ast = parse(code, { sourceType: 'module' });
  const outputs = [];
  function walk(node) {
    if (!node || typeof node !== 'object') return;
    if (node.type === 'StringLiteral' && node.value.length > 10000) {
      const decoded = node.value.includes('self.') ? node.value : Buffer.from(node.value, 'base64').toString();
      if (decoded.includes('onmessage')) outputs.push(decoded);
    }
    for (const value of Object.values(node)) {
      if (Array.isArray(value)) value.forEach(walk);
      else if (value && typeof value === 'object' && value.type) walk(value);
    }
  }
  walk(ast);
  return outputs;
}

/** Static module closure of an entry chunk in a build's module stats
 * (scripts/module-stats-plugin.mjs): module id → rendered code. */
export function staticModuleCode(stats, entryFile) {
  const code = new Map();
  const visited = new Set();
  const pending = [entryFile];
  while (pending.length) {
    const file = pending.pop();
    if (visited.has(file)) continue;
    visited.add(file);
    const chunk = stats.chunks[file];
    if (!chunk) throw new Error(`Module stats lack chunk ${file}`);
    for (const [id, text] of Object.entries(chunk.modules)) code.set(id, `${code.get(id) ?? ''}\n${text}`);
    // External specifiers (node: built-ins, skia-canvas) are not chunks.
    pending.push(...chunk.imports.filter((imported) => stats.chunks[imported]));
  }
  return code;
}

/** Cost of one identifier occurrence, whatever its spelling: identifier text
 * depends on the bundler's per-chunk deconfliction, not on the source. */
const IDENTIFIER_COST = 1;

/**
 * A module's code as a multiset of normalised tokens with their byte cost:
 * every identifier becomes one fixed-cost token (its text depends on chunk-
 * scoped renaming), every other token (keywords, punctuators, string, number,
 * template and regex literals) keeps its exact text and byte length. Comments
 * and whitespace are ignored, and the bundler's per-build chunk-reference
 * placeholders (`!~{…}~`) are normalised.
 */
export function moduleTokenCosts(code) {
  const ast = parse(code, {
    sourceType: 'module', tokens: true, errorRecovery: true,
    allowAwaitOutsideFunction: true, allowReturnOutsideFunction: true, allowUndeclaredExports: true,
  });
  const costs = new Map();
  for (const token of ast.tokens) {
    const label = token.type?.label ?? token.type;
    if (label === 'eof' || label === 'CommentLine' || label === 'CommentBlock') continue;
    // A chunk reference inside a module (a dynamic import of another chunk)
    // is a bundler placeholder whose id is assigned per build.
    const text = code.slice(token.start, token.end).replace(/!~\{[^}]*\}~/g, '!~{chunk}~');
    const key = label === 'name' || label === 'jsxName' || label === '#name' ? 'identifier' : `${label}:${text}`;
    const entry = costs.get(key) ?? { count: 0, cost: key === 'identifier' ? IDENTIFIER_COST : Buffer.byteLength(text) };
    entry.count++;
    costs.set(key, entry);
  }
  return costs;
}

/**
 * Dispatch cost: the code the selected-source dispatch adds to an entry's
 * eager static graph. Per eager module, the enabled build's normalised token
 * multiset is compared with the disabled build's, and every token occurrence
 * the enabled build has beyond the disabled one is charged at its byte cost
 * (a module only the enabled build retains counts whole). The measure is
 * independent of identifier naming and of how the bundler partitions modules
 * into chunks, and growth is charged even when other code of the same module
 * shrinks (only an identical token can cancel). No noise allowance is needed.
 */
export function assertDispatchCost(enabledStats, disabledStats, entryFile, budget, label) {
  const enabled = staticModuleCode(enabledStats, entryFile);
  const disabled = staticModuleCode(disabledStats, entryFile);
  let bytes = 0;
  const grown = [];
  for (const [id, code] of enabled) {
    const before = disabled.has(id) ? moduleTokenCosts(disabled.get(id)) : new Map();
    let added = 0;
    for (const [key, { count, cost }] of moduleTokenCosts(code)) {
      const extra = count - (before.get(key)?.count ?? 0);
      if (extra > 0) added += extra * cost;
    }
    if (added > 0) {
      bytes += added; grown.push(`${id} +${added}`);
      if (process.env.OOXML_DISPATCH_DEBUG) for (const [key, { count }] of moduleTokenCosts(code)) {
        const extra = count - (before.get(key)?.count ?? 0);
        if (extra > 0) console.log(`  ${id}: ${extra} × ${key}`);
      }
    }
  }
  if (bytes > budget) {
    throw new Error(`${label} model-source dispatch cost exceeds ${budget}-byte budget: ${bytes} (${grown.join(', ')})`);
  }
  console.log(`${label} dispatch cost: ${bytes}/${budget} token bytes (${grown.join(', ') || 'no module'})`);
  return bytes;
}

function withComparisonBuilds(check) {
  const root = mkdtempSync(join(process.platform === 'darwin' ? '/private/tmp' : tmpdir(), 'ooxml-source-cost-'));
  const started = performance.now();
  try {
    // Both comparison builds run here with module stats, so the accounting
    // never mixes a published build with a differently configured one.
    const stats = {};
    for (const [label, mode] of [['enabled', 'production'], ['disabled', 'model-sources-off']]) {
      const statsPath = join(root, `${label}.json`);
      execFileSync(process.execPath, [
        'node_modules/vite/bin/vite.js', 'build', '--mode', mode,
        '--outDir', join(root, label), '--emptyOutDir',
      ], { encoding: 'utf8', timeout: 300_000, maxBuffer: 10 * 1024 * 1024, env: { ...process.env, OOXML_MODULE_STATS: statsPath } });
      stats[label] = JSON.parse(readFileSync(statsPath, 'utf8'));
    }
    console.log(`model-source comparison builds: ${((performance.now() - started) / 1000).toFixed(2)}s`);
    check(join(root, 'disabled'), stats);
  } finally {
    // This unique directory is created by this invocation and contains only
    // its generated comparison builds. rmSync does not follow symlink targets.
    rmSync(root, { recursive: true, force: true });
  }
}

export function assertNoSourceRuntime(code, name) {
  if (code.includes('ooxml-model-source-module/v2') || code.includes('model source view default')) {
    throw new Error(`${name} includes optional source runtime in its eager OOXML code`);
  }
}

export function checkBuiltBundles(dist = 'dist', { packages = false, disabledDist, stats } = {}) {
  for (const format of ['docx', 'xlsx', 'pptx', 'node']) {
    const graph = bundleGraph(join(dist, `${format}.mjs`));
    assertNoSourceRuntime(graph.joined, `${format} static entry graph`);
    console.log(`${format} static JS: ${graph.bytes} bytes across ${graph.files} files`);
    if (format !== 'node') {
      for (const payload of graph.codes.flatMap(inlineWorkers)) {
        assertNoSourceRuntime(payload, `${format} inline worker`);
        console.log(`${format} inline worker: ${Buffer.byteLength(payload)} decoded bytes`);
      }
      const sidecar = join(dist, `${format}-source-worker.mjs`);
      if (!existsSync(sidecar)) throw new Error(`Missing optional source sidecar ${sidecar}`);
      const sidecarImports = staticSpecifiers(parse(readFileSync(sidecar, 'utf8'), { sourceType: 'module' }));
      if (sidecarImports.length > 0) throw new Error(`${sidecar} is not self-contained`);
      if (format === 'docx') assertSidecarBudget(sidecar);
      if (format === 'pptx') {
        // The parse workers import the font-demand collector by URL; an inline
        // worker cannot resolve relative chunks, so it must be self-contained.
        for (const file of [join(dist, 'pptx-font-demand.mjs'), ...(packages ? [join('packages', 'pptx', 'dist', 'pptx-font-demand.mjs')] : [])]) {
          if (!existsSync(file)) throw new Error(`Missing font-demand sidecar ${file}`);
          if (staticSpecifiers(parse(readFileSync(file, 'utf8'), { sourceType: 'module' })).length > 0) {
            throw new Error(`${file} is not self-contained`);
          }
        }
        for (const payload of graph.codes.flatMap(inlineWorkers)) {
          if (payload.includes('wordArtVertRtl')) throw new Error('pptx inline worker inlines the font-demand collector');
        }
      }
      const packageSidecar = join('packages', format, 'dist', `${format}-source-worker.mjs`);
      if (packages && !existsSync(packageSidecar)) {
        throw new Error(`Missing optional package source sidecar ${packageSidecar}`);
      }
      if (packages && existsSync(packageSidecar)
        && staticSpecifiers(parse(readFileSync(packageSidecar, 'utf8'), { sourceType: 'module' })).length > 0) {
        throw new Error(`${packageSidecar} is not self-contained`);
      }
      if (packages && format === 'docx') assertSidecarBudget(packageSidecar);
      const packageEntry = join('packages', format, 'dist', 'index.mjs');
      if (packages) {
        if (!existsSync(packageEntry)) throw new Error(`Missing package entry ${packageEntry}`);
        const packageGraph = bundleGraph(packageEntry);
        assertNoSourceRuntime(packageGraph.joined, `${format} package static entry graph`);
        for (const payload of packageGraph.codes.flatMap(inlineWorkers)) {
          assertNoSourceRuntime(payload, `${format} package inline worker`);
        }
      }
    }
  }
  let ordinaryWorkers = 0;
  for (const file of readdirSync(join(dist, 'assets')).filter((name) => /^render-worker-.*\.js$/.test(name))) {
    ordinaryWorkers++;
    assertNoSourceRuntime(readFileSync(join(dist, 'assets', file), 'utf8'), file);
  }
  if (ordinaryWorkers !== 3) {
    throw new Error(`Expected 3 ordinary render workers, found ${ordinaryWorkers}`);
  }
  if (disabledDist) {
    for (const format of Object.keys(DISPATCH_COST_BUDGET)) {
      assertNoSourceRuntime(bundleGraph(join(disabledDist, `${format}.mjs`)).joined, `${format} feature-disabled graph`);
    }
  }
  if (stats) {
    for (const [format, budget] of Object.entries(DISPATCH_COST_BUDGET)) {
      assertDispatchCost(stats.enabled, stats.disabled, `${format}.mjs`, budget, format);
    }
  }
}

export function assertEagerSourceEntries(entries, forbidden = []) {
  const forbiddenPaths = new Set(forbidden.map((module) => resolve(module)));
  for (const entry of entries) {
    for (const module of eagerModules(entry)) {
      // Match future model-source owners as well as today's explicit list.
      // A type-only import is absent from eagerModules, as it should be.
      if (forbiddenPaths.has(module) || /(?:^|\/)(?:[^/]*model-source[^/]*|worker-source|render-worker-source)\.[cm]?[jt]sx?$/.test(module)) {
        throw new Error(`${entry} statically reaches ${module}`);
      }
    }
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  checkSourceDispatchImports(trackedSourceFiles());
  const forbidden = [
    'packages/core/src/source/model-source.ts',
    ...['docx', 'xlsx', 'pptx'].map((format) => `packages/${format}/src/internal/worker-${format === 'docx' ? 'document' : format === 'xlsx' ? 'worksheet' : 'presentation'}-source.ts`),
    ...['docx', 'xlsx'].map((format) => `packages/${format}/src/internal/node-model-source-acquisition.ts`),
    'packages/pptx/src/internal/node-session-acquisition.ts',
    ...['docx', 'xlsx', 'pptx'].map((format) => `packages/${format}/src/internal/model-source-session.ts`),
    'packages/xlsx/src/internal/host-layout-measure.ts',
    'packages/node/src/model-source.ts',
    ...['docx', 'xlsx', 'pptx'].map((format) => `packages/node/src/${format}-model-source.ts`),
    'packages/docx/src/internal/document-model-source.ts',
    'packages/xlsx/src/internal/workbook-model-source.ts',
    'packages/pptx/src/internal/presentation-model-source.ts',
    ...['docx', 'xlsx', 'pptx'].flatMap((format) => [
      `packages/${format}/src/worker-source.ts`,
    ]),
  ];
  assertEagerSourceEntries([
    'packages/core/src/index.ts',
    ...['docx', 'xlsx', 'pptx'].flatMap((format) => [
      `packages/${format}/src/index.ts`,
      `packages/${format}/src/worker.ts`,
      `packages/${format}/src/render-worker.ts`,
    ]),
    'packages/node/src/index.ts',
  ], forbidden);
  if (existsSync('dist/docx.mjs')) {
    withComparisonBuilds((disabledDist, stats) => checkBuiltBundles('dist', {
      packages: process.argv.includes('--packages'), disabledDist, stats,
    }));
  }
  console.log('OOXML entries and workers keep model-source runtime behind dynamic loads.');
}
