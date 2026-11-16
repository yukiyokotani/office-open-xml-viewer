import { build, defineConfig, type Plugin, type ResolvedConfig } from 'vite';
import wasm from 'vite-plugin-wasm';
import { dts } from 'rolldown-plugin-dts';
import { resolve, dirname, basename, posix } from 'path';
import { fileURLToPath } from 'url';
import { readFile } from 'fs/promises';
import { gzipSync } from 'zlib';
import { rolldown } from 'rolldown';
import { docxCompatibilityMetadata } from './scripts/docx-compatibility-build.mjs';
import { optionalRendererMarkers } from './scripts/optional-renderer-markers.mjs';
import { moduleStatsPlugin } from './scripts/module-stats-plugin.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Emit `?url` asset imports as real asset files instead of base64 `data:` URLs.
 * (Named hashlessly here — `rollupOptions.output.assetFileNames` is `[name]`.)
 *
 * Vite **library mode** force-inlines every `?url` asset as a
 * `data:<mime>;base64,…` string regardless of `assetsInlineLimit` (a number or a
 * `() => false` function does NOT override it — `build.lib` unconditionally
 * returns `true` from Vite's internal `shouldInline`). Heavy asset kinds on
 * this path include:
 *   - the three parser WASM modules (`*_parser_bg.wasm?url`, ~0.6–0.7 MB each) —
 *     base64 inflates them +33 % and blocks `WebAssembly.compileStreaming`
 *     (a data URL cannot be fetch-streamed; the worker must `atob` by hand);
 *   - the MathJax + STIX Two Math engine (`assets/mathjax-stix2.js?url`, ~3 MB) —
 *     inlined it turned the opt-in `math.mjs` chunk into a 4.1 MB base64 blob,
 *     even though consumers only import it when a document actually has equations.
 * The parser and math assets use plain `?url` imports in their owner modules
 * (the format main-thread handles and `math/engine.ts`). We intercept those
 * imports here, `emitFile` the bytes
 * as an asset next to the chunk, and hand back the standard ESM asset reference
 * `new URL('<name>', import.meta.url)` — the form Vite / webpack 5 / Rollup /
 * Turbopack rewrite when they re-bundle our `.mjs`, and which resolves
 * correctly for a plain `<script type=module>` too. (esbuild — and therefore
 * the Angular CLI — does NOT process it; those consumers use the `wasmUrl`
 * load option, see the README bundler note.) wasm-bindgen's `--target web`
 * glue then `fetch()`es its URL and hits `instantiateStreaming`; the math engine
 * is lazy-loaded via a `<script src>` pointed at the emitted asset.
 *
 * Runs with `enforce: 'pre'` and claims every `?url` import; a bare-`.wasm`
 * import (owned by `vite-plugin-wasm`) is untouched. Nothing in the tree imports
 * bare `.wasm`, and every `?url` here is a real on-disk asset we want emitted —
 * exactly what Vite's non-lib mode would do anyway.
 */
export function wasmAssetUrl(
  readAsset: (path: string) => Promise<Uint8Array> = readFile,
  modulePlugins: Plugin[] = [],
): Plugin {
  const SUFFIX = '?url';
  return {
    name: 'wasm-asset-url',
    enforce: 'pre',
    // Build-only: emitFile/ROLLUP_FILE_URL are Rollup build machinery and do
    // not exist on the dev server. In dev, Vite's stock `?url` handling serves
    // the file directly (the pre-E4 behavior) — intercepting there returned an
    // unresolvable reference and broke every WASM load (caught by CI smoke).
    apply: 'build',
    async load(id) {
      if (!id.endsWith(SUFFIX)) return null;
      const filePath = id.slice(0, -SUFFIX.length);
      // A `?url` import of a TypeScript module names a self-contained ES
      // module asset (an application-loaded model source module, see
      // packages/legacy-office/src/legacy-*.ts). It is bundled on its own,
      // with every dependency inlined and its exports preserved, because the
      // realm that imports it by URL cannot resolve sibling chunks.
      const moduleAsset = /\.[cm]?ts$/.test(filePath);
      const source = moduleAsset
        ? await bundleSelfContainedModule(filePath, modulePlugins)
        : await readAsset(filePath);
      const referenceId = this.emitFile({
        type: 'asset',
        name: moduleAsset ? basename(filePath).replace(/\.[cm]?ts$/, '.mjs') : basename(filePath),
        source,
      });
      // `import.meta.ROLLUP_FILE_URL_<id>` expands at render time to Rollup's
      // ES default resolution — `new URL('<name>', import.meta.url).href` — an
      // absolute href the worker (or the math engine's `<script>` loader) can
      // fetch from any realm. It must be emitted BARE: wrapping it in another
      // `new URL(…, import.meta.url)` ships a nested pattern that webpack 5
      // compiles into a critical-dependency ContextModule (the outer first arg
      // is now an expression, not a literal), throwing MODULE_NOT_FOUND at
      // module evaluation — `import '@silurus/ooxml/docx'` alone crashed every
      // webpack consumer. The single-level form is the exact shape webpack 5 /
      // Turbopack / Vite statically rewrite when re-bundling our `.mjs`.
      return `export default import.meta.ROLLUP_FILE_URL_${referenceId};`;
    },
    renderChunk(code, _chunk, output) {
      if (output.format !== 'cjs') return null;
      // Oxc lowers `import.meta` to `{}` before Rolldown resolves emitted
      // asset URLs for CJS. Restore the Node entry's own file URL for static
      // sidecars; otherwise importing the CJS package throws immediately.
      const restored = code.replace(
        /new URL\((["'`])([^"'`]+\.(?:wasm|ttf|mjs))\1,\s*\{\}\.url\)\.href/g,
        (_match, _quote, file) =>
          `new URL(${JSON.stringify(file)}, require('node:url').pathToFileURL(__filename)).href`,
      );
      return restored === code ? null : { code: restored, map: null };
    },
  };
}

/** Resolve the emitted chunk graph, including aliases, before publishing. */
export function legacyBundleBoundary(): Plugin {
  return {
    name: 'legacy-bundle-boundary',
    generateBundle(_options, bundle) {
      // Every root entry and worker must stay independent of the companion.
      // Follow static and dynamic edges so a relay outside the guarded
      // source roots cannot hide a reader in a shared chunk.
      const entries = Object.values(bundle).filter((output) =>
        output.type === 'chunk' && output.isEntry
        && !output.fileName.startsWith('.types-work/'));
      const visited = new Set<string>();
      const pending = [...entries];
      while (pending.length > 0) {
        const output = pending.pop()!;
        if (output.type !== 'chunk') continue;
        if (visited.has(output.fileName)) continue;
        visited.add(output.fileName);
        for (const moduleId of Object.keys(output.modules)) {
          if (/(?:^|[\\/])packages[\\/]legacy-office(?:[\\/]|$)/.test(moduleId)
            || /@silurus\/ooxml-legacy-office/.test(moduleId)) {
            this.error(`${output.fileName} contains forbidden legacy module ${moduleId}`);
          }
          // The OpenDocument companion is held to the same boundary.
          if (/(?:^|[\\/])packages[\\/]opendocument(?:[\\/]|$)/.test(moduleId)
            || /@silurus\/ooxml-opendocument/.test(moduleId)) {
            this.error(`${output.fileName} contains forbidden OpenDocument module ${moduleId}`);
          }
        }
        for (const imported of [...output.imports, ...(output.dynamicImports ?? [])]) {
          const child = bundle[imported];
          if (child?.type === 'chunk') pending.push(child);
        }
      }
    },
  };
}

/** The installed viewer owns layout, selected-source sidecars and transport.
 * Check retained modules after tree shaking: importing stateless WASM guards
 * through core's worker barrel must not count its unused transport exports.
 * Apply this to opaque source-module builds too; the outer build sees only
 * their emitted asset bytes, not their dependency graphs. */
export function companionBundleBoundary(): Plugin {
  return {
    name: 'companion-bundle-boundary',
    generateBundle(_options, bundle) {
      for (const output of Object.values(bundle)) {
        if (output.type !== 'chunk') continue;
        for (const [id, module] of Object.entries(output.modules)) {
          if (module.renderedLength === 0) continue;
          const path = id.replaceAll('\\', '/');
          if (/\/packages\/(?:docx|xlsx|pptx|node)\/src\//.test(path)
            || /\/packages\/core\/src\/source\/model-source\.ts$/.test(path)
            || (/\/packages\/core\/src\/worker\//.test(path)
              && !/\/(?:index|wasm-guard)\.ts$/.test(path))) {
            this.error(`${output.fileName} contains viewer-owned runtime ${id}`);
          }
        }
      }
    },
  };
}

/** Bundle one ES module with all of its imports inlined and its exports kept. */
async function bundleSelfContainedModule(entry: string, plugins: Plugin[]): Promise<string> {
  const bundle = await rolldown({
    input: entry,
    plugins: [docxCompatibilityMetadata(), ...plugins],
    platform: 'neutral',
    resolve: {
      alias: {
        '@silurus/ooxml-core/worker': resolve(__dirname, 'packages/core/src/worker/index.ts'),
        '@silurus/ooxml-core/internal/model-source': resolve(__dirname, 'packages/core/src/source/model-source.ts'),
        '@silurus/ooxml-core': resolve(__dirname, 'packages/core/src/index.ts'),
      },
      extensionAlias: { '.js': ['.ts', '.js'] },
    },
    preserveEntrySignatures: 'strict',
  });
  try {
    const { output } = await bundle.generate({ format: 'es', codeSplitting: false, minify: true });
    if (output.length !== 1 || output[0].type !== 'chunk') {
      throw new Error(`${basename(entry)} did not bundle into one self-contained module`);
    }
    return output[0].code;
  } finally {
    await bundle.close();
  }
}

/**
 * Merge the small shared static chunks of the published entries.
 *
 * Lazy boundaries (font sources, model sources, optional renderers) split the
 * modules an entry loads at start-up into many small chunks, one per
 * combination of static and lazy importers. An unbundled (CDN) consumer pays a
 * request and lost compression for each; a re-bundling consumer is unaffected.
 *
 * Which modules an entry loads at start-up is known only after tree shaking,
 * so the library build first runs a probe build (same configuration,
 * automatic chunking, nothing written) and then assigns the probe's chunks to
 * `codeSplitting` groups. Chunks are merged only when
 *  - exactly the same entries, optional entries included, load them
 *    statically, so no entry loads anything at start-up that it did not load
 *    before, and lazy code stays lazy;
 *  - exactly the same entries reach them only through lazy imports, so a
 *    merged file is fetched lazily only by entries that already fetched part
 *    of it lazily (two lazy imports of one entry may now share one file);
 *  - no optional entry loads them only lazily, so those lazy loads stay as
 *    they were;
 *  - none is an entry or contains a dynamic `import()`, so a lazy chunk
 *    importing a group back cannot make a consumer's Rolldown retain the
 *    group through rolldown/rolldown#11213 (see `avoidRedundantChunkLoads`
 *    below);
 *  - none has top-level code that a consumer's bundler keeps whenever the
 *    file is imported at all, so a consumer that re-bundles the package keeps
 *    no more code than before (a `sideEffects: false` module is dropped only
 *    when nothing from it is used);
 *  - the merged file compresses no worse than the separate files.
 *
 * `preserveEntrySignatures: 'strict'` requires `includeDependenciesRecursively`,
 * which captures every retained module a grouped module imports, also through
 * modules tree shaking removed (re-export barrels). A chunk joins a group
 * only when everything it pulls in that way is in its own group or in a group
 * captured before it (the groups form an acyclic graph, captured
 * dependencies first); any other chunk stays with automatic chunking.
 *
 * The published output is checked against the probe: every entry must load
 * exactly the same modules at start-up, and every group must hold exactly its
 * planned modules.
 */
// Marks the nested probe build; passed in its inline configuration so that
// no other build in the process is affected.
const CHUNK_PROBE = 'static-chunk-plan:probe';
const MAX_GROUPS = 256;

type ProbeChunk = {
  fileName: string;
  name: string;
  code: string;
  isEntry: boolean;
  facadeModuleId: string | null;
  imports: string[];
  dynamicImports: string[];
  moduleIds: string[];
};
type ModuleGraph = (id: string) => { importedIds: readonly string[]; dynamicallyImportedIds: readonly string[] } | null;
type ChunkGroup = { name: string; moduleIds: string[]; members: Set<string> };

export function staticChunkPlan(
  defaultEntries: Record<string, string>,
  optionalEntries: Record<string, string>,
): { plugin: Plugin; groups: Array<{ name: (id: string) => string | null; priority: number }> } {
  const allEntries = { ...defaultEntries, ...optionalEntries };
  let resolved: ResolvedConfig | undefined;
  let probing = false;
  let probe: { chunks: Map<string, ProbeChunk>; loads: Map<string, string>; residues: Set<string> } | undefined;
  let groups: ChunkGroup[] | undefined;
  const plugin: Plugin = {
    name: 'static-chunk-plan',
    apply: 'build',
    configResolved(config) {
      resolved = config;
      probing = config.plugins.some((candidate) => candidate.name === CHUNK_PROBE);
    },
    async buildStart(options) {
      probe = undefined;
      groups = undefined;
      if (probing) return;
      // Only the library build: Storybook's build reuses this configuration.
      const input = Object.values(options.input ?? {});
      if (!Object.values(allEntries).every((entry) => input.includes(entry))) return;
      const result = await build({
        configFile: resolved!.configFile,
        mode: resolved!.mode,
        logLevel: 'silent',
        plugins: [{ name: CHUNK_PROBE }],
        build: { write: false, emptyOutDir: false },
      });
      const output = (Array.isArray(result) ? result[0] : result) as { output: Array<ProbeChunk & { type: string }> };
      const chunks = new Map(output.output
        .filter((item) => item.type === 'chunk')
        .map((chunk) => [chunk.fileName, chunk]));
      probe = { chunks, loads: startupLoads(chunks), residues: await topLevelResidues(chunks) };
    },
    buildEnd() {
      if (!probe) return;
      groups = planGroups(probe.chunks, probe.residues, (id) => this.getModuleInfo(id));
      if (groups.length > MAX_GROUPS) this.error(`static-chunk-plan: ${groups.length} groups exceed ${MAX_GROUPS} slots`);
    },
    generateBundle(_options, bundle) {
      if (!probe || !groups) return;
      const chunks = new Map(Object.values(bundle)
        .filter((output) => output.type === 'chunk')
        .map((chunk) => [chunk.fileName, chunk as unknown as ProbeChunk]));
      const loads = startupLoads(chunks);
      for (const entryName of Object.keys(allEntries)) {
        if (loads.get(entryName) !== probe.loads.get(entryName)) {
          this.error(`static-chunk-plan: grouping changed the modules ${entryName}.mjs loads at start-up`);
        }
      }
      for (const { name, moduleIds } of groups) {
        const chunk = [...chunks.values()].find((candidate) => candidate.name === name);
        if (chunk?.moduleIds.slice().sort().join('\n') !== moduleIds.slice().sort().join('\n')) {
          this.error(`static-chunk-plan: group ${name} does not hold exactly its planned modules`);
        }
      }
    },
  };
  // The groups are known only once the module graph is complete, while the
  // output options are fixed before the build: reserve ordered slots, and let
  // slot i hold the i-th planned group. A higher-priority group captures its
  // modules first.
  const slots = Array.from({ length: MAX_GROUPS }, (_, index) => ({
    name: (id: string): string | null => {
      const group = groups?.[index];
      return group?.members.has(id) ? group.name : null;
    },
    priority: MAX_GROUPS - index,
  }));
  return { plugin, groups: slots };

  function entryChunk(chunks: Map<string, ProbeChunk>, id: string): string {
    const chunk = [...chunks.values()].find((candidate) => candidate.isEntry && candidate.facadeModuleId === id);
    if (!chunk) throw new Error(`static-chunk-plan: no entry chunk for ${id}`);
    return chunk.fileName;
  }

  function reach(chunks: Map<string, ProbeChunk>, entry: string, dynamic: boolean): Set<string> {
    const seen = new Set<string>();
    const pending = [entryChunk(chunks, entry)];
    while (pending.length > 0) {
      const fileName = pending.pop()!;
      const chunk = chunks.get(fileName);
      if (seen.has(fileName) || !chunk) continue;
      seen.add(fileName);
      pending.push(...chunk.imports, ...(dynamic ? chunk.dynamicImports : []));
    }
    return seen;
  }

  /**
   * The chunks with top-level code a consumer's bundler keeps whenever the
   * chunk is imported at all (a `sideEffects: false` package drops a module
   * only when nothing from it is used). Merged into a group, that code would
   * be kept for every importer of anything in the group.
   */
  async function topLevelResidues(chunks: Map<string, ProbeChunk>): Promise<Set<string>> {
    const loaded = new Set(Object.values(defaultEntries).flatMap((entry) => [...reach(chunks, entry, false)]));
    const impure = await Promise.all([...loaded].filter((fileName) => !chunks.get(fileName)!.isEntry).map(async (fileName) => {
      const bundle = await rolldown({
        input: '\0residue-probe',
        logLevel: 'silent',
        external: (id) => !id.startsWith('\0') && !chunks.has(posix.normalize(id)),
        treeshake: { moduleSideEffects: (id) => id === fileName },
        plugins: [{
          name: 'probe-chunks',
          resolveId: (id, importer) => {
            if (id === '\0residue-probe') return id;
            const target = importer && !importer.startsWith('\0') ? posix.join(posix.dirname(importer), id) : posix.normalize(id);
            return chunks.has(target) ? target : { id, external: true };
          },
          load: (id) => (id === '\0residue-probe' ? `import ${JSON.stringify(fileName)};` : chunks.get(id)?.code ?? null),
        }],
      });
      try {
        const { output } = await bundle.generate({ format: 'es', minify: true, codeSplitting: false });
        return output[0].code.trim().length > 0 ? fileName : undefined;
      } finally {
        await bundle.close();
      }
    }));
    return new Set(impure.filter((fileName): fileName is string => fileName !== undefined));
  }

  /** Each entry's start-up modules, as one comparable string. */
  function startupLoads(chunks: Map<string, ProbeChunk>): Map<string, string> {
    return new Map(Object.entries(allEntries).map(([entryName, entry]) => [entryName,
      [...reach(chunks, entry, false)].flatMap((fileName) => chunks.get(fileName)!.moduleIds).sort().join('\n')]));
  }

  function planGroups(chunks: Map<string, ProbeChunk>, residues: Set<string>, graph: ModuleGraph): ChunkGroup[] {
    // The entries that load each chunk statically, and those that load it
    // only through a lazy import.
    const signature = new Map<string, string>();
    for (const [entryName, entry] of Object.entries(allEntries)) {
      const loads = reach(chunks, entry, false);
      for (const fileName of reach(chunks, entry, true)) {
        const mark = loads.has(fileName) ? entryName : `lazy:${entryName}`;
        signature.set(fileName, signature.has(fileName) ? `${signature.get(fileName)}+${mark}` : mark);
      }
    }
    const loaded = new Set(Object.values(defaultEntries).flatMap((entry) => [...reach(chunks, entry, false)]));
    const optionalLazy = new Set(Object.values(optionalEntries).flatMap((entry) => {
      const own = reach(chunks, entry, false);
      return [...reach(chunks, entry, true)].filter((fileName) => !own.has(fileName));
    }));
    const owner = new Map([...chunks.values()].flatMap((chunk) =>
      chunk.moduleIds.map((id) => [id, chunk.fileName])));
    // The chunks whose modules capturing `fileName` would pull in.
    const pulls = new Map<string, Set<string>>();
    for (const fileName of loaded) {
      const found = new Set<string>();
      const seen = new Set<string>();
      const pending = chunks.get(fileName)!.moduleIds.flatMap((id) => graph(id)?.importedIds ?? []);
      while (pending.length > 0) {
        const dep = pending.pop()!;
        if (seen.has(dep)) continue;
        seen.add(dep);
        if (owner.has(dep)) found.add(owner.get(dep)!);
        else pending.push(...(graph(dep)?.importedIds ?? []));
      }
      found.delete(fileName);
      pulls.set(fileName, found);
    }
    // Candidate groups: the chunks the same entries load, of those that may
    // be merged at all.
    let units = new Map<string, string[]>();
    for (const fileName of loaded) {
      const chunk = chunks.get(fileName)!;
      if (chunk.isEntry || optionalLazy.has(fileName) || residues.has(fileName) || chunk.dynamicImports.length > 0
        || chunk.moduleIds.some((id) => id.startsWith('\0') || (graph(id)?.dynamicallyImportedIds.length ?? 0) > 0)) continue;
      const key = signature.get(fileName)!;
      units.set(key, [...(units.get(key) ?? []), fileName]);
    }
    // Drop chunks until every chunk pulls only from its own group or from
    // groups that can be captured before it (the group graph is acyclic).
    for (let changed = true; changed;) {
      changed = false;
      const unitOf = new Map([...units].flatMap(([key, files]) => files.map((fileName) => [fileName, key])));
      const keep = (key: string, test: (fileName: string) => boolean) => {
        const files = units.get(key)!.filter(test);
        if (files.length === units.get(key)!.length) return;
        changed = true;
        if (files.length > 0) units.set(key, files);
        else units.delete(key);
      };
      for (const key of [...units.keys()]) keep(key, (fileName) => [...pulls.get(fileName)!].every((dep) => unitOf.has(dep)));
      if (changed) continue;
      // Merging must not make the files compress worse than apart (one
      // request each); smallest chunks are merged first.
      for (const key of [...units.keys()]) {
        const accepted: string[] = [];
        let packed = 0;
        for (const fileName of units.get(key)!.slice().sort((a, b) => chunks.get(a)!.code.length - chunks.get(b)!.code.length)) {
          const together = gzipSize(accepted.concat(fileName).map((member) => chunks.get(member)!.code).join('\n'));
          if (accepted.length === 0 || together <= packed + gzipSize(chunks.get(fileName)!.code)) {
            accepted.push(fileName);
            packed = together;
          }
        }
        keep(key, (fileName) => accepted.includes(fileName));
      }
      if (changed) continue;
      const edges = new Map([...units].map(([key, files]) => [key, new Set(files.flatMap((fileName) =>
        [...pulls.get(fileName)!].map((dep) => unitOf.get(dep)!)).filter((dep) => dep !== key))]));
      for (const cycle of stronglyConnected(edges).filter((component) => component.size > 1)) {
        for (const key of cycle) {
          keep(key, (fileName) => ![...pulls.get(fileName)!].some((dep) => unitOf.get(dep) !== key && cycle.has(unitOf.get(dep)!)));
        }
      }
    }
    // Dependencies first; a group of one chunk only renames it.
    const edges = new Map([...units].map(([key, files]) => [key, new Set(files.flatMap((fileName) =>
      [...pulls.get(fileName)!].map((dep) => [...units].find(([, members]) => members.includes(dep))![0])).filter((dep) => dep !== key))]));
    const order: string[] = [];
    const visit = (key: string) => {
      if (order.includes(key)) return;
      for (const dep of edges.get(key)!) visit(dep);
      order.push(key);
    };
    for (const key of units.keys()) visit(key);
    // A group is identified by its name, which must not repeat an entry's.
    const used = new Set<string>(Object.keys(allEntries));
    return order.map((key) => {
      const files = units.get(key)!;
      const largest = files.reduce((a, b) => (chunks.get(b)!.code.length > chunks.get(a)!.code.length ? b : a));
      let name = chunks.get(largest)!.name;
      for (let index = 2; used.has(name); index++) name = `${chunks.get(largest)!.name}${index}`;
      used.add(name);
      const moduleIds = files.flatMap((fileName) => chunks.get(fileName)!.moduleIds);
      return { name, moduleIds, members: new Set(moduleIds) };
    });
  }
}

function gzipSize(code: string): number {
  return gzipSync(code, { level: 9 }).length;
}

/** Tarjan's strongly connected components. */
function stronglyConnected(edges: Map<string, Set<string>>): Set<string>[] {
  const index = new Map<string, number>();
  const low = new Map<string, number>();
  const stack: string[] = [];
  const onStack = new Set<string>();
  const components: Set<string>[] = [];
  const connect = (node: string) => {
    index.set(node, index.size);
    low.set(node, index.get(node)!);
    stack.push(node);
    onStack.add(node);
    for (const next of edges.get(node) ?? []) {
      if (!index.has(next)) {
        connect(next);
        low.set(node, Math.min(low.get(node)!, low.get(next)!));
      } else if (onStack.has(next)) {
        low.set(node, Math.min(low.get(node)!, index.get(next)!));
      }
    }
    if (low.get(node) === index.get(node)) {
      const component = new Set<string>();
      for (let member = stack.pop()!; ; member = stack.pop()!) {
        onStack.delete(member);
        component.add(member);
        if (member === node) break;
      }
      components.push(component);
    }
  };
  for (const node of edges.keys()) if (!index.has(node)) connect(node);
  return components;
}

const chunkPlan = staticChunkPlan(
  Object.fromEntries(['index', 'docx', 'xlsx', 'pptx', 'node']
    .map((name) => [name, resolve(__dirname, `src/${name}.ts`)])),
  Object.fromEntries(['math', 'three-d', 'region-map', 'chart-ex', 'tiff', 'internal-companion']
    .map((name) => [name, resolve(__dirname, `src/${name}.ts`)])),
);

export default defineConfig(({ command, mode }) => ({
  // The private comparison build removes selected-source dispatch and code
  // reachable only through it. Published and dev builds always use true.
  define: { __OOXML_MODEL_SOURCES__: mode === 'model-sources-off' ? 'false' : 'true' },
  // Published library assets must resolve from the imported module URL, not
  // from the hosting page's origin root. This is especially important for the
  // standalone module workers and sibling assets when consumers serve the
  // package below a subpath or from a CDN.
  base: './',
  experimental: {
    // A consumer bundler re-resolves `new URL('<literal>', import.meta.url)`.
    // Vite resolves a literal that does not start with `.` through the app's
    // aliases first; Nuxt (like any app with an `assets` alias) maps the
    // render workers' `assets/render-worker-*.js` to its own folder, so the
    // worker was never emitted (#2134). JS-hosted built URLs are therefore
    // written dot-relative to the module that references them.
    renderBuiltUrl: (filename, { hostId, hostType }) => {
      if (hostType !== 'js') return { relative: true };
      const path = posix.relative(posix.dirname(hostId), filename);
      return { runtime: `new URL(${JSON.stringify(path.startsWith('.') ? path : `./${path}`)}, import.meta.url).href` };
    },
  },
  plugins: [
    docxCompatibilityMetadata(),
    optionalRendererMarkers(),
    wasmAssetUrl(),
    wasm(),
    legacyBundleBoundary(),
    moduleStatsPlugin(__dirname),
    chunkPlan.plugin,
    // Storybook loads the root Vite config in serve mode. The declaration
    // plugins are build-only: their Rolldown buildStart hooks expect library
    // inputs and fail against Storybook's dev-server graph.
    ...(command === 'build' && mode !== 'runtime' && mode !== 'model-sources-off'
      ? dts({
          // TypeScript 7 is the repository's sole compiler. Its native tsgo
          // declaration generator avoids the removed JavaScript Compiler API.
          generator: 'tsgo',
          tsconfig: './tsconfig.lib.json',
        })
      : []),
  ],
  // rolldown-plugin-dts emits declarations into the same graph. Oxc should
  // leave those declaration modules (and normal JavaScript inputs) untouched.
  oxc: {
    exclude: [/\.js$/, /\.d\.[cm]?ts$/],
  },
  build: {
    lib: {
      // Keep both comparison builds on the same entry set: changing the set
      // would also change shared-chunk factoring and distort the byte delta.
      entry: {
        index: resolve(__dirname, 'src/index.ts'),
        pptx:  resolve(__dirname, 'src/pptx.ts'),
        xlsx:  resolve(__dirname, 'src/xlsx.ts'),
        docx:  resolve(__dirname, 'src/docx.ts'),
        // Opt-in math engine (MathJax + STIX Two Math). Separate entry so the
        // ~3 MB asset stays out of the docx/pptx bundles unless imported.
        math:  resolve(__dirname, 'src/math.ts'),
        // Opt-in model-space 3-D chart mesh/camera painter. Ordinary format
        // bundles keep the lightweight contract and 2-D fallback only.
        'three-d': resolve(__dirname, 'src/three-d.ts'),
        // Opt-in offline ChartEx Region Map geometry/projector.
        'region-map': resolve(__dirname, 'src/region-map.ts'),
        // Opt-in Microsoft ChartEx family renderer. Classic 2-D charts remain
        // part of every format entry.
        'chart-ex': resolve(__dirname, 'src/chart-ex.ts'),
        // Opt-in TIFF 6.0 software decoder. Native raster users retain only the
        // lightweight codec contract and header guard.
        tiff: resolve(__dirname, 'src/tiff.ts'),
        'internal-companion': resolve(__dirname, 'src/internal-companion.ts'),
        // Node-only bounded sessions and server render helpers. Kept as a
        // separate entry so browser consumers never load Node built-ins.
        node:  resolve(__dirname, 'src/node.ts'),
      },
      // ESM-only: the published bundle inlines a large math engine; emitting a
      // duplicate CJS copy of every chunk roughly doubled the package size.
      // Every modern bundler (Vite / webpack / Rollup / esbuild / Next) and
      // Node ≥ 20 consume ESM, so we ship `.mjs` only.
      formats: ['es'],
      fileName: (_format, name) =>
        name.endsWith('.d')
          ? `.types-work/${name.slice(0, -2)}.d.ts`
          : `${name}.mjs`,
    },
    rollupOptions: {
      // Keep consumer PURE annotations; omit only per-module debug regions.
      // The boundary checks retain their explicit ownership markers instead.
      experimental: {
        attachDebugInfo: 'none',
        // Keep modules a lazy chunk shares with its importer in their own
        // chunk instead of folding them into the importer. Folding makes
        // lazy chunks import the format chunk back, and a consumer's Rolldown
        // keeps a chunk that two of its lazy chunks import back even when
        // nothing in it is used: a factories-only import such as
        // `import { googleFonts } from '@silurus/ooxml'` then retained the
        // whole DOCX and PPTX runtime (scripts/check-font-source-tree-shaking.mjs).
        // Upstream: rolldown/rolldown#11213 (still open as of Rolldown 1.2.15);
        // re-test before removing this once it is fixed.
        chunkOptimization: { avoidRedundantChunkLoads: false },
      },
      external: [/^node:/, 'skia-canvas'],
      output: {
        codeSplitting: { groups: chunkPlan.groups },
        assetFileNames: '[name][extname]',
        chunkFileNames: (chunk) =>
          chunk.name.endsWith('.d')
            ? '.types-work/[name]-[hash].d.ts'
            : '[name]-[hash].js',
      },
    },
    target: 'esnext',
  },
  worker: {
    format: 'es',
    // Built-in worker renderers lazy-import the same optional math engine. Keep
    // its ~3 MB `?url` asset external in nested worker builds too; otherwise
    // library mode base64-inlines one copy into every format worker chunk.
    plugins: () => [docxCompatibilityMetadata(), wasmAssetUrl(), wasm(), legacyBundleBoundary()],
    rollupOptions: {
      experimental: { attachDebugInfo: 'none' },
      output: {
        assetFileNames: '[name][extname]',
        // The published format entry is commonly re-bundled by a consumer.
        // A prebuilt worker URL is an opaque asset to that second bundler, so
        // sibling JS chunks imported by the worker would not be copied. Keep
        // each render worker as one self-contained module asset; optional
        // renderer code is fetched only when mode:'worker' loads this asset.
        codeSplitting: false,
      },
    },
  },
}));
