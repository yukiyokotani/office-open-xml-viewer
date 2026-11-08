/**
 * Hand-rolled Formualizer instantiation over the vendored static assets.
 *
 * The `formualizer` npm entry cannot be imported by consumers whose bundler
 * has no WASM plugin (Angular/esbuild): its glue uses wasm-pack
 * `--target bundler` output (`import … from "./formualizer_wasm_bg.wasm"`).
 * The wrapper itself (`vendor/formualizer_wasm_bg.js`) is plain JavaScript,
 * so this module fetches the sibling `.wasm` bytes — emitted as a real file
 * by the repo's `?url` asset handling and staged by consumers exactly like
 * the parser WASMs — and instantiates them manually:
 *
 * - the import object is derived at runtime from the binary
 *   (`WebAssembly.Module.imports`), so added/renamed `__wbg_*` shims keep
 *   working as long as the wrapper still exports them;
 * - then `__wbg_set_wasm(instance.exports)` + `__wbindgen_start()`,
 *   mirroring the package's own glue line-for-line.
 *
 * Memoized: the default (opt-out) path never fetches or compiles anything,
 * and concurrent recalcs share one init. Reset on failure so a later recalc
 * can retry instead of pinning a rejected promise.
 */
import wasmAssetUrl from './vendor/formualizer_wasm_bg.wasm?url';
import * as bg from './vendor/formualizer_wasm_bg.js';
import type { FormualizerEngine } from './vendor/formualizer_wasm_bg.js';

let engineInit: Promise<FormualizerEngine> | null = null;

function resolveHref(url: string, base: string | undefined): string {
  if (/^[a-z][a-z0-9+.-]*:/i.test(url)) return url;
  if (base !== undefined) return new URL(url, base).href;
  return url;
}

export function loadFormualizerEngine(
  input?: string | URL | Uint8Array,
): Promise<FormualizerEngine> {
  if (!engineInit) {
    engineInit = (async () => {
      const bytes =
        input instanceof Uint8Array
          ? input.slice().buffer
          : await (await fetch(resolveHref(
              input?.toString() ?? wasmAssetUrl,
              typeof location !== 'undefined' ? location.href : undefined,
            ))).arrayBuffer();
      const module = await WebAssembly.compile(bytes);
      const shimNames = WebAssembly.Module.imports(module).map((entry) => entry.name);
      const shims: Record<string, WebAssembly.ImportValue> = {};
      const source = bg as unknown as Record<string, unknown>;
      for (const name of shimNames) {
        const shim = source[name];
        if (typeof shim !== 'function') {
          throw new Error(`formualizer shim missing: ${name}`);
        }
        shims[name] = shim;
      }
      const instance = await WebAssembly.instantiate(module, {
        './formualizer_wasm_bg.js': shims,
      });
      bg.__wbg_set_wasm(instance.exports);
      (instance.exports.__wbindgen_start as () => void)();
      return bg as unknown as FormualizerEngine;
    })();
    engineInit.catch(() => {
      engineInit = null;
    });
  }
  return engineInit;
}
