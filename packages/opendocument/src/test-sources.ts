// Test-only OpenDocument sources with explicit URLs: vitest resolves the file:
// source module through its module runner and the WASM from disk.
import { createOdt } from './odt.js';
import { createOdp } from './odp.js';

const here = (path: string): string => new URL(path, import.meta.url).href;

export const testOdtSource = () => createOdt({
  wasmUrl: here('./wasm-odt/opendocument_odt_bg.wasm'),
  moduleUrl: here('./odt-source-module.ts'),
});

export const testOdpSource = () => createOdp({
  wasmUrl: here('./wasm-odp/opendocument_odp_bg.wasm'),
  moduleUrl: here('./odp-source-module.ts'),
});
