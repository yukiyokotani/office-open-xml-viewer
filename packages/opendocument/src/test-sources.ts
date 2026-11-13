// Test-only OpenDocument sources with explicit URLs: vitest resolves the file:
// source module through its module runner and the WASM from disk.
import { createOds } from './ods.js';
import { createOdt } from './odt.js';

const here = (path: string): string => new URL(path, import.meta.url).href;

export const testOdtSource = () => createOdt({
  wasmUrl: here('./wasm-odt/opendocument_odt_bg.wasm'),
  moduleUrl: here('./odt-source-module.ts'),
});

export const testOdsSource = () => createOds({
  wasmUrl: here('./wasm-ods/opendocument_ods_bg.wasm'),
  moduleUrl: here('./ods-source-module.ts'),
});
