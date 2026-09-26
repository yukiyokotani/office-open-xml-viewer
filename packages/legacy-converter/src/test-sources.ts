// Test-only legacy model sources with explicit URLs: vitest resolves the
// file: source modules through its module runner and the WASM from disk.
import { legacyPptSource } from './legacy-ppt.js';
import { legacyXlsSource } from './legacy-xls.js';

const here = (path: string): string => new URL(path, import.meta.url).href;

export const TEST_SOURCE_URLS = {
  xls: { wasmUrl: here('./wasm-direct-xls/legacy_xls_direct_bg.wasm'), moduleUrl: here('./legacy-xls-source-module.ts') },
  ppt: { wasmUrl: here('./wasm-direct-ppt/legacy_ppt_direct_bg.wasm'), moduleUrl: here('./legacy-ppt-source-module.ts') },
} as const;

export const testXlsSource = (options: { maxInputBytes?: number } = {}) => legacyXlsSource({ ...TEST_SOURCE_URLS.xls, ...options });
export const testPptSource = (options: { maxInputBytes?: number } = {}) => legacyPptSource({ ...TEST_SOURCE_URLS.ppt, ...options });
