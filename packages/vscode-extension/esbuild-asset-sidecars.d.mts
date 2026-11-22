import type { Plugin } from 'esbuild';

export declare const bundledAssetSidecars: Plugin;
export declare const singleParserWasmPerFormat: Plugin;
export declare function parserWasmOutputProblems(outputs: Iterable<string>): string[];
