import { readFile } from 'node:fs/promises';

/**
 * Re-bundle Vite library-mode sidecars into the VS Code webview. Vite writes
 * `new URL("parser.wasm", import.meta.url).href` cannot survive esbuild's IIFE
 * output. Converting static asset URLs back to file imports gives esbuild
 * ownership of the files and emits relative URLs under dist/assets. The
 * webview HTML sets its base to the script directory, and VS Code grants that
 * directory as a local root.
 */
export const bundledAssetSidecars = {
  name: 'bundled-asset-sidecars',
  setup(build) {
    // Entry modules are .mjs, but code-split chunks (where the DOCX and PPTX
    // parsers live) are emitted as .js; both carry asset URLs.
    build.onLoad({ filter: /\/packages\/(?:docx|xlsx|pptx)\/dist\/[^/]+\.m?js$/ }, async ({ path }) => {
      const source = await readFile(path, 'utf8');
      const imports = new Map();
      const contents = source.replace(
        /new URL\((["'])([^"']+\.(?:ttf|wasm))\1, import\.meta\.url\)\.href/g,
        (_match, _quote, file) => {
          let identifier = imports.get(file);
          if (!identifier) {
            identifier = `__ooxmlFontAsset${imports.size}`;
            imports.set(file, identifier);
          }
          return identifier;
        },
      );
      if (imports.size === 0) return { contents: source, loader: 'js' };
      const prefix = [...imports].map(([file, id]) => `import ${id} from ${JSON.stringify(`./${file}`)};`).join('\n');
      return { contents: `${prefix}\n${contents}`, loader: 'js' };
    });
  },
};

const PARSER_FORMATS = ['docx', 'xlsx', 'pptx'];

/**
 * The webview bundle must emit exactly one parser WASM per format and no other
 * WASM. The bootstrap imports each parser through the package's `wasm-binary`
 * export while the viewer references the copy in the package's `dist`; esbuild
 * writes one file for both only while their bytes are identical. A stale
 * `dist` would otherwise ship two parsers for a format, and the webview would
 * hand the viewer one it was not built with.
 *
 * @param {Iterable<string>} outputs emitted output paths
 * @returns {string[]} one message per violation; empty when the outputs are valid
 */
export function parserWasmOutputProblems(outputs) {
  const wasm = [...outputs].map((output) => output.replaceAll('\\', '/')).filter((output) => output.endsWith('.wasm'));
  const problems = [];
  for (const format of PARSER_FORMATS) {
    const emitted = wasm.filter((output) => new RegExp(`(?:^|/)${format}_parser_bg-[A-Za-z0-9]+\\.wasm$`).test(output));
    if (emitted.length !== 1) {
      problems.push(`expected one ${format} parser WASM, found ${emitted.length}${emitted.length ? `: ${emitted.join(', ')}` : ''}`);
    }
  }
  const other = wasm.filter((output) => !/(?:^|\/)(?:docx|xlsx|pptx)_parser_bg-[A-Za-z0-9]+\.wasm$/.test(output));
  if (other.length > 0) problems.push(`unexpected WASM output: ${other.join(', ')}`);
  return problems;
}

/** Fail the build when {@link parserWasmOutputProblems} reports a violation. */
export const singleParserWasmPerFormat = {
  name: 'single-parser-wasm-per-format',
  setup(build) {
    build.initialOptions.metafile = true;
    build.onEnd((result) => {
      if (result.errors.length > 0 || !result.metafile) return;
      const problems = parserWasmOutputProblems(Object.keys(result.metafile.outputs));
      if (problems.length === 0) return;
      return {
        errors: problems.map((text) => ({
          pluginName: 'single-parser-wasm-per-format',
          text: `${text}. Rebuild the parser WASM and the library packages so that src/wasm and dist match.`,
        })),
      };
    });
  },
};
