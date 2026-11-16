import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';

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

/** Manifest file name inside the webview output directory (`dist`). */
export const PARSER_WASM_MANIFEST = 'parser-wasm.json';

const PARSER_WASM_OUTPUT = /(?:^|\/)(assets\/(docx|xlsx|pptx)_parser_bg-[A-Za-z0-9]+\.wasm)$/;

/**
 * Record which emitted file holds each format's parser WASM. The extension
 * host reads those bytes from its own install and posts them to the webview
 * (#2212): a webview fetch of an extension resource goes through VS Code's
 * resource service worker, which can time out over Remote SSH. Asset names
 * carry a content hash and `dist/assets` may keep files from earlier builds,
 * so the host must not guess the name from a directory listing.
 */
export const parserWasmManifest = {
  name: 'parser-wasm-manifest',
  setup(build) {
    build.initialOptions.metafile = true;
    build.onEnd(async (result) => {
      if (result.errors.length > 0 || !result.metafile) return;
      const manifest = {};
      for (const output of Object.keys(result.metafile.outputs)) {
        const match = PARSER_WASM_OUTPUT.exec(output.replaceAll('\\', '/'));
        if (!match) continue;
        if (manifest[match[2]]) {
          throw new Error(`[parser-wasm-manifest] two ${match[2]} parser WASM outputs`);
        }
        manifest[match[2]] = match[1];
      }
      const missing = ['docx', 'xlsx', 'pptx'].filter((format) => !manifest[format]);
      if (missing.length > 0) {
        throw new Error(`[parser-wasm-manifest] no parser WASM output for ${missing.join(', ')}`);
      }
      const outdir = build.initialOptions.outdir ?? dirname(build.initialOptions.outfile);
      await writeFile(join(outdir, PARSER_WASM_MANIFEST), `${JSON.stringify(manifest, null, 2)}\n`);
    });
  },
};
