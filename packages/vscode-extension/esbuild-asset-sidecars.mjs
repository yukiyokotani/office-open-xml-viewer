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
    build.onLoad({ filter: /\/packages\/(?:docx|xlsx|pptx)\/dist\/[^/]+\.mjs$/ }, async ({ path }) => {
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
