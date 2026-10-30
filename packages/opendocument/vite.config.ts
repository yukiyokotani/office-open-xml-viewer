import { defineConfig } from 'vite';
import { dts } from 'rolldown-plugin-dts';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { companionBundleBoundary, wasmAssetUrl } from '../../vite.config';

const here = fileURLToPath(new URL('.', import.meta.url));
export default defineConfig({
  base: './',
  // Shared structural types live above this package. Keep tsgo's config/root
  // at repository level so every declaration stays inside its temporary tree.
  plugins: [
    companionBundleBoundary(),
    wasmAssetUrl(undefined, [companionBundleBoundary()]),
    dts({ generator: 'tsgo', tsconfig: '../../tsconfig.opendocument.json' }),
  ],
  oxc: { exclude: [/\.js$/, /\.d\.[cm]?ts$/] },
  build: {
    target: 'esnext',
    lib: {
      entry: { odt: resolve(here, 'src/odt.ts') },
      formats: ['es'],
      fileName: (_format, name) => name.endsWith('.d')
        ? `.types-work/${name.slice(0, -2)}.d.ts` : `${name}.mjs`,
    },
    rolldownOptions: {
      // Keep the runtime assertion/protocol owned by the installed peer. All
      // other runtime dependencies are stateless helpers bundled from core.
      external: ['@silurus/ooxml/internal/companion', /^node:/],
      output: {
        assetFileNames: '[name][extname]',
        chunkFileNames: chunk => chunk.name.endsWith('.d')
          ? '.types-work/[name]-[hash].d.ts' : '[name]-[hash].js',
      },
    },
  },
});
