import { defineConfig } from 'vite';
import wasm from 'vite-plugin-wasm';
import dts from 'vite-plugin-dts';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  plugins: [
    wasm(),
    dts({
      include: [
        'src/**/*',
        'packages/core/src/**/*',
        'packages/pptx/src/**/*',
        'packages/xlsx/src/**/*',
        'packages/docx/src/**/*',
      ],
      outDir: 'dist/types',
      tsconfigPath: './tsconfig.lib.json',
      rollupTypes: true,
      skipDiagnostics: true,
    }),
  ],
  build: {
    lib: {
      entry: {
        index: resolve(__dirname, 'src/index.ts'),
        pptx:  resolve(__dirname, 'src/pptx.ts'),
        xlsx:  resolve(__dirname, 'src/xlsx.ts'),
        docx:  resolve(__dirname, 'src/docx.ts'),
      },
      formats: ['es', 'cjs'],
      fileName: (format, name) => `${name}.${format === 'es' ? 'mjs' : 'cjs'}`,
    },
    rollupOptions: {
      output: { assetFileNames: '[name][extname]' },
    },
    target: 'esnext',
  },
  worker: {
    format: 'es',
    plugins: () => [wasm()],
  },
});
