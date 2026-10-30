#!/usr/bin/env node
// A consumer that imports only the font source factories (the import the
// `useGoogleFonts` deprecation message recommends) must not pay for a viewer.
// Each published browser entry is re-bundled the way an application bundler
// would (Rolldown, the bundler behind Vite), and the result must be one small
// chunk: no parser, renderer or worker code, and no lazily imported chunk.
//
// This guards the library build's chunk graph. Rolldown keeps a chunk that
// two or more of its dynamically imported chunks import back, even when the
// consumer uses nothing from it, so a cycle between a format chunk and its lazy
// chunks makes every import of that chunk retain the whole format.
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { rolldown } from 'rolldown';

const distDir = resolve(process.argv[2] ?? 'dist');
const FACTORIES = ['fontCssApi', 'fontCss', 'fontFiles', 'fontResolver', 'googleFonts', 'googleFontsSubstitutions', 'officeFontSubstitutions'];
// The factories, their validation and the Google preset data minify to about
// 12 KiB. The budget leaves room for growth, not for any viewer code.
const BUDGET_BYTES = 32 * 1024;
const VIEWER_MARKERS = ['_parser_bg.wasm', 'OffscreenCanvas', 'getContext('];

const work = mkdtempSync(join(tmpdir(), 'ooxml-font-shake-'));
try {
  for (const entry of ['index.mjs', 'docx.mjs', 'xlsx.mjs', 'pptx.mjs']) {
    const consumer = join(work, `consumer-${entry}`);
    writeFileSync(consumer, [
      `import { ${FACTORIES.join(', ')} } from ${JSON.stringify(join(distDir, entry))};`,
      `console.log(googleFonts(), fontCssApi, fontCss, fontFiles, fontResolver, googleFontsSubstitutions, officeFontSubstitutions);`,
      '',
    ].join('\n'));
    const bundle = await rolldown({ input: consumer, logLevel: 'silent' });
    try {
      const { output } = await bundle.generate({ format: 'es', minify: true });
      const chunks = output.filter((item) => item.type === 'chunk');
      if (chunks.length !== 1) {
        throw new Error(`${entry}: a factories-only import emitted ${chunks.length} chunks (${chunks.map((chunk) => chunk.fileName).join(', ')}); expected 1`);
      }
      const bytes = Buffer.byteLength(chunks[0].code);
      if (bytes > BUDGET_BYTES) {
        throw new Error(`${entry}: a factories-only import is ${bytes} bytes, over the ${BUDGET_BYTES}-byte budget`);
      }
      const marker = VIEWER_MARKERS.find((text) => chunks[0].code.includes(text));
      if (marker) throw new Error(`${entry}: a factories-only import contains viewer code (${marker})`);
      console.log(`${entry}: factories-only import ${bytes}/${BUDGET_BYTES} bytes, 1 chunk`);
    } finally {
      await bundle.close();
    }
  }
} finally {
  rmSync(work, { recursive: true, force: true });
}
