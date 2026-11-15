import assert from 'node:assert/strict';
import { test } from 'node:test';
import { staticLazyFontSourceModules } from './check-font-source-tree-shaking.mjs';

const chunk = (imports, ...modules) => ({ imports, modules: Object.fromEntries(modules.map((id) => [id, ''])) });

test('a lazy font-source module retained by a static chunk is reported', () => {
  const stats = {
    chunks: {
      'docx.mjs': chunk(['document.js'], 'src/docx.ts'),
      'document.js': chunk(['fonts.js'], 'packages/docx/src/document.ts'),
      // Variable-font instancing pulled into a statically imported chunk.
      'fonts.js': chunk([], 'packages/core/src/fonts/open-type-metrics.ts', 'packages/core/src/fonts/open-type-instances.ts'),
      // Reached only through import(): not part of the static graph.
      'session.js': chunk(['fonts.js'], 'packages/core/src/fonts/source-session.ts'),
    },
  };
  assert.deepEqual(staticLazyFontSourceModules(stats, ['docx.mjs']), [
    ['docx.mjs', 'packages/core/src/fonts/open-type-instances.ts', 'fonts.js'],
  ]);
});

test('lazy font-source modules reached only through import() pass', () => {
  const stats = {
    chunks: {
      'xlsx.mjs': chunk(['workbook.js'], 'src/xlsx.ts'),
      'workbook.js': chunk([], 'packages/xlsx/src/workbook.ts'),
      'font-source-load.js': chunk(['workbook.js'], 'packages/xlsx/src/font-source-load.ts'),
    },
  };
  assert.deepEqual(staticLazyFontSourceModules(stats, ['xlsx.mjs']), []);
  assert.throws(() => staticLazyFontSourceModules(stats, ['pptx.mjs']), /lack the entry pptx\.mjs/);
});

test('a statically retained resolver bridge host is reported; its gate is not', () => {
  const stats = {
    chunks: {
      'pptx.mjs': chunk(['presentation.js'], 'src/pptx.ts'),
      'presentation.js': chunk([], 'packages/pptx/src/presentation.ts', 'packages/core/src/fonts/resolver-bridge-gate.ts',
        'packages/core/src/fonts/resolver-bridge-host.ts', 'packages/core/src/fonts/source-cache.ts'),
    },
  };
  assert.deepEqual(staticLazyFontSourceModules(stats, ['pptx.mjs']), [
    ['pptx.mjs', 'packages/core/src/fonts/resolver-bridge-host.ts', 'presentation.js'],
    ['pptx.mjs', 'packages/core/src/fonts/source-cache.ts', 'presentation.js'],
  ]);
});
