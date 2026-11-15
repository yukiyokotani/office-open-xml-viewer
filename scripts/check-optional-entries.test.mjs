import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import test from 'node:test';
import { checkLegacySource, checkOptionalRenderer, OPTIONAL_RENDERERS } from './check-optional-entries.mjs';

function fixture(t, files) {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-optional-entries-'));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  for (const [path, contents] of Object.entries(files)) {
    mkdirSync(dirname(join(root, path)), { recursive: true });
    writeFileSync(join(root, path), contents);
  }
  return root;
}

const defaultEntries = {
  'dist/index.mjs': "export * from './shared.js';\n",
  'dist/docx.mjs': "import './shared.js';\n",
  'dist/xlsx.mjs': "import './shared.js';\n",
  'dist/pptx.mjs': "import './shared.js';\n",
  'dist/shared.js': 'export const shared = 1;\n',
};

function rendererDist(name, overrides = {}) {
  const { exportName, marker } = OPTIONAL_RENDERERS[name];
  return {
    ...defaultEntries,
    [`dist/${name}.mjs`]: `export { ${exportName} } from './impl.js';\n`,
    'dist/impl.js': `// optional-renderer: ${marker}\nexport const ${exportName} = { render() {} };\n`,
    'packages/core/src/chart/chart-ex-renderer.ts': "import { layout } from './chart-ex-hierarchy.js';\n",
    'packages/core/src/chart/renderer.ts': 'export function renderChart() {}\n',
    ...overrides,
  };
}

const checkRenderer = (name, root) => checkOptionalRenderer(name, { distDir: join(root, 'dist'), sourceRoot: root });

for (const name of Object.keys(OPTIONAL_RENDERERS)) {
  const { exportName, marker } = OPTIONAL_RENDERERS[name];

  test(`${name}: an isolated optional renderer passes`, async (t) => {
    assert.match(await checkRenderer(name, fixture(t, rendererDist(name))), /bundle boundary verified$/);
  });

  test(`${name}: a default entry that statically reaches the implementation fails`, async (t) => {
    for (const entry of ['index', 'docx', 'xlsx', 'pptx']) {
      const root = fixture(t, rendererDist(name, {
        [`dist/${entry}.mjs`]: "import './shared.js';\nimport './leak.js';\n",
        'dist/leak.js': "import{x}from'./impl.js';\n",
      }));
      await assert.rejects(checkRenderer(name, root), new RegExp(`${entry}\\.mjs unexpectedly reaches .* via impl\\.js`));
    }
  });

  test(`${name}: a leak through a chunk whose name contains an apostrophe fails`, async (t) => {
    const root = fixture(t, rendererDist(name, {
      'dist/index.mjs': 'import "./it\'s.js";\n',
      "dist/it's.js": `// optional-renderer: ${marker}\n`,
    }));
    await assert.rejects(checkRenderer(name, root), /index\.mjs unexpectedly reaches/);
    // An apostrophe earlier on the line must not hide the edge that follows it.
    const hidden = fixture(t, rendererDist(name, {
      'dist/index.mjs': "// import './noise'\nimport './leak.js';\n",
      'dist/leak.js': `// optional-renderer: ${marker}\n`,
    }));
    await assert.rejects(checkRenderer(name, hidden));
  });

  test(`${name}: an entry without a usable renderer or without its implementation fails`, async (t) => {
    const unusable = fixture(t, rendererDist(name, {
      'dist/impl.js': `// optional-renderer: ${marker}\nexport const ${exportName} = {};\n`,
    }));
    await assert.rejects(checkRenderer(name, unusable), /does not export a usable/);
    const stub = fixture(t, rendererDist(name, {
      'dist/impl.js': `export const ${exportName} = { render() {} };\n`,
    }));
    await assert.rejects(checkRenderer(name, stub), /does not reach the/);
  });
}

test('a default entry that imports a missing chunk fails', async (t) => {
  for (const source of ["import './missing.js';\n", "// './x'\nimport './missing.js';\n", 'import "./a\'b.js";\nimport "./missing.js";\n']) {
    const root = fixture(t, rendererDist('tiff', { 'dist/index.mjs': source }));
    await assert.rejects(checkRenderer('tiff', root), { code: 'ENOENT' }, source);
  }
});

test('an optional entry with an edge that names no file fails', async (t) => {
  const { exportName, marker } = OPTIONAL_RENDERERS['three-d'];
  const root = fixture(t, rendererDist('three-d', {
    'dist/three-d.mjs': `export const ${exportName} = { render() {} };\n// ${marker}\n// import './missing'; './phantom.mjs'\n`,
  }));
  await assert.rejects(checkRenderer('three-d', root), { code: 'ENOENT' });
});

test('an implementation reached under only one spelling of the edge pattern fails', async (t) => {
  const { exportName, marker } = OPTIONAL_RENDERERS['three-d'];
  const root = fixture(t, rendererDist('three-d', {
    'dist/three-d.mjs': `// import './noise'\nexport { ${exportName} } from './impl.js';\n`,
    [`dist/noise'\nexport { ${exportName} } from './impl.js`]: 'export {};\n',
  }));
  await assert.rejects(checkRenderer('three-d', root), /does not reach the/);
  assert.ok(marker);
});

test('chart-ex: the hierarchy module must stay owned by the optional renderer source', async (t) => {
  const unowned = fixture(t, rendererDist('chart-ex', {
    'packages/core/src/chart/chart-ex-renderer.ts': 'export const chartEx = {};\n',
  }));
  await assert.rejects(checkRenderer('chart-ex', unowned), /does not own the hierarchy module/);
  for (const symbol of OPTIONAL_RENDERERS['chart-ex'].source.forbiddenInDefaultRenderer) {
    const regained = fixture(t, rendererDist('chart-ex', {
      'packages/core/src/chart/renderer.ts': `${symbol} {}\n`,
    }));
    await assert.rejects(checkRenderer('chart-ex', regained), /default renderer unexpectedly owns/);
  }
});

const LEGACY = {
  doc: ['docx', 'legacyDoc', 'createLegacyDoc', 'LegacyDocDocument'],
  xls: ['xlsx', 'legacyXls', 'createLegacyXls', 'LegacyXlsWorkbook'],
  ppt: ['pptx', 'legacyPpt', 'createLegacyPpt', 'LegacyPptPresentation'],
};

function legacyEntry(family) {
  const [target, constant, factory] = LEGACY[family];
  return `
const defaults = {
  moduleUrl: new URL('./assets/legacy-${family}-source-module-abc.mjs', import.meta.url).href,
  wasmUrl: new URL('./legacy_${family}_direct_bg.wasm', import.meta.url).href,
};
export function ${factory}(options = {}) {
  const moduleUrl = options.moduleUrl ?? defaults.moduleUrl;
  const wasmUrl = options.wasmUrl ?? defaults.wasmUrl;
  return Object.freeze({
    target: '${target}',
    claim: () => false,
    beginLoad: () => ({ module: { protocol: 'ooxml-model-source-module/v2', target: '${target}', moduleUrl, config: { wasmUrl } } }),
  });
}
export const ${constant} = ${factory}();
`;
}

function legacyDist(overrides = {}) {
  const files = {
    'dist/index.mjs': "import './worker-host.js';\n",
    'dist/docx.mjs': '', 'dist/xlsx.mjs': '', 'dist/pptx.mjs': '', 'dist/node.mjs': '', 'dist/internal-companion.mjs': '',
    'dist/worker-host.js': 'export const host = 1;\n',
    'dist/assets/render-worker-abc.js': 'self.onmessage = () => {};\n',
  };
  for (const family of Object.keys(LEGACY)) {
    files[`legacy/${family}.mjs`] = legacyEntry(family);
    files[`legacy/legacy_${family}_direct_bg.wasm`] = 'wasm';
    files[`legacy/assets/legacy-${family}-source-module-abc.mjs`] =
      `class ${LEGACY[family][3]} {}\nexport function openModelSource() { return new ${LEGACY[family][3]}(); }\n`;
  }
  return { ...files, ...overrides };
}

const checkLegacy = (family, root) => checkLegacySource(family, { distDir: join(root, 'legacy'), rootDistDir: join(root, 'dist') });

test('legacy sources: isolated lazy entries pass', async (t) => {
  const root = fixture(t, legacyDist());
  for (const family of Object.keys(LEGACY)) {
    assert.match(await checkLegacy(family, root), /model source boundary verified/);
  }
});

test('legacy sources: every leak of a reader is rejected', async (t) => {
  const nativeClass = LEGACY.doc[3];
  const cases = [
    ['entry reaches the reader runtime through import()', {
      'legacy/doc.mjs': `${legacyEntry('doc')}export const load = () => import("./runtime.js");\n`,
      'legacy/runtime.js': '// packages/core/src/source/model-source-runtime.ts\n',
    }, /doc\.mjs unexpectedly reaches packages\/core\/src\/source\/model-source-runtime\.ts/],
    ['entry reaches the WASM glue', {
      'legacy/doc.mjs': `${legacyEntry('doc')}import '../legacy/glue.js';\n`,
      'legacy/glue.js': `export class ${nativeClass} {}\n`,
    }, /doc\.mjs unexpectedly reaches LegacyDocDocument/],
    ['entry does not reference its WASM asset', {
      'legacy/doc.mjs': legacyEntry('doc').replace('legacy_doc_direct_bg.wasm', 'other.wasm'),
    }, /does not reference legacy_doc_direct_bg\.wasm/],
    ['source module imports a sibling', {
      'legacy/assets/legacy-doc-source-module-abc.mjs':
        `import './chunk.js';\nclass ${nativeClass} {}\nexport function openModelSource() {}\n`,
    }, /is not self-contained: imports \.\/chunk\.js/],
    ['source module lacks openModelSource', {
      'legacy/assets/legacy-doc-source-module-abc.mjs': `export class ${nativeClass} {}\n`,
    }, /does not export openModelSource/],
    ['two source module assets', {
      'legacy/assets/legacy-doc-source-module-def.mjs': 'export function openModelSource() {}\n',
    }, /expected one legacy-doc source module asset, found 2/],
    ['empty WASM asset', { 'legacy/legacy_doc_direct_bg.wasm': '' }, /is missing or empty/],
    ['an OOXML entry reaches the reader', {
      'dist/worker-host.js': `export class ${nativeClass} {}\n`,
    }, /index\.mjs unexpectedly reaches LegacyDocDocument/],
    ['another legacy entry reaches the reader asset', {
      'legacy/xls.mjs': `${legacyEntry('xls')}// legacy_doc_direct_bg.wasm\n`,
    }, /xls\.mjs unexpectedly reaches legacy_doc_direct_bg\.wasm/],
    ['a render worker contains legacy code', {
      'dist/assets/render-worker-abc.js': '// legacy-office\n',
    }, /render-worker-abc\.js contains legacy-office/],
  ];
  for (const [name, overrides, expected] of cases) {
    await assert.rejects(checkLegacy('doc', fixture(t, legacyDist(overrides))), expected, name);
  }
});
