import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import nodePath, { join } from 'node:path';
import {
  createTypeScriptResolver, findViolations, isGuardedPath, legacyCargoViolations,
} from './check-core-legacy-boundary.mjs';
import { companionBundleBoundary, legacyBundleBoundary, wasmAssetUrl } from '../vite.config.ts';
import { build } from 'vite';

const rules = (text) => findViolations([{ path: 'packages/docx/src/x.ts', text }]).map((v) => v.rule);

test('flags legacy imports, options, branches and names in the OOXML packages', () => {
  assert.deepEqual(rules("import { x } from '@silurus/ooxml-legacy-office/internal/direct-doc-engine';"),
    ['legacy-package', 'legacy-office-api', 'direct-format-name']);
  assert.deepEqual(rules('await DocxDocument.load(bytes, { legacyConversion: {} });'), ['legacy-office-api']);
  assert.deepEqual(rules("if (kind === 'legacy-xls') return;"), ['legacy-format-name']);
  assert.deepEqual(rules('const nativeDoc = await openNative();'), ['native-legacy-source']);
  assert.deepEqual(rules('archive.revision_markup_in_print?.()'), ['legacy-revision-view']);
  assert.deepEqual(rules('interface T { measureLegacyXlsNormalFont?: unknown; }'), ['legacy-xls-measurement']);
  assert.deepEqual(rules('interface T { langDefault?: string; }'), ['docx-lang-default']);
});

test('flags legacy-producer wire names and binary-format vocabulary', () => {
  for (const [code, rule] of [
    ['// pub native_separators: Option<Box<X>>,', 'legacy-wire-name'],
    ['interface T { readonly nativeTextFlow?: number; }', 'legacy-wire-name'],
    ["code: 'NATIVE_DOC_NESTED_CELL_FRAME_FLOW';", 'legacy-wire-name'],
    ['// [MS-DOC] 2.3.3 reserved stories', 'legacy-binary-spec'],
    ['// [MS-PPT] round-trip color map', 'legacy-binary-spec'],
    ['/// a BIFF chart has no Series record', 'legacy-binary-spec'],
    ['// effective CHPX of the paragraph mark', 'legacy-binary-structure'],
    ['// sprmSTextFlow carries an MSOTXFL', 'legacy-binary-structure'],
  ]) {
    assert.deepEqual(rules(code), [rule], code);
  }
});

test('allows OOXML specifications and the listed binary-format allowances only', () => {
  assert.deepEqual(rules('// [MS-DOCX] §2.6.2.3 and MS-ODRAWXML svgBlip'), []);
  const at = (path, text) => findViolations([{ path, text }]).map((v) => v.rule);
  assert.deepEqual(at('packages/xlsx/src/x.ts', '// [MS-XLSB] and [MS-XLS] semantics'), []);
  assert.deepEqual(at('packages/xlsx/src/x.ts', '// [MS-XLS] and BIFF8'), ['legacy-binary-spec']);
  assert.deepEqual(at('packages/core/src/errors/cfb-sniff.ts', "['WordDocument']; // .doc (§[MS-DOC])"), []);
  assert.deepEqual(at('packages/core/src/errors/cfb-sniff.ts', '// sprmCFSpec'), ['legacy-binary-structure']);
  assert.deepEqual(at('packages/docx/src/layout/diagnostics.ts', "part: 'WordDocument';"), ['legacy-binary-structure']);
  // Allowances name complete tokens and exact files: near misses still fail.
  const pageLimit = 'packages/docx/src/layout/table-acquisition.ts';
  assert.deepEqual(at(pageLimit, '// MS-DOC 2.6.4 sprmSXaPage / sprmSYaPage'), []);
  assert.deepEqual(at(pageLimit, '// sprmSTextFlow'), ['legacy-binary-structure']);
  assert.deepEqual(at(pageLimit, '// sprmSXaPageX'), ['legacy-binary-structure']);
  assert.deepEqual(at(pageLimit, '// sprmSXaPage and sprmCFSpec'), ['legacy-binary-structure']);
  assert.deepEqual(at('packages/core/src/source/cfb-new-reader.ts', "['WordDocument'];"), ['legacy-binary-structure']);
  assert.deepEqual(at('packages/core/src/errors/cfb-sniff2.ts', "['WordDocument'];"), ['legacy-binary-structure']);
  assert.deepEqual(at('packages/xlsx/src/xlsx-cfb-guard.test.ts', "['WordDocument'];"), ['legacy-binary-structure']);
  assert.deepEqual(at('packages/ooxml-common/src/chart/model.rs', '/// BIFF chart'), ['legacy-binary-spec']);
  assert.deepEqual(at('src/docx.ts', 'export { nativeReadingNotices };'), ['legacy-wire-name']);
});

// Each case is a real manifest that Cargo resolves; path dependencies point at
// this repository's crates.
function manifestLines(body) {
  const dir = mkdtempSync(nodePath.join(tmpdir(), 'legacy-cargo-'));
  const crates = nodePath.resolve('packages');
  // Root-level dotted keys must precede every table header.
  const text = [
    ...body.map((line) => line.replaceAll('../..', crates)),
    '[package]', 'name = "legacy-probe"', 'version = "0.0.0"', 'edition = "2021"', '[lib]', 'path = "lib.rs"',
    '[workspace]',
  ].join('\n');
  writeFileSync(nodePath.join(dir, 'Cargo.toml'), text);
  writeFileSync(nodePath.join(dir, 'lib.rs'), '');
  try {
    return legacyCargoViolations(nodePath.join(dir, 'Cargo.toml')).map((v) => v.text);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

test('limits the legacy reader to model crates and the listed parser crates', () => {
  assert.deepEqual(manifestLines([
    '[dependencies]',
    'pptx-model = { path = "../../pptx/model" }',
    'pptx-parser = { path = "../../pptx/parser", optional = true }',
    'docx-parser = { path = "../../docx/parser" }',
    '[dev-dependencies]',
    'xlsx-parser = { path = "../../xlsx/parser" }',
  ]), ['docx-parser resolves to docx-parser']);
});

test('resolves renamed, quoted, dotted, table and target-specific parser dependencies', () => {
  assert.deepEqual(manifestLines([
    '[dependencies]',
    'reader = { "package" = "docx-parser", path = "../../docx/parser" }',
  ]), ['reader resolves to docx-parser']);
  assert.deepEqual(manifestLines([
    'dependencies.reader = { package = "docx-parser", path = "../../docx/parser" }',
  ]), ['reader resolves to docx-parser']);
  assert.deepEqual(manifestLines([
    '[dependencies."docx\\u002dparser"]',
    'path = "../../docx/parser"',
  ]), ['docx-parser resolves to docx-parser']);
  assert.deepEqual(manifestLines([
    "[target.'cfg(unix)'.dependencies.grid]",
    'package = "xlsx-parser"',
    'path = "../../xlsx/parser"',
    "[target.'cfg(unix)'.dev-dependencies]",
    'pptx-model = { path = "../../pptx/model" }',
  ]), ['grid resolves to xlsx-parser']);
  assert.deepEqual(manifestLines([
    '[build-dependencies]',
    'docx-parser = { path = "../../docx/parser" }',
  ]), ['docx-parser resolves to docx-parser']);
});

test('a manifest Cargo cannot read fails closed', () => {
  const [text] = manifestLines(['[dependencies]', '{ odd } = 1']);
  assert.match(text, /^cargo metadata failed/);
});

test('citation allowances admit the token only inside comments', () => {
  const at = (path, text) => findViolations([{ path, text }]).map((v) => v.rule);
  const policy = 'packages/xlsx/src/internal/source-policy.ts';
  assert.deepEqual(at(policy, '// Excel stores this per [MS-XLS] 2.4.1.'), []);
  assert.deepEqual(at(policy, '/**\n * the [MS-XLSB] record keeps the same field.\n */'), []);
  assert.deepEqual(at(policy, "export interface SourceFacts { sourceFormat?: 'MS-XLS'; }"), ['legacy-binary-spec']);
  assert.deepEqual(at(policy, "export const binary = (f: string) => f === 'MS-XLS'; // [MS-XLS]"), ['legacy-binary-spec']);
  const table = 'packages/docx/src/layout/table-acquisition.ts';
  assert.deepEqual(at(table, '/**\n * (MS-DOC 2.6.4 sprmSXaPage / sprmSYaPage)\n */'), []);
  assert.deepEqual(at(table, "const limit = source === 'MS-DOC' ? sprmSXaPage : 0;"), ['legacy-binary-spec', 'legacy-binary-structure']);
  assert.deepEqual(at(policy, "export const binary = (f: string) =>\n  /* source kind */ f === 'MS-XLS';"), ['legacy-binary-spec']);
  assert.deepEqual(at(table, "const limit = /* page */ source === 'MS-DOC' ? 1 : 0;"), ['legacy-binary-spec']);
  assert.deepEqual(at(policy, "const url = 'https://example.test/a'; // [MS-XLS] 2.4"), []);
  assert.deepEqual(at(policy, "export const b = (f: string) =>\n  f.replace(/\\//g, '-') === 'MS-XLS';"), ['legacy-binary-spec']);
  assert.deepEqual(at(table, "const p = /a\\/b/.test(s) && s === 'MS-DOC';"), ['legacy-binary-spec']);
  assert.deepEqual(at(policy, "/*\n  multi-line [MS-XLSB] note\n*/ const kind = 'MS-XLSB';"), ['legacy-binary-spec']);
  const rules = 'packages/docx/src/layout/line-compatibility.ts';
  assert.deepEqual(at(rules, "const rule = {\n  reference: '[MS-DOC] §2.9.146',\n};"), []);
  assert.deepEqual(at(rules, "const binary = format === 'MS-DOC';"), ['legacy-binary-spec']);
});

test('allows the pre-existing generic CFB rejection and unrelated "legacy" wording', () => {
  assert.deepEqual(rules("throw new OoxmlError('legacy-binary-format', 'a legacy binary .doc file');"), []);
  assert.deepEqual(rules('// legacy VML shapes and the legacy chart family fallback'), []);
  assert.deepEqual(rules('export function selectModelSource(sources, target, bytes) {}'), []);
});

test('admits only each surface\'s own opt-in source option key', () => {
  const at = (path, text) => findViolations([{ path, text }]).map((v) => v.rule);
  assert.deepEqual(at('packages/docx/src/x.ts', "interface O { legacyDoc?: ModelSource<'docx'>; }"), []);
  assert.deepEqual(at('packages/xlsx/src/x.ts', 'if (opts.legacyXls !== undefined) {}'), []);
  assert.deepEqual(at('packages/pptx/src/x.ts', 'load(bytes, { legacyPpt });'), []);
  assert.deepEqual(at('packages/node/src/x.ts', 'const { legacyDoc, legacyXls, legacyPpt } = options;'), []);
  // Another format's key, another spelling, or core is still reported.
  assert.deepEqual(at('packages/docx/src/x.ts', 'load(bytes, { legacyXls });'), ['legacy-format-name']);
  assert.deepEqual(at('packages/xlsx/src/x.ts', "const kind = 'legacy-xls';"), ['legacy-format-name']);
  assert.deepEqual(at('packages/pptx/src/x.ts', 'const LegacyPpt = 1;'), ['legacy-format-name']);
  assert.deepEqual(at('packages/core/src/x.ts', 'interface O { legacyDoc?: unknown; }'), ['legacy-format-name']);
});

test('admits only each surface\'s own OpenDocument option key and no ODF vocabulary', () => {
  const at = (path, text) => findViolations([{ path, text }]).map((v) => v.rule);
  assert.deepEqual(at('packages/docx/src/x.ts', "interface O { odt?: ModelSource<'docx'>; }"), []);
  assert.deepEqual(at('packages/xlsx/src/x.ts', 'if (opts.ods !== undefined) {}'), []);
  assert.deepEqual(at('packages/pptx/src/x.ts', 'load(bytes, { odp });'), []);
  assert.deepEqual(at('packages/node/src/x.ts', 'const { odt, ods, odp } = options;'), []);
  assert.deepEqual(at('packages/docx/src/x.ts', 'load(bytes, { ods });'), ['odf-format-name']);
  assert.deepEqual(at('packages/core/src/x.ts', 'interface O { odt?: unknown; }'), ['odf-format-name']);
  // The surface's own key spelling is admitted; the companion import is not.
  assert.deepEqual(at('packages/xlsx/src/x.ts', "import { ods } from '@silurus/ooxml-opendocument/ods';"),
    ['odf-package']);
  assert.deepEqual(at('packages/ooxml-common/src/x.rs', 'const NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";'),
    ['odf-vocabulary']);
});

test('decodes escaped static, dynamic, require and re-export module specifiers', () => {
  for (const code of [
    "import x from '@silurus/ooxml-\\u006cegacy-office';",
    "export * from '@silurus/ooxml-\\u006cegacy-office';",
    "await import('@silurus/ooxml-\\u006cegacy-office');",
    "require('@silurus/ooxml-\\u006cegacy-office');",
    "import x from '../../legacy-\\u006fffice/src/index.js';",
  ]) {
    assert.ok(rules(code).includes('legacy-package'), code);
  }
});

test('rejects computed imports and requires, import attributes and type-level imports', () => {
  for (const code of [
    "import('@silurus/ooxml-' + 'legacy' + '-office');",
    "import(`@silurus/ooxml-${'legacy'}-office`);",
    "const p = '@silurus/ooxml-\\u006cegacy-office'; require(p);",
  ]) {
    assert.ok(rules(code).includes('computed-module'), code);
  }
  assert.deepEqual(rules("import('@silurus/ooxml-\\u006cegacy-office/package.json', { with: { type: 'json' } });"), ['legacy-package']);
  assert.deepEqual(rules("type X = import('@silurus/ooxml-\\u006cegacy-office').X;"), ['legacy-package']);
  assert.ok(rules('require(variable);').includes('computed-module'));
  assert.ok(rules('require(variable);').includes('loader-identifier'));
  assert.deepEqual(rules('import(variable);'), ['computed-module']);
  assert.deepEqual(
    findViolations([{ path: 'packages/docx/src/new-probe.test.ts', text: 'import(variable);' }]).map((v) => v.rule),
    ['computed-module'],
  );
});

test('rejects indirect Node require forms even with escaped module names', () => {
  for (const code of [
    "import { createRequire } from 'node:module'; const load = createRequire(import.meta.url); load(name);",
    "import { createRequire as cr } from 'node:module'; cr(import.meta.url)(name);",
    'require.call(null, name);',
    'require.apply(null, [name]);',
    'require.bind(null)(name);',
    'require?.(name);',
    'module.require(name);',
    "module['require'](name);",
    "module?.['require'](name);",
    'module?.require?.(name);',
  ]) {
    assert.ok(rules(code).some((rule) => rule === 'node-module-import' || rule === 'loader-identifier'), code);
  }
});

test('rejects every Node module import form and loader identifier reference', () => {
  for (const code of [
    "import { 'createRequire' as load } from 'node:module'; load(import.meta.url)(name);",
    "import Module from 'node:module'; const { 'createRequire': load } = Module; load(import.meta.url)(name);",
    "import Module from 'node:module'; Module['create' + 'Require'](import.meta.url)(name);",
    "import * as loader from 'module';",
    "import('node:module');",
    "const { module } = value;",
    "module['require'](name);",
    'require?.(name);',
  ]) {
    assert.ok(rules(code).some((rule) => rule === 'node-module-import' || rule === 'loader-identifier'), code);
  }
});

test('rejects alternate runtime loaders and code generation capabilities', () => {
  for (const code of [
    "process.getBuiltinModule('module').createRequire(import.meta.url)(name);",
    "globalThis['process'].getBuiltinModule('module').createRequire(import.meta.url)(name);",
    "const load = process['getBuiltinModule']('module');",
    "process.binding('natives');",
    "globalThis.process.binding('natives');",
    "eval('process');",
    "(0, eval)('process');",
    "globalThis.eval('process');",
    "new Function('return process')();",
    "Function('return process')();",
    "globalThis.Function('return process')();",
    "const make = Function; make('return process')();",
    "const make = globalThis['Function']; make('return process')();",
    "({}).constructor.constructor('return process')();",
    "importScripts(path);",
    "const load = importScripts; load(path);",
    "const load = self['importScripts']; load(path);",
    "globalThis[key];",
    "self[key];",
    "window[key];",
    "globalThis['eval']('process');",
    "globalThis['process']['binding']('natives');",
  ]) {
    assert.ok(rules(code).length > 0, code);
  }
  for (const code of [
    "importScripts('./worker.js');",
    "globalThis['document'];",
    "self[`location`];",
    "window[0];",
  ]) {
    assert.deepEqual(rules(code), [], code);
  }
  assert.deepEqual(findViolations([{
    path: 'packages/core/src/worker/bridge.test.ts', text: 'expect.any(Function);',
  }]), []);
});

test('follows a relay outside guarded roots to a dynamic legacy import', () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-boundary-relay-'));
  try {
    mkdirSync(join(root, 'packages/core/src'), { recursive: true });
    mkdirSync(join(root, 'packages/legacy-office/src'), { recursive: true });
    mkdirSync(join(root, 'shared'), { recursive: true });
    writeFileSync(join(root, 'packages/core/tsconfig.json'), JSON.stringify({
      compilerOptions: { moduleResolution: 'bundler', module: 'esnext' },
    }));
    writeFileSync(join(root, 'packages/legacy-office/src/index.ts'), 'export const reader = 1;');
    writeFileSync(join(root, 'shared/reader.ts'),
      "export const reader = () => import('../packages/legacy-office/src/index.ts');");
    const violations = findViolations([{
      path: 'packages/core/src/probe.ts',
      text: "export { reader } from '../../../shared/reader.ts';",
    }], { resolveModule: createTypeScriptResolver(root) });
    assert.ok(violations.some((violation) => violation.rule === 'legacy-package'
      && violation.path === 'shared/reader.ts'));
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('TypeScript module resolution catches a path alias into the legacy package', () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-boundary-'));
  try {
    mkdirSync(join(root, 'packages/core/src'), { recursive: true });
    mkdirSync(join(root, 'packages/legacy-office/src'), { recursive: true });
    writeFileSync(join(root, 'packages/legacy-office/src/index.ts'), 'export const reader = 1;');
    writeFileSync(join(root, 'packages/core/tsconfig.json'), JSON.stringify({
      compilerOptions: { baseUrl: '../..', paths: { '@reader': ['packages/legacy-office/src/index.ts'] }, moduleResolution: 'bundler', module: 'esnext' },
    }));
    const violations = findViolations(
      [{ path: 'packages/core/src/probe.ts', text: "import { reader } from '@reader';" }],
      { resolveModule: createTypeScriptResolver(root) },
    );
    assert.deepEqual(violations.map((v) => v.rule), ['legacy-package']);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('built chunk graph rejects a resolved legacy module even after aliasing', () => {
  const hook = legacyBundleBoundary().generateBundle;
  assert.equal(typeof hook, 'function');
  const invoke = (moduleId, fileName = 'entry.js') => hook.call(
    { error(message) { throw new Error(message); } },
    {},
    { [fileName]: { type: 'chunk', fileName, isEntry: true, imports: [], modules: { [moduleId]: {} } } },
  );
  assert.doesNotThrow(() => invoke('/repo/packages/core/src/index.ts'));
  assert.throws(
    () => invoke('/repo/packages/legacy-office/src/index.ts'),
    /forbidden legacy module/,
  );
  assert.throws(() => invoke('/repo/packages/legacy-office/src/index.ts', 'legacy-ppt.mjs'), /forbidden legacy module/);
  assert.throws(() => hook.call(
    { error(message) { throw new Error(message); } },
    {},
    {
      'docx.mjs': { type: 'chunk', fileName: 'docx.mjs', isEntry: true, imports: ['shared.js'], modules: {} },
      'shared.js': { type: 'chunk', fileName: 'shared.js', isEntry: false, imports: [], modules: { '/repo/packages/legacy-office/src/index.ts': {} } },
    },
  ), /forbidden legacy module/);
  assert.throws(() => hook.call(
    { error(message) { throw new Error(message); } },
    {},
    {
      'docx.mjs': { type: 'chunk', fileName: 'docx.mjs', isEntry: true, imports: [], dynamicImports: ['relay.js'], modules: {} },
      'relay.js': { type: 'chunk', fileName: 'relay.js', isEntry: false, imports: [], dynamicImports: [], modules: { '/repo/packages/legacy-office/src/index.ts': {} } },
    },
  ), /forbidden legacy module/);
});

test('a real build rejects a dynamic chunk reached through a relay', async () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-boundary-bundle-'));
  try {
    mkdirSync(join(root, 'packages/core/src'), { recursive: true });
    mkdirSync(join(root, 'packages/legacy-office/src'), { recursive: true });
    mkdirSync(join(root, 'shared'), { recursive: true });
    const entry = join(root, 'packages/core/src/probe.ts');
    writeFileSync(entry, "export { reader } from '../../../shared/reader.ts';");
    writeFileSync(join(root, 'shared/reader.ts'),
      "export const reader = () => import('../packages/legacy-office/src/index.ts');");
    writeFileSync(join(root, 'packages/legacy-office/src/index.ts'), 'export const value = 1;');
    await assert.rejects(() => build({
      configFile: false,
      logLevel: 'silent',
      plugins: [legacyBundleBoundary()],
      build: { write: false, lib: { entry, formats: ['es'] } },
    }), /forbidden legacy module/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('guards source, parser and manifest files but not generated or private output', () => {
  assert.equal(isGuardedPath('packages/core/src/index.ts'), true);
  assert.equal(isGuardedPath('packages/docx/parser/src/parser.rs'), true);
  assert.equal(isGuardedPath('packages/node/package.json'), true);
  assert.equal(isGuardedPath('packages/ooxml-common/src/chart/model.rs'), true);
  assert.equal(isGuardedPath('src/docx.ts'), true);
  assert.equal(isGuardedPath('packages/xlsx/src/wasm/xlsx_parser.js'), false);
  assert.equal(isGuardedPath('packages/pptx/public/demo/readme.md'), false);
  assert.equal(isGuardedPath('packages/legacy-office/src/legacy-doc.ts'), false);
});

// Both the factory graph and the opaque URL-loaded reader assets must keep
// renderer/transport ownership in the installed peer after upstream refactors.
test('companion builds reject viewer runtimes in entries and URL-loaded source modules', async () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-companion-boundary-'));
  try {
    const entry = join(root, 'entry.ts');
    const reader = join(root, 'reader.ts');
    for (const path of [
      'packages/docx/src/internal/document-model-source.ts',
      'packages/xlsx/src/internal/worker-worksheet-source.ts',
      'packages/pptx/src/render-worker.ts',
      'packages/node/src/docx-model-source.ts',
      'packages/core/src/worker/bridge.ts',
      'packages/core/src/source/model-source.ts',
    ]) {
      const owner = join(root, path);
      mkdirSync(join(owner, '..'), { recursive: true });
      writeFileSync(owner, 'export const runtime = 1;');
      writeFileSync(reader, path.includes('document-model-source')
        ? `export const runtime = () => import('./${path}');`
        : `export { runtime } from './${path}';`);
      for (const asset of [false, true]) {
        writeFileSync(entry, asset
          ? "export { default } from './reader.ts?url';"
          : "export { runtime } from './reader.ts';");
        await assert.rejects(() => build({
          configFile: false, logLevel: 'silent',
          plugins: [companionBundleBoundary(), wasmAssetUrl(undefined, [companionBundleBoundary()])],
          build: { write: false, lib: { entry, formats: ['es'] } },
        }), /viewer-owned runtime/);
      }
    }
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
