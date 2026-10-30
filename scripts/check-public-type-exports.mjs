#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import path from 'node:path';
import process from 'node:process';

const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript-compiler-api');
const typesDir = path.resolve(process.cwd(), 'dist/types');
const formats = ['docx', 'pptx', 'xlsx'];
const files = ['index', ...formats, 'math', 'three-d', 'region-map', 'chart-ex', 'tiff', 'internal-companion', 'node']
  .map((entry) => path.join(typesDir, `${entry}.d.ts`));

const legacyTypesDir = path.resolve('packages/legacy-office/dist/types');
files.push(...['doc', 'xls', 'ppt'].map(entry => path.join(legacyTypesDir, `${entry}.d.ts`)));
const program = ts.createProgram(files, {
  module: ts.ModuleKind.ESNext,
  moduleResolution: ts.ModuleResolutionKind.Bundler,
  target: ts.ScriptTarget.ES2022,
  lib: ['lib.es2022.d.ts', 'lib.dom.d.ts', 'lib.dom.iterable.d.ts'],
  strict: true,
  noEmit: true,
});
const diagnostics = ts.getPreEmitDiagnostics(program);
if (diagnostics.length > 0) {
  throw new Error(ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: (file) => file,
    getCurrentDirectory: () => process.cwd(),
    getNewLine: () => '\n',
  }));
}

const checker = program.getTypeChecker();
const moduleExports = (file) => {
  const source = program.getSourceFile(file);
  const symbol = source && checker.getSymbolAtLocation(source);
  assert.ok(symbol, `Cannot resolve declaration module ${path.relative(process.cwd(), file)}.`);
  return new Map(checker.getExportsOfModule(symbol).map((entry) => [entry.name, entry]));
};

const rootExports = moduleExports(files[0]);
const formatExports = [];
for (const [index, format] of formats.entries()) {
  let namespace = rootExports.get(format);
  assert.ok(namespace, `Root declaration does not export the ${format} namespace.`);
  if (namespace.flags & ts.SymbolFlags.Alias) namespace = checker.getAliasedSymbol(namespace);
  const namespaceNames = checker.getExportsOfModule(namespace).map((entry) => entry.name).sort();
  const directExports = moduleExports(files[index + 1]);
  formatExports.push(directExports);
  const directNames = [...directExports.keys()].sort();
  assert.deepEqual(
    namespaceNames,
    directNames,
    `Root ${format} namespace differs from the ./${format} entry point.`,
  );
}

// The `useGoogleFonts` deprecation message tells callers to import the font
// source factories from the root entry as well as from the format entries.
for (const name of ['fontCssApi', 'fontCss', 'fontFiles', 'fontResolver', 'googleFonts', 'googleFontsSubstitutions', 'officeFontSubstitutions']) {
  let root = rootExports.get(name);
  assert.ok(root, `Root declaration does not export font source factory ${name}.`);
  if (root.flags & ts.SymbolFlags.Alias) root = checker.getAliasedSymbol(root);
  for (const [index, exports] of formatExports.entries()) {
    let direct = exports.get(name);
    assert.ok(direct, `./${formats[index]} does not export font source factory ${name}.`);
    if (direct.flags & ts.SymbolFlags.Alias) direct = checker.getAliasedSymbol(direct);
    const rootType = checker.getTypeOfSymbol(root);
    const directType = checker.getTypeOfSymbol(direct);
    assert.ok(
      checker.isTypeAssignableTo(rootType, directType) && checker.isTypeAssignableTo(directType, rootType),
      `Root ${name} differs from the ./${formats[index]} export.`,
    );
  }
}

const sharedOoxmlTypes = [
  'LoadOptions',
  'OoxmlError',
  'OoxmlUnsupportedContentDetails',
  'OoxmlErrorStage',
  'OoxmlFormat',
  'OoxmlResourceLimit',
  'OoxmlResourceLimits',
  'OoxmlResourceMetric',
  'OoxmlResourceName',
  'OoxmlResourceLimitError',
  'OoxmlResourceLimitErrorDetails',
  'OoxmlResourceUsageSnapshot',
  'OoxmlResourceViolation',
];

function declaredType(exports, name, format) {
  let symbol = exports.get(name);
  assert.ok(symbol, `${format} does not export shared OOXML type ${name}.`);
  if (symbol.flags & ts.SymbolFlags.Alias) symbol = checker.getAliasedSymbol(symbol);
  return checker.getDeclaredTypeOfSymbol(symbol);
}

for (const name of sharedOoxmlTypes) {
  const canonical = declaredType(formatExports[0], name, formats[0]);
  for (let index = 1; index < formatExports.length; index += 1) {
    const candidate = declaredType(formatExports[index], name, formats[index]);
    assert.ok(
      checker.isTypeAssignableTo(canonical, candidate)
        && checker.isTypeAssignableTo(candidate, canonical),
      `${name} differs between ${formats[0]} and ${formats[index]}.`,
    );
  }
}

const nodeExports = moduleExports(path.join(typesDir, 'node.d.ts'));
for (const name of ['OoxmlError', 'OoxmlErrorCode', 'OoxmlUnsupportedContentDetails']) {
  const nodeType = declaredType(nodeExports, name, 'Node');
  const formatType = declaredType(formatExports[0], name, 'DOCX');
  assert.ok(checker.isTypeAssignableTo(nodeType, formatType) && checker.isTypeAssignableTo(formatType, nodeType),
    `Node and format entries disagree on ${name}.`);
}

const tiffExports = moduleExports(path.join(typesDir, 'tiff.d.ts'));
assert.deepEqual(
  [...tiffExports.keys()].sort(),
  ['TiffDecodeError', 'TiffRenderOptions', 'TiffRenderer', 'isTiffDecodeError', 'tiff'],
  'The ./tiff declaration entry must expose the runtime codec and its shared contract.',
);

for (const [entry, constant, factory, options] of [
  ['doc', 'legacyDoc', 'createLegacyDoc', 'LegacyDocOptions'],
  ['ppt', 'legacyPpt', 'createLegacyPpt', 'LegacyPptOptions'],
  ['xls', 'legacyXls', 'createLegacyXls', 'LegacyXlsOptions'],
]) {
  const exports = moduleExports(path.join(legacyTypesDir, `${entry}.d.ts`));
  assert.deepEqual(
    [...exports.keys()].sort(),
    [constant, options, factory].sort(),
    `The ./${entry} declaration entry must preserve its default model source, factory and option types.`,
  );
}

assert.deepEqual([...moduleExports(path.join(typesDir, 'internal-companion.d.ts')).keys()].sort(),
  ['CompanionRuntimeMismatchError', 'MODEL_SOURCE_MODULE_PROTOCOL', 'assertCompanionRuntime'].sort());

process.stdout.write(
  'Published declaration entries compile; root namespace exports and shared OOXML contracts match.\n',
);
