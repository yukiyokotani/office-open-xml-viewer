#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import path from 'node:path';
import process from 'node:process';

const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript-compiler-api');
const typesDir = path.resolve(process.cwd(), 'dist/types');
const formats = ['docx', 'pptx', 'xlsx'];
const files = ['index', ...formats, 'math', 'three-d', 'region-map', 'chart-ex', 'tiff', 'legacy-conversion']
  .map((entry) => path.join(typesDir, `${entry}.d.ts`));

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

const sharedOoxmlTypes = [
  'LoadOptions',
  'OoxmlError',
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

const tiffExports = moduleExports(files.at(-2));
assert.deepEqual(
  [...tiffExports.keys()].sort(),
  ['TiffDecodeError', 'TiffRenderOptions', 'TiffRenderer', 'isTiffDecodeError', 'tiff'],
  'The ./tiff declaration entry must expose the runtime codec and its shared contract.',
);

const legacyConversionExports = moduleExports(files.at(-1));
assert.deepEqual(
  [...legacyConversionExports.keys()].sort(),
  [
    'LegacyOfficeConversionError',
    'LegacyOfficeConversionFailureReason',
    'LegacyOfficeConversionInput',
    'LegacyOfficeConversionOptions',
    'LegacyOfficeConversionRecord',
    'LegacyOfficeConversionResult',
    'LegacyOfficeConversionWorker',
    'LegacyOfficeConversionWorkerAdapterOptions',
    'LegacyOfficeConversionWorkerFactory',
    'LegacyOfficeConversionWorkerScope',
    'LegacyOfficeConverter',
    'LegacyOfficeFormat',
    'LegacyOfficeWorkerRequest',
    'LegacyOfficeWorkerResponse',
    'createDisposableWorkerLegacyOfficeConverter',
    'installLegacyOfficeConversionWorkerHandler',
    'validateConvertedOoxml',
  ].sort(),
  'The ./legacy-conversion declaration entry must expose the Worker transport and shared contract.',
);

process.stdout.write(
  'Published declaration entries compile; root namespace exports and shared OOXML contracts match.\n',
);
