#!/usr/bin/env node

import {
  existsSync,
  readFileSync,
  readdirSync,
  statSync,
} from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { dirname, join, relative, resolve, sep } from 'node:path';

// TypeScript 7 intentionally exposes no stable Compiler API yet. This
// architecture guard is root-only tooling, so it uses the explicitly named
// TypeScript 6 compatibility dependency while library builds stay on TS 7.
const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript-compiler-api');

const DOCX_SOURCE = 'packages/docx/src';
const COMPATIBILITY_OWNER = `${DOCX_SOURCE}/layout/compatibility.ts`;
const OBSERVATION_BASELINE =
  'scripts/docx-compatibility-observation-baseline.json';
const MICROSOFT_EVIDENCE_CATALOG =
  'scripts/docx-compatibility-microsoft-evidence.json';
const ALLOWED_DECLARATION_FILES = new Set([
  COMPATIBILITY_OWNER,
  `${DOCX_SOURCE}/layout/anchor-compatibility.ts`,
  `${DOCX_SOURCE}/layout/body-pagination-compatibility.ts`,
  `${DOCX_SOURCE}/layout/line-compatibility.ts`,
  `${DOCX_SOURCE}/layout/paint-compatibility.ts`,
  `${DOCX_SOURCE}/layout/script-compatibility.ts`,
  `${DOCX_SOURCE}/layout/section-compatibility.ts`,
  `${DOCX_SOURCE}/layout/table-compatibility.ts`,
]);

function fail(code, detail) {
  throw new Error(`${code}: ${detail}`);
}

function posixPath(path) {
  return path.split(sep).join('/');
}

function listFiles(root) {
  if (!existsSync(root)) return [];
  return readdirSync(root).flatMap((entry) => {
    const path = join(root, entry);
    return statSync(path).isDirectory() ? listFiles(path) : [path];
  });
}

function isProductionTypeScript(path) {
  return /\.(?:[cm]?[jt]s|[jt]sx)$/.test(path)
    && !path.endsWith('.d.ts')
    && !/\.(test|spec|stories)\.(?:[cm]?[jt]s|[jt]sx)$/.test(path)
    && !path.includes(`${sep}wasm${sep}`);
}

function compatibilityModuleTarget(root, file, specifier) {
  if (!specifier.startsWith('.')) return false;
  const absolute = resolve(root, dirname(file), specifier);
  const candidates = [
    absolute,
    absolute.replace(/\.(?:[cm]?js)$/, '.ts'),
  ];
  return candidates.some((candidate) =>
    posixPath(relative(root, candidate)) === COMPATIBILITY_OWNER);
}

function property(object, name) {
  return object.properties.find((entry) =>
    ts.isPropertyAssignment(entry)
    && ((ts.isIdentifier(entry.name) && entry.name.text === name)
      || (ts.isStringLiteral(entry.name) && entry.name.text === name)));
}

// property() validates the first matching literal entry, but spreads,
// duplicate names, computed or shorthand names, methods and accessors can make
// the runtime object expose a different value. Reject them before any lookup.
function verifyLiteralMembers(object, file) {
  const names = new Set();
  for (const entry of object.properties) {
    if (!ts.isPropertyAssignment(entry)
      || !(ts.isIdentifier(entry.name) || ts.isStringLiteral(entry.name))) {
      fail(
        'COMPATIBILITY_LITERAL_SCHEMA',
        `${file} requires ordinary literal-named object members`,
      );
    }
    if (names.has(entry.name.text)) {
      fail(
        'COMPATIBILITY_LITERAL_SCHEMA',
        `${file} has duplicate object member ${entry.name.text}`,
      );
    }
    names.add(entry.name.text);
  }
}

function stringValue(object, name, file) {
  const entry = property(object, name);
  if (!entry || !ts.isStringLiteralLike(entry.initializer)) {
    fail('COMPATIBILITY_LITERAL_SCHEMA', `${file} requires literal ${name}`);
  }
  if (entry.initializer.text.trim() === '') {
    fail('COMPATIBILITY_LITERAL_SCHEMA', `${file} has empty ${name}`);
  }
  return entry.initializer.text;
}

function objectValue(object, name, file) {
  const entry = property(object, name);
  if (!entry || !ts.isObjectLiteralExpression(entry.initializer)) {
    fail('COMPATIBILITY_LITERAL_SCHEMA', `${file} requires object ${name}`);
  }
  verifyLiteralMembers(entry.initializer, file);
  return entry.initializer;
}

function exportName(call, file) {
  const declaration = call.parent;
  if (!ts.isVariableDeclaration(declaration)
    || !ts.isIdentifier(declaration.name)
    || declaration.initializer !== call) {
    fail(
      'COMPATIBILITY_DECLARATION_SHAPE',
      `${file} must assign each rule directly to a named const`,
    );
  }
  const statement = declaration.parent?.parent;
  if (!ts.isVariableStatement(statement)
    || !statement.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword)
    || (declaration.parent.flags & ts.NodeFlags.Const) === 0) {
    fail(
      'COMPATIBILITY_DECLARATION_SHAPE',
      `${file} must export each rule as a const`,
    );
  }
  return declaration.name.text;
}

function verifyFactoryAccess(root, source, file, calls) {
  const allowedIdentifiers = new Set(calls.map((call) => call.expression));
  const factoryImports = [];

  for (const statement of source.statements) {
    if (ts.isImportDeclaration(statement)
      && ts.isStringLiteral(statement.moduleSpecifier)
      && compatibilityModuleTarget(root, file, statement.moduleSpecifier.text)) {
      const clause = statement.importClause;
      if (clause?.name || (clause?.namedBindings && ts.isNamespaceImport(clause.namedBindings))) {
        fail(
          'COMPATIBILITY_FACTORY_ACCESS',
          `${file} must not default- or namespace-import the compatibility owner`,
        );
      }
      if (clause?.namedBindings && ts.isNamedImports(clause.namedBindings)) {
        for (const element of clause.namedBindings.elements) {
          const importedName = element.propertyName?.text ?? element.name.text;
          if (importedName !== 'defineCompatibilityRule') continue;
          if (element.propertyName !== undefined
            || element.name.text !== 'defineCompatibilityRule'
            || !ALLOWED_DECLARATION_FILES.has(file)) {
            fail(
              'COMPATIBILITY_FACTORY_ACCESS',
              `${file} must not alias or import defineCompatibilityRule outside declaration authority`,
            );
          }
          factoryImports.push(element);
          allowedIdentifiers.add(element.name);
        }
      }
    }

    if (ts.isExportDeclaration(statement)
      && statement.moduleSpecifier
      && ts.isStringLiteral(statement.moduleSpecifier)
      && compatibilityModuleTarget(root, file, statement.moduleSpecifier.text)) {
      if (!statement.exportClause) {
        fail(
          'COMPATIBILITY_FACTORY_ACCESS',
          `${file} must not star-re-export the compatibility owner`,
        );
      }
      if (ts.isNamedExports(statement.exportClause)
        && statement.exportClause.elements.some((element) =>
          (element.propertyName?.text ?? element.name.text) === 'defineCompatibilityRule')) {
        fail(
          'COMPATIBILITY_FACTORY_ACCESS',
          `${file} must not re-export defineCompatibilityRule`,
        );
      }
    }
  }

  const visit = (node) => {
    if (ts.isCallExpression(node) && node.arguments.length > 0) {
      const [argument] = node.arguments;
      const isDynamicImport = node.expression.kind === ts.SyntaxKind.ImportKeyword;
      const isCommonJsRequire = ts.isIdentifier(node.expression)
        && node.expression.text === 'require';
      if ((isDynamicImport || isCommonJsRequire)
        && ts.isStringLiteralLike(argument)
        && compatibilityModuleTarget(root, file, argument.text)) {
        fail(
          'COMPATIBILITY_FACTORY_ACCESS',
          `${file} must not dynamically load the compatibility owner`,
        );
      }
    }
    if (ts.isElementAccessExpression(node)
      && node.argumentExpression
      && ts.isStringLiteralLike(node.argumentExpression)
      && node.argumentExpression.text === 'defineCompatibilityRule') {
      fail(
        'COMPATIBILITY_FACTORY_ACCESS',
        `${file} accesses defineCompatibilityRule through a string key`,
      );
    }
    if (ts.isIdentifier(node) && node.text === 'defineCompatibilityRule') {
      const isOwnerDeclaration = file === COMPATIBILITY_OWNER
        && ts.isFunctionDeclaration(node.parent)
        && node.parent.name === node;
      if (!isOwnerDeclaration && !allowedIdentifiers.has(node)) {
        fail(
          'COMPATIBILITY_FACTORY_ACCESS',
          `${file} accesses defineCompatibilityRule outside a direct checked call`,
        );
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return factoryImports;
}

function verifyFactoryImport(file, factoryImports) {
  if (file === COMPATIBILITY_OWNER) return;
  if (factoryImports.length !== 1) {
    fail(
      'COMPATIBILITY_FACTORY_IMPORT',
      `${file} must import defineCompatibilityRule exactly and without an alias`,
    );
  }
}

function microsoftEvidenceCatalog(root) {
  const path = resolve(root, MICROSOFT_EVIDENCE_CATALOG);
  if (!existsSync(path)) {
    fail('COMPATIBILITY_MICROSOFT_CATALOG', `${MICROSOFT_EVIDENCE_CATALOG} is missing`);
  }
  const catalog = JSON.parse(readFileSync(path, 'utf8'));
  if (catalog?.version !== 1 || !Array.isArray(catalog.documents)) {
    fail('COMPATIBILITY_MICROSOFT_CATALOG', MICROSOFT_EVIDENCE_CATALOG);
  }
  const documents = new Map();
  for (const document of catalog.documents) {
    if (typeof document?.id !== 'string'
      || typeof document?.revision !== 'string'
      || typeof document?.published !== 'string'
      || typeof document?.sourceUrl !== 'string'
      || !document.sourceUrl.startsWith('https://learn.microsoft.com/')
      || !Array.isArray(document?.sections)) {
      fail('COMPATIBILITY_MICROSOFT_CATALOG', `invalid document ${document?.id ?? '<missing>'}`);
    }
    if (documents.has(document.id)) {
      fail('COMPATIBILITY_MICROSOFT_CATALOG', `duplicate document ${document.id}`);
    }
    const sections = new Set();
    for (const section of document.sections) {
      if (typeof section?.id !== 'string'
        || typeof section?.title !== 'string'
        || section.title.trim() === ''
        || !/^\d+(?:\.\d+)+$/.test(section.id)
        || sections.has(section.id)) {
        fail(
          'COMPATIBILITY_MICROSOFT_CATALOG',
          `invalid section ${document.id} ${section?.id ?? '<missing>'}`,
        );
      }
      sections.add(section.id);
    }
    documents.set(document.id, sections);
  }
  return documents;
}

function verifyEvidence(root, file, ruleId, evidence, microsoftCatalog) {
  const kind = stringValue(evidence, 'kind', file);
  if (kind === 'regression-test') {
    const reference = stringValue(evidence, 'reference', file);
    const separator = reference.indexOf('#');
    if (separator <= 0 || separator === reference.length - 1) {
      fail(
        'COMPATIBILITY_REGRESSION_REFERENCE',
        `${ruleId} must use path#test-title`,
      );
    }
    const path = reference.slice(0, separator);
    const title = reference.slice(separator + 1);
    if (!path.startsWith(`${DOCX_SOURCE}/`) || !/\.(test|spec)\.tsx?$/.test(path)) {
      fail(
        'COMPATIBILITY_REGRESSION_REFERENCE',
        `${ruleId} references non-DOCX test ${path}`,
      );
    }
    const absolute = resolve(root, path);
    if (!existsSync(absolute)) {
      fail('COMPATIBILITY_EVIDENCE_MISSING', `${ruleId} references missing ${path}`);
    }
    if (!readFileSync(absolute, 'utf8').includes(title)) {
      fail(
        'COMPATIBILITY_EVIDENCE_STALE',
        `${ruleId} test title is absent from ${path}: ${title}`,
      );
    }
    return;
  }
  if (kind === 'microsoft-note') {
    const reference = stringValue(evidence, 'reference', file);
    const match = reference.match(/^\[(MS-[A-Z0-9]+)\] §§?(.+)$/);
    if (!match) {
      fail(
        'COMPATIBILITY_MICROSOFT_REFERENCE',
        `${ruleId} has invalid Microsoft note ${reference}`,
      );
    }
    const sectionIds = [...match[2].matchAll(/\b\d+(?:\.\d+)+\b/g)]
      .map((entry) => entry[0]);
    const knownSections = microsoftCatalog.get(match[1]);
    if (sectionIds.length === 0
      || !knownSections
      || sectionIds.some((sectionId) => !knownSections.has(sectionId))) {
      fail(
        'COMPATIBILITY_MICROSOFT_EVIDENCE',
        `${ruleId} references an uncatalogued Microsoft section: ${reference}`,
      );
    }
    return;
  }
  if (kind === 'office-observation') {
    const fixtureId = stringValue(evidence, 'syntheticFixtureId', file);
    stringValue(evidence, 'application', file);
    stringValue(evidence, 'version', file);
    stringValue(evidence, 'platform', file);
    if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(fixtureId)) {
      fail(
        'COMPATIBILITY_FIXTURE_ID',
        `${ruleId} has invalid synthetic fixture id ${fixtureId}`,
      );
    }
    return;
  }
  fail('COMPATIBILITY_EVIDENCE_KIND', `${ruleId} has unknown evidence kind ${kind}`);
}

function inspectSource(root, absolute, rules, microsoftCatalog) {
  const file = posixPath(relative(root, absolute));
  const source = ts.createSourceFile(
    absolute,
    readFileSync(absolute, 'utf8'),
    ts.ScriptTarget.Latest,
    true,
  );
  const calls = [];
  const visit = (node) => {
    if (ts.isCallExpression(node)
      && ts.isIdentifier(node.expression)
      && node.expression.text === 'defineCompatibilityRule') {
      calls.push(node);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  const factoryImports = verifyFactoryAccess(root, source, file, calls);
  if (calls.length === 0) return;
  if (!ALLOWED_DECLARATION_FILES.has(file)) {
    fail(
      'COMPATIBILITY_DECLARATION_AUTHORITY',
      `${file} declares a compatibility rule`,
    );
  }
  verifyFactoryImport(file, factoryImports);
  for (const call of calls) {
    if (call.arguments.length !== 1 || !ts.isObjectLiteralExpression(call.arguments[0])) {
      fail(
        'COMPATIBILITY_LITERAL_SCHEMA',
        `${file} must pass one literal rule object`,
      );
    }
    const object = call.arguments[0];
    const name = exportName(call, file);
    if (!/^[A-Z][A-Z0-9_]+$/.test(name)) {
      fail('COMPATIBILITY_EXPORT_NAME', `${file} exports invalid rule name ${name}`);
    }
    verifyLiteralMembers(object, file);
    const id = stringValue(object, 'id', file);
    stringValue(object, 'description', file);
    verifyEvidence(
      root,
      file,
      id,
      objectValue(object, 'evidence', file),
      microsoftCatalog,
    );
    const prior = rules.get(id);
    if (prior) {
      fail('COMPATIBILITY_DUPLICATE_ID', `${id} is declared by ${prior} and ${file}`);
    }
    rules.set(id, file);
  }
}

function options(argv) {
  let root = process.cwd();
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === '--root' && argv[index + 1]) {
      root = resolve(argv[index + 1]);
      index += 1;
      continue;
    }
    fail('UNKNOWN_ARGUMENT', argv[index] ?? '<missing>');
  }
  return { root };
}

const observationMarker =
  /\[MS-[A-Z0-9]+\]|\b(?:[Oo]bserved|[Mm]easured|[Vv]erified|[Aa]djudicated|[Rr]ecorded)\b[^\n]{0,120}\b(?:Word|Office)\b|\b(?:matching|following|exactly\s+like|where)\s+(?:Word|Office)\b|\b(?:Word|Office)(?:'s)?[- ](?:observed|verified|measured|adjudicated|compatible|runtime|specific|defined)\b|\b(?:Word|Office)(?:'s)?\s+[^.\n]{0,80}\b(?:default|pen)\b|\b(?:Word|Office)(?:'s)?\b[^.\n]{0,40}\b(?:uses|applies|adds|treats|admits|keeps|places|starts|does|doesn't|ignores|clips|clamps|pins|measures|renders|draws|fills|resolves|confirms|defers|emits|paints|centres|centers|collapses|aligns|reserves|lays|flows|widens|advances|classifies|breaks|counts|runs|excludes|honors|honours|formats|swaps|writes|stores|layers|anchors|moves|suppresses|never)\b|\b(?:Word|Office)(?:'s)?\s+(?:GT|ground truth|markup|output PDFs?|PDFs?|PDF export|behavior|behaviour|compatibility)\b/;

function observationKeys(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  return listFiles(sourceRoot)
    .filter(isProductionTypeScript)
    .flatMap((absolute) => {
      const file = posixPath(relative(root, absolute));
      if (ALLOWED_DECLARATION_FILES.has(file)) return [];
      return readFileSync(absolute, 'utf8').split(/\r?\n/).flatMap((line) =>
        observationMarker.test(line)
          ? [`${file}::${line.trim().replace(/\s+/g, ' ')}`]
          : []);
    })
    .sort();
}

function verifyObservationBaseline(root, observations) {
  const path = resolve(root, OBSERVATION_BASELINE);
  if (existsSync(path)) {
    fail('FINAL_COMPATIBILITY_OBSERVATION_BASELINE', OBSERVATION_BASELINE);
  }
  if (observations.length > 0) {
    fail(
      'INLINE_COMPATIBILITY_OBSERVATION',
      observations.join('\n'),
    );
  }
}

export function checkDocxCompatibilityEvidence(root) {
  const rules = new Map();
  const microsoftCatalog = microsoftEvidenceCatalog(root);
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const file of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    inspectSource(root, file, rules, microsoftCatalog);
  }
  if (rules.size === 0) fail('COMPATIBILITY_REGISTRY_EMPTY', DOCX_SOURCE);
  const observations = observationKeys(root);
  verifyObservationBaseline(root, observations);
  return Object.freeze([...rules].map(([id, file]) => Object.freeze({ id, file })));
}

if (process.argv[1]
  && resolve(process.argv[1]) === resolve(fileURLToPath(import.meta.url))) {
  try {
    const { root } = options(process.argv.slice(2));
    const rules = checkDocxCompatibilityEvidence(root);
    process.stdout.write(`DOCX compatibility evidence verified (${rules.length} rules).\n`);
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  }
}
