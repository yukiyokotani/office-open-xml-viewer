#!/usr/bin/env node
// Compares the declaration surface generated from the build with the
// committed baseline. A pass means the two agree; it does not cover runtime
// behaviour, and a signature change committed together with a matching
// baseline edit also passes, so the baseline diff is what a reviewer reads.

import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const toolRequire = createRequire(new URL('../package.json', import.meta.url));
const ts = toolRequire('typescript-compiler-api');

function parseArgs(argv) {
  const result = {
    root: process.cwd(),
    entry: 'docx.d.ts',
    baseline: 'packages/docx/api/public-api-baseline.d.ts',
    label: 'DOCX',
    baseRef: undefined,
    writeBaseline: false,
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--root') result.root = path.resolve(argv[++index]);
    else if (arg === '--entry') result.entry = argv[++index];
    else if (arg === '--baseline') result.baseline = argv[++index];
    else if (arg === '--label') result.label = argv[++index];
    else if (arg === '--base-ref') result.baseRef = argv[++index];
    else if (arg === '--write-baseline') result.writeBaseline = true;
    else throw new Error(`Unknown argument: ${arg}`);
  }
  return result;
}

function normalizeText(source) {
  return source
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .filter((line) => !/^\/\/# sourceMappingURL=/.test(line))
    .map((line) => line.replace(/[ \t]+$/g, ''))
    .join('\n')
    .trimEnd();
}

export function normalizeDeclaration(source, fileName) {
  const parsed = ts.createSourceFile(
    fileName,
    source,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  const localExportNames = new Set();
  const exportAliases = new Map();
  for (const statement of parsed.statements) {
    if (!ts.isExportDeclaration(statement) || statement.moduleSpecifier || !statement.exportClause
        || !ts.isNamedExports(statement.exportClause)) continue;
    for (const element of statement.exportClause.elements) {
      const localName = element.propertyName?.text ?? element.name.text;
      localExportNames.add(element.name.text);
      if (localName !== element.name.text) exportAliases.set(localName, element.name.text);
    }
  }
  const collisionGroups = new Map();
  for (const statement of parsed.statements) {
    if (!('name' in statement) || !statement.name || !ts.isIdentifier(statement.name)) continue;
    if (localExportNames.has(statement.name.text)) continue;
    const match = /^(.*?)(?:_|\$)\d+$/.exec(statement.name.text);
    if (!match) continue;
    const names = collisionGroups.get(match[1]) ?? [];
    names.push(statement.name.text);
    collisionGroups.set(match[1], names);
  }
  const collisionAliases = new Map();
  for (const [base, names] of collisionGroups) {
    names.sort((left, right) => {
      const leftOrdinal = Number(/\d+$/.exec(left)?.[0]);
      const rightOrdinal = Number(/\d+$/.exec(right)?.[0]);
      return leftOrdinal - rightOrdinal || left.localeCompare(right);
    });
    names.forEach((name, index) => {
      collisionAliases.set(name, `${base}__emitterCollision${index + 1}`);
    });
  }
  const transformed = ts.transform(parsed, [(context) => {
    const hasModifier = (member, kind) => ts.getModifiers(member)?.some(
      (modifier) => modifier.kind === kind,
    ) ?? false;
    const normalizeClassMembers = (members) => {
      let hasInstancePrivate = false;
      let hasStaticPrivate = false;
      let hasHardPrivate = false;
      let hasPrivateConstructor = false;
      const visibleMembers = members.filter((member) => {
        if (member.name && ts.isPrivateIdentifier(member.name)) {
          hasHardPrivate = true;
          return false;
        }
        if (!hasModifier(member, ts.SyntaxKind.PrivateKeyword)) return true;
        if (ts.isConstructorDeclaration(member)) hasPrivateConstructor = true;
        else if (hasModifier(member, ts.SyntaxKind.StaticKeyword)) hasStaticPrivate = true;
        else hasInstancePrivate = true;
        return false;
      });
      if (visibleMembers.length === members.length) return members;
      if (hasInstancePrivate) {
        visibleMembers.push(ts.factory.createPropertyDeclaration(
          [ts.factory.createModifier(ts.SyntaxKind.PrivateKeyword)],
          '__privatePresence',
        ));
      }
      if (hasStaticPrivate) {
        visibleMembers.push(ts.factory.createPropertyDeclaration(
          [
            ts.factory.createModifier(ts.SyntaxKind.PrivateKeyword),
            ts.factory.createModifier(ts.SyntaxKind.StaticKeyword),
          ],
          '__staticPrivatePresence',
        ));
      }
      if (hasHardPrivate) {
        visibleMembers.push(ts.factory.createPropertyDeclaration(
          undefined,
          ts.factory.createPrivateIdentifier('#private'),
        ));
      }
      if (hasPrivateConstructor) {
        visibleMembers.push(ts.factory.createConstructorDeclaration(
          [ts.factory.createModifier(ts.SyntaxKind.PrivateKeyword)],
          [],
          undefined,
        ));
      }
      return ts.factory.createNodeArray(visibleMembers);
    };
    const statementName = (statement) => {
      if ('name' in statement && statement.name && ts.isIdentifier(statement.name)) {
        return statement.name.text;
      }
      if (ts.isVariableStatement(statement)) {
        return statement.declarationList.declarations
          .map((declaration) => ts.isIdentifier(declaration.name) ? declaration.name.text : '')
          .join(',');
      }
      return '';
    };
    const normalizeSourceFile = (sourceFile) => {
      // API Extractor emits `export declare interface Foo`, whereas Rolldown's
      // declaration bundler emits `interface Foo` plus `export { Foo }`. Treat
      // those equivalent spellings alike so the guard protects the API rather
      // than coupling the project to one declaration emitter.
      const localExports = new Set();
      for (const statement of sourceFile.statements) {
        if (!ts.isExportDeclaration(statement) || statement.moduleSpecifier || !statement.exportClause) continue;
        if (!ts.isNamedExports(statement.exportClause)) continue;
        for (const element of statement.exportClause.elements) {
          const localName = element.propertyName?.text ?? element.name.text;
          if (localName === element.name.text) localExports.add(localName);
        }
      }

      const statements = [];
      for (const statement of sourceFile.statements) {
        if (ts.isExportDeclaration(statement) && !statement.moduleSpecifier && statement.exportClause
            && ts.isNamedExports(statement.exportClause)
            && statement.exportClause.elements.every((element) => {
              const localName = element.propertyName?.text ?? element.name.text;
              return localName === element.name.text;
            })) {
          continue;
        }
        const name = statementName(statement);
        if (ts.canHaveModifiers(statement)) {
          // `declare` is optional in an ambient .d.ts module and emitters differ
          // on whether they spell it explicitly.
          const modifiers = (ts.getModifiers(statement) ?? [])
            .filter((modifier) => modifier.kind !== ts.SyntaxKind.DeclareKeyword);
          if (name && localExports.has(name)
              && !modifiers.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword)) {
            statements.push(ts.factory.replaceModifiers(statement, [
              ts.factory.createModifier(ts.SyntaxKind.ExportKeyword),
              ...modifiers,
            ]));
            continue;
          }
          statements.push(ts.factory.replaceModifiers(statement, modifiers));
          continue;
        }
        statements.push(statement);
      }

      // Declaration order is not part of a module's public surface. Keep
      // overloads and merged declarations with the same name stable while
      // canonicalizing the emitter-specific top-level ordering.
      const indexed = statements.map((statement, index) => ({ statement, index }));
      indexed.sort((left, right) => {
        const leftName = statementName(left.statement);
        const rightName = statementName(right.statement);
        const byName = leftName.localeCompare(rightName);
        return byName || left.index - right.index;
      });
      return ts.factory.updateSourceFile(
        sourceFile,
        ts.factory.createNodeArray(indexed.map(({ statement }) => statement)),
      );
    };
    const visit = (node) => {
      if (ts.isIdentifier(node) && exportAliases.has(node.text)) {
        return ts.factory.createIdentifier(exportAliases.get(node.text));
      }
      if (ts.isIdentifier(node) && collisionAliases.has(node.text)) {
        return ts.factory.createIdentifier(collisionAliases.get(node.text));
      }
      if (ts.isParenthesizedTypeNode(node)) return visit(node.type);
      if (ts.isUnionTypeNode(node)) {
        return ts.factory.createUnionTypeNode(node.types.flatMap((type) => {
          const visited = visit(type);
          return ts.isUnionTypeNode(visited) ? [...visited.types] : [visited];
        }));
      }
      if (ts.isIntersectionTypeNode(node)) {
        return ts.factory.createIntersectionTypeNode(node.types.flatMap((type) => {
          const visited = visit(type);
          return ts.isIntersectionTypeNode(visited) ? [...visited.types] : [visited];
        }));
      }
      if (ts.isStringLiteral(node)) return ts.factory.createStringLiteral(node.text, true);
      const visited = ts.visitEachChild(node, visit, context);
      if (ts.isClassDeclaration(visited)) {
        return ts.factory.updateClassDeclaration(
          visited,
          visited.modifiers,
          visited.name,
          visited.typeParameters,
          visited.heritageClauses,
          normalizeClassMembers(visited.members),
        );
      }
      if (ts.isClassExpression(visited)) {
        return ts.factory.updateClassExpression(
          visited,
          visited.modifiers,
          visited.name,
          visited.typeParameters,
          visited.heritageClauses,
          normalizeClassMembers(visited.members),
        );
      }
      if (ts.isSourceFile(visited)) return normalizeSourceFile(visited);
      return visited;
    };
    return (root) => ts.visitNode(root, visit);
  }]);
  const printer = ts.createPrinter({
    newLine: ts.NewLineKind.LineFeed,
    removeComments: true,
  });
  const normalized = normalizeText(printer.printFile(transformed.transformed[0]));
  transformed.dispose();
  return normalized;
}

function localSpecifiers(source) {
  const info = ts.preProcessFile(source, true, true);
  const values = [
    ...info.importedFiles,
    ...info.referencedFiles,
  ].map((reference) => reference.fileName);
  return [...new Set(values.filter((specifier) => specifier.startsWith('.')))];
}

function resolveDeclaration(fromFile, specifier) {
  const absolute = path.resolve(path.dirname(fromFile), specifier);
  const withoutRuntimeExtension = absolute.replace(/\.(?:mjs|cjs|js|jsx|ts|tsx)$/, '');
  const candidates = [
    absolute,
    `${absolute}.d.ts`,
    `${withoutRuntimeExtension}.d.ts`,
    path.join(absolute, 'index.d.ts'),
    path.join(withoutRuntimeExtension, 'index.d.ts'),
  ];
  return candidates.find((candidate) => existsSync(candidate));
}

function collectDeclarations(typesRoot, entryName) {
  const entry = path.resolve(typesRoot, entryName);
  const entryRelative = path.relative(typesRoot, entry);
  if (entryRelative.startsWith('..') || path.isAbsolute(entryRelative)) {
    throw new Error(`Generated declaration entry must stay inside dist/types (${entryName}).`);
  }
  if (!existsSync(entry)) {
    throw new Error(`Generated declaration entry is missing (${entryName}); build the published package first.`);
  }

  const pending = [entry];
  const sources = new Map();
  while (pending.length > 0) {
    const file = pending.pop();
    if (sources.has(file)) continue;
    const source = readFileSync(file, 'utf8');
    sources.set(file, normalizeDeclaration(source, file));
    for (const specifier of localSpecifiers(source)) {
      const resolved = resolveDeclaration(file, specifier);
      if (!resolved) {
        const relative = path.relative(typesRoot, file);
        throw new Error(`Cannot resolve local declaration ${specifier} from ${relative}.`);
      }
      pending.push(resolved);
    }
  }

  return [...sources]
    .map(([file, source]) => ({ file: path.relative(typesRoot, file).split(path.sep).join('/'), source }))
    .sort((left, right) => left.file.localeCompare(right.file));
}

function renderBaseline(declarations, label) {
  const header = [
    '// Generated by scripts/check-public-api.mjs.',
    `// This file records every local declaration reachable from the ${label} public entry.`,
    '// Do not edit by hand.',
  ].join('\n');
  const modules = declarations.map(({ file, source }) => `// --- file: ${file} ---\n${source}`);
  return `${header}\n\n${modules.join('\n\n')}\n`;
}

function normalizeRenderedBaseline(source) {
  return normalizeText(source)
    .split(/\n\n(?=\/\/ --- file: )/)
    .map((section, index) => {
      if (index === 0) return section;
      const declarationStart = section.indexOf('\n');
      const heading = section.slice(0, declarationStart);
      const fileName = heading.slice('// --- file: '.length, -' ---'.length);
      return `${heading}\n${normalizeDeclaration(section.slice(declarationStart + 1), fileName)}`;
    })
    .join('\n\n');
}

function resolveMergeBase(root, explicit) {
  for (const candidate of explicit ? [explicit] : ['origin/main', 'main']) {
    try {
      return execFileSync('git', ['merge-base', candidate, 'HEAD'], {
        cwd: root,
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'ignore'],
      }).trim();
    } catch {}
  }
  throw new Error('Cannot resolve the merge base; fetch origin/main or pass --base-ref.');
}

function refContains(root, ref, relativePath) {
  try {
    execFileSync('git', ['cat-file', '-e', `${ref}:${relativePath}`], {
      cwd: root,
      stdio: 'ignore',
    });
    return true;
  } catch {
    return false;
  }
}

export function checkPublicApi(options) {
  const typesRoot = path.join(options.root, 'dist/types');
  const baselineRelative = options.baseline ?? 'packages/docx/api/public-api-baseline.d.ts';
  const baselinePath = path.join(options.root, baselineRelative);
  const resolvedBaselineRelative = path.relative(options.root, baselinePath);
  if (resolvedBaselineRelative.startsWith('..') || path.isAbsolute(resolvedBaselineRelative)) {
    throw new Error(`Public API baseline must stay inside the repository (${baselineRelative}).`);
  }
  const label = options.label ?? 'DOCX';
  const actual = renderBaseline(collectDeclarations(typesRoot, options.entry), label);
  const mergeBase = resolveMergeBase(options.root, options.baseRef);

  if (options.writeBaseline) {
    if (refContains(options.root, mergeBase, baselineRelative)) {
      throw new Error('--write-baseline is only permitted before the merge base contains the public API baseline.');
    }
    mkdirSync(path.dirname(baselinePath), { recursive: true });
    writeFileSync(baselinePath, actual);
    process.stdout.write(`Wrote ${baselineRelative}.\n`);
    return;
  }

  if (!existsSync(baselinePath)) {
    throw new Error(`Public API baseline is missing (${baselineRelative}).`);
  }
  const expected = normalizeText(readFileSync(baselinePath, 'utf8'));
  if (normalizeRenderedBaseline(actual) !== normalizeRenderedBaseline(expected)) {
    throw new Error(
      `${label} public API declaration baseline differs. Rebuild, inspect the reachable declarations, and update the committed baseline only for an intentional compatible API change.`,
    );
  }
  process.stdout.write(`${label} public API declaration baseline matches.\n`);
}

function isMainModule() {
  return process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
}

if (isMainModule()) {
  try {
    checkPublicApi(parseArgs(process.argv.slice(2)));
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  }
}
