#!/usr/bin/env node

import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import process from 'node:process';
import { rolldown } from 'rolldown';
import { dts } from 'rolldown-plugin-dts';

const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript-compiler-api');

const entries = [
  'index',
  'docx',
  'xlsx',
  'pptx',
  'math',
  'three-d',
  'region-map',
  'chart-ex',
  'node',
];
const dist = path.resolve(process.cwd(), 'dist');
const workDir = path.join(dist, '.types-work');
const outDir = path.join(dist, 'types');

function stripInternalMembers(sourceFile) {
  const isInternal = (node) => ts.getJSDocTags(node)
    .some((tag) => tag.tagName.text === 'internal');
  const ranges = [];
  const visit = (node) => {
    if (ts.isClassDeclaration(node) || ts.isClassExpression(node) || ts.isInterfaceDeclaration(node)) {
      for (const member of node.members) {
        if (isInternal(member)) ranges.push([member.getFullStart(), member.getEnd()]);
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
  return ranges
    .sort((left, right) => right[0] - left[0])
    .reduce(
      (source, [start, end]) => `${source.slice(0, start)}${source.slice(end)}`,
      sourceFile.getFullText(),
    );
}

function stripComments(source) {
  // Rolldown's Oxc declaration resolver can reattach JSDoc between modifiers
  // and a member name, which changes the parsed public surface under tsgo.
  // Strip comments only after @internal members have been identified and
  // removed; the declarations themselves remain the compiler-owned source of
  // truth and are compiled again by check-public-type-exports.mjs.
  const scanner = ts.createScanner(ts.ScriptTarget.Latest, false, undefined, source);
  const tokens = [];
  for (let kind = scanner.scan(); kind !== ts.SyntaxKind.EndOfFileToken; kind = scanner.scan()) {
    if (kind !== ts.SyntaxKind.SingleLineCommentTrivia
        && kind !== ts.SyntaxKind.MultiLineCommentTrivia) {
      tokens.push(scanner.getTokenText());
    }
  }
  return tokens.join('');
}

async function declarationFiles(root) {
  const entries = await readdir(root, { recursive: true, withFileTypes: true });
  return entries
    .filter((entry) => entry.isFile() && entry.name.endsWith('.d.ts'))
    .map((entry) => path.join(entry.parentPath ?? entry.path, entry.name));
}

async function prepareDeclarationInputs(files) {
  await Promise.all(files.map(async (file) => {
    const source = await readFile(file, 'utf8');
    const sourceFile = ts.createSourceFile(
      file,
      source,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const withoutInternals = stripInternalMembers(sourceFile);
    await writeFile(file, stripComments(withoutInternals));
  }));
}

await mkdir(outDir, { recursive: true });

try {
  await prepareDeclarationInputs(await declarationFiles(workDir));

  await Promise.all(entries.map(async (entry) => {
    const build = await rolldown({
      input: path.join(workDir, `${entry}.d.ts`),
      plugins: [dts({ dtsInput: true })],
    });
    try {
      await build.write({
        file: path.join(outDir, `${entry}.d.ts`),
        format: 'es',
        codeSplitting: false,
      });
    } finally {
      await build.close();
    }
  }));
} finally {
  await rm(workDir, { recursive: true, force: true });
}
