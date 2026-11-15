#!/usr/bin/env node
// Fails when something listed in scripts/tombstones.mjs has come back.
import { existsSync } from 'node:fs';
import { basename, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  isProductionTypeScript, listFiles, posixPath, sourceFile, ts,
} from './check-docx-layout-boundaries.mjs';
import { DELETED_IDENTIFIERS, DELETED_PATHS } from './tombstones.mjs';

function isStaticPropertyName(node) {
  const parent = node.parent;
  return (ts.isPropertyAccessExpression(parent) && parent.name === node)
    || (ts.isElementAccessExpression(parent) && parent.argumentExpression === node)
    || (ts.isPropertyAssignment(parent) && parent.name === node)
    || (ts.isShorthandPropertyAssignment(parent) && parent.name === node)
    || (ts.isPropertyDeclaration(parent) && parent.name === node)
    || (ts.isPropertySignature(parent) && parent.name === node)
    || (ts.isMethodDeclaration(parent) && parent.name === node)
    || (ts.isGetAccessorDeclaration(parent) && parent.name === node)
    || (ts.isSetAccessorDeclaration(parent) && parent.name === node)
    || (ts.isBindingElement(parent) && (parent.propertyName ?? parent.name) === node);
}

/** Every resurrected tombstone under `root`, as sorted `kind: detail` lines. */
export function tombstoneViolations(root) {
  const violations = new Set();
  for (const entry of DELETED_PATHS) {
    if (entry.path) {
      if (existsSync(resolve(root, entry.path))) violations.add(`DELETED_PATH: ${entry.path}`);
      continue;
    }
    for (const path of listFiles(resolve(root, entry.under))) {
      if (basename(path) === entry.basename) violations.add(`DELETED_PATH: ${posixPath(relative(root, path))}`);
    }
  }
  for (const entry of DELETED_IDENTIFIERS) {
    const names = new Set(entry.names);
    const exceptions = new Set(entry.patternExceptions);
    const propertyNames = new Set(entry.propertyNames);
    for (const path of listFiles(resolve(root, entry.root)).filter(isProductionTypeScript)) {
      const file = posixPath(relative(root, path));
      const visit = (node) => {
        if (ts.isIdentifier(node)
          && (names.has(node.text) || (entry.pattern.test(node.text) && !exceptions.has(node.text)))) {
          violations.add(`DELETED_IDENTIFIER: ${file}#${node.text}`);
        }
        if ((ts.isIdentifier(node) || ts.isStringLiteralLike(node))
          && propertyNames.has(node.text) && isStaticPropertyName(node)) {
          violations.add(`DELETED_PROPERTY: ${file}#${node.text}`);
        }
        ts.forEachChild(node, visit);
      };
      visit(sourceFile(path));
    }
  }
  return [...violations].sort();
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const violations = tombstoneViolations(resolve(process.argv[2] ?? '.'));
  if (violations.length > 0) {
    console.error(`${violations.join('\n')}\nThese were deleted on purpose (scripts/tombstones.mjs) and must not return.`);
    process.exitCode = 1;
  } else {
    console.log('no deleted path or identifier has returned');
  }
}
