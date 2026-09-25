#!/usr/bin/env node
// Keep legacy Office support out of the OOXML packages.
//
// The legacy DOC/XLS/PPT readers are an opt-in module (packages/legacy-converter)
// that plugs into the renderers only through the format-generic ModelSource
// contract. Core, DOCX, XLSX, PPTX and the Node facade must therefore carry no
// legacy-specific names, branches, defaults, dependencies or imports: anything
// they need must be expressible as a generic contract or capability.
//
// Usage: node scripts/check-core-legacy-boundary.mjs [--ref <git-ref>]
// Without --ref the working tree's tracked files are checked.
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export const GUARDED_ROOTS = [
  'packages/core/',
  'packages/docx/',
  'packages/xlsx/',
  'packages/pptx/',
  'packages/node/',
];

// Generated or local-only content inside the guarded roots.
const IGNORED = [
  /\/src\/wasm\//,
  /\/public\//,
  /\/tests\/visual\/(baseline|screenshots|diffs|references|report)\//,
];

const CHECKED_EXTENSIONS = /\.(?:[cm]?[jt]sx?|rs|json|toml|md|astro|html)$/;

/**
 * Each rule names a legacy-only concept. `legacy-binary-format` (the OOXML
 * loaders' typed rejection of a CFB container) and the container sniffer's
 * stream-name table predate the legacy module and stay allowed.
 */
export const RULES = [
  { id: 'legacy-package', pattern: /ooxml-legacy-converter|legacy-converter|legacy_office_converter/ },
  { id: 'legacy-format-name', pattern: /legacy[-_ ]?(?:doc|xls|ppt)(?![a-z])/i },
  { id: 'legacy-office-api', pattern: /LegacyOffice|legacyConversion|legacy-office|legacy_office/ },
  { id: 'legacy-sniffer', pattern: /sniffLegacyOfficeFormat|LegacyCfbFormat/ },
  { id: 'direct-format-name', pattern: /\bdirect[-_]?(?:doc|xls|ppt)(?![a-z])/i },
  { id: 'native-legacy-source', pattern: /native(?:Doc|Xls|Ppt)\b|nativeSource\b/ },
  { id: 'legacy-revision-view', pattern: /sourceRevisionView|sourceRevisionMarkup|revision_markup_in_print/ },
  { id: 'legacy-xls-measurement', pattern: /measureLegacy|xls-font|XLS_FONT_|configure_mdw|measurement_request/ },
  // Owner decision: DOCX w:lang/@w:val (langDefault) is not modeled.
  { id: 'docx-lang-default', pattern: /langDefault|lang_default/ },
];

export function findViolations(files) {
  const violations = [];
  for (const { path, text } of files) {
    const lines = text.split('\n');
    lines.forEach((line, index) => {
      for (const rule of RULES) {
        if (rule.pattern.test(line)) {
          violations.push({ path, line: index + 1, rule: rule.id, text: line.trim().slice(0, 160) });
        }
      }
    });
  }
  return violations;
}

export function isGuardedPath(path) {
  return GUARDED_ROOTS.some((root) => path.startsWith(root))
    && CHECKED_EXTENSIONS.test(path)
    && !IGNORED.some((pattern) => pattern.test(path));
}

function trackedFiles(ref) {
  const git = (args) => execFileSync('git', args, { encoding: 'utf8', maxBuffer: 256 * 1024 * 1024 });
  const paths = (ref
    ? git(['ls-tree', '-r', '--name-only', ref, '--', ...GUARDED_ROOTS])
    : git(['ls-files', '--', ...GUARDED_ROOTS]))
    .split('\n')
    .filter(Boolean)
    .filter(isGuardedPath)
    // A tracked file deleted in the working tree is not part of the tree checked.
    .filter((path) => ref || existsSync(path));
  return paths.map((path) => ({
    path,
    text: ref ? git(['show', `${ref}:${path}`]) : readFileSync(path, 'utf8'),
  }));
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const refIndex = process.argv.indexOf('--ref');
  const ref = refIndex >= 0 ? process.argv[refIndex + 1] : undefined;
  const violations = findViolations(trackedFiles(ref));
  if (violations.length > 0) {
    for (const violation of violations.slice(0, 200)) {
      console.error(`${violation.path}:${violation.line} [${violation.rule}] ${violation.text}`);
    }
    console.error(`\n${violations.length} legacy-specific reference(s) in the OOXML packages.`);
    console.error('Express the need as a format-generic contract (see core source/model-source.ts).');
    process.exit(1);
  }
  console.log('OOXML packages carry no legacy-specific names, branches or imports.');
}
