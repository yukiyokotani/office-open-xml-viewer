// Agent knowledge stays structured, indexed and reachable from both Codex and
// Claude Code.
//
// AGENTS.md files are short maps; the knowledge lives in docs/, the system of
// record (see the root AGENTS.md, "How the knowledge is organized"). Codex
// concatenates AGENTS.md from the repository root down to the working
// directory and stops reading at project_doc_max_bytes (32 KiB by default),
// so text past that limit silently vanishes. Skills in .agents/skills (where
// Codex finds them) are thin pointers into docs/workflow/; Claude Code reaches
// each one through a relative link in .claude/skills, so the link works in
// every checkout and worktree.
//
// The same suite keeps the documentation from accumulating: docs/ may hold
// only procedures, the core beliefs and user guides something consumer-facing
// links to, Markdown may not appear elsewhere under other names, and the
// paths and scripts a page names must exist (docs/workflow/documentation.md).

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, lstatSync, readFileSync, readdirSync, readlinkSync, statSync } from 'node:fs';
import { dirname, join, posix, relative, resolve, sep } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const CODEX_PROJECT_DOC_MAX_BYTES = 32 * 1024;
const ROOT_AGENTS_MAX_LINES = 120;
const PACKAGE_AGENTS_MAX_LINES = 100;
const SKILL_BODY_MAX_LINES = 12;
const SKIP = new Set(['node_modules', '.git', 'dist', 'target', 'spec', 'worktrees', 'private', 'wasm']);

function walk(dir, accept) {
  const found = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (SKIP.has(entry.name)) continue;
    const path = join(dir, entry.name);
    if (entry.isDirectory()) found.push(...walk(path, accept));
    else if (accept(entry.name)) found.push(path);
  }
  return found;
}

const agentsFiles = () => walk(ROOT, (name) => name === 'AGENTS.md');
const docsFiles = () => walk(join(ROOT, 'docs'), (name) => name.endsWith('.md'));
// Repository-relative path with forward slashes on every platform.
const rel = (path) => relative(ROOT, path).split(sep).join('/');
// File text with LF line endings, whatever the checkout's autocrlf setting.
const read = (file) => readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
const isInside = (dir, path) => {
  const offset = relative(dir, path);
  return offset !== '' && offset.split(sep)[0] !== '..' && offset === relative(dir, resolve(dir, offset));
};

// Relative Markdown link targets outside fenced code blocks, without anchors.
function relativeLinks(file) {
  const text = read(file).replace(/^```[\s\S]*?^```/gm, '');
  const targets = [];
  for (const match of text.matchAll(/\]\(([^)\s]+)\)/g)) {
    const target = match[1].split('#')[0];
    if (!target || /^[a-z][a-z0-9+.-]*:/i.test(target)) continue;
    targets.push(target);
  }
  return targets;
}

test('every AGENTS.md chain fits within the Codex instruction limit', () => {
  const over = [];
  for (const file of agentsFiles()) {
    let bytes = 0;
    for (let dir = dirname(file); ; dir = dirname(dir)) {
      const candidate = join(dir, 'AGENTS.md');
      if (existsSync(candidate)) bytes += statSync(candidate).size;
      if (dir === ROOT) break;
    }
    if (bytes > CODEX_PROJECT_DOC_MAX_BYTES) over.push(`${rel(file)}: ${bytes} bytes`);
  }
  assert.deepEqual(over, [], `move content into docs/ and link it (limit ${CODEX_PROJECT_DOC_MAX_BYTES} bytes per chain)`);
});

test('every AGENTS.md stays a short map', () => {
  const over = [];
  for (const file of agentsFiles()) {
    const limit = dirname(file) === ROOT ? ROOT_AGENTS_MAX_LINES : PACKAGE_AGENTS_MAX_LINES;
    const lines = read(file).split('\n').length - 1;
    if (lines > limit) over.push(`${rel(file)}: ${lines} lines (limit ${limit})`);
  }
  assert.deepEqual(over, [], 'move the detail into a docs/ page and leave a one-line pointer in AGENTS.md');
});

test('every docs/ page is listed in an index', () => {
  // An index is docs/README.md or the README.md of the page's own directory.
  const indexed = new Set();
  for (const file of docsFiles().filter((path) => path.endsWith('README.md'))) {
    for (const target of relativeLinks(file)) indexed.add(resolve(dirname(file), target));
  }
  const missing = docsFiles()
    .filter((path) => path !== join(ROOT, 'docs', 'README.md') && !indexed.has(path))
    .map(rel);
  assert.deepEqual(missing, [], 'link each page from docs/README.md (or its directory README.md) with a one-line summary');
});

test('relative links in agent knowledge resolve', () => {
  const sources = [
    ...agentsFiles(),
    ...docsFiles(),
    ...walk(join(ROOT, '.agents'), (name) => name === 'SKILL.md'),
  ];
  const broken = [];
  for (const file of sources) {
    for (const target of relativeLinks(file)) {
      if (!existsSync(resolve(dirname(file), target))) broken.push(`${rel(file)} -> ${target}`);
    }
  }
  assert.deepEqual(broken, [], 'fix or remove each broken link');
});

test('every repository skill is a linked pointer into docs/', () => {
  const shared = join(ROOT, '.agents', 'skills');
  const links = join(ROOT, '.claude', 'skills');
  const problems = [];
  const names = existsSync(shared) ? readdirSync(shared) : [];
  for (const name of names) {
    const skill = join(shared, name, 'SKILL.md');
    if (!existsSync(skill)) {
      problems.push(`${name}: no SKILL.md`);
      continue;
    }
    const text = read(skill);
    const front = /^---\n([\s\S]*?)\n---\n/.exec(text);
    if (!front || !new RegExp(`^name: ${name}$`, 'm').test(front[1]) || !/^description: \S/m.test(front[1])) {
      problems.push(`${name}: SKILL.md front matter needs name: ${name} and a description`);
    }
    const body = front ? text.slice(front[0].length) : text;
    if (body.trim().split('\n').length > SKILL_BODY_MAX_LINES) {
      problems.push(`${name}: keep the procedure in docs/workflow/ and only point to it from the skill`);
    }
    if (!relativeLinks(skill).some((target) => isInside(join(ROOT, 'docs'), resolve(dirname(skill), target)))) {
      problems.push(`${name}: SKILL.md must link the docs/ page it loads`);
    }
    const link = join(links, name);
    const target = `../../.agents/skills/${name}`;
    if (!existsSync(link) || !lstatSync(link).isSymbolicLink() || readlinkSync(link).split(sep).join('/') !== target) {
      problems.push(`${name}: link .claude/skills/${name} -> ${target}`);
    }
  }
  for (const name of existsSync(links) ? readdirSync(links) : []) {
    if (!names.includes(name)) problems.push(`.claude/skills/${name}: no matching .agents/skills/${name}`);
  }
  assert.deepEqual(problems, []);
});

// Tracked files only, so local scratch files and other worktrees do not count.
const tracked = () =>
  execFileSync('git', ['ls-files', '-z'], { cwd: ROOT, encoding: 'utf8' }).split('\0').filter(Boolean);

test('Markdown exists only under the allowed names and places', () => {
  const allowedNames = new Set([
    'README.md',
    'AGENTS.md',
    'CLAUDE.md',
    'SKILL.md',
    'CHANGELOG.md',
    'CODE_OF_CONDUCT.md',
    'THIRD_PARTY_NOTICES.md',
  ]);
  const stray = tracked().filter((path) => {
    if (!/\.mdx?$/.test(path)) return false;
    if (path.startsWith('docs/')) return false;
    if (path === '.storybook/Introduction.mdx') return false; // the Storybook landing page
    return !allowedNames.has(path.split('/').at(-1));
  });
  assert.deepEqual(stray, [], 'record the knowledge in source, a script header or docs/ (docs/workflow/documentation.md)');
});

// Repository paths of the docs/ pages a consumer-facing file links to. A
// Markdown file counts its link destinations; site and package source count
// string content outside comments, where a GitHub URL or a path names the page.
function linkedDocs(path, text) {
  const found = new Set();
  const add = (target) => {
    const clean = target.split(/[?#]/)[0];
    if (/\.md$/.test(clean) && clean.startsWith('docs/')) found.add(clean);
  };
  if (path.endsWith('.md')) {
    const body = text.replace(/^```[\s\S]*?^```/gm, '');
    for (const [, target] of body.matchAll(/\]\(<?([^)\s>]+)>?\)/g)) {
      const github = /^https:\/\/github\.com\/[^/]+\/[^/]+\/blob\/[^/]+\/(.+)$/.exec(target);
      if (github) add(github[1]);
      else if (!/^[a-z][a-z0-9+.-]*:/i.test(target)) add(posix.normalize(posix.join(posix.dirname(path), target)));
    }
    return found;
  }
  const code = text
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/(^|\s)\/\/.*$/gm, '$1');
  for (const [, target] of code.matchAll(/(?:^|[\s"'`(/])(docs\/(?:[\w.-]+\/)*[\w-]+\.md)(?![\w-]|\.\w)/g)) {
    add(posix.normalize(target));
  }
  return found;
}

test('a consumer link to a guide is a real link destination, not incidental text', () => {
  const links = (path, text) => [...linkedDocs(path, text)];
  // Markdown: a relative or GitHub link destination counts, wherever it is written from.
  assert.deepEqual(links('README.md', 'See the [guide](docs/./example.md#part).'), ['docs/example.md']);
  assert.deepEqual(links('packages/x/README.md', '[guide](../../docs/example.md)'), ['docs/example.md']);
  assert.deepEqual(
    links('README.md', '[guide](https://github.com/o/r/blob/main/docs/example.md)'),
    ['docs/example.md'],
  );
  assert.deepEqual(links('README.md', 'The file `docs/example.md` is mentioned, [old](docs/example.md.old).'), []);
  // Site and source: a string naming the page counts; a comment or a longer name does not.
  assert.deepEqual(
    links('site/src/pages/a.astro', '<a href="https://github.com/o/r/blob/main/docs/example.md">guide</a>'),
    ['docs/example.md'],
  );
  assert.deepEqual(links('site/src/pages/a.astro', '<!-- docs/example.md --> <p>docs/example.md.old</p>'), []);
  assert.deepEqual(
    links('packages/core/src/a.ts', "export const MESSAGE = 'See https://github.com/o/r/blob/main/docs/example.md.';"),
    ['docs/example.md'],
  );
  assert.deepEqual(links('packages/core/src/a.ts', '// see docs/example.md\n/* docs/example.md */\nexport {};'), []);
});

test('docs/ holds only procedures, the core beliefs and linked user guides', () => {
  const files = tracked().filter((path) => path.startsWith('docs/') && /\.mdx?$/.test(path));
  const misplaced = files.filter(
    (path) =>
      path.endsWith('.mdx') ||
      (path !== 'docs/README.md' &&
        path !== 'docs/design-docs/core-beliefs.md' &&
        !/^docs\/workflow\/[^/]+\.md$/.test(path) &&
        !/^docs\/[^/]+\.md$/.test(path)),
  );
  assert.deepEqual(misplaced, [], 'docs/ has no other directories or formats; design rationale belongs beside the source');

  // A user guide earns its place by being linked from something a library
  // consumer reads: a README, the site, or a string in shipped source.
  const surfaces = tracked().filter(
    (path) =>
      path === 'README.md' ||
      /^packages\/[^/]+\/README\.md$/.test(path) ||
      (/^site\/src\//.test(path) && !/\.test\.[cm]?[jt]s$/.test(path)) ||
      (/^packages\/[^/]+\/src\/.*\.[cm]?[jt]s$/.test(path) && !/\.test\.[cm]?[jt]s$/.test(path)),
  );
  const linked = new Set();
  for (const surface of surfaces) {
    const text = read(join(ROOT, surface));
    if (!text.includes('docs/')) continue;
    for (const target of linkedDocs(surface, text)) linked.add(target);
  }
  const guides = files.filter((path) => /^docs\/[^/]+\.md$/.test(path) && path !== 'docs/README.md');
  assert.deepEqual(
    guides.filter((guide) => !linked.has(guide)),
    [],
    'no README, site page or shipped source links this guide: delete it, or move design rationale beside the source',
  );
});

test('docs/ carries no task lists', () => {
  const found = [];
  for (const file of docsFiles()) {
    const lines = read(file).replace(/^```[\s\S]*?^```/gm, '').split('\n');
    if (lines.some((line) => /^\s*[-*] \[[ xX]\] /.test(line))) found.push(rel(file));
  }
  assert.deepEqual(found, [], 'plans and trackers are GitHub issues, not documents');
});

// What a page names that the repository does not have: a backticked repository
// path, or a pnpm script. `scriptsOf(filter)` returns the script names a
// command can run: those of the package a `--filter` names, otherwise of any
// package.
function missingReferences(file, text, { exists, scriptsOf }) {
  const pnpmBuiltins = new Set(['install', 'exec', 'add', 'run', 'dlx', 'why', 'pack', 'publish']);
  // Generated, git-ignored or local-only locations a page may legitimately name.
  const localOnly = [/^spec(\/|$)/, /(^|\/)src\/wasm(\/|$)/, /\/public\/private(\/|$)/, /^node_modules\//, /(^|\/)dist(\/|$)/, /^\/?tmp\//];
  // Part names inside an OOXML or ODF package, not repository paths.
  const packageParts = /^(word|xl|ppt|docProps|_rels|customXml|META-INF)\//;
  const sourceFile = /\.(md|mjs|cjs|js|ts|tsx|json|ya?ml|rs|py|sh|ps1|toml|html|astro|applescript)$/;
  const missing = [];
  for (const [, span] of text.matchAll(/`([^`\n]+)`/g)) {
    // A repository path: at least one slash, no spaces, placeholders or globs.
    if (!/^\.?[\w@.-]+(\/[\w@.-]+)+\/?$/.test(span) || /^(https?:|@silurus\/|origin\/|refs\/)/.test(span)) continue;
    if (localOnly.some((pattern) => pattern.test(span)) || packageParts.test(span)) continue;
    const first = span.replace(/^\.\//, '').split('/')[0];
    const bases = ['', posix.dirname(file)].filter((base) => exists(posix.join(base, first)));
    // Slash-separated identifiers (`lIns/tIns`) are not paths. A span is one
    // when its first segment exists, or when it names a source file: the
    // second case catches a path whose whole directory was removed.
    if (bases.length === 0 ? sourceFile.test(span) : !bases.some((base) => exists(posix.join(base, span)))) {
      missing.push(`${file}: ${span}`);
    }
  }
  for (const [, filter, name] of text.matchAll(/\bpnpm (?:--filter (\S+) )?(?:run )?([a-z][\w:.-]*)/g)) {
    if (pnpmBuiltins.has(name)) continue;
    // A placeholder filter (`@silurus/ooxml-<fmt>`) names no single package.
    const scoped = filter && !/[<*{]/.test(filter) ? filter : undefined;
    if (!scriptsOf(scoped).has(name)) missing.push(`${file}: pnpm ${filter ? `--filter ${filter} ` : ''}${name}`);
  }
  return missing;
}

test('a removed path or a script of another package is reported', () => {
  const present = new Set(['scripts', 'scripts/a.mjs', 'packages', 'packages/docx']);
  const scripts = { '@scope/docx': new Set(['conformance', 'vrt']), '@scope/pptx': new Set(['vrt']) };
  const all = new Set(Object.values(scripts).flatMap((names) => [...names]));
  const context = {
    exists: (path) => present.has(path.replace(/\/$/, '')),
    scriptsOf: (filter) => (filter ? (scripts[filter] ?? new Set()) : all),
  };
  const check = (text) => missingReferences('AGENTS.md', text, context);
  assert.deepEqual(check('`scripts/a.mjs`, `packages/docx/`, `lIns/tIns`, `word/fontTable.xml`, `spec/part1`'), []);
  assert.deepEqual(check('`scripts/gone.mjs` and `rules/example.yml`'), [
    'AGENTS.md: scripts/gone.mjs',
    'AGENTS.md: rules/example.yml',
  ]);
  assert.deepEqual(check('`pnpm vrt`, `pnpm --filter @scope/docx conformance`, `pnpm --filter @scope/<fmt> conformance`, `pnpm install`'), []);
  assert.deepEqual(check('`pnpm --filter @scope/pptx conformance` and `pnpm nope`'), [
    'AGENTS.md: pnpm --filter @scope/pptx conformance',
    'AGENTS.md: pnpm nope',
  ]);
});

test('paths and pnpm scripts named in agent knowledge exist', () => {
  const sources = [
    ...agentsFiles(),
    ...walk(join(ROOT, 'docs', 'workflow'), (name) => name.endsWith('.md')),
    ...walk(join(ROOT, 'docs', 'design-docs'), (name) => name.endsWith('.md')),
    ...walk(join(ROOT, '.agents'), (name) => name === 'SKILL.md'),
  ];
  const byName = new Map();
  const all = new Set();
  const manifests = [ROOT, ...readdirSync(join(ROOT, 'packages')).map((name) => join(ROOT, 'packages', name))];
  for (const dir of manifests) {
    const manifest = join(dir, 'package.json');
    if (!existsSync(manifest)) continue;
    const { name, scripts = {} } = JSON.parse(read(manifest));
    const names = new Set(Object.keys(scripts));
    byName.set(name, names);
    byName.set(`./${rel(dir)}`, names);
    for (const script of names) all.add(script);
  }
  const context = {
    exists: (path) => existsSync(join(ROOT, path)),
    scriptsOf: (filter) => (filter ? (byName.get(filter) ?? new Set()) : all),
  };
  const missing = sources.flatMap((file) => missingReferences(rel(file), read(file), context));
  assert.deepEqual(missing, [], 'the page names something that no longer exists: update or delete the statement');
});
