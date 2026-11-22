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
    const workflow = join(ROOT, 'docs', 'workflow');
    if (!relativeLinks(skill).some((target) => target.endsWith('.md') && isInside(workflow, resolve(dirname(skill), target)))) {
      problems.push(`${name}: SKILL.md must link the docs/workflow/ page it loads`);
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

// The URL prefix under which a page on `main` is published on GitHub, taken
// from the root manifest so a fork or rename needs no edit here.
const MAIN_BLOB_URL = `${JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'))
  .repository.url.replace(/^git\+/, '')
  .replace(/\.git$/, '')}/blob/main/`;

// Guides a supported release still sends users to, by a URL its shipped code
// prints. Such a guide stays even when nothing on `main` links to it any more.
// Remove the entry, and then the guide, when the release leaves support.
const RELEASE_RETAINED_GUIDES = new Map([
  ['docs/migration-0.90.md', '0.90.x prints this URL in the useGoogleFonts deprecation warning'],
]);

// Repository paths of the docs/ pages a consumer-facing file links to.
// Markdown: inline and reference link destinations, relative or the page's
// GitHub URL on `main`. Site and package source: that GitHub URL outside
// comments. Prose that merely names a path is not a link.
function linkedDocs(path, text, blobUrl) {
  const found = new Set();
  const add = (target) => {
    const clean = target.split(/[?#]/)[0];
    if (/^docs\/.+\.md$/.test(clean)) found.add(clean);
  };
  const absolute = (target) => {
    if (target.startsWith(blobUrl)) add(posix.normalize(target.slice(blobUrl.length)));
  };
  if (path.endsWith('.md')) {
    const body = text.replace(/^```[\s\S]*?^```/gm, '');
    const targets = [
      ...[...body.matchAll(/\]\(<?([^)\s>]+)>?(?:\s+"[^"]*")?\)/g)].map((match) => match[1]),
      ...[...body.matchAll(/^ {0,3}\[[^\]]+\]:\s*<?(\S+?)>?(?:\s|$)/gm)].map((match) => match[1]),
    ];
    for (const target of targets) {
      if (/^[a-z][a-z0-9+.-]*:/i.test(target)) absolute(target);
      else add(posix.normalize(posix.join(posix.dirname(path), target)));
    }
    return found;
  }
  const code = text
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/(^|\s)\/\/.*$/gm, '$1');
  for (const [url] of code.matchAll(/https:\/\/github\.com\/[^\s"'`<>)]+?\.md(?![\w-]|\.\w)/g)) absolute(url);
  return found;
}

// Guides that no consumer surface links to and no supported release retains.
const guidesToDelete = (guides, linked, retained) =>
  guides.filter((guide) => !linked.has(guide) && !retained.has(guide));

test('a consumer link to a guide is a real link destination, not incidental text', () => {
  const blob = 'https://github.com/o/r/blob/main/';
  const links = (path, text) => [...linkedDocs(path, text, blob)];
  // Markdown: inline and reference destinations, relative or this repository's URL on main.
  assert.deepEqual(links('README.md', 'See the [guide](docs/./example.md#part).'), ['docs/example.md']);
  assert.deepEqual(links('packages/x/README.md', '[guide](../../docs/example.md)'), ['docs/example.md']);
  assert.deepEqual(links('README.md', '[guide](https://github.com/o/r/blob/main/docs/example.md)'), ['docs/example.md']);
  assert.deepEqual(links('README.md', 'Read the [guide][g].\n\n[g]: docs/example.md\n'), ['docs/example.md']);
  assert.deepEqual(links('README.md', 'The file `docs/example.md` is mentioned, [old](docs/example.md.old).'), []);
  assert.deepEqual(links('README.md', '[other](https://github.com/unrelated/project/blob/old/docs/example.md)'), []);
  assert.deepEqual(links('README.md', '[tag](https://github.com/o/r/blob/v1.0.0/docs/example.md)'), []);
  // Site and source: only the page's URL on main, outside comments.
  assert.deepEqual(
    links('site/src/pages/a.astro', '<a href="https://github.com/o/r/blob/main/docs/example.md">guide</a>'),
    ['docs/example.md'],
  );
  assert.deepEqual(links('site/src/pages/a.astro', '<p>See docs/example.md for details</p>'), []);
  assert.deepEqual(links('site/src/pages/a.astro', '<!-- https://github.com/o/r/blob/main/docs/example.md -->'), []);
  assert.deepEqual(
    links('packages/core/src/a.ts', "export const MESSAGE = 'See https://github.com/o/r/blob/main/docs/example.md.';"),
    ['docs/example.md'],
  );
  assert.deepEqual(
    links('packages/core/src/a.ts', '// https://github.com/o/r/blob/main/docs/example.md\nexport {};'),
    [],
  );
});

test('a guide is deleted only when nothing links to it and no supported release retains it', () => {
  const guides = ['docs/linked.md', 'docs/retained.md', 'docs/orphan.md'];
  assert.deepEqual(
    guidesToDelete(guides, new Set(['docs/linked.md']), new Map([['docs/retained.md', 'printed by 1.x']])),
    ['docs/orphan.md'],
  );
});

test('docs/ holds only procedures, the core beliefs, linked user guides and their images', () => {
  const files = tracked().filter((path) => path.startsWith('docs/'));
  const misplaced = files.filter(
    (path) =>
      path !== 'docs/README.md' &&
      path !== 'docs/design-docs/core-beliefs.md' &&
      !/^docs\/workflow\/[^/]+\.md$/.test(path) &&
      !/^docs\/[^/]+\.md$/.test(path) &&
      !/^docs\/images\/[^/]+\.png$/.test(path),
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
    for (const target of linkedDocs(surface, text, MAIN_BLOB_URL)) linked.add(target);
  }
  const guides = files.filter((path) => /^docs\/[^/]+\.md$/.test(path) && path !== 'docs/README.md');
  assert.deepEqual(
    guidesToDelete(guides, linked, RELEASE_RETAINED_GUIDES),
    [],
    'no README, site page or shipped source links this guide: delete it, or move design rationale beside the source',
  );
  assert.deepEqual(
    [...RELEASE_RETAINED_GUIDES.keys()].filter((guide) => !guides.includes(guide)),
    [],
    'remove the entry of a guide that no longer exists',
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
// command can run: those of the package a `--filter` names, otherwise those
// of the page's own working directory (the repository root, plus the package
// a package-level page belongs to).
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
  // Every workspace manifest, by package name and by `./path` filter.
  const byName = new Map();
  const byDir = new Map();
  for (const manifest of tracked().filter((path) => /(^|\/)package\.json$/.test(path))) {
    const { name, scripts = {} } = JSON.parse(read(join(ROOT, manifest)));
    const names = new Set(Object.keys(scripts));
    const dir = posix.dirname(manifest);
    byDir.set(dir, names);
    byName.set(`./${dir}`, names);
    if (name) byName.set(name, names);
  }
  const scriptsAt = (file) => {
    const names = new Set(byDir.get('.'));
    for (let dir = posix.dirname(file); dir !== '.'; dir = posix.dirname(dir)) {
      for (const name of byDir.get(dir) ?? []) names.add(name);
    }
    return names;
  };
  const missing = sources.flatMap((source) => {
    const file = rel(source);
    return missingReferences(file, read(source), {
      exists: (path) => existsSync(join(ROOT, path)),
      scriptsOf: (filter) => (filter ? (byName.get(filter) ?? new Set()) : scriptsAt(file)),
    });
  });
  assert.deepEqual(missing, [], 'the page names something that no longer exists: update or delete the statement');
});
