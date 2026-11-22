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

// A fenced code block and its body. The fence may be indented, as it is under
// a list item.
const FENCED_BLOCK = /^[ \t]*```[^\n]*\n([\s\S]*?)^[ \t]*```[ \t]*$/gm;

// Link destinations a Markdown reader would follow: inline links, and
// reference links that are used and defined. Fenced code, code spans and HTML
// comments hold no links, and an unused definition is not one.
function markdownLinks(text) {
  const body = text
    .replace(FENCED_BLOCK, '')
    // An indented code block: lines indented four or more columns at the
    // start of the file or after a blank line. Under a list item, whose
    // continuation text is itself indented, six or more columns.
    .replace(/(^|^\S[^\n]*\n(?:[ \t]*\n)+)((?:(?: {4,}|\t)[^\n]*(?:\n|$)|[ \t]*\n)+)/gm, (block, lead, code, offset) => {
      if (lead === '' && offset !== 0) return block;
      if (!/^(?:[-*+]|\d+[.)])\s/.test(lead)) return lead;
      return /^(?: {6,}|\t\t)/.test(code) ? lead : block;
    })
    .replace(/<!--[\s\S]*?-->/g, '')
    // A code span may wrap across lines.
    .replace(/`[^`]*`/g, '');
  const label = (raw) => raw.trim().replace(/\s+/g, ' ').toLowerCase();
  const definitions = new Map();
  const prose = body.replace(/^ {0,3}\[([^\]]+)\]:[ \t]*<?(\S+?)>?(?:[ \t]+.*)?$/gm, (_, name, target) => {
    if (!definitions.has(label(name))) definitions.set(label(name), target);
    return '';
  });
  const targets = [
    ...[...prose.matchAll(/\[[^\]]*\]\(<?([^)\s>]+)>?(?:\s+"[^"]*")?\)/g)].map((match) => match[1]),
    // Raw HTML: <img src="…">, <a href="…">.
    ...[...prose.matchAll(/<(?:img|a|source|video)\b[^>]*?(?<![\w-])(?:src|href)=["']([^"']+)["']/gi)].map((match) => match[1]),
    // Autolinks: <https://…>.
    ...[...prose.matchAll(/<([a-z][a-z0-9+.-]*:\/\/[^\s<>]+)>/gi)].map((match) => match[1]),
  ];
  // [text][label], [label][] and the shortcut form [label].
  for (const [, text, explicit] of prose.matchAll(/\[([^\]]+)\](?:\[([^\]]*)\])?(?![(:])/g)) {
    const target = definitions.get(label(explicit || text));
    if (target !== undefined) targets.push(target);
  }
  return targets;
}

// Relative link targets of a Markdown file, without anchors.
function relativeLinks(file) {
  return markdownLinks(read(file))
    .map((target) => target.split('#')[0])
    .filter((target) => target && !/^[a-z][a-z0-9+.-]*:/i.test(target));
}

// The anchors GitHub generates for a Markdown file's headings.
function headingAnchors(text) {
  const anchors = new Set();
  for (const [, title] of text.replace(FENCED_BLOCK, '').matchAll(/^#{1,6}[ \t]+(.+?)[ \t]*#*$/gm)) {
    const base = title
      .replace(/`/g, '')
      .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
      .toLowerCase()
      .replace(/[^\p{L}\p{N} _-]/gu, '')
      .replace(/ /g, '-');
    let anchor = base;
    for (let n = 1; anchors.has(anchor); n += 1) anchor = `${base}-${n}`;
    anchors.add(anchor);
  }
  return anchors;
}

// Relative links of a Markdown file whose `#fragment` names no heading of the
// Markdown file they point to (or of the file itself, for `#fragment` alone).
function brokenFragments(file) {
  const broken = [];
  for (const target of markdownLinks(read(file))) {
    if (/^[a-z][a-z0-9+.-]*:/i.test(target) || !target.includes('#')) continue;
    const [path, fragment] = target.split('#');
    const destination = path ? resolve(dirname(file), path) : file;
    if (!fragment || !destination.endsWith('.md') || !existsSync(destination)) continue;
    if (!headingAnchors(read(destination)).has(fragment.toLowerCase())) broken.push(`${rel(file)} -> ${target}`);
  }
  return broken;
}

test('heading anchors follow the generated form', () => {
  assert.deepEqual(
    [...headingAnchors('# Top\n\n## Keeping `docs/` current\n\n### Gardening pass\n\n## Git, PRs and Issues\n\n## Top\n')],
    ['top', 'keeping-docs-current', 'gardening-pass', 'git-prs-and-issues', 'top-1'],
  );
});

test('Markdown links are the destinations a reader can follow', () => {
  assert.deepEqual(markdownLinks('[a](x.md) and [b][ref] and [ref] and [c][].\n\n[ref]: y.md\n[c]: z.md "t"\n'), [
    'x.md', 'y.md', 'y.md', 'z.md',
  ]);
  assert.deepEqual(markdownLinks('<img src="docs/images/icon.png" width="80" alt="icon" />'), ['docs/images/icon.png']);
  assert.deepEqual(markdownLinks('See <https://example.org/docs/a.md> and <span>text</span>.'), ['https://example.org/docs/a.md']);
  assert.deepEqual(markdownLinks('Example:\n\n    [unused](docs/orphan.md)\n\nAfter [a](a.md).\n'), ['a.md']);
  assert.deepEqual(markdownLinks('- item\n\n    continued [b](b.md)\n'), ['b.md']);
  assert.deepEqual(markdownLinks('    [start](docs/orphan.md)\n\nText [c](c.md).\n'), ['c.md']);
  assert.deepEqual(markdownLinks('- item\n\n      [code](docs/orphan.md)\n'), []);
  assert.deepEqual(markdownLinks('<a data-href="docs/orphan.md">x</a> <a href="d.md">d</a>'), ['d.md']);
  assert.deepEqual(markdownLinks('Prose ](docs/missing.md) and ![alt](a.png).'), ['a.png']);
  assert.deepEqual(markdownLinks('Example: `wrapped\n[fake](docs/missing.md)` end.'), []);
  assert.deepEqual(markdownLinks('<!-- [hidden](h.md) -->\n`[code](c.md)`\n\n[unused]: u.md\n'), []);
  assert.deepEqual(markdownLinks('```\n[fenced](f.md)\n```\n[guide][g]\n\n[g]: missing.md\n'), ['missing.md']);
  assert.deepEqual(markdownLinks('1. Run:\n\n   ```bash\n   echo "[guide](docs/example.md)"\n   ```\n'), []);
});

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
    ...tracked().filter((path) => /(^|\/)README\.md$/.test(path) && !path.startsWith('docs/')).map((path) => join(ROOT, path)),
  ];
  const broken = [];
  for (const file of sources) {
    for (const target of relativeLinks(file)) {
      if (!existsSync(resolve(dirname(file), target))) broken.push(`${rel(file)} -> ${target}`);
    }
    broken.push(...brokenFragments(file));
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

// Source text without block, HTML and line comments.
const withoutComments = (text) =>
  text
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/(^|\s)\/\/.*$/gm, '$1');

// Shipped source of a package or the site: not a test, story, bench or
// test-support module.
const isShippedSource = (path) => !/(\.(test|spec|stories|test-support|bench)|-test-dom)\.[cm]?[jt]sx?$/.test(path);

// Repository paths of the docs/ pages a consumer-facing file links to.
// Markdown: inline and reference link destinations, relative or the page's
// GitHub URL on `main`. Site and package source: that GitHub URL outside
// comments. Prose that merely names a path is not a link.
function linkedDocs(path, text, blobUrl, renderedProperties = new Set()) {
  const found = new Set();
  const add = (target) => {
    const clean = target.split(/[?#]/)[0];
    if (/^docs\/.+\.md$/.test(clean)) found.add(clean);
  };
  const absolute = (target) => {
    if (target.startsWith(blobUrl)) add(posix.normalize(target.slice(blobUrl.length)));
  };
  if (path.endsWith('.md')) {
    for (const target of markdownLinks(text)) {
      if (/^[a-z][a-z0-9+.-]*:/i.test(target)) absolute(target);
      else add(posix.normalize(posix.join(posix.dirname(path), target)));
    }
    return found;
  }
  const code = withoutComments(text);
  const url = String.raw`(https:\/\/github\.com\/[^\s"'\`<>)]+?\.md)(?![\w-]|\.\w)`;
  if (path.startsWith('site/')) {
    // A site file links a page where the URL is rendered as a link: an `href`
    // attribute, a constant an `href={name}` of the same file uses, or a data
    // property that some site component renders as `href={item.property}`
    // (`renderedProperties`). A URL in prose or in an unused constant is not
    // counted.
    for (const [, target] of code.matchAll(new RegExp(String.raw`(?<![\w-])href=["']` + url, 'g'))) absolute(target);
    for (const [, name, target] of code.matchAll(new RegExp(String.raw`\b(\w+)\s*:\s*["'\`]` + url, 'g'))) {
      if (renderedProperties.has(name)) absolute(target);
    }
    for (const [, name, target] of code.matchAll(new RegExp(String.raw`\bconst\s+(\w+)\s*=\s*["'\`]` + url, 'g'))) {
      if (new RegExp(String.raw`(?<![\w-])href=\{\s*${name}\s*\}`).test(code)) absolute(target);
    }
    return found;
  }
  // Package source: a URL in a string the shipped code can print.
  for (const [, target] of code.matchAll(new RegExp(url, 'g'))) absolute(target);
  return found;
}

// Names of the data properties a site component renders as a link target:
// `href={item.property}`.
const renderedLinkProperties = (text) =>
  new Set([...withoutComments(text).matchAll(/(?<![\w-])href=\{\s*[\w$]+(?:\?)?\.([\w$]+)\s*\}/g)].map((match) => match[1]));

// A repository path without its source extension, as a relative import names it.
const moduleName = (path) => path.replace(/\.[cm]?[jt]sx?$|\.astro$/, '');
// Repository paths (without extension) of the modules `file` imports by a
// relative specifier. Package imports name no file of the site.
const importedModuleNames = (file, text) =>
  new Set(
    [...withoutComments(text).matchAll(/\bfrom\s+["'](\.[^"']+)["']/g)].map((match) =>
      moduleName(posix.normalize(posix.join(posix.dirname(file), match[1]))),
    ),
  );

// Guides that no consumer surface links to and no supported release retains.
const guidesToDelete = (guides, linked, retained) =>
  guides.filter((guide) => !linked.has(guide) && !retained.has(guide));

test('a consumer link to a guide is a real link destination, not incidental text', () => {
  const blob = 'https://github.com/o/r/blob/main/';
  const links = (path, text, rendered) => [...linkedDocs(path, text, blob, rendered)];
  // Markdown: inline and reference destinations, relative or this repository's URL on main.
  assert.deepEqual(links('README.md', 'See the [guide](docs/./example.md#part).'), ['docs/example.md']);
  assert.deepEqual(links('packages/x/README.md', '[guide](../../docs/example.md)'), ['docs/example.md']);
  assert.deepEqual(links('README.md', '[guide](https://github.com/o/r/blob/main/docs/example.md)'), ['docs/example.md']);
  assert.deepEqual(links('README.md', 'Read the [guide][g].\n\n[g]: docs/example.md\n'), ['docs/example.md']);
  assert.deepEqual(links('README.md', 'The file `docs/example.md` is mentioned, [old](docs/example.md.old).'), []);
  assert.deepEqual(links('README.md', '[other](https://github.com/unrelated/project/blob/old/docs/example.md)'), []);
  assert.deepEqual(links('README.md', '<!-- [unused](docs/hidden.md) -->\n\n[unused]: docs/orphan.md\n'), []);
  assert.deepEqual(links('README.md', '[tag](https://github.com/o/r/blob/v1.0.0/docs/example.md)'), []);
  // Site and source: only the page's URL on main, outside comments.
  assert.deepEqual(
    links('site/src/pages/a.astro', '<a href="https://github.com/o/r/blob/main/docs/example.md">guide</a>'),
    ['docs/example.md'],
  );
  assert.deepEqual(links('site/src/pages/a.astro', '<p>See docs/example.md for details</p>'), []);
  assert.deepEqual(
    links('site/src/pages/a.astro', '<button data-href="https://github.com/o/r/blob/main/docs/example.md">x</button>'),
    [],
  );
  const constant = "const guide = 'https://github.com/o/r/blob/main/docs/example.md';\n";
  assert.deepEqual(links('site/src/pages/a.astro', `${constant}<a href={guide}>guide</a>`), ['docs/example.md']);
  assert.deepEqual(links('site/src/pages/a.astro', `${constant}<p>no link</p>`), []);
  const entry = "export const entry = { detailsHref: 'https://github.com/o/r/blob/main/docs/example.md' };";
  assert.deepEqual(links('site/src/lib/a.ts', entry), []);
  assert.deepEqual(links('site/src/lib/a.ts', entry, new Set(['detailsHref'])), ['docs/example.md']);
  assert.deepEqual(
    [...renderedLinkProperties('<a class="x" href={o.detailsHref}>d</a> <a href={guide}>g</a> <a href={item?.url}>u</a>')],
    ['detailsHref', 'url'],
  );
  assert.deepEqual([...renderedLinkProperties('<!-- <a href={o.detailsHref}>d</a> -->')], []);
  assert.deepEqual(
    [
      ...importedModuleNames(
        'site/src/components/A.astro',
        "import { A } from '../lib/api-reference';\nimport b from './b.js';\nimport { D } from '@scope/pkg/docx';\n// import c from './c';",
      ),
    ],
    ['site/src/lib/api-reference', 'site/src/components/b'],
  );
  assert.equal(moduleName('site/src/lib/api-reference.ts'), 'site/src/lib/api-reference');
  assert.equal(moduleName('site/src/pages/api/docx.astro'), 'site/src/pages/api/docx');
  assert.equal(isShippedSource('packages/docx/src/test-support/tab-fitting.test-support.ts'), false);
  assert.equal(isShippedSource('packages/docx/src/a.test.ts'), false);
  assert.equal(isShippedSource('packages/xlsx/src/viewer-destroy-test-dom.ts'), false);
  assert.equal(isShippedSource('packages/core/src/fonts/deprecations.ts'), true);
  assert.deepEqual(links('README.md', 'See <https://github.com/o/r/blob/main/docs/example.md>.'), ['docs/example.md']);
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

// Linked docs/ pages that are not in the repository.
const staleGuideLinks = (linked, existing) => [...linked].filter((target) => !existing.has(target)).sort();

test('a consumer link to a deleted guide is reported', () => {
  assert.deepEqual(staleGuideLinks(new Set(['docs/kept.md', 'docs/deleted.md']), new Set(['docs/kept.md'])), ['docs/deleted.md']);
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
      (/^site\/src\//.test(path) && isShippedSource(path)) ||
      (/^(packages\/[^/]+\/)?src\/.*\.[cm]?[jt]s$/.test(path) && isShippedSource(path)),
  );
  const texts = new Map(surfaces.map((surface) => [surface, read(join(ROOT, surface))]));
  const siteTexts = [...texts].filter(([surface]) => surface.startsWith('site/'));
  const rendered = new Set(siteTexts.flatMap(([, text]) => [...renderedLinkProperties(text)]));
  // Data properties count only in a module that another site file imports.
  const imported = new Set(siteTexts.flatMap(([surface, text]) => [...importedModuleNames(surface, text)]));
  const linked = new Set();
  for (const [surface, text] of texts) {
    if (!text.includes('docs/')) continue;
    const properties = imported.has(moduleName(surface)) ? rendered : new Set();
    for (const target of linkedDocs(surface, text, MAIN_BLOB_URL, properties)) linked.add(target);
  }
  const guides = files.filter((path) => /^docs\/[^/]+\.md$/.test(path) && path !== 'docs/README.md');
  assert.deepEqual(
    staleGuideLinks(linked, new Set(files)),
    [],
    'a README, site page or shipped source links a docs/ page that does not exist',
  );
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
    const lines = read(file).replace(FENCED_BLOCK, '').split('\n');
    if (lines.some((line) => /^\s*[-*] \[[ xX]\] /.test(line))) found.push(rel(file));
  }
  assert.deepEqual(found, [], 'plans and trackers are GitHub issues, not documents');
});

// What a page names that the repository does not have: a backticked repository
// path, or a pnpm script. `scriptsOf(filter)` returns the script names a
// command can run: those of the package a `--filter` names, otherwise those
// of the page's own working directory (the repository root, plus the package
// a package-level page belongs to).
function missingReferences(file, text, { exists, hasFileNamed, scriptsOf }) {
  const pnpmBuiltins = new Set(['install', 'exec', 'add', 'run', 'dlx', 'why', 'pack', 'publish']);
  // Generated, git-ignored or local-only locations a page may legitimately name.
  const localOnly = [/^(\.\.\/)*spec(\/|$)/, /(^|\/)src\/wasm(\/|$)/, /(^|\/)public\/private(\/|$)/, /^node_modules\//, /(^|\/)dist(\/|$)/, /^\/?tmp\//];
  // Part names inside an OOXML or ODF package, not repository paths.
  const packageParts = /^(word|xl|ppt|docProps|_rels|customXml|META-INF)\//;
  const sourceFile = /\.(md|mjs|cjs|js|ts|tsx|json|ya?ml|rs|py|sh|ps1|toml|lock|html|astro|applescript)$/;
  const missing = [];
  // Words of inline code and of fenced examples: a command names its script as
  // one of its words (`node scripts/x.mjs`).
  const words = [
    ...[...text.matchAll(FENCED_BLOCK)].map((match) => match[1]),
    // An inline code span may wrap across lines.
    ...[...text.replace(FENCED_BLOCK, '').matchAll(/`([^`]+)`/g)].map((match) => match[1]),
  ]
    .flatMap((code) => code.split(/\s+/))
    .map((word) => word.replace(/^["'(]+|["'),;:]+$/g, ''))
    // `packages/{docx,xlsx}/package.json` names one path per alternative.
    .flatMap((word) => {
      const list = /^([^{}]*)\{([\w@.-]+(?:,[\w@.-]+)+)\}([^{}]*)$/.exec(word);
      return list ? list[2].split(',').map((item) => `${list[1]}${item}${list[3]}`) : [word];
    });
  for (const span of new Set(words)) {
    // A bare file name (`sgconfig.yml`): some tracked file must carry it.
    // Files of the git-ignored private corpus are not tracked by design.
    const privateCorpusFile = /\.local\.|^reference-windows-manifest\.json$/;
    if (/^[\w@-][\w@.-]*$/.test(span) && sourceFile.test(span) && !privateCorpusFile.test(span)) {
      if (!hasFileNamed(span)) missing.push(`${file}: ${span}`);
      continue;
    }
    // A repository path: at least one slash, no placeholders, globs or assignments.
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
  for (const [, filter, name] of text.matchAll(/\bpnpm\s+(?:--filter\s+(\S+)\s+)?(?:run\s+)?([a-z][\w:.-]*)/g)) {
    if (pnpmBuiltins.has(name)) continue;
    // `<fmt>` stands for each format package; any other placeholder names no
    // single package.
    const filters = !filter
      ? [undefined]
      : filter.includes('<fmt>')
        ? ['docx', 'xlsx', 'pptx'].map((format) => filter.replace('<fmt>', format))
        : [/[<*{]/.test(filter) ? undefined : filter];
    for (const scoped of filters) {
      if (!scriptsOf(scoped).has(name)) missing.push(`${file}: pnpm ${scoped ? `--filter ${scoped} ` : ''}${name}`);
    }
  }
  return missing;
}

test('a removed path or a script of another package is reported', () => {
  const present = new Set(['scripts', 'scripts/a.mjs', 'packages', 'packages/docx']);
  const scripts = {
    '@scope/docx': new Set(['conformance', 'vrt']),
    '@scope/xlsx': new Set(['vrt']),
    '@scope/pptx': new Set(['vrt']),
  };
  const all = new Set(Object.values(scripts).flatMap((names) => [...names]));
  const context = {
    exists: (path) => present.has(path.replace(/\/$/, '')),
    hasFileNamed: (name) => [...present].some((path) => path.split('/').at(-1) === name),
    scriptsOf: (filter) => (filter ? (scripts[filter] ?? new Set()) : all),
  };
  const check = (text) => missingReferences('AGENTS.md', text, context);
  assert.deepEqual(check('`scripts/a.mjs`, `packages/docx/`, `lIns/tIns`, `word/fontTable.xml`, `spec/part1`'), []);
  assert.deepEqual(check('Run `node scripts/missing.mjs --flag`.\n\n```bash\nnode scripts/a.mjs && python3 scripts/absent.py\n```\n'), [
    'AGENTS.md: scripts/absent.py',
    'AGENTS.md: scripts/missing.mjs',
  ]);
  assert.deepEqual(check('4. Run the gates:\n\n   ```bash\n   mise exec -- node scripts/removed.mjs --final\n   ```\n'), [
    'AGENTS.md: scripts/removed.mjs',
  ]);
  assert.deepEqual(check('`a.mjs`, `notes.local.md`, `packages/{docx}/x`, `Array.isArray`'), []);
  assert.deepEqual(check('`Cargo.lock`'), ['AGENTS.md: Cargo.lock']);
  assert.deepEqual(check('`sgconfig.yml` and `packages/{docx,xlsx}/package.json`'), [
    'AGENTS.md: sgconfig.yml',
    'AGENTS.md: packages/docx/package.json',
    'AGENTS.md: packages/xlsx/package.json',
  ]);
  assert.deepEqual(check('`scripts/gone.mjs` and `rules/example.yml`'), [
    'AGENTS.md: scripts/gone.mjs',
    'AGENTS.md: rules/example.yml',
  ]);
  assert.deepEqual(check('`pnpm vrt`, `pnpm --filter @scope/docx conformance`, `pnpm --filter @scope/<fmt> vrt`, `pnpm --filter @scope/<any> conformance`, `pnpm install`'), []);
  assert.deepEqual(check('`pnpm --filter @scope/<fmt> conformance`'), [
    'AGENTS.md: pnpm --filter @scope/xlsx conformance',
    'AGENTS.md: pnpm --filter @scope/pptx conformance',
  ]);
  assert.deepEqual(check('Rebuild (`pnpm\ngone-script` or `node\nscripts/gone.mjs`).'), [
    'AGENTS.md: scripts/gone.mjs',
    'AGENTS.md: pnpm gone-script',
  ]);
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
  // Members of the root pnpm workspace (`pnpm-workspace.yaml`: packages/*, site).
  for (const manifest of tracked().filter((path) => /^(packages\/[^/]+\/|site\/)?package\.json$/.test(path))) {
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
  const basenames = new Set(tracked().map((path) => path.split('/').at(-1)));
  const missing = sources.flatMap((source) => {
    const file = rel(source);
    return missingReferences(file, read(source), {
      exists: (path) => existsSync(join(ROOT, path)),
      hasFileNamed: (name) => basenames.has(name),
      scriptsOf: (filter) => (filter ? (byName.get(filter) ?? new Set()) : scriptsAt(file)),
    });
  });
  assert.deepEqual(missing, [], 'the page names something that no longer exists: update or delete the statement');
});
