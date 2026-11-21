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

import assert from 'node:assert/strict';
import { existsSync, lstatSync, readFileSync, readdirSync, readlinkSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
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
const rel = (path) => relative(ROOT, path);

// Relative Markdown link targets outside fenced code blocks, without anchors.
function relativeLinks(file) {
  const text = readFileSync(file, 'utf8').replace(/^```[\s\S]*?^```/gm, '');
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
    const lines = readFileSync(file, 'utf8').split('\n').length - 1;
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
    const text = readFileSync(skill, 'utf8');
    const front = /^---\n([\s\S]*?)\n---\n/.exec(text);
    if (!front || !new RegExp(`^name: ${name}$`, 'm').test(front[1]) || !/^description: \S/m.test(front[1])) {
      problems.push(`${name}: SKILL.md front matter needs name: ${name} and a description`);
    }
    const body = front ? text.slice(front[0].length) : text;
    if (body.trim().split('\n').length > SKILL_BODY_MAX_LINES) {
      problems.push(`${name}: keep the procedure in docs/workflow/ and only point to it from the skill`);
    }
    if (!relativeLinks(skill).some((target) => resolve(dirname(skill), target).startsWith(join(ROOT, 'docs') + '/'))) {
      problems.push(`${name}: SKILL.md must link the docs/ page it loads`);
    }
    const link = join(links, name);
    const target = `../../.agents/skills/${name}`;
    if (!existsSync(link) || !lstatSync(link).isSymbolicLink() || readlinkSync(link) !== target) {
      problems.push(`${name}: link .claude/skills/${name} -> ${target}`);
    }
  }
  for (const name of existsSync(links) ? readdirSync(links) : []) {
    if (!names.includes(name)) problems.push(`.claude/skills/${name}: no matching .agents/skills/${name}`);
  }
  assert.deepEqual(problems, []);
});
