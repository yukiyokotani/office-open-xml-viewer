// Agent instruction files reach both Codex and Claude Code intact.
//
// Codex concatenates AGENTS.md from the repository root down to the working
// directory and stops reading at project_doc_max_bytes (32 KiB by default).
// Text past that limit silently vanishes from the agent's instructions, so
// every chain must fit. Procedures live in shared skills under .agents/skills
// (where Codex finds them); Claude Code reaches each one through a relative
// link in .claude/skills, so the link works in every checkout and worktree.

import assert from 'node:assert/strict';
import { existsSync, lstatSync, readFileSync, readdirSync, readlinkSync, statSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const CODEX_PROJECT_DOC_MAX_BYTES = 32 * 1024;
const SKIP = new Set(['node_modules', '.git', 'dist', 'target', 'spec', 'worktrees', 'private', 'wasm']);

function agentsFiles(dir) {
  const found = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (SKIP.has(entry.name)) continue;
    const path = join(dir, entry.name);
    if (entry.isDirectory()) found.push(...agentsFiles(path));
    else if (entry.name === 'AGENTS.md') found.push(path);
  }
  return found;
}

test('every AGENTS.md chain fits within the Codex instruction limit', () => {
  const over = [];
  for (const file of agentsFiles(ROOT)) {
    let bytes = 0;
    for (let dir = dirname(file); ; dir = dirname(dir)) {
      const candidate = join(dir, 'AGENTS.md');
      if (existsSync(candidate)) bytes += statSync(candidate).size;
      if (dir === ROOT) break;
    }
    if (bytes > CODEX_PROJECT_DOC_MAX_BYTES) over.push(`${relative(ROOT, file)}: ${bytes} bytes`);
  }
  assert.deepEqual(over, [], `move procedures out of AGENTS.md into skills or READMEs (limit ${CODEX_PROJECT_DOC_MAX_BYTES} bytes)`);
});

test('every repository skill is reachable from Codex and Claude Code', () => {
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
    const front = /^---\n([\s\S]*?)\n---\n/.exec(readFileSync(skill, 'utf8'));
    if (!front || !new RegExp(`^name: ${name}$`, 'm').test(front[1]) || !/^description: \S/m.test(front[1])) {
      problems.push(`${name}: SKILL.md front matter needs name: ${name} and a description`);
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
