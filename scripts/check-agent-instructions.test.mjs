// Agent instruction files reach both Codex and Claude Code intact.
//
// Codex concatenates AGENTS.md from the repository root down to the working
// directory and stops reading at project_doc_max_bytes (32 KiB by default).
// Text past that limit silently vanishes from the agent's instructions, so
// every chain must fit. Repository skills live in .agents/skills (where Codex
// finds them), and Claude Code reaches each one through a link in
// .claude/skills.

import assert from 'node:assert/strict';
import { existsSync, readdirSync, realpathSync, statSync } from 'node:fs';
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

test('every repository skill is reachable from Claude Code', () => {
  const shared = join(ROOT, '.agents', 'skills');
  if (!existsSync(shared)) return;
  const missing = [];
  for (const name of readdirSync(shared)) {
    const link = join(ROOT, '.claude', 'skills', name);
    if (!existsSync(link) || realpathSync(link) !== realpathSync(join(shared, name))) missing.push(name);
  }
  assert.deepEqual(missing, [], 'link each .agents/skills/<name> as .claude/skills/<name> -> ../../.agents/skills/<name>');
});
