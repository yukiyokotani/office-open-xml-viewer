#!/usr/bin/env node
import { execFileSync, spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../', import.meta.url));
execFileSync(process.execPath, ['scripts/sync-companion-version.mjs'], { cwd: root, stdio: 'inherit' });
for (const directory of ['.', 'packages/legacy-office', 'packages/opendocument']) {
  const cwd = fileURLToPath(new URL(`${directory}/`, new URL('../', import.meta.url)));
  const { name, version } = JSON.parse(readFileSync(new URL(`${directory}/package.json`, new URL('../', import.meta.url)), 'utf8'));
  const query = spawnSync('npm', ['view', `${name}@${version}`, 'version', '--json'], { encoding: 'utf8' });
  if (query.status === 0) {
    if (JSON.parse(query.stdout) !== version) throw new Error(`Unexpected registry version for ${name}`);
    console.log(`Skipping already published ${name}@${version}`);
    continue;
  }
  // npm distinguishes an absent exact version from authentication, network,
  // or registry errors. Only absence permits attempting a new publication.
  let error;
  try { error = JSON.parse(query.stdout || query.stderr).error; } catch {}
  if (error?.code !== 'E404') throw new Error(`Cannot verify ${name}@${version}: ${query.stderr}`);
  execFileSync('pnpm', ['publish', '--no-git-checks'], { cwd, stdio: 'inherit' });
}
