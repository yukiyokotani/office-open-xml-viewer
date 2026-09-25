import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFileSync, mkdtempSync, writeFileSync, mkdirSync, symlinkSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { execFileSync } from 'node:child_process';

const source = readFileSync(new URL('./legacy-office-export.applescript', import.meta.url), 'utf8');
function sameFile(actual, expected) {
  const helper = source.match(/^on sameLocalFile\(actualPath, expectedPath\)\n[\s\S]*?^end sameLocalFile$/m)?.[0];
  assert.ok(helper, 'the export script must own the tested file-identity handler');
  const input = `${helper}\non run arguments\nreturn my sameLocalFile(item 1 of arguments, item 2 of arguments)\nend run\n`;
  return execFileSync('/usr/bin/osascript', ['-', actual, expected], {
    input, encoding: 'utf8', timeout: 10000, maxBuffer: 1024 * 1024,
  }).trim() === 'true';
}
function fixture(fn) {
  const directory = mkdtempSync(join(tmpdir(), 'office-export-identity-'));
  try {
    const original = join(directory, "input 'quoted' 日本語.ppt");
    writeFileSync(original, 'passive test bytes');
    const link = join(directory, 'alias.ppt');
    symlinkSync(original, link);
    fn({ directory, original, link });
  } finally {
    // Only this test's freshly created directory, never a sample directory.
    rmSync(directory, { recursive: true, force: true });
  }
}

test('PowerPoint file guard owns its cleanup reference only after verification', () => {
  const ppt = source.slice(source.indexOf('else if family is "ppt" then'));
  const tokens = [
    'set presentationRef to missing value',
    'set candidateRef to active presentation',
    'if not my sameLocalFile(full name of candidateRef, sourcePath) then error',
    'set presentationRef to candidateRef',
    'save presentationRef',
  ];
  const positions = tokens.map(token => ppt.indexOf(token));
  assert.ok(positions.every(p => p >= 0), 'missing fail-closed ownership step');
  assert.ok(positions.every((p, i) => i === 0 || positions[i - 1] < p));
  assert.doesNotMatch(ppt, /set presentationRef to active presentation/);
  assert.match(ppt, /if presentationRef is not missing value then\s+close presentationRef saving no/);
});

const mac = { skip: process.platform !== 'darwin' };
test('identical existing files match, including spaces and non-ASCII paths', mac, () => fixture(({ original }) => {
  assert.equal(sameFile(original, original), true);
}));
test('a symbolic-link spelling refers to the same file in both directions', mac, () => fixture(({ original, link }) => {
  assert.equal(sameFile(link, original), true);
  assert.equal(sameFile(original, link), true);
}));
test('equal content and equal basenames in different folders do not establish identity', mac, () => fixture(({ directory, original }) => {
  const other = join(directory, 'different'); mkdirSync(other);
  const path = join(other, "input 'quoted' 日本語.ppt"); writeFileSync(path, 'passive test bytes');
  assert.equal(sameFile(path, original), false);
}));
test('a missing actual file fails closed', mac, () => fixture(({ directory, original }) => {
  assert.equal(sameFile(join(directory, 'missing.ppt'), original), false);
}));
test('a missing expected file fails closed', mac, () => fixture(({ directory, original }) => {
  assert.equal(sameFile(original, join(directory, 'missing.ppt')), false);
}));
test('a dangling symbolic link fails closed', mac, () => fixture(({ directory, original }) => {
  const dangling = join(directory, 'dangling.ppt'); symlinkSync(join(directory, 'missing.ppt'), dangling);
  assert.equal(sameFile(dangling, original), false);
}));
