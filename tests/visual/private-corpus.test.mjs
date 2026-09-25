import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { PNG } from 'pngjs';
import { clearPrivateCandidateItemOutput, pngPixelsEqual } from './private-corpus.mjs';

test('private corpus self-VRT compares decoded pixels, not encoder bytes', () => {
  const image = new PNG({ width: 2, height: 1 });
  image.data.set([255, 0, 0, 255, 0, 0, 255, 255]);
  const fast = PNG.sync.write(image, { deflateLevel: 0 });
  const compact = PNG.sync.write(image, { deflateLevel: 9 });

  assert.equal(fast.equals(compact), false);
  assert.equal(pngPixelsEqual(fast, compact), true);
});

test('private corpus self-VRT rejects a one-channel pixel change', () => {
  const left = new PNG({ width: 1, height: 1 });
  left.data.set([1, 2, 3, 255]);
  const right = new PNG({ width: 1, height: 1 });
  right.data.set([1, 2, 4, 255]);

  assert.equal(pngPixelsEqual(PNG.sync.write(left), PNG.sync.write(right)), false);
});

test('candidate capture discards stale pages without following local evidence symlinks', () => {
  const root = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-private-vrt-output-'));
  try {
    const directory = join(root, 'docx', 'case');
    mkdirSync(directory, { recursive: true });
    writeFileSync(join(directory, 'page-29.png'), 'stale');
    writeFileSync(join(directory, 'notes.json'), 'retain');
    clearPrivateCandidateItemOutput({ stem: 'docx/case', itemKind: 'page', outputRoot: root });
    assert.equal(existsSync(join(directory, 'page-29.png')), false);
    assert.equal(readFileSync(join(directory, 'notes.json'), 'utf8'), 'retain');

    symlinkSync(directory, join(root, 'docx', 'linked'), 'dir');
    assert.throws(() => clearPrivateCandidateItemOutput({
      stem: 'docx/linked', itemKind: 'page', outputRoot: root,
    }), /symlinked private corpus output/);
    assert.equal(readFileSync(join(directory, 'notes.json'), 'utf8'), 'retain');
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
