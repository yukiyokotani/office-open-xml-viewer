import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PNG } from 'pngjs';
import { listPrivateCorpus, pngPixelsEqual } from './private-corpus.mjs';

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

test('private corpus listing covers the format folder and the top level', () => {
  const root = mkdtempSync(join(tmpdir(), 'private-corpus-'));
  try {
    mkdirSync(join(root, 'docx'));
    mkdirSync(join(root, 'doc'));
    for (const file of ['sample-10.docx', 'sample-2.docx', '~$lock.docx', 'notes.md']) {
      writeFileSync(join(root, 'docx', file), '');
    }
    writeFileSync(join(root, 'doc', 'sample-3.docx'), '');
    writeFileSync(join(root, 'top.docx'), '');
    assert.deepEqual(listPrivateCorpus('docx', root), [
      'docx/sample-2.docx',
      'docx/sample-10.docx',
      'top.docx',
    ]);
    assert.deepEqual(listPrivateCorpus('docx', join(root, 'missing')), []);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
