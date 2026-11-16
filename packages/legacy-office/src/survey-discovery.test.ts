import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, expect, it } from 'vitest';
import { corpusFiles, referencePdf, surveyOutput } from '../tests/survey/survey.js';

// The Office-PDF survey is local-only; its input discovery decides which
// corpus files are compared at all, so it is pinned here on a synthetic tree.
const root = mkdtempSync(join(tmpdir(), 'legacy-survey-'));
afterAll(() => rmSync(root, { recursive: true, force: true }));

function file(path: string): void {
  mkdirSync(join(root, path, '..'), { recursive: true });
  writeFileSync(join(root, path), '');
}

it('finds inputs in subfolders and keeps the top-level selection rule', () => {
  for (const path of [
    'b.doc', 'A.DOC', '.dot.doc', '~$b.doc', 'b.pdf', 'note.txt',
    'gaps/c.doc', 'gaps/~$c.doc', 'gaps/deep/d.doc', '.hidden/e.doc',
  ]) file(path);
  expect(corpusFiles(root, 'doc')).toEqual(['.dot.doc', 'A.DOC', 'b.doc', 'gaps/c.doc', 'gaps/deep/d.doc']);
  expect(corpusFiles(root, 'doc', 'gaps/')).toEqual(['gaps/c.doc', 'gaps/deep/d.doc']);
  expect(corpusFiles(join(root, 'missing'), 'doc')).toEqual([]);
});

it('takes the PDF beside the input, else the highest numbered windows export folder', () => {
  for (const path of [
    'gaps/c.pdf', 'gaps/windows/c.pdf',
    'gaps/deep/windows/d.pdf', 'gaps/deep/windows-2/d.pdf', 'gaps/deep/windows-10/d.pdf', 'gaps/deep/other/d.pdf', 'gaps/deep/windows-old/d.pdf',
    'reference-mac/A.pdf',
  ]) file(path);
  expect(referencePdf(root, 'b.doc')).toBe(join(root, 'b.pdf'));
  expect(referencePdf(root, 'gaps/c.doc')).toBe(join(root, 'gaps/c.pdf'));
  expect(referencePdf(root, 'gaps/deep/d.doc')).toBe(join(root, 'gaps/deep/windows-10/d.pdf'));
  expect(referencePdf(root, 'A.DOC')).toBeUndefined();
});

it('gives every input its own output directory', () => {
  expect(surveyOutput('/out', 'a b.doc')).toBe('/out/a_b.doc');
  expect(surveyOutput('/out', 'a/b.doc')).toBe('/out/a/b.doc');
  expect(surveyOutput('/out', 'a_b.doc')).not.toBe(surveyOutput('/out', 'a/b.doc'));
  expect(surveyOutput('/out', 'a b/c d.doc')).toBe('/out/a b/c_d.doc');
  expect(surveyOutput('/out', 'a b/c.doc')).not.toBe(surveyOutput('/out', 'a_b/c.doc'));
});
