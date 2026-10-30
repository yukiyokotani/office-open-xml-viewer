import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, it } from 'vitest';
import { OoxmlError, type ModelSource } from '@silurus/ooxml-core';
import { MODEL_SOURCE_MODULE_PROTOCOL } from '@silurus/ooxml-core/internal/model-source';
import { resolveNodeSessionInput } from './model-source.js';

it('normalizes source open and lazy failures, preserving frozen methods and their receiver', async () => {
  const directory = mkdtempSync(join(tmpdir(), 'ooxml-source-rejection-'));
  const path = join(directory, 'source.mjs');
  writeFileSync(path, `
    const failure = 'MODEL_SOURCE_UNSUPPORTED:{"format":"custom","reason":"source-subset"}';
    export async function openModelSource(_bytes, config) {
      if (config.failOpen) throw failure;
      const archive = Object.freeze({ value: 7, read() { if (this.value !== 7) throw new Error('wrong receiver'); throw failure; } });
      return { archive, close() {} };
    }`);
  const source = (failOpen: boolean): ModelSource<'docx'> => ({
    target: 'docx', claim: () => true,
    beginLoad: () => ({ module: { protocol: MODEL_SOURCE_MODULE_PROTOCOL, target: 'docx',
      moduleUrl: pathToFileURL(path).href, config: { failOpen } }, release() {} }),
  });
  const open = (failOpen: boolean) => resolveNodeSessionInput(new Uint8Array(), 'docx',
    { sources: [source(failOpen)] }, value => value as { read(): void });
  try {
    await expect(open(true)).rejects.toBeInstanceOf(OoxmlError);
    const result = await open(false);
    expect(result.kind).toBe('model-source');
    if (result.kind !== 'model-source') throw new Error('source was not selected');
    expect(() => result.opened.archive.read()).toThrow(expect.objectContaining({ code: 'unsupported-content',
      details: { format: 'custom', reason: 'source-subset' } }));
    result.opened.close();
  } finally { rmSync(directory, { recursive: true, force: true }); }
});

it('asks the named sources in order and hands the claimed module the session resource policy', async () => {
  const directory = mkdtempSync(join(tmpdir(), 'ooxml-source-context-'));
  const path = join(directory, 'source.mjs');
  writeFileSync(path, `
    export async function openModelSource(_bytes, _config, _signal, _transfer, context) {
      return { archive: { context }, close() {} };
    }`);
  const order: string[] = [];
  const source = (name: string, claims: boolean): ModelSource<'docx'> => ({
    target: 'docx', claim: () => { order.push(name); return claims; },
    beginLoad: () => ({ module: { protocol: MODEL_SOURCE_MODULE_PROTOCOL, target: 'docx',
      moduleUrl: pathToFileURL(path).href, config: {} }, release() {} }),
  });
  try {
    const result = await resolveNodeSessionInput(new Uint8Array(), 'docx', {
      sources: [source('legacyDoc', false), undefined, source('odt', true)],
      resourceLimits: { maxArchiveEntryBytes: 1024, maxTotalInflatedBytes: null },
    }, value => value as { context: unknown });
    expect(order).toEqual(['legacyDoc', 'odt']);
    if (result.kind !== 'model-source') throw new Error('source was not selected');
    expect(result.opened.archive.context).toEqual({
      resourcePolicy: expect.objectContaining({ maxArchiveEntryBytes: 1024, maxTotalInflatedBytes: null }),
    });
    result.opened.close();
  } finally { rmSync(directory, { recursive: true, force: true }); }
});
