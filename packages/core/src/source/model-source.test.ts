import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterAll, describe, expect, it, vi } from 'vitest';
import {
  beginModelSourceLoad,
  copyModelSourceBytes,
  MODEL_SOURCE_MODULE_PROTOCOL,
  openModelSourceModule,
  selectModelSource,
  validateModelSourceModuleDescriptor,
  type ModelSource,
  type ModelSourceModuleDescriptor,
} from './model-source';

const directory = mkdtempSync(join(tmpdir(), 'ooxml-model-source-'));
afterAll(() => rmSync(directory, { recursive: true, force: true }));

function moduleUrl(name: string, source: string): string {
  const path = join(directory, `${name}.mjs`);
  writeFileSync(path, source);
  return pathToFileURL(path).href;
}

function descriptor(url: string, config: Record<string, unknown> = {}): ModelSourceModuleDescriptor {
  return { protocol: MODEL_SOURCE_MODULE_PROTOCOL, target: 'docx', moduleUrl: url, config } as ModelSourceModuleDescriptor;
}

function source(target: 'docx' | 'xlsx' | 'pptx', claim: (bytes: Uint8Array) => unknown): ModelSource {
  return {
    target,
    claim: claim as (bytes: Uint8Array) => boolean,
    beginLoad: () => ({ module: descriptor('https://example.test/source.mjs'), release() {} }),
  };
}

describe('selectModelSource', () => {
  const bytes = new Uint8Array([1, 2, 3]);

  it('returns undefined without sources and picks the first source that claims the bytes', () => {
    expect(selectModelSource(undefined, 'docx', bytes)).toBeUndefined();
    const declined = source('docx', () => false);
    const first = source('docx', () => true);
    const second = source('docx', () => true);
    expect(selectModelSource([declined, first, second], 'docx', bytes)).toBe(first);
    expect(selectModelSource([declined], 'docx', bytes)).toBeUndefined();
  });

  it('fails closed on a source for another target, a non-boolean claim and a throwing claim', () => {
    expect(() => selectModelSource([source('xlsx', () => true)], 'docx', bytes)).toThrow(TypeError);
    expect(() => selectModelSource([source('docx', () => 1)], 'docx', bytes)).toThrow(TypeError);
    const failure = new Error('rejected by source');
    expect(() => selectModelSource([source('docx', () => { throw failure; })], 'docx', bytes))
      .toThrow(failure);
  });
});

describe('validateModelSourceModuleDescriptor', () => {
  it('detaches and freezes a valid descriptor', () => {
    const input = descriptor('https://example.test/m.mjs', { wasmUrl: 'https://example.test/a.wasm', limit: 4, flag: true, none: null });
    const result = validateModelSourceModuleDescriptor(input, 'docx');
    expect(result).toEqual(input);
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.config)).toBe(true);
    expect(result.config).not.toBe(input.config);
  });

  it.each([
    ['relative URL', descriptor('/source.mjs')],
    ['data URL', descriptor('data:text/javascript,export {}')],
    ['javascript URL', descriptor('javascript:alert(1)')],
    ['extra field', { ...descriptor('https://example.test/m.mjs'), extra: 1 }],
    ['wrong protocol', { ...descriptor('https://example.test/m.mjs'), protocol: 'v0' }],
    ['object config value', descriptor('https://example.test/m.mjs', { nested: {} })],
    ['non-finite number', descriptor('https://example.test/m.mjs', { n: Number.NaN })],
    ['bad key', descriptor('https://example.test/m.mjs', { 'bad key': 1 })],
  ])('rejects a descriptor with a %s', (_label, value) => {
    expect(() => validateModelSourceModuleDescriptor(value, 'docx')).toThrow();
  });

  it('rejects a descriptor for another target and oversized config', () => {
    expect(() => validateModelSourceModuleDescriptor(descriptor('https://example.test/m.mjs'), 'pptx'))
      .toThrow(TypeError);
    const many = Object.fromEntries(Array.from({ length: 33 }, (_, index) => [`k${index}`, index]));
    expect(() => validateModelSourceModuleDescriptor(descriptor('https://example.test/m.mjs', many)))
      .toThrow(RangeError);
  });
});

describe('beginModelSourceLoad', () => {
  it('releases once when the load descriptor is invalid', () => {
    const release = vi.fn();
    const invalid: ModelSource = {
      target: 'docx',
      claim: () => true,
      beginLoad: () => ({ module: descriptor('relative.mjs'), release }),
    };
    expect(() => beginModelSourceLoad(invalid, 'docx')).toThrow(TypeError);
    expect(release).toHaveBeenCalledTimes(1);
  });

  it('returns an idempotent release for a valid load', () => {
    const release = vi.fn();
    const valid: ModelSource = {
      target: 'docx',
      claim: () => true,
      beginLoad: () => ({ module: descriptor('https://example.test/m.mjs'), release }),
    };
    const admitted = beginModelSourceLoad(valid, 'docx');
    admitted.release();
    admitted.release();
    expect(release).toHaveBeenCalledTimes(1);
    expect(admitted.transfer).toEqual([]);
  });

  it('rejects unsupported, repeated and unbounded transfer lists and releases each failed admission', () => {
    const buffer = new ArrayBuffer(1);
    for (const transfer of [[{}], [buffer, buffer], Array.from({ length: 33 }, () => new ArrayBuffer(1))]) {
      const release = vi.fn();
      const selected: ModelSource = {
        target: 'docx', claim: () => true,
        beginLoad: () => ({ module: descriptor('https://example.test/m.mjs'), transfer: transfer as Transferable[], release }),
      };
      expect(() => beginModelSourceLoad(selected, 'docx')).toThrow();
      expect(release).toHaveBeenCalledOnce();
    }
  });
});

it('copies exactly the bytes returned by a model source into a detachable buffer', () => {
  const backing = new Uint8Array([99, 1, 2, 3, 88]);
  const result = copyModelSourceBytes(backing.subarray(1, 4));
  expect(new Uint8Array(result)).toEqual(new Uint8Array([1, 2, 3]));
  expect(result).not.toBe(backing.buffer);
  structuredClone(result, { transfer: [result] });
  expect(result.byteLength).toBe(0);
  expect(backing).toEqual(new Uint8Array([99, 1, 2, 3, 88]));
});

describe('openModelSourceModule', () => {
  const archiveOf = (value: unknown) => {
    if (typeof value !== 'object' || value === null || !('kind' in value)) {
      throw new TypeError('archive rejected');
    }
    return value as { kind: string };
  };

  it('imports the module, passes the frozen config and validates the result', async () => {
    const url = moduleUrl('ok', `
      export async function openModelSource(bytes, config) {
        return { archive: { kind: config.kind, length: bytes.length }, viewDefaults: { showTrackedChanges: true }, close() {} };
      }`);
    const opened = await openModelSourceModule(descriptor(url, { kind: 'fake' }), new Uint8Array(4), archiveOf);
    expect(opened.archive).toEqual({ kind: 'fake', length: 4 });
    expect(opened.viewDefaults).toEqual({ showTrackedChanges: true });
  });

  it('hands the module a detached, frozen copy of the caller policy context', async () => {
    const url = moduleUrl('context', `
      export async function openModelSource(_bytes, _config, _signal, _transfer, context) {
        return { archive: { kind: 'x', context, frozen: Object.isFrozen(context) && Object.isFrozen(context.resourcePolicy) }, close() {} };
      }`);
    const resourcePolicy = { maxArchiveEntryBytes: 10, maxTotalInflatedBytes: null, maxArchiveEntries: 3, extra: 'dropped' };
    const opened = await openModelSourceModule(
      descriptor(url, {}), new Uint8Array(), (value) => value as { context: unknown; frozen: boolean },
      undefined, [], { resourcePolicy },
    );
    expect(opened.archive.frozen).toBe(true);
    expect(opened.archive.context).toEqual({
      resourcePolicy: { maxArchiveEntryBytes: 10, maxTotalInflatedBytes: null, maxArchiveEntries: 3 },
    });
    await expect(openModelSourceModule(
      descriptor(url, {}), new Uint8Array(), (value) => value,
      undefined, [], { resourcePolicy: { ...resourcePolicy, maxArchiveEntries: -1 } },
    )).rejects.toThrow(TypeError);
  });

  it('closes what the module opened when the archive or view defaults are invalid', async () => {
    const url = moduleUrl('invalid', `
      export let closed = 0;
      export async function openModelSource(_bytes, config) {
        return {
          archive: config.valid ? { kind: 'x' } : {},
          viewDefaults: config.valid ? { showTrackedChanges: 'yes' } : undefined,
          close() { closed += 1; },
        };
      }`);
    await expect(openModelSourceModule(descriptor(url, { valid: false }), new Uint8Array(), archiveOf))
      .rejects.toThrow('archive rejected');
    await expect(openModelSourceModule(descriptor(url, { valid: true }), new Uint8Array(), archiveOf))
      .rejects.toThrow(TypeError);
    const loaded = await import(/* @vite-ignore */ url) as { closed: number };
    expect(loaded.closed).toBe(2);
  });

  it('rejects a module without openModelSource and an aborted load', async () => {
    const url = moduleUrl('empty', 'export const unrelated = 1;');
    await expect(openModelSourceModule(descriptor(url), new Uint8Array(), archiveOf))
      .rejects.toThrow('openModelSource');
    const controller = new AbortController();
    controller.abort();
    await expect(openModelSourceModule(descriptor(url), new Uint8Array(), archiveOf, controller.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
  });
});
