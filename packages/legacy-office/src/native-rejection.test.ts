import { expect, it, vi } from 'vitest';
import { parseTypedParserError } from '@silurus/ooxml-core/worker';
import { createDirectSourceRuntime } from '@silurus/ooxml-core/internal/model-source-runtime';
import { nativeRejection } from './native-rejection.js';

it('maps native constructor and lazy failures, never initialization/network errors', async () => {
  const format = 'xls';
  const create = (construct: () => { free(): void; read(): void }, initialize = async () => {}) => createDirectSourceRuntime({
    label: 'test', mapNativeError: (error) => nativeRejection(error, format), loadGlue: async () => ({ default: initialize }),
    resolveWasm: async () => new Uint8Array(), construct, closeNative: () => {},
  });
  const check = (error: unknown) => expect(parseTypedParserError(error)).toMatchObject({ code: 'unsupported-content',
    details: { format, reason: 'source-subset' } });
  const open = (runtime: ReturnType<typeof create>) => runtime.open(new Uint8Array(), 'https://example.test/reader.wasm');
  const failure = 'UNSUPPORTED:field containing PRIVATE document text';
  const constructor = await open(create(() => { throw failure; })).catch(e => e);
  check(constructor);
  expect(String(constructor)).not.toContain('PRIVATE');
  const free = vi.fn();
  const source = await open(create(() => ({ free, read() { throw failure; } })));
  try { source.archive.read(); throw new Error('expected rejection'); } catch (error) { check(error); }
  source.closeArchive(); source.closeArchive();
  expect(free).toHaveBeenCalledOnce();
  const network = new Error(failure);
  await expect(open(create(() => { throw new Error('unreachable'); }, async () => { throw network; }))).rejects.toBe(network);
});

it('keeps encryption, native budgets, cursor errors and Rust resource envelopes separate', () => {
  expect(parseTypedParserError(nativeRejection('LEGACY_UNSUPPORTED_ENCRYPTION:private', 'doc')))
    .toMatchObject({ code: 'unsupported-encryption' });
  for (const error of ['LEGACY_RESOURCE_LIMIT:budget', 'OUTPUT_TOO_LARGE', 'document cursor is not active', 'OOXML_RESOURCE_LIMIT:{}']) {
    const mapped = nativeRejection(error, 'doc');
    expect(mapped).toBeInstanceOf(Error);
    expect(parseTypedParserError(mapped)).toBeUndefined();
  }
  const trap = new WebAssembly.RuntimeError('trap');
  expect(nativeRejection(trap, 'doc')).toBe(trap);
});
