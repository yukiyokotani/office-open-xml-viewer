import { describe, expect, it, vi } from 'vitest';
import { createDirectSourceRuntime, resolveDirectWasmInput } from './direct-source-runtime.js';

class Archive {
  free = vi.fn();
  close = vi.fn();
  call = vi.fn<() => string>(() => 'ok');
  __destroy_into_raw = vi.fn(() => 1);
}

const descriptor = { wasmUrl: 'https://example.test/runtime.wasm' };

function engine(glue: object, construct: () => Archive) {
  return createDirectSourceRuntime({
    label: 'test', maximumSourceBytes: 100,
    validate: () => descriptor,
    loadGlue: async () => glue as { default(input: { module_or_path: unknown }): Promise<unknown> },
    resolveWasm: async () => new Uint8Array(),
    construct,
    closeNative: archive => archive.close(),
  });
}

describe('direct source WASM trap containment', () => {
  it.each([
    'https://example.test/direct.wasm',
    'blob:https://example.test/01234567-89ab-cdef-0123-456789abcdef',
    'data:application/wasm;base64,AGFzbQ==',
  ])('retains validated absolute WASM input %s without a module-relative base', async (wasmUrl) => {
    const resolved = await resolveDirectWasmInput(wasmUrl);
    expect(resolved).toBeInstanceOf(URL);
    expect((resolved as URL).href).toBe(wasmUrl);
  });

  it('rejects a relative WASM input instead of resolving it against worker code', async () => {
    await expect(resolveDirectWasmInput('./direct.wasm')).rejects.toThrow();
  });

  it('poisons sibling live handles and never calls poisoned destructors', async () => {
    const glue = { default: vi.fn(async () => undefined) };
    const firstArchive = new Archive(); const secondArchive = new Archive();
    const archives = [firstArchive, secondArchive];
    const runtime = engine(glue, () => archives.shift()!);
    const first = await runtime.open(new Uint8Array(), descriptor);
    const second = await runtime.open(new Uint8Array(), descriptor);
    const cached = second.archive.call;
    firstArchive.call.mockImplementation(() => { throw new WebAssembly.RuntimeError('trap'); });
    let failure: unknown;
    try { first.archive.call(); } catch (error) { failure = error; }
    expect(failure).toBeInstanceOf(Error);
    expect(capture(() => second.archive.call())).toBe(failure);
    expect(capture(cached)).toBe(failure);
    expect(firstArchive.__destroy_into_raw).toHaveBeenCalledTimes(1);
    expect(secondArchive.__destroy_into_raw).toHaveBeenCalledTimes(1);
    first.closeArchive(); second.closeArchive();
    expect(firstArchive.free).not.toHaveBeenCalled();
    expect(secondArchive.free).not.toHaveBeenCalled();
    await expect(runtime.open(new Uint8Array(), descriptor)).rejects.toThrow('WASM runtime trapped');
  });

  it('shares poisoning by generated glue identity across engine wrappers', async () => {
    const glue = { default: vi.fn(async () => undefined) };
    const firstArchive = new Archive(); const secondArchive = new Archive();
    const firstRuntime = engine(glue, () => firstArchive);
    const secondRuntime = engine(glue, () => secondArchive);
    const first = await firstRuntime.open(new Uint8Array(), descriptor);
    const second = await secondRuntime.open(new Uint8Array(), descriptor);
    secondArchive.call.mockImplementation(() => { throw new WebAssembly.RuntimeError('trap'); });
    const failure = capture(() => second.archive.call());
    expect(capture(() => first.archive.call())).toBe(failure);
    expect(await captureAsync(() => firstRuntime.open(new Uint8Array(), descriptor))).toBe(failure);
    expect(firstArchive.__destroy_into_raw).toHaveBeenCalledOnce();
  });

  it('keeps ordinary Result-style exceptions nonfatal and preserves exact close', async () => {
    const archive = new Archive();
    const runtime = engine({ default: vi.fn(async () => undefined) }, () => archive);
    const source = await runtime.open(new Uint8Array(), descriptor);
    const ordinary = new Error('ordinary');
    archive.call.mockImplementationOnce(() => { throw ordinary; });
    expect(() => source.archive.call()).toThrow(ordinary);
    expect(source.archive.call()).toBe('ok');
    const cached = source.archive.call;
    source.closeArchive(); source.closeArchive();
    expect(() => cached()).toThrow('closed');
    expect(archive.close).toHaveBeenCalledOnce();
    expect(archive.free).toHaveBeenCalledOnce();
  });

  it('makes constructor and close traps sticky without re-entering free', async () => {
    const glue = { default: vi.fn(async () => undefined) };
    const constructorRuntime = engine(glue, () => { throw new WebAssembly.RuntimeError('ctor'); });
    await expect(constructorRuntime.open(new Uint8Array(), descriptor)).rejects.toThrow('WASM runtime trapped');
    await expect(constructorRuntime.open(new Uint8Array(), descriptor)).rejects.toThrow('WASM runtime trapped');

    const archive = new Archive();
    const closeRuntime = engine({ default: vi.fn(async () => undefined) }, () => archive);
    const source = await closeRuntime.open(new Uint8Array(), descriptor);
    archive.close.mockImplementation(() => { throw new WebAssembly.RuntimeError('close'); });
    expect(() => source.closeArchive()).toThrow('WASM runtime trapped');
    expect(archive.free).not.toHaveBeenCalled();
    expect(archive.__destroy_into_raw).toHaveBeenCalledOnce();
    expect(() => source.closeArchive()).not.toThrow();

    const freeArchive = new Archive();
    const freeRuntime = engine({ default: vi.fn(async () => undefined) }, () => freeArchive);
    const freeSource = await freeRuntime.open(new Uint8Array(), descriptor);
    freeArchive.free.mockImplementation(() => { throw new WebAssembly.RuntimeError('free'); });
    expect(() => freeSource.closeArchive()).toThrow('WASM runtime trapped');
    expect(freeArchive.close).toHaveBeenCalledOnce();
    expect(freeArchive.free).toHaveBeenCalledOnce();
    expect(freeArchive.__destroy_into_raw).toHaveBeenCalledOnce();
    await expect(freeRuntime.open(new Uint8Array(), descriptor)).rejects.toThrow('WASM runtime trapped');
  });

  it('makes an initialization trap sticky without constructing', async () => {
    const glue = { default: vi.fn(async () => { throw new WebAssembly.RuntimeError('init'); }) };
    const construct = vi.fn(() => new Archive());
    const runtime = engine(glue, construct);
    const failure = await captureAsync(() => runtime.open(new Uint8Array(), descriptor));
    expect(failure).toBeInstanceOf(Error);
    expect(await captureAsync(() => runtime.open(new Uint8Array(), descriptor))).toBe(failure);
    expect(glue.default).toHaveBeenCalledOnce();
    expect(construct).not.toHaveBeenCalled();
  });
});

function capture(operation: () => unknown): unknown {
  try { operation(); } catch (error) { return error; }
  throw new Error('expected operation to throw');
}

async function captureAsync(operation: () => Promise<unknown>): Promise<unknown> {
  try { await operation(); } catch (error) { return error; }
  throw new Error('expected operation to reject');
}
