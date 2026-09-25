import { describe, expect, it, vi } from 'vitest';

// The source modules keep their real engines and direct runtime; only the
// generated glue is replaced by a native fake, so native close/free calls are
// observable and admission failures can be shown to happen before any load.
const fake = vi.hoisted(() => {
  class Native {
    closes = 0;
    frees = 0;
    constructor(readonly bytes: Uint8Array) { fake.opened.push(this); }
    free() { this.frees += 1; }
    close_document_session() { this.closes += 1; }
    close_workbook_session() { this.closes += 1; }
    close_presentation_session() { this.closes += 1; }
    revision_markup_in_print(): boolean { return fake.revision.inPrint(); }
    has_revision_marks(): boolean { return fake.revision.hasMarks; }
  }
  const glue = {
    default: async () => undefined,
    LegacyDocDocument: Native,
    LegacyXlsWorkbook: Native,
    LegacyPptPresentation: Native,
  };
  return {
    Native,
    glue,
    loads: [] as string[],
    resolved: [] as string[],
    opened: [] as InstanceType<typeof Native>[],
    revision: { inPrint: (): boolean => false, hasMarks: false },
  };
});

function engineMock(create: string) {
  return async (importOriginal: () => Promise<Record<string, unknown>>) => {
    const original = await importOriginal();
    const real = original[create] as (load: () => Promise<unknown>, resolve: (url: string) => Promise<unknown>) => unknown;
    return {
      ...original,
      [create]: () => real(
        async () => { fake.loads.push(create); return fake.glue; },
        async (url: string) => { fake.resolved.push(url); return new Uint8Array(); },
      ),
    };
  };
}
vi.mock('./direct-doc-engine.js', engineMock('createLegacyDocSourceEngine'));
vi.mock('./direct-xls-engine.js', engineMock('createLegacyXlsSourceEngine'));
vi.mock('./direct-ppt-engine.js', engineMock('createLegacyPptSourceEngine'));

import { openModelSource as openDoc } from './legacy-doc-source-module.js';
import { openModelSource as openPpt } from './legacy-ppt-source-module.js';
import { MAX_LEGACY_SOURCE_BYTES } from './legacy-source.js';
import { openModelSource as openXls } from './legacy-xls-source-module.js';

const WASM = 'https://cdn.example.test/legacy.wasm';
const modules = [
  { family: 'DOC', open: openDoc },
  { family: 'XLS', open: openXls },
  { family: 'PPT', open: openPpt },
] as const;

describe.each(modules)('legacy $family source module lifecycle', ({ open }) => {
  it('fails closed on invalid config and oversize input before loading the runtime', async () => {
    const loads = fake.loads.length;
    const bytes = new Uint8Array(8);
    for (const config of [null, 'config', { wasmUrl: WASM, extra: 1 }, {}, { wasmUrl: 42 }]) {
      await expect(open(bytes, config)).rejects.toThrow(TypeError);
    }
    for (const maxInputBytes of [0, 1.5, '8', MAX_LEGACY_SOURCE_BYTES + 1]) {
      await expect(open(bytes, { wasmUrl: WASM, maxInputBytes })).rejects.toThrow(RangeError);
    }
    await expect(open(bytes, { wasmUrl: './relative.wasm' })).rejects.toThrow(TypeError);
    await expect(open(bytes, { wasmUrl: WASM, maxInputBytes: 7 })).rejects.toThrow(RangeError);
    const aborted = AbortSignal.abort();
    await expect(open(bytes, { wasmUrl: WASM }, aborted)).rejects.toMatchObject({ name: 'AbortError' });
    expect(fake.loads.length).toBe(loads);
  });

  it('opens with the configured WASM URL and closes the native archive exactly once', async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const opened = await open(bytes, { wasmUrl: WASM, maxInputBytes: 3 });
    const native = fake.opened.at(-1)!;
    expect(native.bytes).toBe(bytes);
    expect(fake.resolved).toContain(WASM);
    expect(Object.isFrozen(opened)).toBe(true);
    opened.close();
    opened.close();
    expect([native.closes, native.frees]).toEqual([1, 1]);
    expect(() => (opened.archive as unknown as { free(): void }).free()).toThrow(/closed/);
  });
});

describe('legacy DOC source module view defaults', () => {
  const open = (bytes = new Uint8Array(1)) => openDoc(bytes, { wasmUrl: WASM });

  it('asks for the markup view only when the DOC prints revision markup and has marks', async () => {
    const cases = [
      { inPrint: true, hasMarks: true, expected: { showTrackedChanges: true } },
      { inPrint: true, hasMarks: false, expected: {} },
      { inPrint: false, hasMarks: true, expected: {} },
    ];
    for (const { inPrint, hasMarks, expected } of cases) {
      fake.revision = { inPrint: () => inPrint, hasMarks };
      const opened = await open();
      expect(opened.viewDefaults).toEqual(expected);
      opened.close();
    }
  });

  it('closes the archive and rethrows when view defaults cannot be read', async () => {
    const failure = new Error('settings unavailable');
    fake.revision = { inPrint: () => { throw failure; }, hasMarks: true };
    await expect(open()).rejects.toBe(failure);
    const native = fake.opened.at(-1)!;
    expect([native.closes, native.frees]).toEqual([1, 1]);
    fake.revision = { inPrint: () => false, hasMarks: false };
  });
});
