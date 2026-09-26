import { describe, expect, it, vi } from 'vitest';
import {
  MODEL_SOURCE_MODULE_PROTOCOL,
  validateModelSourceModuleDescriptor,
  type ModelSource,
  type ModelSourceTarget,
} from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';

// Creating, claiming and describing a source must never load native code:
// the glue and source modules may be imported only by the archive realm.
const loaded = vi.hoisted(() => [] as string[]);
vi.mock('./wasm-direct-xls/legacy_xls_direct.js', () => { loaded.push('xls glue'); return {}; });
vi.mock('./wasm-direct-ppt/legacy_ppt_direct.js', () => { loaded.push('ppt glue'); return {}; });
vi.mock('./legacy-xls-source-module.ts', () => { loaded.push('xls module'); return {}; });
vi.mock('./legacy-ppt-source-module.ts', () => { loaded.push('ppt module'); return {}; });

import { legacyPptSource } from './legacy-ppt.js';
import { MAX_LEGACY_SOURCE_BYTES, createLegacySource, type LegacyFamily } from './legacy-source.js';
import { legacyXlsSource } from './legacy-xls.js';
import { buildStoredZip } from './zip-fixture.js';

const cfb = (...streams: string[]) => new Uint8Array(buildCfbFixture(['Root Entry', ...streams]));
const WASM = 'https://cdn.example.test/legacy.wasm';
const MODULE = 'https://cdn.example.test/legacy-source-module.js';

const factories = [
  { family: 'xls', target: 'xlsx', create: legacyXlsSource, streams: [['Workbook'], ['Book']] },
  { family: 'ppt', target: 'pptx', create: legacyPptSource, streams: [['PowerPoint Document']] },
] as const satisfies ReadonlyArray<{
  family: LegacyFamily;
  target: ModelSourceTarget;
  create: (options?: { wasmUrl?: string; moduleUrl?: string; maxInputBytes?: number }) => ModelSource;
  streams: ReadonlyArray<readonly string[]>;
}>;
// Binary families whose readers are not in this package yet; a source must
// still recognise their containers as foreign.
const readerlessFamilies = [
  { family: 'doc', streams: [['WordDocument']] },
] as const satisfies ReadonlyArray<{ family: LegacyFamily; streams: ReadonlyArray<readonly string[]> }>;

describe.each(factories)('legacy $family source factory', ({ family, target, create, streams }) => {
  it('describes its own target with frozen, admissible default asset URLs', () => {
    const source = create();
    expect(source.target).toBe(target);
    expect(Object.isFrozen(source)).toBe(true);
    const load = source.beginLoad();
    const { module } = load;
    expect(Object.isFrozen(module)).toBe(true);
    expect(Object.isFrozen(module.config)).toBe(true);
    expect(module).toEqual({
      protocol: MODEL_SOURCE_MODULE_PROTOCOL,
      target,
      moduleUrl: expect.any(String),
      config: { wasmUrl: expect.any(String), maxInputBytes: MAX_LEGACY_SOURCE_BYTES },
    });
    expect(new URL(module.moduleUrl).pathname).toContain(`legacy-${family}-source-module`);
    expect(new URL(String(module.config.wasmUrl)).pathname)
      .toContain(`/wasm-direct-${family}/legacy_${family}_direct_bg.wasm`);
    // The core loader admits the descriptor for this target and no other.
    expect(validateModelSourceModuleDescriptor(module, target)).toEqual(module);
    expect(() => load.release()).not.toThrow();
  });

  it('passes explicit URLs and budget through verbatim', () => {
    const { module } = create({ wasmUrl: WASM, moduleUrl: MODULE, maxInputBytes: 4096 }).beginLoad();
    expect(module).toEqual({
      protocol: MODEL_SOURCE_MODULE_PROTOCOL,
      target,
      moduleUrl: MODULE,
      config: { wasmUrl: WASM, maxInputBytes: 4096 },
    });
  });

  it('claims only an unencrypted CFB whose single binary family is its own', () => {
    const source = create();
    for (const own of streams) expect(source.claim(cfb(...own))).toBe(true);
    const others = [...factories.filter((other) => other.family !== family), ...readerlessFamilies];
    for (const other of others) {
      for (const foreign of other.streams) expect(source.claim(cfb(...foreign))).toBe(false);
      // Embedded objects can make the family ambiguous; do not guess.
      expect(source.claim(cfb(...streams[0], ...other.streams[0]))).toBe(false);
    }
    // Encrypted packages stay on the OOXML path, which reports them.
    expect(source.claim(cfb('EncryptionInfo', 'EncryptedPackage', ...streams[0]))).toBe(false);
    expect(source.claim(cfb())).toBe(false);
    expect(source.claim(buildStoredZip({ '[Content_Types].xml': '<Types/>' }))).toBe(false);
    expect(source.claim(new Uint8Array([1, 2, 3]))).toBe(false);
  });

  it('rejects oversize claimed input and ignores oversize unclaimed input', () => {
    const own = cfb(...streams[0]);
    const source = create({ maxInputBytes: own.byteLength - 1 });
    expect(() => source.claim(own)).toThrow(RangeError);
    expect(create({ maxInputBytes: own.byteLength }).claim(own)).toBe(true);
    const foreign = [...factories, ...readerlessFamilies].find((other) => other.family !== family)!.streams[0];
    expect(source.claim(cfb(...foreign))).toBe(false);
    expect(source.claim(buildStoredZip({ big: new Uint8Array(own.byteLength) }))).toBe(false);
  });
});

describe('createLegacySource option validation', () => {
  const defaults = { wasmUrl: WASM, moduleUrl: MODULE };
  const make = (options: unknown) => createLegacySource('doc', options as never, defaults);

  it('fails closed on non-object options', () => {
    for (const options of [null, 'https://example.test/doc.wasm', 1]) {
      expect(() => make(options)).toThrow(TypeError);
    }
  });

  it('admits absolute URLs of the supported protocols; data: only for WASM', () => {
    for (const wasmUrl of [
      'http://example.test/doc.wasm',
      'file:///opt/assets/doc.wasm',
      'blob:https://example.test/1234',
      'data:application/wasm;base64,AGFzbQ==',
    ]) {
      expect(make({ wasmUrl }).beginLoad().module.config.wasmUrl).toBe(wasmUrl);
    }
    for (const moduleUrl of ['http://example.test/m.js', 'file:///opt/m.js', 'blob:https://example.test/1234']) {
      expect(make({ moduleUrl }).beginLoad().module.moduleUrl).toBe(moduleUrl);
    }
    expect(() => make({ moduleUrl: 'data:text/javascript,export%20default%201' })).toThrow(/protocol/);
  });

  it.each(['wasmUrl', 'moduleUrl'] as const)('rejects a malformed %s', (field) => {
    for (const value of ['', './relative', ' https://example.test/x', 'https://example.test/x\n', 'javascript:alert(1)', 42]) {
      expect(() => make({ [field]: value })).toThrow(TypeError);
    }
  });

  it('bounds maxInputBytes to a positive safe integer up to the 256 MiB policy', () => {
    for (const maxInputBytes of [0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, MAX_LEGACY_SOURCE_BYTES + 1]) {
      expect(() => make({ maxInputBytes })).toThrow(RangeError);
    }
    expect(make({ maxInputBytes: 1 }).beginLoad().module.config.maxInputBytes).toBe(1);
    expect(make({ maxInputBytes: MAX_LEGACY_SOURCE_BYTES }).beginLoad().module.config.maxInputBytes)
      .toBe(MAX_LEGACY_SOURCE_BYTES);
  });

  it('never loaded a glue or source module while sources were created and used', () => {
    expect(loaded).toEqual([]);
  });
});
