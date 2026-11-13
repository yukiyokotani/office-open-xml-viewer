import { describe, expect, it, vi } from 'vitest';
import { validateModelSourceModuleDescriptor } from '@silurus/ooxml-core/internal/model-source';
import { storedZip } from '@silurus/ooxml-core/testing';

// Creating, claiming and describing a source must never load native code.
const loaded = vi.hoisted(() => [] as string[]);
vi.mock('./wasm-odp/opendocument_odp.js', () => { loaded.push('odp glue'); return {}; });
vi.mock('./odp-source-module.ts', () => { loaded.push('odp module'); return {}; });

import { createOdp, odp } from './odp.js';
import { buildOdpFixture } from './odp-fixture.js';

const PRESENTATION = 'application/vnd.oasis.opendocument.presentation';

describe('ODP source claim', () => {
  it('claims packaged presentations and templates by the leading mimetype only', () => {
    expect(odp.claim(buildOdpFixture({ pages: '<draw:page/>' }))).toBe(true);
    expect(odp.claim(storedZip([['mimetype', `${PRESENTATION}-template`]]))).toBe(true);
    for (const other of ['application/vnd.oasis.opendocument.text', 'application/vnd.oasis.opendocument.graphics']) {
      expect(odp.claim(storedZip([['mimetype', other]]))).toBe(false);
    }
    expect(odp.claim(storedZip([['[Content_Types].xml', '<Types/>'], ['mimetype', PRESENTATION]]))).toBe(false);
    expect(loaded).toEqual([]);
  });

  it('describes a PPTX module descriptor and validates options', () => {
    const source = createOdp({ wasmUrl: 'https://cdn.example.test/odp.wasm', moduleUrl: 'https://cdn.example.test/odp.mjs' });
    const descriptor = validateModelSourceModuleDescriptor(source.beginLoad().module, 'pptx');
    expect(descriptor).toMatchObject({ target: 'pptx', moduleUrl: 'https://cdn.example.test/odp.mjs', config: { wasmUrl: 'https://cdn.example.test/odp.wasm' } });
    expect(() => createOdp({ wasmUrl: './relative.wasm' })).toThrow(TypeError);
    expect(odp.target).toBe('pptx');
    expect(loaded).toEqual([]);
  });
});
