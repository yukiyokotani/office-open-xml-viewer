import { describe, expect, it, vi } from 'vitest';
import { validateModelSourceModuleDescriptor } from '@silurus/ooxml-core/internal/model-source';
import { storedZip } from '@silurus/ooxml-core/testing';

// Creating, claiming and describing a source must never load native code.
const loaded = vi.hoisted(() => [] as string[]);
vi.mock('./wasm-odt/opendocument_odt.js', () => { loaded.push('odt glue'); return {}; });
vi.mock('./odt-source-module.ts', () => { loaded.push('odt module'); return {}; });

import { leadingOdfMediaType } from './odf-source.js';
import { createOdt, odt } from './odt.js';
import { buildOdtFixture } from './odt-fixture.js';

const TEXT = 'application/vnd.oasis.opendocument.text';

describe('ODT source claim', () => {
  it('claims packaged text documents and templates by the leading mimetype only', () => {
    expect(odt.claim(buildOdtFixture({ body: '<text:p/>' }))).toBe(true);
    expect(odt.claim(storedZip([['mimetype', `${TEXT}-template`]]))).toBe(true);
    for (const other of ['application/vnd.oasis.opendocument.spreadsheet', 'application/vnd.oasis.opendocument.presentation', `${TEXT}-master`]) {
      expect(odt.claim(storedZip([['mimetype', other]]))).toBe(false);
    }
    // Not first, not named mimetype, an OOXML package, or not a ZIP at all.
    expect(odt.claim(storedZip([['[Content_Types].xml', '<Types/>'], ['mimetype', TEXT]]))).toBe(false);
    expect(odt.claim(storedZip([['mimetypes', TEXT]]))).toBe(false);
    expect(odt.claim(new Uint8Array([1, 2, 3]))).toBe(false);
    expect(loaded).toEqual([]);
  });

  it('reads only a stored, unencrypted leading entry', () => {
    const bytes = storedZip([['mimetype', TEXT]]);
    expect(leadingOdfMediaType(bytes)).toBe(TEXT);
    const compressed = bytes.slice(); compressed[8] = 8;
    expect(leadingOdfMediaType(compressed)).toBeUndefined();
    const encrypted = bytes.slice(); encrypted[6] = 1;
    expect(leadingOdfMediaType(encrypted)).toBeUndefined();
    const truncated = bytes.slice(0, 40);
    expect(leadingOdfMediaType(truncated)).toBeUndefined();
  });

  it('describes a frozen v2 module descriptor with absolute URLs and validates options', () => {
    const source = createOdt({ wasmUrl: 'https://cdn.example.test/odt.wasm', moduleUrl: 'https://cdn.example.test/odt.mjs' });
    const descriptor = validateModelSourceModuleDescriptor(source.beginLoad().module, 'docx');
    expect(descriptor).toMatchObject({ target: 'docx', moduleUrl: 'https://cdn.example.test/odt.mjs', config: { wasmUrl: 'https://cdn.example.test/odt.wasm' } });
    expect(() => createOdt({ wasmUrl: './relative.wasm' })).toThrow(TypeError);
    expect(() => createOdt({ moduleUrl: 'data:text/javascript,0' })).toThrow(TypeError);
    expect(() => createOdt({ maxInputBytes: 1 } as never)).toThrow(TypeError);
    expect(odt.target).toBe('docx');
    expect(Object.isFrozen(odt)).toBe(true);
    expect(loaded).toEqual([]);
  });
});
