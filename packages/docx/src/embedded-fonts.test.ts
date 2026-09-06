import { afterEach, describe, expect, it, vi } from 'vitest';
import { deobfuscateOdttf } from '@silurus/ooxml-core';
import { loadEmbeddedFonts } from './embedded-fonts.js';
import type { DocxDocumentModel, EmbeddedFontRef } from './types';

// `loadEmbeddedFonts` maps `doc.embeddedFonts` → `EmbeddedFontFace[]` and calls
// the real core `registerEmbeddedFonts`. We stub the global FontFace +
// FontFaceSet so the faces the mapper produces surface as `added` entries on a
// fake set — asserting the derived family / weight / style / odttf-plaintext,
// exactly as core's own embedded.test.ts does.

const G = globalThis as Record<string, unknown>;
const ORIG = { document: G.document, self: G.self, FontFace: G.FontFace };

afterEach(() => {
  G.document = ORIG.document;
  G.self = ORIG.self;
  G.FontFace = ORIG.FontFace;
  vi.restoreAllMocks();
});

interface FakeFace {
  family: string;
  source: ArrayBuffer;
  weight: string;
  style: string;
  descriptors?: { weight?: string; style?: string };
  load: () => Promise<FakeFace>;
}

function installFontFaceSet() {
  const added: FakeFace[] = [];
  class FakeFontFace implements FakeFace {
    family: string;
    source: ArrayBuffer;
    weight: string;
    style: string;
    constructor(
      family: string,
      source: ArrayBuffer,
      public descriptors?: { weight?: string; style?: string },
    ) {
      this.family = family;
      this.source = source;
      this.weight = descriptors?.weight ?? 'normal';
      this.style = descriptors?.style ?? 'normal';
    }
    load(): Promise<FakeFace> {
      return Promise.resolve(this);
    }
  }
  const set = {
    add: (f: FakeFace) => {
      added.push(f);
    },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  delete G.self;
  return added;
}

// A minimal, valid sfnt header (TrueType 0x00010000) so `FontFace(source)`
// would accept the bytes — mirrors core's `validHeader`.
const validHeader = () =>
  new Uint8Array([
    0x00, 0x01, 0x00, 0x00, 0x00, 0x10, 0x01, 0x00, 0x00, 0x40, 0x00, 0x30,
    0x47, 0x53, 0x55, 0x42, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  ]);

function metricSfnt(): Uint8Array {
  const tableCount = 4;
  const headOffset = 12 + tableCount * 16;
  const hheaOffset = headOffset + 54;
  const os2Offset = hheaOffset + 36;
  const cmapOffset = os2Offset + 78;
  const bytes = new Uint8Array(cmapOffset + 40);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, tableCount);
  const record = (index: number, tag: string, offset: number, length: number) => {
    const at = 12 + index * 16;
    for (let i = 0; i < 4; i++) bytes[at + i] = tag.charCodeAt(i);
    view.setUint32(at + 8, offset);
    view.setUint32(at + 12, length);
  };
  record(0, 'head', headOffset, 54);
  record(1, 'hhea', hheaOffset, 36);
  record(2, 'OS/2', os2Offset, 78);
  record(3, 'cmap', cmapOffset, 40);
  view.setUint16(headOffset + 18, 2048);
  view.setInt16(hheaOffset + 4, 1802);
  view.setInt16(hheaOffset + 6, -455);
  view.setUint16(cmapOffset + 2, 1);
  view.setUint16(cmapOffset + 4, 3);
  view.setUint16(cmapOffset + 6, 10);
  view.setUint32(cmapOffset + 8, 12);
  view.setUint16(cmapOffset + 12, 12);
  view.setUint32(cmapOffset + 16, 28);
  view.setUint32(cmapOffset + 24, 1);
  view.setUint32(cmapOffset + 28, 0x56fd);
  view.setUint32(cmapOffset + 32, 0x56fd);
  view.setUint32(cmapOffset + 36, 1);
  return bytes;
}

const GUID = '{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}';

function modelWith(embeddedFonts?: EmbeddedFontRef[]): DocxDocumentModel {
  const emptyHf = { default: null, first: null, even: null };
  return {
    section: {} as DocxDocumentModel['section'],
    body: [],
    headers: emptyHf,
    footers: emptyHf,
    embeddedFonts,
  };
}

describe('loadEmbeddedFonts (ECMA-376 §17.8.1 / §17.8.3)', () => {
  it('derives East-Asian line metrics from an arbitrary embedded face, not its name', async () => {
    installFontFaceSet();
    const loaded = await loadEmbeddedFonts(modelWith([{
      fontName: 'Unlisted CJK Face', style: 'regular',
      partPath: 'word/fonts/font1.ttf', fontKey: '',
    }]), async () => metricSfnt());

    expect(loaded.metrics['unlisted cjk face']).toMatchObject({
      family: 'Unlisted CJK Face',
      requestedFamily: 'Unlisted CJK Face',
      weight: 400,
      style: 'normal',
      eastAsianLineHeightRatio: ((1802 + 455) * 1.3) / 2048,
    });
  });

  it('does not assign East-Asian metrics to a face whose cmap lacks East-Asian glyphs', async () => {
    installFontFaceSet();
    const latinOnly = metricSfnt();
    const view = new DataView(latinOnly.buffer);
    const cmapOffset = 12 + 4 * 16 + 54 + 36 + 78;
    view.setUint32(cmapOffset + 28, 0x41);
    view.setUint32(cmapOffset + 32, 0x5a);
    const loaded = await loadEmbeddedFonts(modelWith([{
      fontName: 'Latin Embedded Face', style: 'regular',
      partPath: 'word/fonts/font1.ttf', fontKey: '',
    }]), async () => latinOnly);

    expect(loaded.faces).toHaveLength(1);
    expect(loaded.metrics).toEqual({});
  });

  it('does not mix resources when a malformed document repeats one CSS face tuple', async () => {
    const added = installFontFaceSet();
    const latinOnly = metricSfnt();
    const cmapOffset = 12 + 4 * 16 + 54 + 36 + 78;
    new DataView(latinOnly.buffer).setUint32(cmapOffset + 28, 0x41);
    new DataView(latinOnly.buffer).setUint32(cmapOffset + 32, 0x5a);
    const cjk = metricSfnt();
    const fetchFontBytes = vi.fn(async (path: string) =>
      path.endsWith('first.ttf') ? latinOnly : cjk);
    const loaded = await loadEmbeddedFonts(modelWith([
      {
        fontName: 'Duplicate Face', style: 'regular',
        partPath: 'word/fonts/first.ttf', fontKey: '',
      },
      {
        fontName: 'Duplicate Face', style: 'regular',
        partPath: 'word/fonts/second.ttf', fontKey: '',
      },
    ]), fetchFontBytes);

    expect(fetchFontBytes).toHaveBeenCalledTimes(1);
    expect(fetchFontBytes).toHaveBeenCalledWith('word/fonts/first.ttf');
    expect(added).toHaveLength(1);
    expect(loaded.faces).toHaveLength(1);
    expect(loaded.metrics).toEqual({});
  });

  it('maps a 4-slot font to 4 faces with the correct weight/style descriptors', async () => {
    const added = installFontFaceSet();
    const refs: EmbeddedFontRef[] = [
      { fontName: 'Ubuntu', style: 'regular', partPath: 'word/fonts/font1.odttf', fontKey: GUID },
      { fontName: 'Ubuntu', style: 'bold', partPath: 'word/fonts/font2.odttf', fontKey: GUID },
      { fontName: 'Ubuntu', style: 'italic', partPath: 'word/fonts/font3.odttf', fontKey: GUID },
      { fontName: 'Ubuntu', style: 'boldItalic', partPath: 'word/fonts/font4.odttf', fontKey: GUID },
    ];
    // Every part is a valid header obfuscated with the GUID (so de-obfuscation
    // yields a valid sfnt), keyed by path.
    const bytesByPath = new Map(
      refs.map((r) => [r.partPath, deobfuscateOdttf(validHeader(), GUID)]),
    );
    await loadEmbeddedFonts(modelWith(refs), async (p) => bytesByPath.get(p)!);

    expect(added).toHaveLength(4);
    expect(added.every((f) => f.family === 'Ubuntu')).toBe(true);
    const byDesc = added.map((f) => `${f.descriptors?.weight}/${f.descriptors?.style}`).sort();
    expect(byDesc).toEqual([
      'bold/italic',
      'bold/normal',
      'normal/italic',
      'normal/normal',
    ]);
  });

  it('de-obfuscates a .odttf part once before registration', async () => {
    const added = installFontFaceSet();
    const refs: EmbeddedFontRef[] = [
      { fontName: 'Ubuntu', style: 'regular', partPath: 'word/fonts/font1.ODTTF', fontKey: GUID },
    ];
    // Obfuscated on the wire; the loader recognizes the extension
    // case-insensitively and hands the same plaintext bytes to metrics + FontFace.
    await loadEmbeddedFonts(modelWith(refs), async () => deobfuscateOdttf(validHeader(), GUID));
    expect(added).toHaveLength(1);
    // The first 4 bytes are the plaintext sfnt tag after de-obfuscation.
    expect(Array.from(new Uint8Array(added[0].source).slice(0, 4))).toEqual([
      0x00, 0x01, 0x00, 0x00,
    ]);
  });

  it('does not de-obfuscate a non-.odttf part (odttf=false)', async () => {
    const added = installFontFaceSet();
    const refs: EmbeddedFontRef[] = [
      { fontName: 'Roboto', style: 'regular', partPath: 'word/fonts/font1.ttf', fontKey: '' },
    ];
    // A raw sfnt part: odttf must be false so the bytes reach FontFace verbatim.
    await loadEmbeddedFonts(modelWith(refs), async () => validHeader());
    expect(added).toHaveLength(1);
    expect(Array.from(new Uint8Array(added[0].source).slice(0, 4))).toEqual([
      0x00, 0x01, 0x00, 0x00,
    ]);
  });

  it('skips a face whose fetch rejects, keeping the rest', async () => {
    const added = installFontFaceSet();
    const refs: EmbeddedFontRef[] = [
      { fontName: 'Good', style: 'regular', partPath: 'word/fonts/good.ttf', fontKey: '' },
      { fontName: 'Missing', style: 'regular', partPath: 'word/fonts/missing.ttf', fontKey: '' },
    ];
    await loadEmbeddedFonts(modelWith(refs), async (p) => {
      if (p.endsWith('missing.ttf')) throw new Error('no such part');
      return validHeader();
    });
    expect(added.map((f) => f.family)).toEqual(['Good']);
  });

  it('no-ops (no fetch) when embeddedFonts is empty or undefined', async () => {
    installFontFaceSet();
    const fetchSpy = vi.fn(async () => validHeader());

    await loadEmbeddedFonts(modelWith([]), fetchSpy);
    expect(fetchSpy).not.toHaveBeenCalled();

    await loadEmbeddedFonts(modelWith(undefined), fetchSpy);
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});
