import { afterEach, describe, expect, it, vi } from 'vitest';
import { createHash } from 'node:crypto';
import { deobfuscateOdttf, unregisterEmbeddedFonts } from '@silurus/ooxml-core';
import { loadEmbeddedFonts } from './embedded-fonts.js';
import { buildSegments, paragraphMarkLineMetrics, type LayoutTextSeg } from './line-layout.js';
import { snapshotFontMetrics } from './layout/text.js';
import { createTextLayoutService } from './layout/text.js';
import { createFontResolver } from './layout/font-service.js';
import type { DocParagraph, DocRun, DocxDocumentModel, EmbeddedFontRef } from './types';

// `loadEmbeddedFonts` maps `doc.embeddedFonts` → `EmbeddedFontFace[]` and calls
// the real core `registerEmbeddedFonts`. We stub the global FontFace +
// FontFaceSet so the faces the mapper produces surface as `added` entries on a
// fake set — asserting the derived family / weight / style / odttf-plaintext,
// exactly as core's own embedded.test.ts does.

const G = globalThis as Record<string, unknown>;
const ORIG = { document: G.document, self: G.self, FontFace: G.FontFace };

afterEach(() => {
  vi.unstubAllGlobals();
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

function installFontFaceSet(load?: (face: FakeFace) => Promise<FakeFace>) {
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
      return load?.(this) ?? Promise.resolve(this);
    }
  }
  const set = {
    add: (f: FakeFace) => {
      added.push(f);
    },
    delete: (f: FakeFace) => {
      const index = added.indexOf(f);
      if (index >= 0) added.splice(index, 1);
      return index >= 0;
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

function metricSfnt(codePageRange1 = 1 << 17): Uint8Array {
  const tableCount = 4;
  const headOffset = 12 + tableCount * 16;
  const hheaOffset = headOffset + 54;
  const os2Offset = hheaOffset + 36;
  const cmapOffset = os2Offset + 86;
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
  record(2, 'OS/2', os2Offset, 86);
  record(3, 'cmap', cmapOffset, 40);
  view.setUint16(headOffset + 18, 2048);
  view.setInt16(hheaOffset + 4, 1802);
  view.setInt16(hheaOffset + 6, -455);
  view.setUint16(os2Offset, 1);
  view.setUint32(os2Offset + 78, codePageRange1);
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
  it('isolates different resources with the same authored face while both documents are open', async () => {
    const added = installFontFaceSet();
    const refs = [{
      fontName: 'Shared Authored Name', style: 'regular' as const,
      partPath: 'word/fonts/shared.ttf', fontKey: '',
    }];
    const firstBytes = metricSfnt();
    const secondBytes = metricSfnt();
    new DataView(secondBytes.buffer).setInt16(12 + 4 * 16 + 54 + 4, 1600);
    const first = await loadEmbeddedFonts(modelWith(refs), async () => firstBytes);
    const second = await loadEmbeddedFonts(modelWith(refs), async () => secondBytes);
    expect(added).toHaveLength(2);
    expect(first.routes[0].resolvedFamily).not.toBe(second.routes[0].resolvedFamily);
    expect(first.routes[0].resourceIdentity).not.toBe(second.routes[0].resourceIdentity);
    expect(first.metrics['shared authored name'].family).toBe(first.routes[0].resolvedFamily);
    expect(second.metrics['shared authored name'].family).toBe(second.routes[0].resolvedFamily);
    const resolve = (routes: typeof first.routes) => createFontResolver(routes.map((route) => ({
      ...route, source: 'embedded' as const,
    }))).resolve({ requestedFamily: 'Shared Authored Name' });
    expect(resolve(first.routes).resolvedFamily).toBe(first.faces[0].family);
    expect(resolve(second.routes).resolvedFamily).toBe(second.faces[0].family);
    unregisterEmbeddedFonts(first.faces);
    expect(added).toHaveLength(1);
    expect(resolve(second.routes).resolvedFamily).toBe(added[0].family);
    unregisterEmbeddedFonts(second.faces);
    expect(added).toHaveLength(0);
  });

  it('derives the same isolated family when SubtleCrypto is unavailable', async () => {
    installFontFaceSet();
    const refs = [{
      fontName: 'Offline Face', style: 'regular' as const,
      partPath: 'word/fonts/offline.ttf', fontKey: '',
    }];
    const online = await loadEmbeddedFonts(modelWith(refs), async () => metricSfnt());
    vi.stubGlobal('crypto', undefined);
    const offline = await loadEmbeddedFonts(modelWith(refs), async () => metricSfnt());
    expect(offline.routes[0].resolvedFamily).toBe(online.routes[0].resolvedFamily);
    expect(offline.routes[0].resolvedFamily).toContain(
      createHash('sha256').update(metricSfnt()).digest('hex'),
    );
    unregisterEmbeddedFonts(online.faces);
    unregisterEmbeddedFonts(offline.faces);
  });

  it('derives East-Asian line metrics from an arbitrary embedded face, not its name', async () => {
    installFontFaceSet();
    const loaded = await loadEmbeddedFonts(modelWith([{
      fontName: 'Unlisted CJK Face', style: 'regular',
      partPath: 'word/fonts/font1.ttf', fontKey: '',
    }]), async () => metricSfnt());

    expect(loaded.metrics['unlisted cjk face']).toMatchObject({
      family: loaded.routes[0].resolvedFamily,
      requestedFamily: 'Unlisted CJK Face',
      weight: 400,
      style: 'normal',
    });
    expect(loaded.metrics['unlisted cjk face'].lineHeightRatio)
      .toBeCloseTo(((1802 + 455) * 1.3) / 2048, 12);
    expect(loaded.metrics['unlisted cjk face'].eastAsianLineHeightRatio)
      .toBeCloseTo(((1802 + 455) * 1.3) / 2048, 12);
    const retained = snapshotFontMetrics(loaded.metrics)['unlisted cjk face'];
    expect(retained.designAscentRatio).toBeCloseTo((1802 + (2257 * 0.3) / 2) / 2048, 12);
    expect(retained.designDescentRatio).toBeCloseTo((455 + (2257 * 0.3) / 2) / 2048, 12);
  });

  it('keeps code-page line allocation but withholds East-Asian coverage when cmap lacks glyphs', async () => {
    installFontFaceSet();
    const latinOnly = metricSfnt();
    const view = new DataView(latinOnly.buffer);
    const cmapOffset = 12 + 4 * 16 + 54 + 36 + 86;
    view.setUint32(cmapOffset + 28, 0x41);
    view.setUint32(cmapOffset + 32, 0x5a);
    const loaded = await loadEmbeddedFonts(modelWith([{
      fontName: 'Latin Embedded Face', style: 'regular',
      partPath: 'word/fonts/font1.ttf', fontKey: '',
    }]), async () => latinOnly);

    expect(loaded.faces).toHaveLength(1);
    expect(loaded.metrics['latin embedded face'].lineHeightRatio)
      .toBeCloseTo(((1802 + 455) * 1.3) / 2048, 12);
    expect(loaded.metrics['latin embedded face'].eastAsianLineHeightRatio).toBeUndefined();
  });

  it('uses ordinary hhea allocation when the same CJK cmap has no Far East code-page bit', async () => {
    installFontFaceSet();
    const loaded = await loadEmbeddedFonts(modelWith([{
      fontName: 'Ordinary CJK Face', style: 'regular',
      partPath: 'word/fonts/font1.ttf', fontKey: '',
    }]), async () => metricSfnt(1));
    expect(loaded.metrics['ordinary cjk face']).toMatchObject({
      lineHeightRatio: (1802 + 455) / 2048,
      eastAsianLineHeightRatio: (1802 + 455) / 2048,
    });
  });

  it('does not lend a subset face line height to an uncovered fallback glyph', async () => {
    installFontFaceSet();
    const loaded = await loadEmbeddedFonts(modelWith([{
      fontName: 'Subset CJK Face', style: 'regular',
      partPath: 'word/fonts/subset.ttf', fontKey: '',
    }]), async () => metricSfnt());
    const metrics = snapshotFontMetrics(loaded.metrics);
    const textService = createTextLayoutService({
      fonts: createFontResolver([{
        ...loaded.routes[0], source: 'embedded',
      }]),
      measurer: {
        fingerprint: 'subset-embedded-coverage',
        measure: (request) => ({
          advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2,
        }),
      },
      fontMetrics: metrics,
    });
    const segment = (text: string): LayoutTextSeg => buildSegments([{
      type: 'text', text, fontFamily: 'Subset CJK Face', fontFamilyEastAsia: 'Subset CJK Face',
      fontSize: 10, bold: false, italic: false, underline: false, strikethrough: false,
    } as DocRun], {
      pageIndex: 0, totalPages: 1,
      layoutServices: { text: textService } as unknown as
        NonNullable<Parameters<typeof buildSegments>[1]['layoutServices']>,
    })[0] as LayoutTextSeg;
    expect(segment('国').resolvedLineHeightRatio).toBeCloseTo(1.3 * 2257 / 2048, 12);
    expect(segment('日').resolvedLineHeightRatio).toBeUndefined();
    expect(segment('日').resolvedEastAsianLineHeightRatio).toBeUndefined();

    const context = {
      font: '10px serif',
      measureText: () => ({
        width: 5, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
      }),
    } as unknown as CanvasRenderingContext2D;
    const paragraph = {
      runs: [], defaultFontFamily: 'Subset CJK Face', defaultFontSize: 10,
      lineSpacing: null,
    } as unknown as DocParagraph;
    const mark = paragraphMarkLineMetrics(
      paragraph, 1, undefined, false, false, context, {}, null, metrics,
    );
    expect(mark.advancePx).toBe(10);
  });

  it('does not mix resources when a malformed document repeats one CSS face tuple', async () => {
    const added = installFontFaceSet();
    const latinOnly = metricSfnt();
    const cmapOffset = 12 + 4 * 16 + 54 + 36 + 86;
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
    expect(loaded.metrics['duplicate face'].lineHeightRatio)
      .toBeCloseTo(((1802 + 455) * 1.3) / 2048, 12);
    expect(loaded.metrics['duplicate face'].eastAsianLineHeightRatio).toBeUndefined();
  });

  it('bounds fetch fan-out and retained faces for a document with many declarations', async () => {
    let activeLoads = 0;
    let peakLoads = 0;
    const added = installFontFaceSet(async (face) => {
      activeLoads++;
      peakLoads = Math.max(peakLoads, activeLoads);
      await Promise.resolve();
      activeLoads--;
      return face;
    });
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const refs = Array.from({ length: 80 }, (_, index): EmbeddedFontRef => ({
      fontName: `Face ${index}`, style: 'regular',
      partPath: `word/fonts/font${index}.ttf`, fontKey: '',
    }));
    let active = 0;
    let peak = 0;
    const fetchFontBytes = vi.fn(async () => {
      active++;
      peak = Math.max(peak, active);
      await Promise.resolve();
      active--;
      return validHeader();
    });

    const loaded = await loadEmbeddedFonts(modelWith(refs), fetchFontBytes);
    expect(fetchFontBytes).toHaveBeenCalledTimes(64);
    expect(peak).toBe(2);
    expect(peakLoads).toBe(2);
    expect(activeLoads).toBe(0);
    expect(added).toHaveLength(64);
    expect(loaded.faces).toHaveLength(64);
    expect(warning).toHaveBeenCalledWith(expect.stringContaining('skipped 16 embedded font face(s)'));
  });

  it('admits faces in document order under the aggregate decoded-byte ceiling', async () => {
    const added = installFontFaceSet();
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const refs = Array.from({ length: 5 }, (_, index): EmbeddedFontRef => ({
      fontName: `Large Face ${index}`, style: 'regular',
      partPath: `word/fonts/large${index}.ttf`, fontKey: '',
    }));
    const fetchFontBytes = vi.fn(async () => {
      // A declared length lets this boundary test exercise the 128 MiB budget
      // without allocating five large font binaries. FontFace still receives
      // the small valid header held in the backing buffer.
      const bytes = validHeader();
      Object.defineProperty(bytes, 'byteLength', { value: 30 * 1024 * 1024 });
      return bytes;
    });

    const loaded = await loadEmbeddedFonts(modelWith(refs), fetchFontBytes);
    expect(fetchFontBytes).toHaveBeenCalledTimes(5);
    expect(added.map((face) => face.family)).toEqual(loaded.routes.map((route) => route.resolvedFamily));
    expect(loaded.faces).toHaveLength(4);
    expect(warning).toHaveBeenCalledWith(expect.stringContaining('skipped 1 embedded font face(s)'));
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
    expect(new Set(added.map((f) => f.family))).toEqual(new Set([added[0].family]));
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
    expect(added.map((f) => f.family)).toEqual([expect.stringMatching(/^__ooxml_docx_embedded_/)]);
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
