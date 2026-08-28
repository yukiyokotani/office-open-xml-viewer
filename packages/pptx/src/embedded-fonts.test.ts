import { afterEach, describe, expect, it, vi } from 'vitest';
import { excludeEmbeddedFontFamilies, loadEmbeddedFonts } from './embedded-fonts.js';
import type { PptxEmbeddedFontRef } from './worker-protocol';

const globals = globalThis as Record<string, unknown>;
const original = { document: globals.document, self: globals.self, FontFace: globals.FontFace };

afterEach(() => {
  globals.document = original.document;
  globals.self = original.self;
  globals.FontFace = original.FontFace;
  vi.restoreAllMocks();
});

function installFontFaceSet(failLoad = false) {
  const added: Array<{ family: string; source: ArrayBuffer; descriptors: FontFaceDescriptors }> = [];
  class FakeFontFace {
    constructor(
      public family: string,
      public source: ArrayBuffer,
      public descriptors: FontFaceDescriptors,
    ) {}
    load() {
      return failLoad
        ? Promise.reject(new Error('load failed'))
        : Promise.resolve(this);
    }
  }
  globals.FontFace = FakeFontFace;
  globals.document = { fonts: { add: (face: typeof added[number]) => added.push(face), ready: Promise.resolve() } };
  delete globals.self;
  return added;
}

const bytes = () => new Uint8Array([0, 1, 0, 0, 1]);

describe('loadEmbeddedFonts (ECMA-376 §19.2.1.9 / §15.2.13)', () => {
  it('maps all four PresentationML slots to CSS weight and style', async () => {
    const added = installFontFaceSet();
    const refs: PptxEmbeddedFontRef[] = ['regular', 'bold', 'italic', 'boldItalic'].map(
      (style, index) => ({
        fontName: 'Deck Sans',
        style: style as PptxEmbeddedFontRef['style'],
        partPath: `ppt/fonts/font${index + 1}.fntdata`,
        contentType: 'application/x-font-ttf',
      }),
    );
    const loaded = await loadEmbeddedFonts(refs, async () => bytes());
    expect(added.map((face) => `${face.descriptors.weight}/${face.descriptors.style}`).sort()).toEqual([
      'bold/italic', 'bold/normal', 'normal/italic', 'normal/normal',
    ]);
    expect(new Set(added.map((face) => face.family)).size).toBe(1);
    expect(loaded.aliases.get('deck sans')).toBe(added[0].family);
    expect(loaded.authoredFamilies.get(added[0].family)).toBe('deck sans');
  });

  it('keeps raw PPTX bytes and skips an unreadable part without aborting siblings', async () => {
    const added = installFontFaceSet();
    const refs: PptxEmbeddedFontRef[] = [
      { fontName: 'Good', style: 'regular', partPath: 'ppt/fonts/good.fntdata', contentType: 'application/x-font-ttf' },
      { fontName: 'Missing', style: 'regular', partPath: 'ppt/fonts/missing.fntdata', contentType: 'application/x-fontdata' },
    ];
    const loaded = await loadEmbeddedFonts(refs, async (path) => {
      if (path.includes('missing')) throw new Error('missing');
      return bytes();
    });
    expect(added).toHaveLength(1);
    expect(loaded.aliases.has('good')).toBe(true);
    expect(loaded.aliases.has('missing')).toBe(false);
    expect(Array.from(new Uint8Array(added[0].source))).toEqual(Array.from(bytes()));
  });

  it('does not fetch when there are no embedded fonts', async () => {
    installFontFaceSet();
    const fetchFont = vi.fn(async () => bytes());
    await loadEmbeddedFonts([], fetchFont);
    expect(fetchFont).not.toHaveBeenCalled();
  });

  it('bounds concurrent extraction to two font parts', async () => {
    installFontFaceSet();
    let active = 0;
    let peak = 0;
    const refs: PptxEmbeddedFontRef[] = Array.from({ length: 5 }, (_, index) => ({
      fontName: `Deck Font ${index}`,
      style: 'regular',
      partPath: `ppt/fonts/font${index}.fntdata`,
      contentType: 'application/x-font-ttf',
    }));
    await loadEmbeddedFonts(refs, async () => {
      active++;
      peak = Math.max(peak, active);
      await Promise.resolve();
      active--;
      return bytes();
    });
    expect(peak).toBe(2);
  });

  it('isolates the same authored family across concurrently open presentations', async () => {
    const added = installFontFaceSet();
    const refs: PptxEmbeddedFontRef[] = [{
      fontName: 'Shared Family', style: 'regular', partPath: 'ppt/fonts/font1.fntdata',
      contentType: 'application/x-font-ttf',
    }];
    const first = await loadEmbeddedFonts(refs, async () => bytes());
    const second = await loadEmbeddedFonts(refs, async () => new Uint8Array([0, 1, 0, 0, 2]));
    expect(added).toHaveLength(2);
    expect(first.aliases.get('shared family')).not.toBe(second.aliases.get('shared family'));
    expect(new Set(added.map((face) => face.family)).size).toBe(2);
  });

  it('keeps successfully loaded embedded families ahead of optional Google-font substitutes', async () => {
    installFontFaceSet();
    const loaded = await loadEmbeddedFonts([{
      fontName: 'Calibri', style: 'regular', partPath: 'ppt/fonts/font1.fntdata',
      contentType: 'application/x-font-ttf',
    }], async () => bytes());
    expect(excludeEmbeddedFontFamilies(['Aptos', 'calibri', null], loaded.aliases)).toEqual(['Aptos', null]);
  });

  it('keeps the web substitute eligible when an embedded face fails to load', async () => {
    installFontFaceSet(true);
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    const loaded = await loadEmbeddedFonts([{
      fontName: 'Calibri', style: 'regular', partPath: 'ppt/fonts/font1.fntdata',
      contentType: 'application/x-font-ttf',
    }], async () => bytes());
    expect(loaded.faces).toEqual([]);
    expect(excludeEmbeddedFontFamilies(['calibri'], loaded.aliases)).toEqual(['calibri']);
  });
});
