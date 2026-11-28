import { describe, expect, it, vi } from 'vitest';
import type { DocxDocumentModel } from './types.js';
import { documentRenderCapabilities } from './vertical-render-capability.js';

const documentRequiresDomVerticalGlyphLayout = (document: DocxDocumentModel) =>
  documentRenderCapabilities(document).requiresDomVerticalGlyphLayout;

function model(): DocxDocumentModel {
  return {
    section: { textDirection: null },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

describe('vertical render capability', () => {
  it.each(['tbRl', 'tbRlV', 'tbLrV'])('routes a %s section through the DOM glyph path', (textDirection) => {
    const document = model();
    document.section.textDirection = textDirection;
    expect(documentRequiresDomVerticalGlyphLayout(document)).toBe(true);
  });

  it('finds vertical section transitions', () => {
    const sectioned = model();
    sectioned.body = [{ type: 'sectionBreak', kind: 'nextPage', textDirection: 'tbRl' }];
    expect(documentRequiresDomVerticalGlyphLayout(sectioned)).toBe(true);
  });

  it('keeps horizontal, all-rotated, and text-box-only vertical text on the OffscreenCanvas path', () => {
    const document = model() as unknown as Record<string, unknown>;
    document.section = { textDirection: 'btLr' };
    document.body = [
      { textVert: 'vert' },
      { textVert: 'vert270' },
      { textVert: 'eaVert' },
      { textVert: 'mongolianVert' },
    ];
    expect(documentRequiresDomVerticalGlyphLayout(document as unknown as DocxDocumentModel)).toBe(false);
  });

  it('yields both capability facts from one walk, safely on a cyclic caller model', () => {
    const document = model() as unknown as Record<string, unknown>;
    const run: Record<string, unknown> = { type: 'text', text: 'a', lang: 'sr-rs' };
    run.self = run;
    document.body = [{ type: 'paragraph', runs: [run, run] }, { type: 'sectionBreak', textDirection: 'tbRl' }];
    expect(documentRenderCapabilities(document as unknown as DocxDocumentModel))
      .toEqual({ requiresDomVerticalGlyphLayout: true, carriesRunLanguage: true });
    const keys = vi.spyOn(Object, 'keys');
    try {
      // A language and a vertical section found first: the walk stops there.
      document.body = [{ type: 'sectionBreak', textDirection: 'tbRl', lang: 'sr-rs' },
        ...Array.from({ length: 1000 }, () => ({ type: 'paragraph', runs: [] }))];
      (document.body as unknown[]).reverse();
      documentRenderCapabilities(document as unknown as DocxDocumentModel);
      expect(keys.mock.calls.length).toBeLessThan(10);
    } finally {
      keys.mockRestore();
    }
  });

  it('stops worker admission at a vertical section without collecting languages', () => {
    const document = model() as unknown as Record<string, unknown>;
    // The depth-first walk pops the last body block first.
    document.body = [
      ...Array.from({ length: 50_000 }, () => ({ type: 'paragraph', runs: [{ type: 'text', text: 'a' }] })),
      { type: 'sectionBreak', kind: 'nextPage', textDirection: 'tbRl' },
    ];
    const keys = vi.spyOn(Object, 'keys');
    try {
      expect(documentRenderCapabilities(
        document as unknown as DocxDocumentModel,
        { stopOnVerticalLayout: true },
      )).toEqual({ requiresDomVerticalGlyphLayout: true });
      expect(keys.mock.calls.length).toBeLessThan(10);
    } finally {
      keys.mockRestore();
    }
    // Without the option the walk still collects the language fact.
    expect(documentRenderCapabilities(document as unknown as DocxDocumentModel))
      .toEqual({ requiresDomVerticalGlyphLayout: true, carriesRunLanguage: false });
  });

  it('visits a subtree shared by several parents once', () => {
    let shared: Record<string, unknown> = { type: 'text', text: 'a' };
    for (let level = 0; level < 19; level += 1) shared = { left: shared, right: shared };
    const document = model() as unknown as Record<string, unknown>;
    document.body = [{ type: 'paragraph', runs: [shared] }];
    const keys = vi.spyOn(Object, 'keys');
    try {
      expect(documentRenderCapabilities(document as unknown as DocxDocumentModel))
        .toEqual({ requiresDomVerticalGlyphLayout: false, carriesRunLanguage: false });
      // 20 shared levels + the paragraph and the model's own records, not 2^19.
      expect(keys.mock.calls.length).toBeLessThan(40);
    } finally {
      keys.mockRestore();
    }
  });
});
