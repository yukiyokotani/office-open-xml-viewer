import { afterEach, describe, expect, it } from 'vitest';
import { acquireNativeReadingWordBreaking } from './native-reading-word-breaking.js';
import { nativeReadingNotices, hasNativeReadingRequests, nativeReadingRequests } from './native-reading-notice.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { DocxFindController } from './find.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunsForPage } from './text-run-projection.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';
const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
afterEach(() => {
  if (originalCanvas) Object.defineProperty(globalThis, 'OffscreenCanvas', originalCanvas);
  else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
});
const text = 'A😀e\u{301} literal-hyphen tail repeated word';
function model(reading: boolean) {
  const result = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 1 });
  const paragraph = result.body[0];
  if (paragraph.type !== 'paragraph') throw new Error('invented paragraph absent');
  const run = paragraph.runs[0];
  if (run.type !== 'text') throw new Error('invented text absent');
  run.text = text;
  if (reading) Object.assign(run, { __nativeReadingWordBreaking: { rawHres: 0, rawChHres: 1 } });
  result.section = { ...result.section, pageWidth: 90, marginLeft: 5, marginRight: 5 };
  return result;
}
describe('native reading unresolved word-breaking ownership', () => {
  it('acquires an immutable exact two-byte owner and refuses malformed wire records', () => {
    const raw = { rawHres: 0, rawChHres: 1 };
    const owner = acquireNativeReadingWordBreaking(raw);
    raw.rawChHres = 120;
    expect(owner).toEqual({ rawHres: 0, rawChHres: 1 });
    expect(Object.isFrozen(owner)).toBe(true);
    expect(acquireNativeReadingWordBreaking(undefined)).toBeUndefined();
    for (const value of [null, [], {}, { rawHres: 0 }, { rawHres: 1, rawChHres: 0 },
      { rawHres: -1, rawChHres: 1 }, { rawHres: 0, rawChHres: 256 },
      { rawHres: 0, rawChHres: 1, normalized: true }])
      expect(() => acquireNativeReadingWordBreaking(value)).toThrow(TypeError);
  });
  it('acquires the separate final numbering owner without rewriting marker text', () => {
    const result = model(false);
    const paragraph = result.body[0];
    if (paragraph.type !== 'paragraph') throw new Error('invented paragraph absent');
    paragraph.numbering = { numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 18, tab: 18, suff: 'tab' };
    Object.assign(paragraph.numbering, { __nativeReadingWordBreaking: { rawHres: 2, rawChHres: 120 } });
    const source = layoutSourceStore(result);
    const acquired = source.blocks.resolve(source.blocks.sources[0]);
    if (acquired.type !== 'paragraph') throw new Error('acquired paragraph absent');
    expect(acquired.numbering?.text).toBe('1.');
    expect(acquired.numbering).not.toHaveProperty('__nativeReadingWordBreaking');
    expect(acquired.nativeReadingNumberingWordBreaking).toEqual({ rawHres: 2, rawChHres: 120 });
    expect(hasNativeReadingRequests(source)).toBe(true);
  });
  it('discloses a field-only unresolved final owner after real acquisition and page layout', () => {
    installStubCanvas();
    const result = model(false), paragraph = result.body[0];
    if (paragraph.type !== 'paragraph') throw new Error('invented paragraph absent');
    paragraph.runs = [{ type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '7',
      bold: false, italic: false, underline: false, strikethrough: false, fontSize: 10,
      color: null, fontFamily: null, background: null, vertAlign: null }];
    Object.assign(paragraph.runs[0], { __nativeReadingWordBreaking: { rawHres: 0, rawChHres: 1 } });
    const source = layoutSourceStore(result), acquired = source.blocks.resolve(source.blocks.sources[0]);
    if (acquired.type !== 'paragraph' || acquired.runs[0].type !== 'field') throw new Error('field owner absent');
    expect(acquired.runs[0].nativeReadingWordBreaking).toEqual({ rawHres: 0, rawChHres: 1 });
    expect(acquired.runs[0]).not.toHaveProperty('__nativeReadingWordBreaking');
    expect(hasNativeReadingRequests(source)).toBe(true);
    const layout = layoutDocument(result, createLayoutServices(source));
    expect(layout.pages.length).toBeGreaterThan(0);
    expect(nativeReadingNotices(layout, nativeReadingRequests(source))).toEqual([
      { code: 'WORD_BREAKING_SIMPLIFIED_FOR_READING', message: 'Word breaking is simplified for readability; line and page breaks may differ' },
    ]);
    expect(textRunsForPage(layout, 0, { scale: 1 }).map(run => run.text).join('')).toBe('1');
  });
  it('retains source UTF-16 text and raw ownership through real acquisition and wrapping', async () => {
    installStubCanvas();
    const ordinary = model(false), reading = model(true);
    const source = layoutSourceStore(reading);
    expect(hasNativeReadingRequests(source)).toBe(true);
    const paragraph = source.blocks.resolve(source.blocks.sources[0]);
    if (paragraph.type !== 'paragraph') throw new Error('acquired paragraph absent');
    const run = paragraph.runs[0];
    if (run.type !== 'text') throw new Error('acquired text absent');
    expect(run.text).toBe(text);
    expect(run.nativeReadingWordBreaking).toEqual({ rawHres: 0, rawChHres: 1 });
    expect(run).not.toHaveProperty('__nativeReadingWordBreaking');
    const a = layoutDocument(ordinary, createLayoutServices(ordinary));
    const b = layoutDocument(reading, createLayoutServices(reading));
    const copy = (layout: typeof a) => layout.pages.flatMap((_, i) => textRunsForPage(layout, i, { scale: 1 }));
    expect(copy(b)).toEqual(copy(a));
    expect(copy(b).map(run => run.text).join('')).toBe(text);
    const ranges = (layout: typeof a) => layout.pages.flatMap((_, i) => textRunGeometryForPage(layout, i).map(run => ({
      source: run.source, sourceRunIndex: run.placement.sourceRunIndex,
      range: run.placement.range, clusters: run.placement.clusters,
      bounds: run.placement.bounds, transform: run.pointToPage,
    })));
    expect(ranges(b)).toEqual(ranges(a));
    expect(Math.max(...ranges(b).map(run => run.range.end))).toBe(text.length);
    expect(nativeReadingNotices(b, nativeReadingRequests(source))).toEqual([
      { code: 'WORD_BREAKING_SIMPLIFIED_FOR_READING', message: 'Word breaking is simplified for readability; line and page breaks may differ' },
    ]);
    expect(nativeReadingNotices(a)).toEqual([]);
    const find = (layout: typeof a) => new DocxFindController(() => layout.pages.length,
      async page => textRunsForPage(layout, page, { scale: 1 }));
    const ordinaryFind = find(a), readingFind = find(b);
    for (const query of ['A😀e\u{301}', 'literal-hyphen', 'tail']) {
      const expected = await ordinaryFind.find(query);
      expect(expected).toHaveLength(1);
      expect(await readingFind.find(query)).toEqual(expected);
      for (let page = 0; page < a.pages.length; page++)
        expect(readingFind.pageHighlights(page)).toEqual(ordinaryFind.pageHighlights(page));
    }
  });
});
