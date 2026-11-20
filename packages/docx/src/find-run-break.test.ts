import { describe, expect, it } from 'vitest';
import type {
  BodyElement, DocParagraph, DocxTextRun, DocxDocumentModel, SectionProps,
} from './types';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { textRunsForPage } from './text-run-projection.js';
import { DocxFindController } from './find.js';
import { docxRunsBreak } from './find-run-break.js';

// Whole-word matching must tell an automatic wrap inside a word from the end
// of a line of text. The layout reports each run's range in its paragraph's
// text; this pins that contract against the real layout, not hand-made runs.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

type DocRun = DocParagraph['runs'][number];
function textRun(text: string): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
  return { type: 'text', ...run } as DocRun;
}
function para(...runs: DocRun[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

function pageRuns(body: BodyElement[], pageWidth: number) {
  const section: SectionProps = {
    pageWidth, pageHeight: 800,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  const model = {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const layout = layoutDocument(
    model, createLayoutServices(model, { measureContext: makeCtx() }), { currentDateMs: 0 },
  );
  return textRunsForPage(layout, 0, { scale: 1 }).filter((run) => run.text.length > 0);
}

describe('DOCX layout breaks for whole-word matching', () => {
  it('joins a word that wraps inside the word, and separates paragraphs, manual breaks and tabs', async () => {
    const runs = pageRuns([
      // Too long for the 100pt line: the layout breaks inside the word.
      para(textRun('catalogue')),
      para(textRun('cat'), { type: 'break', breakType: 'line' } as DocRun, textRun('alog')),
      para(textRun('cat\talog')),
    ], 140);
    const text = runs.map((run) => run.text).join('|');
    const first = runs.filter((run) => run.source?.path[0] === 0);
    expect(first.map((run) => run.text).join('')).toBe('catalogue');
    // The word is on more than one line, and its pieces are adjacent in the text.
    expect(new Set(first.map((run) => run.y)).size).toBeGreaterThan(1);
    for (let i = 1; i < first.length; i++) {
      expect(first[i].textRange?.start, text).toBe(first[i - 1].textRange?.end);
      expect(docxRunsBreak(first[i - 1], first[i]), text).toBe(false);
    }

    const pair = (paragraph: number) => runs.filter((run) => run.source?.path[0] === paragraph && run.text.trim() !== '');
    // A manual line break and a tab each lie between the two runs in the text.
    for (const paragraph of [1, 2]) {
      const [cat, after] = pair(paragraph);
      expect(cat.text, text).toBe('cat');
      expect(after.text.startsWith('al'), text).toBe(true);
      expect(after.textRange!.start, text).toBeGreaterThan(cat.textRange!.end);
      expect(docxRunsBreak(cat, after), text).toBe(true);
    }
    // After the tab the rest of the line is too short for "alog": it wraps
    // inside the word and stays one word.
    const [, head, tail] = pair(2);
    expect(head.text + tail.text, text).toBe('alog');
    expect(tail.y, text).toBeGreaterThan(head.y);
    expect(docxRunsBreak(head, tail), text).toBe(false);
    // The last run of a paragraph and the first of the next are separate.
    expect(docxRunsBreak(first.at(-1)!, pair(1)[0])).toBe(true);

    const find = new DocxFindController(() => 1, async () => runs);
    // "cat" stands alone twice (before the break and before the tab), and is
    // not found at the start of the wrapped "catalogue".
    expect(await find.find('cat', { wholeWord: true })).toHaveLength(2);
    expect(await find.find('cat')).toHaveLength(3);
    expect(await find.find('catalogue', { wholeWord: true })).toHaveLength(1);
    expect((await find.highlight([{ id: 'a', terms: ['alog'], wholeWord: true }])).map((match) => match.text))
      .toEqual(['alog', 'alog']);
  });

  it('keeps a word in one piece when distributed alignment spreads its runs apart', async () => {
    const bold = { ...(textRun('cat') as object), bold: true } as DocRun;
    const distributed = para(bold, textRun('alog'));
    (distributed as unknown as { alignment: string }).alignment = 'distribute';
    // A second paragraph keeps the distributed one from being a last line only.
    const runs = pageRuns([distributed, para(textRun('cat'))], 400);
    const text = runs.map((run) => `${run.text}@${Math.round(run.x)}`).join('|');
    const [cat, alog] = runs;
    expect([cat.text, alog.text], text).toEqual(['cat', 'alog']);
    // The runs are adjacent in the text although the line is spread out.
    expect(alog.textRange?.start, text).toBe(cat.textRange?.end);
    expect(alog.x - (cat.x + cat.w), text).toBeGreaterThan(0.25 * cat.fontSize);
    expect(docxRunsBreak(cat, alog), text).toBe(false);

    const find = new DocxFindController(() => 1, async () => runs);
    // Only the second paragraph's "cat" is a word.
    const matches = await find.find('cat', { wholeWord: true });
    expect(matches, text).toHaveLength(1);
    expect(find.pageHighlights(0)[0].slices[0].runIndex).toBe(runs.length - 1);
    expect(await find.find('catalog', { wholeWord: true })).toHaveLength(1);
  });

  it.each(['tab', 'space', 'nothing'] as const)(
    'keeps a numbering marker apart from the paragraph text (suffix: %s)', async (suff) => {
      const numbered = para(textRun('cat catalog'));
      (numbered as unknown as { numbering: unknown; indentLeft: number; indentFirst: number }).numbering =
        { numId: 1, level: 0, format: 'upperLetter', text: 'A', indentLeft: 36, tab: 36, suff, jc: 'left' };
      const runs = pageRuns([numbered], 400);
      const text = runs.map((run) => `${run.text}[${run.textRange?.start},${run.textRange?.end})`).join('|');
      const [marker, first] = runs;
      expect(marker.text, text).toBe('A');
      expect(marker.textRange!.start, text).toBeLessThan(0);
      expect(first.textRange!.start, text).toBe(0);
      // The marker is generated: it is not the first letter of "Acat".
      expect(docxRunsBreak(marker, first), text).toBe(true);
      const find = new DocxFindController(() => 1, async () => runs);
      expect((await find.find('cat', { wholeWord: true })).map((match) => match.text), text).toEqual(['cat']);
      expect(await find.find('a', { wholeWord: true }), text).toHaveLength(1);
    },
  );

  it('lets the text range alone decide when the layout reports it', () => {
    const run = (text: string, x: number, y: number, start: number) => ({
      text, x, y, w: text.length * 5, h: 12, fontSize: 10, font: '10px serif',
      textRange: { start, end: start + text.length },
    });
    // Adjacent in the text: one word, also when justification or a wrap moved
    // the second run away, and in right-to-left order.
    expect(docxRunsBreak(run('cat', 0, 0, 0), run('alog', 15, 0, 3))).toBe(false);
    expect(docxRunsBreak(run('cat', 0, 0, 0), run('alog', 60, 0, 3))).toBe(false);
    expect(docxRunsBreak(run('cat', 0, 0, 0), run('alog', 0, 12, 3))).toBe(false);
    expect(docxRunsBreak(run('alog', 15, 0, 3), run('cat', 0, 0, 0))).toBe(false);
    // Something lies between them in the text, however close they are drawn.
    expect(docxRunsBreak(run('cat', 0, 0, 0), run('alog', 15, 0, 4))).toBe(true);
    // The spans of a numbering marker join each other, and not the text.
    expect(docxRunsBreak(run('1', 0, 0, -2), run('.', 5, 0, -1))).toBe(false);
    expect(docxRunsBreak(run('.', 5, 0, -1), run('cat', 10, 0, 0))).toBe(true);
  });

  it('falls back to positions for runs without a text range', () => {
    const run = (text: string, x: number, y: number) =>
      ({ text, x, y, w: text.length * 5, h: 12, fontSize: 10, font: '10px serif' });
    expect(docxRunsBreak(run('cat', 0, 0), run('alog', 15, 0))).toBe(false);
    expect(docxRunsBreak(run('cat', 0, 0), run('alog', 0, 12))).toBe(true);
    expect(docxRunsBreak(run('cat', 0, 0), run('alog', 60, 0))).toBe(true);
  });
});
