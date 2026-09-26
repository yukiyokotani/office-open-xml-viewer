import { describe, it, expect } from 'vitest';
import { formatNoteNumber } from './line-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement, DocNote, DocParagraph, DocxDocumentModel, SectionProps,
} from './types';

// ECMA-376 §17.11.17/.18 numFmt and §17.11.20 numStart: document-wide note
// numbering format and starting value for automatic reference marks, applied
// to both the body reference and the in-note number.

const TEST_FONT = 'Times New Roman';

interface Call { text: string; y: number; }
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    strokeText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function textRun(text: string, extra: Record<string, unknown> = {}) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false,
    strikethrough: false, fontSize: 10, color: null, fontFamily: TEST_FONT,
    fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...extra,
  };
}

function para(runs: Array<Record<string, unknown>>): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: runs as DocParagraph['runs'],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT, widowControl: false,
  } as unknown as DocParagraph;
}

function docWith(
  body: BodyElement[],
  footnotes: DocNote[],
  noteLayoutSettings?: Record<string, unknown>,
): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 400,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage',
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
    footnotes,
    ...(noteLayoutSettings ? { __noteLayoutSettings: noteLayoutSettings } : {}),
  } as unknown as DocxDocumentModel;
}

async function renderPage0(doc: DocxDocumentModel): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400 });
  return calls;
}

describe('note numbering format and start (ECMA-376 §17.11.17/.18/.20)', () => {
  it('formats ordinals with the note kind numbering', () => {
    expect(formatNoteNumber(2, undefined)).toBe('2');
    expect(formatNoteNumber(1, { format: 'lowerRoman', start: 1 })).toBe('i');
    expect(formatNoteNumber(3, { format: 'upperLetter', start: 4 })).toBe('F');
  });

  function notesDoc(settings?: Record<string, unknown>) {
    const footnotes: DocNote[] = [
      { id: '1', content: [para([
        textRun('', { noteRef: { kind: 'footnote', id: '' }, vertAlign: 'super' }),
        textRun(' NOTEA'),
      ]) as unknown as BodyElement] },
      { id: '2', content: [para([
        textRun('', { noteRef: { kind: 'footnote', id: '' }, vertAlign: 'super' }),
        textRun(' NOTEB'),
      ]) as unknown as BodyElement] },
    ];
    const body: BodyElement[] = [para([
      textRun('BODY'),
      textRun('1', { noteRef: { kind: 'footnote', id: '1' }, vertAlign: 'super' }),
      textRun(' MORE'),
      textRun('2', { noteRef: { kind: 'footnote', id: '2' }, vertAlign: 'super' }),
    ]) as unknown as BodyElement];
    return docWith(body, footnotes, settings);
  }

  it('keeps decimal numbering from 1 without authored settings', async () => {
    const texts = (await renderPage0(notesDoc())).map((call) => call.text);
    expect(texts.filter((text) => text === '1')).toHaveLength(2);
    expect(texts.filter((text) => text === '2')).toHaveLength(2);
  });

  it('applies the document-wide footnote format and start to references and notes', async () => {
    const texts = (await renderPage0(notesDoc({
      footnoteNumberFormat: 'lowerRoman',
      footnoteNumberStart: 3,
    }))).map((call) => call.text);
    expect(texts.filter((text) => text === 'iii')).toHaveLength(2);
    expect(texts.filter((text) => text === 'iv')).toHaveLength(2);
    expect(texts).not.toContain('1');
  });
});
