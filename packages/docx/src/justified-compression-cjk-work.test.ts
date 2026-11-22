import { describe, expect, it, vi } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocxDocumentModel, SectionProps } from './types';

// Work bound for the distribute/lowKashida CJK refusal of
// WORD_JUSTIFIED_INTERWORD_COMPRESSION: a refused multi-piece CJK unit is
// recorded once, so its later pieces are not collected again on that line.
const collected = vi.hoisted(() => ({ followers: 0 }));
vi.mock('./line-breaker/line-gaps.js', async (importOriginal) => {
  const original = await importOriginal<typeof import('./line-breaker/line-gaps.js')>();
  return {
    ...original,
    candidateUnit: (...args: Parameters<typeof original.candidateUnit>) => {
      const unit = original.candidateUnit(...args);
      collected.followers += unit.length;
      return unit;
    },
  };
});

function canvas(): HTMLCanvasElement {
  let font = '12px Test';
  const width = (text: string) => [...text].length * Number(/([\d.]+)px/.exec(font)?.[1] ?? 12);
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: width(text), fontBoundingBoxAscent: 10, fontBoundingBoxDescent: 3,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      actualBoundingBoxLeft: 0, actualBoundingBoxRight: width(text),
    }) as TextMetrics,
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {}, scale() {},
    translate() {}, setTransform() {}, setLineDash() {}, drawImage() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {}, fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left', direction: 'ltr', globalAlpha: 1, lineCap: 'butt', lineJoin: 'miter',
  };
  return { width: 0, height: 0, style: {}, getContext: () => ctx } as unknown as HTMLCanvasElement;
}

async function followersFor(runCount: number, alignment: string): Promise<number> {
  const runs = Array.from({ length: runCount }, (_, index) => ({
    type: 'text', text: '丙', bold: index % 2 === 1, italic: false, underline: false,
    strikethrough: false, fontSize: 12, color: null, fontFamily: 'Test', isLink: false,
    background: null, vertAlign: null, hyperlink: null,
  }));
  const element = {
    type: 'paragraph', alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ ...runs[0], text: 'Lead ' }, ...runs],
    defaultFontSize: 12, defaultFontFamily: 'Test', widowControl: false,
  } as unknown as BodyElement;
  const widthPt = 400;
  const doc = {
    section: {
      pageWidth: widthPt, pageHeight: 4000, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    settings: { compatibilityMode: 15 },
    body: [element],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
  collected.followers = 0;
  await renderDocumentToCanvas(doc, canvas(), 0, { dpr: 1, width: widthPt });
  return collected.followers;
}

describe('distribute CJK compression refusal work', () => {
  it('collects styled CJK pieces no more often than a justified paragraph does', async () => {
    // A refused multi-piece unit is recorded once, as the both/justify
    // refusal records it; without that record the distribute/lowKashida
    // refusal re-collected every later piece (about 30 times the work here).
    const justified = await followersFor(512, 'both');
    for (const alignment of ['distribute', 'lowKashida']) {
      expect.soft(await followersFor(512, alignment), alignment).toBeLessThanOrEqual(justified);
    }
  });
});
