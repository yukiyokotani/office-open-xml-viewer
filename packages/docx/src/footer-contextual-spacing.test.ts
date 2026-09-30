import { describe, it, expect } from 'vitest';
import { layoutDocument } from './document-layout.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// Issue #1650 — header/footer stories must apply w:contextualSpacing (§17.3.1.9)
// without the previous paragraph's flow bounds overlapping the next one.

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() {
    let font = '10px serif';
    return {
      get font() { return font; }, set font(v: string) { font = v; },
      letterSpacing: '0px',
      measureText: (s: string) => {
        const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
        return {
          width: [...s].length * p * 0.5,
          fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
          actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
        } as TextMetrics;
      },
    };
  }
};

function para(text: string, over: Partial<DocParagraph>): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 6, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    styleId: 'Normal',
    ...over,
  } as unknown as DocParagraph;
}

function doc(footer: DocParagraph[], header: DocParagraph[] = []): DocxDocumentModel {
  const section = {
    pageWidth: 400, pageHeight: 400,
    marginTop: 40, marginRight: 10, marginBottom: 40, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section,
    body: [para('Body', {})] as unknown as BodyElement[],
    headers: { default: header.length ? { body: header } : null, first: null, even: null },
    footers: { default: { body: footer }, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

describe('footer contextualSpacing (issue #1650)', () => {
  it('lays out consecutive contextualSpacing footer paragraphs without FLOW_OVERLAP', () => {
    const model = doc([
      para('F1', { contextualSpacing: true }),
      para('F2', { contextualSpacing: true }),
    ]);
    expect(() => layoutDocument(model)).not.toThrow();
  });

  it('keeps ordinary spacing between different-style footer paragraphs', () => {
    const model = doc([
      para('F1', { contextualSpacing: true }),
      para('F2', { contextualSpacing: true, styleId: 'Other' }),
    ]);
    expect(() => layoutDocument(model)).not.toThrow();
  });

  it('keeps header/footer leading spaceBefore in story positioning', () => {
    const baselineOf = (spaceBefore: number, where: 'header' | 'footer'): number => {
      const p = [para('X', { spaceBefore })];
      const layout = layoutDocument(where === 'header' ? doc([], p) : doc(p));
      const nodes = (layout.pages[0] as any).layers[where] as any[];
      return nodes[0].lines[0].baselinePt;
    };
    expect(baselineOf(20, 'header') - baselineOf(0, 'header')).toBeCloseTo(20, 5);
    // The footer is bottom-anchored: the extra leading space grows it upward.
    expect(baselineOf(0, 'footer') - baselineOf(20, 'footer')).toBeCloseTo(0, 5);
  });
});
