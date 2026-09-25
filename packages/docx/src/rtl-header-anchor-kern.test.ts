import { describe, expect, it } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

function paragraph(runs: DocParagraph['runs'], bidi = false): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 14, defaultFontFamily: 'serif', widowControl: false, bidi,
  } as DocParagraph;
}

function anchorFacts(occurrenceId: string) {
  const missingEdges = {
    topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
    bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing',
  };
  return {
    occurrenceId,
    simplePosition: {
      enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid',
    },
    horizontal: {
      relativeFrom: 'column', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 80 },
    },
    vertical: {
      relativeFrom: 'paragraph', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 0 },
    },
    extent: { widthPt: 40, heightPt: 12, widthStatus: 'valid', heightStatus: 'valid' },
    parentEffectExtent: missingEdges,
    anchorDistances: missingEdges,
    relativeSize: { horizontal: null, vertical: null },
    wrap: {
      kind: 'none', authoredKinds: ['wrapNone'], side: null,
      distances: missingEdges, effectExtent: null, polygon: null,
    },
    behavior: {
      behindDoc: false, behindDocStatus: 'valid',
      relativeHeight: 1, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid',
      allowOverlap: true, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
  } as const;
}

describe('RTL header anchor kerning projection', () => {
  it('retains the anchored image while an absent w:kern disables optional pair kerning', () => {
    const occurrenceId = 'wp-anchor-rtl-header';
    const facts = anchorFacts(occurrenceId);
    const header = paragraph([
      {
        type: 'anchorHost', fontSize: 14, fontFamily: 'serif',
        __anchorOccurrenceId: occurrenceId,
      },
      {
        type: 'image', imagePath: 'word/media/header.png', mimeType: 'image/png',
        widthPt: 40, heightPt: 12, anchor: true,
        __anchorAcquisition: facts,
      },
      {
        type: 'text', text: 'نص عنوان', bold: false, italic: false,
        underline: false, strikethrough: false, fontSize: 14,
        color: null, fontFamily: 'serif', fontFamilyCs: 'serif', fontSizeCs: 14,
        isLink: false, background: null, vertAlign: null, hyperlink: null,
        rtl: true, cs: true,
      },
    ] as unknown as DocParagraph['runs'], true);
    const body = paragraph([{
      type: 'text', text: 'body', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'serif',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }] as unknown as DocParagraph['runs']);
    const section = {
      pageWidth: 200, pageHeight: 160,
      marginTop: 30, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [body as unknown as BodyElement],
      headers: {
        default: { body: [header as unknown as BodyElement] }, first: null, even: null,
      },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    let fontKerning: CanvasFontKerning = 'auto';
    const rtlMeasurementStates: CanvasFontKerning[] = [];
    const measureContext = {
      font: '10px serif', letterSpacing: '0px',
      get fontKerning() { return fontKerning; },
      set fontKerning(value: CanvasFontKerning) { fontKerning = value; },
      measureText(text: string) {
        if (/\p{Script=Arabic}/u.test(text)) rtlMeasurementStates.push(fontKerning);
        return {
          width: [...text].length * 6,
          actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
          fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;

    const layout = layoutDocument(
      model,
      createLayoutServices(model, { measureContext }),
      { currentDateMs: 0 },
    );
    const headerParagraph = layout.pages[0]?.layers.header.find(
      (node) => node.kind === 'paragraph',
    );

    expect(headerParagraph?.kind).toBe('paragraph');
    if (headerParagraph?.kind !== 'paragraph') throw new Error('expected header paragraph');
    expect(headerParagraph.drawings).toEqual([
      expect.objectContaining({
        commands: [expect.objectContaining({ kind: 'resource', resourceKind: 'image' })],
      }),
    ]);
    expect(rtlMeasurementStates.length).toBeGreaterThan(0);
    // Font discovery may probe under Canvas `auto`; retained WML text must be
    // measured under the spec's absent-kerning default.
    expect(rtlMeasurementStates).toContain('none');
  });
});
