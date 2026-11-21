import { describe, expect, it } from 'vitest';
import {
  type InternalDocParagraph,
  type InternalNumberingInfo,
  bodyAcquisitionInputProjections,
  normalizeInternalDocumentModel,
  numberingMarkerShapeInput,
  paragraphAcquisitionInput,
  paragraphMarkShapeInput,
  tableColumnLayoutInput,
  tableFormatInput,
  tableParticipatesInOrdinaryFlow,
  tableSourceAcquisitionInput,
} from './parser-model.js';
import type { DocParagraph, DocxDocumentModel } from './types.js';
import type { SourceRef } from './layout/types.js';

const paragraph = (): DocParagraph => ({
  type: 'paragraph',
  alignment: 'left',
  indentLeft: 0,
  indentRight: 0,
  indentFirst: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [{
    type: 'text',
    text: 'cached parser fact',
    fontSize: 10,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: null,
    fontFamily: null,
    isLink: false,
    background: null,
    vertAlign: null,
  }],
} as unknown as DocParagraph);

const document = (bodyParagraph: DocParagraph): DocxDocumentModel => ({
  section: {
    pageWidth: 612,
    pageHeight: 792,
    marginTop: 72,
    marginRight: 72,
    marginBottom: 72,
    marginLeft: 72,
    headerDistance: 36,
    footerDistance: 36,
  },
  body: [bodyParagraph],
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
} as unknown as DocxDocumentModel);

const bodySource = (path: number[] = [0]): SourceRef => ({
  story: 'body',
  storyInstance: 'body',
  path,
});

describe('parser-to-body-acquisition projection capability', () => {
  it('is one frozen identity-preserving record without compatibility wrappers', () => {
    expect(Object.isFrozen(bodyAcquisitionInputProjections)).toBe(true);
    // The boundary gate pins capability names; this test checks runtime
    // immutability and direct projection identities rather than its key list.
    expect(bodyAcquisitionInputProjections.numberingMarkerShapeInput)
      .toBe(numberingMarkerShapeInput);
    expect(bodyAcquisitionInputProjections.paragraphAcquisitionInput)
      .toBe(paragraphAcquisitionInput);
    expect(bodyAcquisitionInputProjections.paragraphMarkShapeInput)
      .toBe(paragraphMarkShapeInput);
    expect(bodyAcquisitionInputProjections.tableColumnLayoutInput)
      .toBe(tableColumnLayoutInput);
    expect(bodyAcquisitionInputProjections.tableFormatInput).toBe(tableFormatInput);
    expect(bodyAcquisitionInputProjections.tableSourceAcquisitionInput).toBe(tableSourceAcquisitionInput);
    expect(bodyAcquisitionInputProjections.tableParticipatesInOrdinaryFlow)
      .toBe(tableParticipatesInOrdinaryFlow);
  });

  it.each([
    { threshold: undefined, expected: false },
    { threshold: 0, expected: true },
    { threshold: 8, expected: true },
  ])('preserves unmeasured marker/mark zero semantics with threshold $threshold', ({ threshold, expected }) => {
    const facts = { fontSize: 10, fontFamily: 'serif', kerning: threshold };
    const source: InternalDocParagraph = { ...paragraph(), paragraphMarkFontFacts: facts };
    const numbering: InternalNumberingInfo = {
      numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 0, tab: 0, suff: 'space',
      fontFacts: facts,
    };
    expect(paragraphMarkShapeInput(source)?.kerning).toBe(expected);
    expect(numberingMarkerShapeInput(numbering, 10)?.kerning).toBe(expected);
  });

  it('reuses only one document-scoped paragraph/source parser-fact projection', () => {
    const modelParagraph = paragraph();
    const firstDocument = normalizeInternalDocumentModel(document(modelParagraph));
    const project = firstDocument.bodyModelGateway.acquisitionInputs.paragraphAcquisitionInput;

    const first = project(modelParagraph, bodySource());
    const equivalentSource = project(modelParagraph, bodySource());
    const otherSource = project(modelParagraph, bodySource([1]));
    const equivalentParagraph = project(paragraph(), bodySource());

    expect(equivalentSource).toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(otherSource).not.toBe(first);
    expect(equivalentParagraph).not.toBe(first);

    const secondDocument = normalizeInternalDocumentModel(document(modelParagraph));
    expect(secondDocument.bodyModelGateway.acquisitionInputs.paragraphAcquisitionInput(
      modelParagraph,
      bodySource(),
    )).not.toBe(first);
  });

  it('keeps unavailable drawing geometry clone-safe without widening the public run union', () => {
    const inputParagraph = paragraph();
    const textRun = inputParagraph.runs[0]!;
    inputParagraph.runs = [
      textRun,
      {
        type: 'unavailableDrawing',
        resourceKind: 'image',
        widthPt: 24,
        heightPt: 12,
      },
      { ...textRun, text: 'after unavailable drawing' },
    ] as unknown as DocParagraph['runs'];

    const raw = document(inputParagraph);
    const normalized = normalizeInternalDocumentModel(raw);
    const publicParagraph = normalized.document.body[0] as DocParagraph;
    const acquired = normalized.bodyModelGateway.acquisitionInputs
      .paragraphAcquisitionInput(publicParagraph, bodySource());

    expect(publicParagraph.runs.map((run) => run.type))
      .toEqual(['text', 'text']);
    expect(JSON.stringify(structuredClone(normalized.document)))
      .not.toContain('unavailableDrawing');
    expect(acquired.runs.map((run) => run.type))
      .toEqual(['text', 'unavailableDrawing', 'text']);
    expect(acquired.runs[1]).toMatchObject({
      resourceKind: 'image',
      widthPt: 24,
      heightPt: 12,
    });
    const cloned = normalizeInternalDocumentModel(
      structuredClone(raw),
    );
    const clonedParagraph = cloned.document.body[0] as DocParagraph;
    expect(cloned.bodyModelGateway.acquisitionInputs
      .paragraphAcquisitionInput(clonedParagraph, bodySource()).runs)
      .toEqual(acquired.runs);
  });

  it('projects noBreakHyphen wire provenance into immutable parser-independent ranges', () => {
    const inputParagraph = paragraph();
    Object.assign(inputParagraph.runs[0]!, {
      text: 'a-b',
      __noBreakBefore: true,
      __noBreakAfter: true,
      __noBreakHyphenOffsets: [2],
    });

    const acquired = paragraphAcquisitionInput(inputParagraph, bodySource());
    expect(acquired.runs[0]).toMatchObject({
      text: 'a-b',
      noBreakBefore: true,
      noBreakAfter: true,
      noBreakRanges: [{ start: 1, end: 2 }],
    });
    expect(acquired.runs[0]).not.toHaveProperty('__noBreakBefore');
    expect(acquired.runs[0]).not.toHaveProperty('__noBreakAfter');
    expect(acquired.runs[0]).not.toHaveProperty('__noBreakHyphenOffsets');
    expect(Object.isFrozen((acquired.runs[0] as { noBreakRanges: readonly object[] }).noBreakRanges[0]))
      .toBe(true);
  });
});

describe('inline object run position ownership (§17.3.2.24)', () => {
  it('keeps the parser-private position off the public model and hands it to acquisition', () => {
    const image = {
      type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png', widthPt: 40, heightPt: 16,
      anchor: false, anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
      __inlinePosition: -15,
    };
    const doc = {
      body: [{ ...paragraph(), runs: [image] }],
      section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72, marginBottom: 72,
        marginLeft: 72, headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false },
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;
    const normalized = normalizeInternalDocumentModel(doc).document;
    const publicParagraph = normalized.body[0] as DocParagraph;
    const publicRun = publicParagraph.runs[0]!;
    expect(Object.keys(publicRun)).not.toContain('__inlinePosition');
    expect(JSON.stringify(normalized)).not.toContain('__inlinePosition');
    const source: SourceRef = { story: 'body', storyInstance: 'body', path: [0] };
    const acquired = paragraphAcquisitionInput(publicParagraph, source);
    expect((acquired.runs[0] as { inlinePositionPt?: number }).inlinePositionPt).toBe(-15);
    // Normalizing the public model again keeps the fact on the same run.
    const again = normalizeInternalDocumentModel(normalized).document.body[0] as DocParagraph;
    expect((paragraphAcquisitionInput(again, source).runs[0] as { inlinePositionPt?: number }).inlinePositionPt).toBe(-15);
  });

  it('keeps the position when normalization replaces a shape whose text box holds math', () => {
    const shape = {
      type: 'shape', inline: true, widthPt: 40, heightPt: 16, anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: false, zOrder: 0, subpaths: [], presetGeometry: 'rect',
      fill: null, stroke: null, textBlocks: [],
      textBoxContent: [{ ...paragraph(), runs: [{ type: 'math', display: false, nodes: [] }] }],
      __inlinePosition: -5,
    };
    const doc = {
      body: [{ ...paragraph(), runs: [shape] }],
      section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72, marginBottom: 72,
        marginLeft: 72, headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false },
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;
    const source: SourceRef = { story: 'body', storyInstance: 'body', path: [0] };
    const positionOf = (model: DocxDocumentModel) => (paragraphAcquisitionInput(
      model.body[0] as DocParagraph, source,
    ).runs[0] as { inlinePositionPt?: number }).inlinePositionPt;
    const first = normalizeInternalDocumentModel(doc).document;
    expect(JSON.stringify(first)).not.toContain('__inlinePosition');
    expect(positionOf(first)).toBe(-5);
    const second = normalizeInternalDocumentModel(first).document;
    // The nested math gets a new identity on every pass, so the shape is
    // cloned again: the side-table fact must follow the new run.
    expect((second.body[0] as DocParagraph).runs[0]).not.toBe((first.body[0] as DocParagraph).runs[0]);
    expect(positionOf(second)).toBe(-5);
    expect(positionOf(normalizeInternalDocumentModel(second).document)).toBe(-5);
  });
});
