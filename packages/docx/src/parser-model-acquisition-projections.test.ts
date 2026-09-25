import { describe, expect, it } from 'vitest';
import {
  bodyAcquisitionInputProjections,
  normalizeInternalDocumentModel,
  numberingMarkerShapeInput,
  paragraphAcquisitionInput,
  paragraphMarkShapeInput,
  tableColumnLayoutInput,
  tableFormatInput,
  tableParticipatesInOrdinaryFlow,
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
    expect(Object.keys(bodyAcquisitionInputProjections).sort()).toEqual([
      'numberingMarkerShapeInput',
      'paragraphAcquisitionInput',
      'paragraphMarkShapeInput',
      'tableColumnLayoutInput',
      'tableFormatInput',
      'tableParticipatesInOrdinaryFlow',
    ]);
    expect(bodyAcquisitionInputProjections.numberingMarkerShapeInput)
      .toBe(numberingMarkerShapeInput);
    expect(bodyAcquisitionInputProjections.paragraphAcquisitionInput)
      .toBe(paragraphAcquisitionInput);
    expect(bodyAcquisitionInputProjections.paragraphMarkShapeInput)
      .toBe(paragraphMarkShapeInput);
    expect(bodyAcquisitionInputProjections.tableColumnLayoutInput)
      .toBe(tableColumnLayoutInput);
    expect(bodyAcquisitionInputProjections.tableFormatInput).toBe(tableFormatInput);
    expect(bodyAcquisitionInputProjections.tableParticipatesInOrdinaryFlow)
      .toBe(tableParticipatesInOrdinaryFlow);
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
