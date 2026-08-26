import { describe, expect, it } from 'vitest';
import type {
  BodyElement, DocParagraph, DocxDocumentModel, DocxTextRun, RunRevision, SectionProps,
} from './types.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { buildSegments, type LineLayoutEnvironment } from './line-layout.js';
import { normalizeInternalDocumentModel, type InternalDocParagraph } from './parser-model.js';

function makeCtx(): CanvasRenderingContext2D {
  return {
    font: '10px serif',
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * 10,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

type DocRun = DocParagraph['runs'][number];

function textRun(text: string, revision?: RunRevision): DocRun {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: null,
    fontFamily: 'NotInMetrics',
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
    ...(revision ? { revision } : {}),
  } as unknown as DocxTextRun & { type: 'text' };
}

function paragraph(runs: DocRun[]): BodyElement {
  return {
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
    runs,
    defaultFontSize: 20,
    defaultFontFamily: 'NotInMetrics',
    widowControl: false,
  } as unknown as BodyElement;
}

function documentModel(): DocxDocumentModel {
  const section = {
    pageWidth: 600,
    pageHeight: 400,
    marginTop: 20,
    marginRight: 20,
    marginBottom: 20,
    marginLeft: 20,
    headerDistance: 0,
    footerDistance: 0,
    titlePage: false,
    evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section,
    body: [paragraph([
      textRun('kept '),
      textRun('added', { kind: 'insertion', author: 'Alice' }),
      textRun('gone', { kind: 'deletion', author: 'Alice' }),
      textRun('moved-away', { kind: 'moveFrom', author: 'Bob' }),
      textRun('moved-in', { kind: 'moveTo', author: 'Bob' }),
    ])],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

describe('tracked changes final-state projection (§17.13.5)', () => {
  it('keeps insertion/moveTo and omits deletion/moveFrom from layout', () => {
    const model = documentModel();
    const layout = layoutDocument(
      model,
      createLayoutServices(model, { measureContext: makeCtx() }),
      { currentDateMs: 0 },
    );
    const text = layout.pages.flatMap((page) => page.layers.body).flatMap((node) =>
      node.kind === 'paragraph'
        ? node.lines.flatMap((line) => line.placements)
            .filter((placement) => placement.kind === 'text')
            .map((placement) => placement.text)
        : []).join('');
    expect(text).toBe('kept addedmoved-in');
  });

  it.each(['deletion', 'moveFrom'] as const)(
    'omits every inline run kind carried by a %s container',
    (kind) => {
      const revision: RunRevision = { kind, author: 'Reviewer' };
      const deletedRuns = [
        { type: 'break', breakType: 'line', revision },
        { type: 'field', revision },
        { type: 'math', revision },
        { type: 'ptab', revision },
        { type: 'image', revision },
        { type: 'chart', revision },
        { type: 'shape', revision },
        { type: 'anchorHost', revision },
      ] as unknown as DocRun[];

      expect(buildSegments(deletedRuns, {} as LineLayoutEnvironment)).toEqual([]);
    },
  );

  it.each(['insertion', 'moveTo'] as const)(
    'keeps non-text inline content carried by an accepted %s container',
    (kind) => {
      const revision: RunRevision = { kind, author: 'Reviewer' };
      const acceptedRuns = [
        { type: 'break', breakType: 'line', revision },
        {
          type: 'field', fieldType: 'other', instruction: '', fallbackText: 'field',
          bold: false, italic: false, underline: false, strikethrough: false,
          fontSize: 10, color: null, fontFamily: null, background: null,
          vertAlign: null, revision,
        },
        { type: 'math', nodes: [], display: false, fontSize: 10, revision },
        {
          type: 'ptab', alignment: 'left', relativeTo: 'margin', leader: 'none',
          fontSize: 10, revision,
        },
        {
          type: 'image', imagePath: 'word/media/image.png', mimeType: 'image/png',
          widthPt: 10, heightPt: 10, rotation: 0, flipH: false, flipV: false,
          revision,
        },
        {
          type: 'shape', inline: true, widthPt: 10, heightPt: 10,
          anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false,
          anchorYFromPara: false, revision,
        },
      ] as unknown as DocRun[];

      expect(buildSegments(acceptedRuns, {} as LineLayoutEnvironment)).toHaveLength(6);
    },
  );

  it('projects parser revision sidecars onto every public inline run kind', () => {
    const raw = documentModel();
    const paragraph = raw.body[0] as unknown as InternalDocParagraph;
    paragraph.runs = [
      { type: 'text' },
      { type: 'break', breakType: 'line' },
      { type: 'field' },
      { type: 'math', nodes: [], display: false, fontSize: 10 },
      { type: 'ptab' },
      { type: 'image' },
      { type: 'chart' },
      { type: 'shape' },
      { type: 'anchorHost' },
    ] as unknown as DocRun[];
    paragraph.__runRevisions = paragraph.runs.map(() => ({
      kind: 'deletion', author: 'Reviewer',
    }));

    const normalized = normalizeInternalDocumentModel(raw).document;
    const normalizedParagraph = normalized.body[0] as DocParagraph;
    expect(normalizedParagraph.runs.map((run) => run.revision?.kind)).toEqual(
      paragraph.runs.map(() => 'deletion'),
    );
    expect((normalizedParagraph as InternalDocParagraph).__runRevisions).toBeUndefined();
  });
});
