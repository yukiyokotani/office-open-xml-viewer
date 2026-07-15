import { describe, expect, it } from 'vitest';
import { bodySectionIndexInput } from '../parser-model.js';
import {
  createBodySectionIndex,
  logicalSectionGeometry,
  physicalSectionGeometry,
  sectionBodyInsetPt,
  sectionGeometry,
  type BodySectionOccurrence,
} from './context.js';
import type {
  BodyElement,
  ColumnsSpec,
  DocxDocumentModel,
  HeadersFooters,
  LineNumbering,
  PageNumType,
  SectionGeom,
  SectionProps,
} from '../types.js';

interface PrivateSectionPlacementWire {
  readonly sectionId: string;
  readonly vAlign: string | null;
  readonly lineNumbering: LineNumbering | null;
}

type PrivateSectionBreak = Extract<BodyElement, { type: 'sectionBreak' }> & {
  readonly __sectionPlacement: PrivateSectionPlacementWire;
};

const EMPTY_HF: HeadersFooters = { default: null, first: null, even: null };

function paragraph(label: string): BodyElement {
  return { type: 'paragraph', runs: [{ type: 'text', text: label }] } as BodyElement;
}

function geometry(overrides: Partial<SectionGeom> = {}): SectionGeom {
  return {
    pageWidth: 612,
    pageHeight: 792,
    marginTop: 72,
    marginRight: 72,
    marginBottom: 72,
    marginLeft: 72,
    headerDistance: 36,
    footerDistance: 36,
    ...overrides,
  };
}

function columns(count: number): ColumnsSpec {
  return {
    count,
    spacePt: 18,
    equalWidth: true,
    sep: false,
    cols: [],
  };
}

function marker(input: Readonly<{
  sectionId: string;
  kind: string;
  columns?: ColumnsSpec | null;
  geom?: SectionGeom;
  textDirection?: string | null;
  pageNumType?: PageNumType | null;
  headers?: HeadersFooters;
  footers?: HeadersFooters;
  titlePage?: boolean;
  vAlign?: string | null;
  lineNumbering?: LineNumbering | null;
}>): BodyElement {
  return {
    type: 'sectionBreak',
    kind: input.kind,
    columns: input.columns ?? null,
    geom: input.geom,
    textDirection: input.textDirection ?? null,
    pageNumType: input.pageNumType ?? null,
    headers: input.headers,
    footers: input.footers,
    titlePage: input.titlePage,
    __sectionPlacement: {
      sectionId: input.sectionId,
      vAlign: input.vAlign ?? null,
      lineNumbering: input.lineNumbering ?? null,
    },
  } as PrivateSectionBreak;
}

function document(body: BodyElement[], overrides: Partial<SectionProps> = {}): DocxDocumentModel {
  return {
    body,
    section: {
      ...geometry(),
      titlePage: false,
      evenAndOddHeaders: false,
      sectionStart: 'nextPage',
      columns: null,
      pageNumType: null,
      textDirection: null,
      vAlign: null,
      lineNumbering: null,
      ...overrides,
    },
    headers: EMPTY_HF,
    footers: EMPTY_HF,
    fontFamilyClasses: {},
  } as DocxDocumentModel;
}

function ids(occurrences: readonly BodySectionOccurrence[]): string[] {
  return occurrences.map((occurrence) => occurrence.sectionOccurrenceId);
}

describe('pre-indexed body section ownership', () => {
  it('assigns each paragraph-owned marker to the section it terminates', () => {
    const doc = document([
      paragraph('first'),
      marker({ sectionId: 'section:cover', kind: 'nextPage' }),
      paragraph('middle'),
      marker({ sectionId: 'section:middle', kind: 'continuous' }),
      paragraph('final'),
    ], { sectionStart: 'oddPage' });

    const index = createBodySectionIndex(bodySectionIndexInput(doc));

    expect(ids(index.occurrences)).toEqual([
      'section:cover',
      'section:middle',
      'section:2',
    ]);
    expect(index.sectionAtBodyIndex(0)).toBe(index.occurrences[0]);
    expect(index.sectionAtBodyIndex(1)).toBe(index.occurrences[0]);
    expect(index.sectionAtBodyIndex(2)).toBe(index.occurrences[1]);
    expect(index.sectionAtBodyIndex(3)).toBe(index.occurrences[1]);
    expect(index.sectionAtBodyIndex(4)).toBe(index.occurrences[2]);
    expect(index.sectionAtBodyIndex(doc.body.length)).toBe(index.occurrences[2]);
    expect(index.occurrences.map(({ startType }) => startType)).toEqual([
      'nextPage',
      'continuous',
      'oddPage',
    ]);
  });

  it('retains every section-scoped layout fact from its owning sectPr projection', () => {
    const endingHeaders = { ...EMPTY_HF };
    const endingFooters = { ...EMPTY_HF };
    const endingGeometry = geometry({ pageWidth: 792, pageHeight: 612, marginTop: 48 });
    const endingColumns = columns(3);
    const endingNumbering = { start: 7, fmt: 'upperRoman' };
    const endingLineNumbering: LineNumbering = {
      countBy: 2,
      start: 5,
      distance: 10,
      restart: 'newSection',
    };
    const finalHeaders = { ...EMPTY_HF };
    const finalFooters = { ...EMPTY_HF };
    const finalColumns = columns(2);
    const finalNumbering = { start: 20, fmt: 'decimal' };
    const doc = document([
      paragraph('ending'),
      marker({
        sectionId: 'section:landscape',
        kind: 'evenPage',
        columns: endingColumns,
        geom: endingGeometry,
        textDirection: 'tbRl',
        pageNumType: endingNumbering,
        headers: endingHeaders,
        footers: endingFooters,
        titlePage: true,
        vAlign: 'center',
        lineNumbering: endingLineNumbering,
      }),
      paragraph('final'),
    ], {
      sectionStart: 'continuous',
      columns: finalColumns,
      pageNumType: finalNumbering,
      textDirection: 'btLr',
      titlePage: true,
      vAlign: 'bottom',
      lineNumbering: { countBy: 1, start: 9, restart: 'newPage' },
    });
    doc.headers = finalHeaders;
    doc.footers = finalFooters;

    const [ending, final] = createBodySectionIndex(bodySectionIndexInput(doc)).occurrences;

    expect(ending).toMatchObject({
      sectionOccurrenceId: 'section:landscape',
      ordinal: 0,
      startBodyIndex: 0,
      endBodyIndex: 1,
      markerBodyIndex: 1,
      final: false,
      startType: 'evenPage',
      columns: endingColumns,
      geometry: endingGeometry,
      textDirection: 'tbRl',
      pageNumType: endingNumbering,
      headers: endingHeaders,
      footers: endingFooters,
      titlePage: true,
      vAlign: 'center',
      lineNumbering: endingLineNumbering,
    });
    expect(final).toMatchObject({
      sectionOccurrenceId: 'section:1',
      ordinal: 1,
      startBodyIndex: 2,
      endBodyIndex: 2,
      markerBodyIndex: null,
      final: true,
      startType: 'continuous',
      columns: finalColumns,
      textDirection: 'btLr',
      pageNumType: finalNumbering,
      headers: finalHeaders,
      footers: finalFooters,
      titlePage: true,
      vAlign: 'bottom',
      lineNumbering: { countBy: 1, start: 9, restart: 'newPage' },
    });
  });

  it('uses the final physical page geometry when a non-final sectPr inherits its page box', () => {
    const finalGeometry = geometry({ pageWidth: 700, marginLeft: 54 });
    const doc = document([
      paragraph('inherited'),
      marker({ sectionId: 'section:0', kind: 'nextPage' }),
      paragraph('final'),
    ], finalGeometry);

    const index = createBodySectionIndex(bodySectionIndexInput(doc));

    expect(index.occurrences[0]?.geometry).toEqual(finalGeometry);
    expect(index.occurrences[0]?.headers).toEqual(EMPTY_HF);
    expect(index.occurrences[0]?.footers).toEqual(EMPTY_HF);
    expect(index.occurrences[0]?.titlePage).toBe(false);
  });

  it('serves lookups from the built index without rescanning a subsequently changed body', () => {
    const doc = document([
      paragraph('first'),
      marker({ sectionId: 'section:0', kind: 'nextPage' }),
      paragraph('final'),
    ]);
    const index = createBodySectionIndex(bodySectionIndexInput(doc));
    const final = index.sectionAtBodyIndex(2);

    doc.body.splice(0, doc.body.length);

    expect(index.sectionAtBodyIndex(2)).toBe(final);
    expect(() => index.sectionAtBodyIndex(-1)).toThrow(RangeError);
    expect(() => index.sectionAtBodyIndex(4)).toThrow(RangeError);
  });
});

describe('section geometry coordinate boundary', () => {
  it('round-trips a physical page box through the vertical logical frame', () => {
    const physical = geometry({
      pageWidth: 612,
      pageHeight: 792,
      marginTop: 36,
      marginRight: 54,
      marginBottom: 72,
      marginLeft: 90,
    });

    expect(physicalSectionGeometry(logicalSectionGeometry(physical))).toEqual(physical);
  });

  it('projects only page-box facts and preserves signed-margin body distance', () => {
    const props = document([], { marginTop: -36, marginBottom: -54 }).section;

    expect(sectionGeometry(props)).toEqual(geometry({ marginTop: -36, marginBottom: -54 }));
    expect(sectionBodyInsetPt(props.marginTop)).toBe(36);
    expect(sectionBodyInsetPt(props.marginBottom)).toBe(54);
  });
});
