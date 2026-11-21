import { describe, expect, it } from 'vitest';
import { bodySectionIndexInput } from '../parser-model.js';
import {
  createBodySectionIndex,
  effectiveHeaderFooterInlineMargins,
  effectivePhysicalSectionGeometry,
  logicalSectionGeometry,
  sectionPageBox,
  sectionBodyInsetPt,
  physicalSectionGeometry,
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
  geom?: Partial<SectionGeom>;
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
    geom: input.geom as SectionGeom | undefined,
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

describe('section flow normalization', () => {
  const flowMarker = (textDirection: string | null, sectionFlow?: string): BodyElement => ({
    type: 'sectionBreak',
    kind: 'nextPage',
    columns: null,
    textDirection,
    pageNumType: null,
    __sectionPlacement: {
      sectionId: 'section:flow',
      vAlign: null,
      lineNumbering: null,
      ...(sectionFlow === undefined ? {} : { sectionFlow }),
    },
  }) as unknown as BodyElement;

  it('normalizes the section flow fact per section', () => {
    const doc = document(
      [paragraph('a'), flowMarker('btLr', 'bottomToTop'), paragraph('b')],
      { textDirection: 'btLr' },
    );
    const { occurrences } = bodySectionIndexInput(doc);
    expect(occurrences.map((occurrence) => occurrence.sectionFlow ?? null))
      .toEqual(['bottomToTop', null]);
    expect(occurrences.map((occurrence) => occurrence.textDirection)).toEqual(['btLr', 'btLr']);
  });

  it.each(['btLr', 'bottomToTop', 'sideways-lr'])
    ('never derives the counter-clockwise frame from an authored %s token', (textDirection) => {
      const doc = document(
        [paragraph('a'), flowMarker(textDirection), paragraph('b')],
        { textDirection },
      );
      for (const occurrence of bodySectionIndexInput(doc).occurrences) {
        expect(occurrence.sectionFlow).toBeUndefined();
      }
    });

  it.each([
    ['tbRl', 'bottomToTop'],
    [null, 'bottomToTop'],
    ['btLr', 'topToBottom'],
  ] as const)('rejects a wire whose direction is %s for section flow %s', (textDirection, flow) => {
    const doc = document([paragraph('a'), flowMarker(textDirection, flow), paragraph('b')]);
    expect(() => bodySectionIndexInput(doc)).toThrow(TypeError);
  });
});

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

  it('inherits omitted page geometry backward only for a continuous section', () => {
    const finalGeometry = geometry({ pageWidth: 700, marginLeft: 54 });
    const doc = document([
      paragraph('inherited'),
      marker({ sectionId: 'section:0', kind: 'continuous' }),
      paragraph('final'),
    ], finalGeometry);

    const index = createBodySectionIndex(bodySectionIndexInput(doc));

    expect(index.occurrences[0]?.geometry).toEqual(finalGeometry);
    expect(index.occurrences[0]?.headers).toEqual(EMPTY_HF);
    expect(index.occurrences[0]?.footers).toEqual(EMPTY_HF);
    expect(index.occurrences[0]?.titlePage).toBe(false);
  });

  it('inherits the following page box for omitted non-continuous geometry', () => {
    const doc = document([
      paragraph('defaulted'),
      marker({ sectionId: 'section:0', kind: 'nextPage' }),
      paragraph('final'),
    ], geometry({ pageWidth: 700, marginLeft: 54 }));

    expect(createBodySectionIndex(bodySectionIndexInput(doc)).occurrences[0]?.geometry)
      .toEqual(geometry({ pageWidth: 700, marginLeft: 54 }));
  });

  it('preserves authored fields and inherits omitted fields for a non-continuous section', () => {
    const doc = document([
      paragraph('partially authored'),
      marker({
        sectionId: 'section:0',
        kind: 'nextPage',
        geom: { pageWidth: 660, marginTop: -24 },
      }),
      paragraph('final'),
    ], geometry({ pageWidth: 700, marginTop: 48, marginLeft: 54 }));

    expect(createBodySectionIndex(bodySectionIndexInput(doc)).occurrences[0]?.geometry).toEqual({
      pageWidth: 660, pageHeight: 792,
      marginTop: -24, marginRight: 72, marginBottom: 72, marginLeft: 54,
      headerDistance: 36, footerDistance: 36,
    });
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
  const pagePolicy = (
    overrides: Partial<Parameters<typeof effectivePhysicalSectionGeometry>[0]> = {},
  ): Parameters<typeof effectivePhysicalSectionGeometry>[0] => ({
    physicalGeometry: geometry({ pageWidth: 792, pageHeight: 612 }),
    columns: null,
    textDirection: 'lrTb',
    gutterPt: 18,
    rtlGutter: false,
    sheetOrientation: 'portrait' as const,
    mirrorMargins: false,
    gutterAtTop: false,
    bookFoldPrinting: false,
    bookFoldRevPrinting: false,
    printTwoOnOne: false,
    ...overrides,
  });

  // `word-two-page-sheet-imposition`: Windows Word controls with left 90,
  // right 27, top 45, bottom 63 and a 21.6pt gutter, values chosen so that no
  // two candidate rules coincide. Each row lists the body margins
  // [left, right, top, bottom] and the header/footer [left, right] of physical
  // pages 1 and 2. The authored page box is never divided.
  type Margins = readonly [number, number, number, number];
  type Inline = readonly [number, number];
  it.each([
    ['landscape book fold', 'landscape', { bookFoldPrinting: true },
      [48.6, 90, 45, 63], [111.6, 27], [90, 48.6, 45, 63], [111.6, 27]],
    ['landscape reverse book fold', 'landscape', { bookFoldPrinting: true, bookFoldRevPrinting: true },
      [90, 48.6, 45, 63], [111.6, 27], [48.6, 90, 45, 63], [111.6, 27]],
    ['landscape book fold with gutterAtTop', 'landscape', { bookFoldPrinting: true, gutterAtTop: true },
      [48.6, 68.4, 66.6, 63], [90, 27], [90, 27, 66.6, 63], [90, 27]],
    ['landscape reverse book fold with gutterAtTop', 'landscape',
      { bookFoldRevPrinting: true, gutterAtTop: true },
      [90, 27, 66.6, 63], [90, 27], [48.6, 68.4, 66.6, 63], [90, 27]],
    ['landscape book fold with mirrorMargins', 'landscape', { bookFoldPrinting: true, mirrorMargins: true },
      [48.6, 90, 45, 63], [111.6, 27], [90, 48.6, 45, 63], [27, 111.6]],
    ['portrait book fold', 'portrait', { bookFoldPrinting: true },
      [90, 27, 45, 63], [111.6, 27], [90, 27, 45, 63], [111.6, 27]],
    ['landscape two-up', 'landscape', { printTwoOnOne: true },
      [48.6, 90, 45, 63], [48.6, 90], [90, 48.6, 45, 63], [90, 48.6]],
    ['landscape two-up with gutterAtTop', 'landscape', { printTwoOnOne: true, gutterAtTop: true },
      [48.6, 68.4, 66.6, 63], [48.6, 90], [90, 27, 66.6, 63], [90, 48.6]],
    ['portrait two-up', 'portrait', { printTwoOnOne: true },
      [90, 27, 84.6, 45], [90, 27], [90, 27, 45, 84.6], [90, 27]],
    ['portrait two-up with gutterAtTop and mirrorMargins', 'portrait',
      { printTwoOnOne: true, gutterAtTop: true, mirrorMargins: true },
      [90, 27, 84.6, 45], [90, 27], [90, 27, 45, 84.6], [90, 27]],
    // A gutter wider than the left margin: Word keeps the mirrored width and
    // runs past the page edge (right margin -11.6); columns stop at the edge.
    ['landscape book fold with gutterAtTop and a 10pt left margin', 'landscape',
      {
        bookFoldPrinting: true, gutterAtTop: true,
        physicalGeometry: geometry({
          pageWidth: 792, pageHeight: 612,
          marginLeft: 10, marginRight: 27, marginTop: 45, marginBottom: 63,
        }),
      },
      [48.6, 0, 66.6, 63], [10, 27], [10, 27, 66.6, 63], [10, 27]],
    // tbRl sections bind on the right.
    ['tbRl single page', 'landscape', { textDirection: 'tbRl' },
      [90, 48.6, 45, 63], [90, 48.6], [90, 48.6, 45, 63], [90, 48.6]],
    ['tbRl landscape book fold', 'landscape', { textDirection: 'tbRl', bookFoldPrinting: true },
      [90, 48.6, 45, 63], [90, 48.6], [90, 48.6, 45, 63], [90, 48.6]],
    ['tbRl landscape two-up', 'landscape', { textDirection: 'tbRl', printTwoOnOne: true },
      [90, 48.6, 45, 63], [90, 48.6], [48.6, 90, 45, 63], [48.6, 90]],
    ['tbRl portrait two-up', 'portrait', { textDirection: 'tbRl', printTwoOnOne: true },
      [90, 27, 84.6, 45], [90, 27], [90, 27, 45, 84.6], [90, 27]],
  ] as const)(
    'imposes the Word-observed %s margins on the authored page box',
    (_label, sheetOrientation, settings, ...pages) => {
      const policy = pagePolicy({
        physicalGeometry: geometry({
          pageWidth: 792, pageHeight: 612,
          marginLeft: 90, marginRight: 27, marginTop: 45, marginBottom: 63,
        }),
        gutterPt: 21.6,
        sheetOrientation,
        ...settings,
      });
      for (const pageIndex of [0, 1]) {
        const body = pages[pageIndex * 2] as Margins;
        const story = pages[pageIndex * 2 + 1] as Inline;
        const page = effectivePhysicalSectionGeometry(policy, pageIndex);
        const header = effectiveHeaderFooterInlineMargins(policy, pageIndex);
        expect([page.pageWidth, page.pageHeight]).toEqual([792, 612]);
        [page.marginLeft, page.marginRight, page.marginTop, page.marginBottom]
          .forEach((value, index) => expect(value).toBeCloseTo(body[index]!, 9));
        expect(header.marginLeft).toBeCloseTo(story[0], 9);
        expect(header.marginRight).toBeCloseTo(story[1], 9);
      }
    },
  );

  it.each([
    ['left edge', {}, 0, { marginTop: 72, marginRight: 72, marginLeft: 90 }],
    ['right edge', { rtlGutter: true }, 0, { marginTop: 72, marginRight: 90, marginLeft: 72 }],
    ['top edge', { gutterAtTop: true }, 0, { marginTop: 90, marginRight: 72, marginLeft: 72 }],
    ['first mirrored edge', { mirrorMargins: true }, 0, {
      marginTop: 72, marginRight: 72, marginLeft: 90,
    }],
    ['second mirrored edge', { mirrorMargins: true }, 1, {
      marginTop: 72, marginRight: 90, marginLeft: 72,
    }],
  ] as const)(
    'places the authored gutter on the %s',
    (_label, settings, pageIndex, expected) => {
      expect(effectivePhysicalSectionGeometry(pagePolicy(settings), pageIndex))
        .toMatchObject(expected);
    },
  );

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

  it('keeps the clockwise margin mapping and gives a native BtoT frame its own edges', () => {
    const physical = geometry({ marginTop: 36, marginRight: 54, marginBottom: 72, marginLeft: 90 });
    // Established Transitional vertical frame: logical left/top/right/bottom
    // are the physical top/right/bottom/left margins.
    expect(logicalSectionGeometry(physical)).toEqual(geometry({
      pageWidth: 792, pageHeight: 612,
      marginLeft: 36, marginTop: 54, marginRight: 72, marginBottom: 90,
    }));
    // Native BtoT: lines start at the physical bottom and later lines move
    // right, so logical left/right/top/bottom are physical bottom/top/left/right.
    const native = logicalSectionGeometry(physical, 'bottomToTop');
    expect(native).toEqual(geometry({
      pageWidth: 792, pageHeight: 612,
      marginLeft: 72, marginRight: 36, marginTop: 90, marginBottom: 54,
    }));
    expect(physicalSectionGeometry(native, 'bottomToTop')).toEqual(physical);
  });

  it('projects only page-box facts and preserves signed-margin body distance', () => {
    const props = document([], { marginTop: -36, marginBottom: -54 }).section;

    expect(sectionPageBox(props)).toEqual(geometry({ marginTop: -36, marginBottom: -54 }));
    expect(sectionBodyInsetPt(props.marginTop)).toBe(36);
    expect(sectionBodyInsetPt(props.marginBottom)).toBe(54);
  });
});
