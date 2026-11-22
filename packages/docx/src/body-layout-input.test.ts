import { describe, expect, it } from 'vitest';
import { bodyLayoutAcquisitionInput } from './parser-model.js';
import { projectBodyLayoutInput } from './layout/body-layout-input.js';
import type { BodyElement, DocxDocumentModel, SectionProps } from './types.js';

const paragraph = (text: string, spaceBefore = 0, spaceAfter = 0): BodyElement => ({
  type: 'paragraph',
  runs: [{ type: 'text', text }],
  alignment: 'left',
  indentLeft: 0,
  indentRight: 0,
  indentFirst: 0,
  spaceBefore,
  spaceAfter,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
} as unknown as BodyElement);

const finalSection = (overrides: Partial<SectionProps> = {}): SectionProps => ({
  pageWidth: 612,
  pageHeight: 792,
  marginTop: 72,
  marginRight: 72,
  marginBottom: 72,
  marginLeft: 72,
  headerDistance: 36,
  footerDistance: 36,
  titlePage: false,
  evenAndOddHeaders: false,
  sectionStart: 'nextPage',
  columns: null,
  pageNumType: null,
  textDirection: null,
  vAlign: null,
  lineNumbering: null,
  ...overrides,
});

const createBodyLayoutInput = (document: DocxDocumentModel) =>
  projectBodyLayoutInput(bodyLayoutAcquisitionInput(document));

describe('canonical body layout input', () => {
  it('distinguishes textual paragraphs and empty marks from inline objects', () => {
    const image = { type: 'image', src: 'word/media/picture.png' };
    const withRuns = (runs: readonly unknown[]): BodyElement => ({
      ...paragraph(''), runs,
    }) as unknown as BodyElement;
    const document = {
      body: [
        paragraph('body text'),
        withRuns([image]),
        withRuns([{ type: 'text', text: 'caption' }, image]),
        paragraph(''),
        withRuns([{ type: 'break', breakType: 'line' }]),
      ],
      section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as DocxDocumentModel;

    expect(bodyLayoutAcquisitionInput(document).sequence.map((entry) =>
      entry.kind === 'body-block' && entry.block.kind === 'paragraph'
        ? entry.block.onlyTextualContent
        : undefined,
    )).toEqual([true, false, false, true, true]);
  });

  it('acquires clone-safe parser facts before resolving layout section owners', () => {
    const document = {
      body: [
        paragraph('first'),
        {
          type: 'sectionBreak',
          kind: 'continuous',
          geom: { ...finalSection(), pageWidth: 500 },
          columns: null,
          textDirection: null,
          pageNumType: null,
          headers: { default: null, first: null, even: null },
          footers: { default: null, first: null, even: null },
          titlePage: false,
        },
        paragraph('second'),
      ],
      section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as DocxDocumentModel;

    const acquired = bodyLayoutAcquisitionInput(document);
    const boundary = acquired.sequence.find((entry) => entry.kind === 'begin-section');

    expect(structuredClone(acquired)).toEqual(acquired);
    expect(acquired.sectionIndex.occurrences).toHaveLength(2);
    expect(boundary).toMatchObject({
      kind: 'begin-section',
      section: { sectionOccurrenceId: expect.any(String), startType: 'nextPage' },
    });
    expect(boundary && 'section' in boundary && 'context' in boundary.section).toBe(false);
    expect(acquired.noteLayoutSettings).toEqual({
      footnotePosition: 'pageBottom',
      endnotePosition: 'docEnd',
      // §17.11.18/.20 defaults; endnotes take Word's lowerRoman default
      // (WORD_NOTE_NUMBER_FORMAT_DEFAULTS, [MS-OI29500] §2.1.306).
      footnoteNumbering: { format: 'decimal', start: 1, restart: 'continuous' },
      endnoteNumbering: { format: 'lowerRoman', start: 1, restart: 'continuous' },
      footnoteSeparator: 'default',
      endnoteSeparator: 'default',
      endnoteContinuationSeparator: 'default',
      footnoteContinuationSeparator: 'default',
    });
  });

  it('projects parser-private authored note placement facts', () => {
    const document = {
      body: [paragraph('body')],
      section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
      __noteLayoutSettings: {
        footnotePosition: 'beneathText',
        endnotePosition: 'sectEnd',
        footnoteNumberFormat: 'upperLetter',
        footnoteNumberStart: 4,
        endnoteNumberFormat: 'lowerRoman',
        footnoteSeparator: 'none',
        footnoteContinuationSeparator: 'short',
      },
    } as unknown as DocxDocumentModel;

    expect(createBodyLayoutInput(document).noteLayoutSettings).toEqual({
      footnotePosition: 'beneathText',
      endnotePosition: 'sectEnd',
      footnoteNumbering: { format: 'upperLetter', start: 4, restart: 'continuous' },
      endnoteNumbering: { format: 'lowerRoman', start: 1, restart: 'continuous' },
      footnoteSeparator: 'none',
      endnoteSeparator: 'default',
      endnoteContinuationSeparator: 'default',
      footnoteContinuationSeparator: 'short',
    });
  });

  it('preserves an authored short rule in both note roles across serialization', () => {
    const document = {
      body: [paragraph('body')], section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null }, fontFamilyClasses: {},
      __noteLayoutSettings: { footnoteSeparator: 'short', footnoteContinuationSeparator: 'short' },
    } as unknown as DocxDocumentModel;
    const input = createBodyLayoutInput(structuredClone(document));
    expect(input.noteLayoutSettings).toMatchObject({
      footnoteSeparator: 'short', footnoteContinuationSeparator: 'short',
    });
    expect(structuredClone(input)).toEqual(input);
    const invalid = { ...document, __noteLayoutSettings: { footnoteContinuationSeparator: 'custom' } };
    expect(() => createBodyLayoutInput(invalid as DocxDocumentModel)).toThrow('Unsupported note separator');
  });

  it('retains reserved separator formatting once as an independent frozen snapshot', () => {
    const rule = (mark: string, probe: number) => ({
      rule: {
        mark,
        control: { run: { text: '', fontSize: 14, __typographyAcquisition: { probe } } },
      },
      paragraph: {
        paragraph: { runs: [], defaultFontSize: 14 },
        contentMark: { run: { text: '', fontSize: 9 } },
      },
    });
    const reservedSeparators = {
      footnote: {
        separator: rule('short', 0),
        continuationSeparator: rule('full', 3),
        continuationNotice: {},
      },
    };
    const document = {
      body: [paragraph('body')], section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null }, fontFamilyClasses: {},
      __noteLayoutSettings: {
        footnoteSeparator: 'short', footnoteContinuationSeparator: 'full', reservedSeparators,
      },
    } as unknown as DocxDocumentModel;
    const acquired = bodyLayoutAcquisitionInput(document);
    const settings = projectBodyLayoutInput(acquired).noteLayoutSettings;
    // Normalized once: projection reuses the frozen acquisition snapshot.
    expect(settings).toBe(acquired.noteLayoutSettings);
    expect(settings).toMatchObject({ footnoteSeparator: 'short', footnoteContinuationSeparator: 'full' });
    const retained = settings?.reservedSeparators?.footnote;
    const authoredControl = reservedSeparators.footnote.separator.rule.control.run;
    expect(retained?.separator.rule?.control.run).toEqual(authoredControl);
    // The content paragraph mark keeps its own character properties.
    expect(retained?.separator.paragraph?.contentMark?.run?.fontSize).toBe(9);
    expect(retained?.continuationNotice).toEqual({});
    authoredControl.fontSize = 99;
    expect(retained?.separator.rule?.control.run?.fontSize).toBe(14);
    expect(Object.isFrozen(retained?.separator.rule?.control.run)).toBe(true);

    const withSettings = (value: unknown) =>
      ({ ...document, __noteLayoutSettings: value }) as DocxDocumentModel;
    expect(createBodyLayoutInput(withSettings({ footnoteSeparator: 'short' })).noteLayoutSettings)
      .not.toHaveProperty('reservedSeparators');
    const footnote = { ...reservedSeparators.footnote, continuationNotice: { rule: { mark: 'dotted', control: {} } } };
    expect(() => createBodyLayoutInput(withSettings({ reservedSeparators: { footnote } })))
      .toThrow('Unsupported reserved note separator story');
  });

  it('projects section ownership and authored transitions without parser handles', () => {
    const body: BodyElement[] = [
      paragraph('first'),
      {
        type: 'sectionBreak',
        kind: 'continuous',
        geom: { ...finalSection(), pageWidth: 500 },
        columns: null,
        textDirection: null,
        pageNumType: null,
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        titlePage: false,
      },
      paragraph('second'),
      { type: 'columnBreak' },
    ];
    const document = {
      body,
      section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as DocxDocumentModel;

    const input = createBodyLayoutInput(document);
    const cloned = structuredClone(input);

    expect(input.initialSection.source.path).toEqual([1]);
    expect(input.initialSection.context.geometry.pageWidth).toBe(500);
    expect(input.sequence).toMatchObject([
      { kind: 'body-block', block: { kind: 'paragraph', source: { path: [0] } } },
      { kind: 'begin-section', source: { path: [1] }, section: { source: { path: [] } } },
      { kind: 'body-block', block: { kind: 'paragraph', source: { path: [2] } } },
      { kind: 'authored-break', break: 'column', source: { path: [3] } },
    ]);
    expect(cloned).toEqual(input);
    expect(JSON.stringify(input)).not.toContain('__sectionPlacement');
  });

  it('projects parser-private sectPr bidi into the retained section context', () => {
    const endingSection = {
      type: 'sectionBreak',
      kind: 'nextColumn',
      geom: finalSection(),
      columns: {
        count: 2,
        spacePt: 20,
        equalWidth: true,
        sep: false,
        cols: [],
      },
      textDirection: null,
      pageNumType: null,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      titlePage: false,
      __sectionPlacement: {
        sectionId: 'section:rtl',
        sectionBidi: true,
      },
    } as unknown as BodyElement;
    const document = {
      body: [paragraph('first'), endingSection, paragraph('second')],
      section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as DocxDocumentModel;

    const input = createBodyLayoutInput(document);
    const nextSection = input.sequence.find((entry) => entry.kind === 'begin-section');

    expect(input.initialSection.context.sectionBidi).toBe(true);
    expect(nextSection).toMatchObject({
      kind: 'begin-section',
      section: { context: { sectionBidi: false } },
    });
    expect(structuredClone(input)).toEqual(input);
    expect(JSON.stringify(input)).not.toContain('__sectionPlacement');
  });

  it('projects a parser-private pageBreakBefore transition without retaining the wire field', () => {
    const document = {
      body: [
        { type: 'pageBreak', origin: 'authored', __pageBreakBefore: true },
        { type: 'columnBreak' },
        paragraph('mark'),
      ],
      section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const input = createBodyLayoutInput(document);
    expect(input.sequence.slice(0, 2)).toMatchObject([
      { kind: 'authored-break', break: 'pageBreakBefore' },
      { kind: 'authored-break', break: 'column' },
    ]);
    expect(JSON.stringify(input)).not.toContain('__pageBreakBefore');
  });

  it('retains page-break parity and authored versus synthetic provenance during acquisition', () => {
    const document = {
      body: [
        { type: 'pageBreak', parity: 'odd', origin: 'authored' },
        { type: 'pageBreak', origin: 'coverPageSynthetic' },
        { type: 'pageBreak', origin: 'authored' },
        { type: 'columnBreak', parity: 'even' },
      ] as unknown as BodyElement[],
      section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as DocxDocumentModel;

    expect(bodyLayoutAcquisitionInput(document).sequence).toMatchObject([
      { kind: 'authored-break', break: 'page', parity: 'odd', origin: 'authored' },
      { kind: 'authored-break', break: 'page', origin: 'coverPageSynthetic' },
      { kind: 'authored-break', break: 'page', origin: 'authored' },
      { kind: 'authored-break', break: 'column' },
    ]);
    expect(bodyLayoutAcquisitionInput(document).sequence[1]).not.toHaveProperty('parity');
    expect(bodyLayoutAcquisitionInput(document).sequence[3]).not.toHaveProperty('parity');
  });

  it('consumes a vanished empty paragraph without admitting a body block', () => {
    const hidden = { ...paragraph(''), runs: [], markVanish: true } as BodyElement;
    const document = {
      body: [hidden, paragraph('visible')],
      section: finalSection(),
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as DocxDocumentModel;

    expect(createBodyLayoutInput(document).sequence.map((entry) => entry.kind)).toEqual([
      'consume-source', 'body-block',
    ]);
  });

  it('projects the suppress-before role from the resolved incoming continuous section', () => {
    const geometry = finalSection({ sectionStart: 'continuous' });
    const document = {
      body: [
        paragraph('A'),
        paragraph('', 20),
        { type: 'sectionBreak', kind: 'nextPage', geom: geometry } as BodyElement,
        paragraph('B'),
      ],
      section: geometry,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as DocxDocumentModel;

    const roles = createBodyLayoutInput(document).sequence.flatMap((entry) =>
      entry.kind === 'body-block' && entry.block.kind === 'paragraph'
        ? [entry.block.continuousSectionRole]
        : []);
    expect(roles).toEqual([
      undefined, 'suppress-before', undefined,
    ]);
  });

  it('projects mutually exclusive collapsed-mark and drop-previous-after roles', () => {
    const geometry = finalSection({ sectionStart: 'continuous' });
    const document = {
      body: [
        paragraph('A', 0, 40),
        paragraph(''),
        paragraph(''),
        { type: 'sectionBreak', kind: 'continuous', geom: geometry } as BodyElement,
        paragraph('B'),
      ],
      section: geometry,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as DocxDocumentModel;

    const roles = createBodyLayoutInput(document).sequence.flatMap((entry) =>
      entry.kind === 'body-block' && entry.block.kind === 'paragraph'
        ? [entry.block.continuousSectionRole]
        : []);
    expect(roles).toEqual([
      undefined, 'drop-previous-after', 'collapse-mark', undefined,
    ]);
  });
});
