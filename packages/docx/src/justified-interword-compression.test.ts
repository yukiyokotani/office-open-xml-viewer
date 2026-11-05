import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, DocxTextRun, SectionProps } from './types';

// WORD_JUSTIFIED_INTERWORD_COMPRESSION. Expected partitions and gap widths are
// read from Word 16.113 (macOS 27) PDF exports: the line-fit controls and the
// public demo document. The stub measures with the Georgia / Georgia Italic
// hmtx advances (2048 upm) of the faces embedded in those PDFs, so no
// renderer constant is restated here.
const GEORGIA_ASCII = [494,678,843,1317,1249,1674,1455,441,768,768,967,1317,552,766,552,960,1257,880,1144,1130,1157,1082,1159,1029,1221,1159,640,640,1317,1317,1317,980,1902,1374,1339,1315,1534,1338,1227,1485,1669,798,1060,1422,1236,1899,1571,1524,1249,1524,1437,1149,1267,1549,1365,1998,1455,1260,1232,768,960,768,1317,1317,1024,1032,1147,930,1176,990,666,1043,1192,600,598,1097,586,1804,1210,1104,1170,1146,839,885,707,1178,1017,1510,1034,1008,909,881,768,881,1317];
const GEORGIA: Record<string, number> = {
  ...Object.fromEntries(GEORGIA_ASCII.map((value, index) => [String.fromCharCode(32 + index), value])),
  'ê': 990,
  '\u00a0': 494,
  '\t': 494,
};
const GEORGIA_ITALIC: Record<string, number> = {" ":494,"'":441,",":552,".":552,"F":1227,"Q":1496,"T":1267,"a":1173,"b":1134,"c":929,"d":1178,"e":966,"f":673,"g":1173,"h":1152,"i":609,"k":1081,"l":584,"m":1801,"n":1208,"o":1100,"p":1184,"r":945,"s":883,"t":711,"u":1178,"v":1102,"w":1684,"y":1146,"z":909};

function georgiaCanvas(spacePairUnits = 0, painted?: { text: string; x: number }[]): HTMLCanvasElement {
  let font = '11px Georgia';
  const advance = (text: string): number => {
    const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 11);
    const table = /italic/.test(font) ? GEORGIA_ITALIC : GEORGIA;
    let units = 0;
    for (const character of text) {
      // Unlisted characters (controls, marks, dashes) get a fixed advance;
      // only the partition comparison with the previous layout matters there.
      units += table[character] ?? 600;
    }
    // Synthetic pair kerning between adjacent letters, so every kerned word
    // measures differently from its unkerned advance.
    if (ctx.fontKerning === 'normal') {
      units -= 10 * (text.match(/[A-Za-z](?=[A-Za-z])/g)?.length ?? 0);
      units -= spacePairUnits * (text.match(/ T/g)?.length ?? 0);
    }
    return (units * px) / 2048;
  };
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => {
      const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 11);
      return {
        width: advance(text),
        fontBoundingBoxAscent: px * 0.9,
        fontBoundingBoxDescent: px * 0.25,
        actualBoundingBoxAscent: px * 0.7,
        actualBoundingBoxDescent: px * 0.2,
        actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: advance(text),
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, setTransform() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number) { painted?.push({ text, x }); }, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { width: 0, height: 0, style: {}, getContext: () => ctx } as unknown as HTMLCanvasElement;
}

function paragraph(
  text: string,
  alignment: DocParagraph['alignment'],
  fontSize = 11,
  italic = false,
  extra: Partial<DocxTextRun> = {},
): BodyElement {
  const run: DocxTextRun = {
    text, bold: false, italic, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'Georgia', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
  return {
    type: 'paragraph',
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...run }],
    defaultFontSize: fontSize, defaultFontFamily: 'Georgia',
    widowControl: false,
  } as unknown as BodyElement;
}

async function renderLines(
  element: BodyElement,
  widthPt: number,
  settings: Record<string, unknown> = { compatibilityMode: 15, characterSpacingControl: 'compressPunctuation' },
  sectionExtra: Partial<SectionProps> = {},
  canvas = georgiaCanvas(),
): Promise<DocxTextRunInfo[][]> {
  const section = {
    pageWidth: widthPt, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...sectionExtra,
  } as SectionProps;
  const doc = {
    section, settings, body: [element],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc, canvas, 0, {
    dpr: 1,
    width: widthPt,
    onTextRun: (run) => { if (run.text) runs.push(run); },
  });
  const lines = new Map<number, DocxTextRunInfo[]>();
  for (const run of runs) {
    const key = Math.round(run.y);
    lines.set(key, [...(lines.get(key) ?? []), run]);
  }
  return [...lines.entries()].sort(([a], [b]) => a - b)
    .map(([, line]) => line.sort((a, b) => a.x - b.x));
}

const words = (line: readonly DocxTextRunInfo[]): string[] =>
  line.map((run) => run.text).join('').trim().split(/[ \t\u00a0]+/);

const CONTROL = 'Quiet rivers carry morning light across the valley beyond the distant hills today';
const MANY_GAPS = `${Array.from({ length: 24 }, () => 'a').join(' ')} beyond the distant hills today`;

describe('WORD_JUSTIFIED_INTERWORD_COMPRESSION', () => {
  it.each([
    { alignment: 'start' as const, stop: 216, count: 6, extra: 'seventh', admitted: 7 },
    { alignment: 'end' as const, stop: 360, count: 16, extra: '', admitted: 13 },
    { alignment: 'center' as const, stop: 216, count: 16, extra: '', admitted: 13 },
  ])('admits the same near-overflow tab cell at a $alignment stop', async ({ alignment, stop, count, extra, admitted }) => {
    // Independent Courier hmtx: every source glyph advances 1229/2048 em.
    // A 481.9pt band leaves 265.9pt after the start stop, or 460.3pt after
    // the three-glyph label when an end/center cell fits back. Word admits
    // the next word under both, distribute and lowKashida in mode 15; mode
    // 14 keeps it on the next line in all three original tab geometries.
    const text = `${Array.from({ length: count }, () => 'word').join(' ')}${extra ? ` ${extra}` : ''}`;
    const firstCellWords = async (jc: DocParagraph['alignment'], mode: number) => {
      const canvas = georgiaCanvas();
      const context = canvas.getContext('2d') as CanvasRenderingContext2D;
      const measure = context.measureText.bind(context);
      context.measureText = value => {
        const width = [...value].length * Number(/([\d.]+)px/.exec(context.font)?.[1] ?? 12) * 1229 / 2048;
        return { ...measure(value), width, actualBoundingBoxRight: width };
      };
      const element = paragraph(`Row\t${text}`, jc, 12, false, { fontFamily: 'Courier New' }) as DocParagraph;
      element.tabStops = [{ pos: stop, alignment, leader: 'none' }];
      const [first] = await renderLines(element as BodyElement, 481.9, { compatibilityMode: mode }, {}, canvas);
      const last = first!.at(-1)!;
      const hanging = (last.text.match(/ +$/)?.[0].length ?? 0) * last.fontSize * 1229 / 2048;
      expect(last.x + last.w - hanging).toBeLessThanOrEqual(481.9 + 1e-6);
      return first!.map(run => run.text).join('').replace(/\t/g, '').replace(/^Row/, '').trim().split(/ +/).length;
    };
    for (const jc of ['both', 'distribute', 'lowKashida'] as const) {
      expect.soft(await firstCellWords(jc, 15), `${jc}, mode 15`).toBe(admitted);
      expect.soft(await firstCellWords(jc, 14), `${jc}, mode 14`).toBe(admitted - 1);
    }
  });

  it.each([
    { label: 'authored RTL mark', word: 'word', extra: 'seventh', rtl: true },
    { label: 'unmarked RTL text', word: 'אבגד', extra: 'אבגדהוז', rtl: false },
  ])('preserves natural-width partitions outside the LTR capability: $label', async ({ word, extra, rtl }) => {
    // Six four-glyph words and their following spaces occupy 216.04pt;
    // the seven-glyph next word exceeds the remaining 265.9pt cell at its
    // natural Courier advance. The new alignment capability must preserve
    // that partition when source gaps need RTL physical-edge ownership.
    for (const alignment of ['distribute', 'lowKashida'] as const) {
      const canvas = georgiaCanvas();
      const context = canvas.getContext('2d') as CanvasRenderingContext2D;
      const measure = context.measureText.bind(context);
      context.measureText = value => {
        const width = [...value].length * Number(/([\d.]+)px/.exec(context.font)?.[1] ?? 12) * 1229 / 2048;
        return { ...measure(value), width, actualBoundingBoxRight: width };
      };
      const element = paragraph(`Row\t${Array(6).fill(word).join(' ')} ${extra}`, alignment, 12, false,
        { fontFamily: 'Courier New', rtl }) as DocParagraph;
      element.tabStops = [{ pos: 216, alignment: 'start', leader: 'none' }];
      const [first, second] = await renderLines(element as BodyElement, 481.9, { compatibilityMode: 15 }, {}, canvas);
      const visibleWords = (line: readonly DocxTextRunInfo[]) => line.map(run => run.text)
        .join(' ').replace('Row', '').trim().split(/\s+/).filter(Boolean);
      expect.soft(visibleWords(first!), alignment).toHaveLength(6);
      expect.soft(visibleWords(second!), alignment).toEqual([extra]);
    }
  });

  it('preserves a native space-to-letter pair only when its tokens share a line', async () => {
    // An independent whole-string measure owns this boundary: the stub adds
    // a 200/2048 em space-T pair, in addition to its ordinary letter pairs.
    // Token sums exceed the quarter-space bound; continuous text does not.
    const prefix = 'Quiet To rivers';
    const nativeAdvance = (GEORGIA_SUM(prefix) - 10 * 10 - 200) * 11 / 2048;
    const width = nativeAdvance - 1;
    const element = paragraph(`${prefix} beyond the valley`, 'both', 11, false, { kerning: 8 });
    const painted: { text: string; x: number }[] = [];
    const [first] = await renderLines(element, width, undefined, {}, georgiaCanvas(200, painted));
    expect(words(first)).toEqual(['Quiet', 'To', 'rivers']);
    const toOnFirst = painted.find(run => run.text.trim() === 'To')!;
    const quietAdvance = (GEORGIA_SUM('Quiet') - 40) * 11 / 2048;
    const spaceAdvance = 494 * 11 / 2048;
    expect(toOnFirst.x).toBeCloseTo(quietAdvance + spaceAdvance - 0.5 - 200 * 11 / 2048, 6);
    const last = first.at(-1)!;
    const lastBareAdvance = (GEORGIA_SUM(last.text.trimEnd()) - 10 * 5) * 11 / 2048;
    expect(last.x + lastBareAdvance).toBeCloseTo(width, 6);

    const runs = (element as DocParagraph).runs;
    const original = runs[0];
    const split = { ...element, runs: ['Quiet ', 'To ', 'rivers beyond the valley']
      .map(text => ({ ...original, text })) } as BodyElement;
    const [splitFirst] = await renderLines(split, width, undefined, {}, georgiaCanvas(200));
    expect(splitFirst.map(run => ({ text: run.text, x: run.x })))
      .toEqual(first.map(run => ({ text: run.text, x: run.x })));

    // Rejected candidates start with their independent origin: the previous
    // line's separator must not move the first glyph of the next line.
    const narrow = (GEORGIA_SUM('Quiet') - 40) * 11 / 2048 + 0.1;
    const narrowPaint: { text: string; x: number }[] = [];
    await renderLines(element, narrow, undefined, {}, georgiaCanvas(200, narrowPaint));
    const to = narrowPaint.find(run => run.text.trim() === 'To')!;
    expect(to.x).toBeCloseTo(0, 6);
  });


  it('rejects a positive native boundary advance before isolated fit can override it', async () => {
    // Without the positive pair, the complete visible candidate fits naturally.
    // With it, overflow exceeds one quarter of the genuine interword space.
    const independentWidth = (GEORGIA_SUM('Quiet To') - 50) * 11 / 2048;
    const painted: { text: string; x: number }[] = [];
    const [first] = await renderLines(paragraph('Quiet To beyond', 'both', 11, false, { kerning: 8 }),
      independentWidth + 0.3, undefined, {}, georgiaCanvas(-200, painted));
    expect(words(first)).toEqual(['Quiet']);
    expect(painted.find(run => run.text.trim() === 'To')?.x).toBeCloseTo(0, 6);
  });

  it('pulls a word onto a justified line by shrinking its gaps, and paints those gaps', async () => {
    // Word keeps `valley` on line one at 233.30pt (natural 236.31pt) and
    // draws seven equal gaps of about 2.21pt instead of the 2.65pt space.
    const [first] = await renderLines(paragraph(CONTROL, 'both'), 233.3);
    expect(words(first).at(-1)).toBe('valley');
    const bare = (run: DocxTextRunInfo) => (GEORGIA_SUM(run.text.trimEnd()) * 11) / 2048;
    const gaps = first.slice(1).map((run, index) => run.x - (first[index].x + bare(first[index])));
    for (const gap of gaps) expect(gap).toBeCloseTo(2.21, 1);
    const last = first.at(-1)!;
    expect(last.x + bare(last)).toBeCloseTo(233.3, 3);
  });

  it('does not count the candidate separator once the visible word fits', async () => {
    // Word: `valley` stays on line one at 237.30pt (0.99pt visible room,
    // less than its 2.65pt separator), and its last glyph meets the band edge.
    const [first] = await renderLines(paragraph(CONTROL, 'both'), 237.3);
    expect(words(first).at(-1)).toBe('valley');
    const last = first.at(-1)!;
    expect(last.x + (GEORGIA_SUM(last.text.trimEnd()) * 11) / 2048).toBeCloseTo(237.3, 3);
  });

  it('never compresses a left-aligned line', async () => {
    const [first] = await renderLines(paragraph(CONTROL, 'left'), 236.05);
    expect(words(first).at(-1)).toBe('the');
  });

  it('rejects a short word whose overflow exceeds half the expansion alternative', async () => {
    // 23 gaps can absorb 3pt, yet Word moves the 24th one-letter word at
    // 191.05pt and keeps it at 193.05pt.
    const [rejected] = await renderLines(paragraph(MANY_GAPS, 'both'), 191.05);
    expect(words(rejected)).toHaveLength(23);
    const [accepted] = await renderLines(paragraph(MANY_GAPS, 'both'), 193.05);
    expect(words(accepted)).toHaveLength(24);
  });

  it('applies the same rule to the final line of a both paragraph', async () => {
    // Word lays this 12pt italic paragraph out in two lines at 451.3pt; its
    // final line overflows naturally by 5.19pt and is drawn with shrunk gaps.
    const text = "Four of the world's great forest biomes, sketched in their broadest strokes. "
      + 'Figures are illustrative, drawn from openly published estimates and rounded for readability.';
    const lines = await renderLines(paragraph(text, 'both', 12, true), 451.3);
    expect(lines).toHaveLength(2);
    expect(words(lines[0]).at(-1)).toBe('Figures');
    expect(words(lines[1]).at(-1)).toBe('readability.');
  });

  it('separates Word-accepted and Word-rejected candidates at the expansion boundary', async () => {
    // Demo document, Georgia 11pt, 451.3pt band. Word keeps `fire.` (gap
    // compression 0.46 of the per-space expansion alternative) and moves `of`
    // (0.52), although `of` needs less compression per gap than `fire.`.
    const accepted = await renderLines(paragraph(
      'Two hundred feet tall and counting, with bark thick enough to shrug off all but the fiercest fire. A single mature tree',
      'both',
    ), 451.3);
    expect(words(accepted[0]).at(-1)).toBe('fire.');
    const rejected = await renderLines(paragraph(
      'phosphorus and water for sugars made above. The same network carries chemical warnings of insect attack',
      'both',
    ), 451.3);
    expect(words(rejected[0]).at(-1)).toBe('warnings');
  });

  it('admits or rejects an elided word with its accented continuation as one unit', async () => {
    // The elision d' alone would fit by compression at 214pt, but the
    // complete word d'être needs more than the rule allows: it moves whole.
    const text = "Quiet rivers carry morning light across the d'être beyond the distant hills";
    const rejected = await renderLines(paragraph(text, 'both'), 214);
    expect(words(rejected[0]).at(-1)).toBe('the');
    expect(words(rejected[1])[0]).toBe("d'être");
    // At 233pt the complete word overflows by 2.9pt over seven gaps and stays.
    const accepted = await renderLines(paragraph(text, 'both'), 233);
    expect(words(accepted[0]).at(-1)).toBe("d'être");
  });

  it('admits the measured one-gap boundary and rejects the wider overflow', async () => {
    const text = 'Quiet valley beyond the distant hills today';
    expect(words((await renderLines(paragraph(text, 'both'), 57.3))[0]).at(-1)).toBe('valley');
    expect(words((await renderLines(paragraph(text, 'both'), 56.55))[0])).toEqual(['Quiet']);
  });

  it('keeps natural visible fit but excludes compression under a character grid', async () => {
    const grid = { docGridType: 'linesAndChars', docGridLinePitch: 18, docGridCharSpace: 0 };
    // A prospective line's edge separator cannot reject visible text that
    // naturally fits, with or without a grid. The grid excludes the observed
    // compression allowance only when the visible prefix really overflows.
    const visible = GEORGIA_SUM('Quiet rivers carry morning light across') * 11 / 2048;
    const [natural] = await renderLines(paragraph(CONTROL, 'both'), 190, undefined, grid);
    expect(words(natural).at(-1)).toBe('across');
    const [plain] = await renderLines(paragraph(CONTROL, 'both'), visible - 0.5);
    expect(words(plain).at(-1)).toBe('across');
    const [gridded] = await renderLines(paragraph(CONTROL, 'both'), visible - 0.5, undefined, grid);
    expect(words(gridded).at(-1)).toBe('light');
  });

  const textRun = (text: string, extra: Partial<DocxTextRun> = {}) => ({
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 11, color: null, fontFamily: 'Georgia', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  }) as DocParagraph['runs'][number];
  const withRuns = (runs: DocParagraph['runs'], extra: Partial<DocParagraph> = {}): BodyElement => ({
    ...(paragraph('', 'both') as DocParagraph), runs, ...extra,
  }) as unknown as BodyElement;

  it('fits a word split at an authored style seam as one candidate', async () => {
    const [first] = await renderLines(withRuns([
      textRun('Quiet rivers carry morning light across the val'),
      textRun('ley beyond the distant hills today', { color: 'ff0000' }),
    ]), 233.3);
    expect(words(first).at(-1)).toBe('valley');
  });

  it('keeps later compression available after a refused source-split word', async () => {
    const element = withRuns([
      textRun('Quiet rivers carry morning light across the valley be'),
      textRun('yond the distant hills today Quiets rivers carry morning light across the valley beyond the distant hills today', { color: 'ff0000' }),
    ]);
    (element as DocParagraph).alignment = 'distribute';
    const result = await renderLines(element, 233.3);
    // The style-split next word is refused on the first compressed line.
    // On the next line the independently measured visible phrase advances
    // 235.667pt, so retaining its final word also needs real contraction.
    expect(words(result[0]).at(-1)).toBe('valley');
    expect(words(result[1]).join(' ')).toBe('beyond the distant hills today Quiets rivers carry');
    expect(result[1].at(-1)!.x + GEORGIA_SUM('carry') * 11 / 2048).toBeCloseTo(233.3, 3);
  });

  it('counts double spaces and separators in their own source runs', async () => {
    const text = 'Quiet  rivers carry light across valley beyond the distant hills today';
    const double = await renderLines(paragraph(text, 'both'), 173.2);
    expect(words(double[0]).at(-1)).toBe('valley');
    const separate = await renderLines(withRuns([
      textRun('Quiet'), textRun('  '), textRun('rivers carry light across valley beyond the distant hills today'),
    ]), 173.2);
    expect(words(separate[0])).toEqual(words(double[0]));
    expect(separate[0].at(-1)!.x + GEORGIA_SUM(separate[0].at(-1)!.text.trimEnd()) * 11 / 2048)
      .toBeCloseTo(173.2, 3);
  });

  it('paints mixed-size spaces proportionally under compression and expansion', async () => {
    const element = withRuns([
      textRun('Quiet '), textRun('rivers ', { fontSize: 14 }),
      textRun('carry light across valley beyond the distant hills today'),
    ]);
    for (const width of [178.85, 184.85]) {
      const [first] = await renderLines(element, width);
      expect(words(first).at(-1)).toBe('valley');
      const fractions = first.slice(0, -1).map((run, index) => {
        const naturalSpace = 494 * run.fontSize / 2048;
        const wordWidth = GEORGIA_SUM(run.text.trimEnd()) * run.fontSize / 2048;
        const drawnGap = first[index + 1].x - run.x - wordWidth;
        return (drawnGap - naturalSpace) / naturalSpace;
      });
      for (const factor of fractions) expect(factor).toBeCloseTo(fractions[0], 5);
    }
  });

  it('keeps spaces next to ruby fixed while the other spaces absorb slack', async () => {
    const element = withRuns([
      textRun('Quiet '), textRun('rivers', { ruby: { text: 'rise', fontSizePt: 5.5 } } as Partial<DocxTextRun>),
      textRun(' carry light across valley beyond the distant hills today'),
    ]);
    expect(words((await renderLines(element, 170.55))[0]).at(-1)).toBe('across');
    const [accepted] = await renderLines(element, 172.55);
    expect(words(accepted).at(-1)).toBe('valley');
    expect(accepted[1].x - accepted[0].x).toBeCloseTo(GEORGIA_SUM('Quiet ') * 11 / 2048, 4);
  });

  it('preserves the fixed width and atomic wrapping of a linked fitText region', async () => {
    const element = withRuns([
      textRun('Quiet rivers '),
      textRun('carry ', { fitTextVal: 1600, fitTextId: 7 }),
      textRun('morning light', { fitTextVal: 1600, fitTextId: 7, color: 'ff0000' }),
      textRun(' across the valley beyond the distant hills today'),
    ]);
    const lines = await renderLines(element, 100);
    const regionLine = lines.find(line => line.some(run => run.text.includes('carry')))!;
    expect(words(lines[0])).toEqual(['Quiet', 'rivers']);
    expect(words(regionLine).slice(0, 3)).toEqual(['carry', 'morning', 'light']);
    const first = regionLine.find(run => run.text.startsWith('carry'))!;
    const last = regionLine.find(run => run.text.includes('light'))!;
    expect(last.x + last.w - first.x).toBeCloseTo(80, 5);
  });

  it('keeps a space with its combining mark atomic during fit and paint', async () => {
    const element = paragraph(CONTROL.replace('Quiet ', 'Quiet \u0301'), 'both');
    for (const width of [233.3, 237.3]) {
      const lines = await renderLines(element, width);
      expect(lines.flat().map(run => run.text).join('')).toBe(element.type === 'paragraph'
        ? (element as DocParagraph).runs.map(run => 'text' in run ? run.text : '').join('') : '');
      const first = lines[0][0];
      const mark = lines[0].find(run => run.text.startsWith('\u0301'))!;
      expect(mark.x - first.x).toBeCloseTo(GEORGIA_SUM('Quiet ') * 11 / 2048, 5);
    }
  });

  it('acquires a scalar space after a grapheme prepend without splitting the cluster', async () => {
    const [first] = await renderLines(paragraph(CONTROL.replace('Quiet ', 'Quiet \u0600 '), 'both'), 250);
    const sign = first.find(run => run.text.includes('\u0600'))!;
    const following = first[first.indexOf(sign) + 1];
    expect(sign).toBeDefined();
    expect(following.x).toBeGreaterThan(sign.x);
    expect(sign.text.endsWith(' ')).toBe(true);
  });

  it.each(['§', '°', '×', '€', '$', '!', '漢'])('allows any text glyph (%s) beside a gap', async (glyph) => {
    // A font-slot/width-balance classification is not a text-cell boundary.
    // Bracket a one-point overflow with the actual fixture advances.
    const prefix = `Quiet rivers carry morning light across ${glyph} valley`;
    const units = [...prefix].reduce((sum, character) => sum + (GEORGIA[character] ?? 600), 0);
    const [first] = await renderLines(paragraph(`${prefix} beyond the distant hills today`, 'both'),
      units * 11 / 2048 - 1, { compatibilityMode: 15, balanceSingleByteDoubleByteWidth: true });
    expect(words(first).at(-1)).toBe('valley');
  });

  it('retains raised text and authored run advances without a provenance gate', async () => {
    const [raised] = await renderLines(paragraph(CONTROL, 'both', 11, false, { position: 3 }), 233.3);
    expect(words(raised).at(-1)).toBe('valley');
    const [kerned] = await renderLines(paragraph(CONTROL, 'both', 11, false, { kerning: 1 }), 233.3);
    expect(words(kerned).at(-1)).toBe('valley');
  });

  // Closed paragraph/layout classes are outside the measured Word rule.
  // Compare their actual partitions with the previous mode-14 policy.
  it.each([
    ['linesAndChars', paragraph(CONTROL, 'both'), { docGridType: 'linesAndChars', docGridLinePitch: 18, docGridCharSpace: 0 }, {}],
    ['snapToChars', paragraph(CONTROL, 'both'), { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0 }, {}],
    ['medium kashida', paragraph(CONTROL, 'mediumKashida'), {}, {}],
    ['RTL', withRuns([textRun(CONTROL)], { bidi: true }), {}, {}],
    ['left', paragraph(CONTROL, 'left'), {}, {}],
  ] as const)('preserves the previous partition for %s', async (_name, element, section, settings) => {
    for (const width of [233.3, 237.3]) {
      const partition = (lines: DocxTextRunInfo[][]) => lines.map(line => words(line));
      const previous = await renderLines(element, width, { compatibilityMode: 14, ...settings }, section as Partial<SectionProps>);
      const current = await renderLines(element, width, { compatibilityMode: 15, ...settings }, section as Partial<SectionProps>);
      expect(partition(current)).toEqual(partition(previous));
    }
  });

  // Windows Word exports identical compressed justified lines with and
  // without lineWrapLikeWord6 in mode 15.
  it('compresses justified mode-15 lines regardless of lineWrapLikeWord6', async () => {
    const [first] = await renderLines(paragraph(CONTROL, 'both'), 233.3,
      { compatibilityMode: 15, lineWrapLikeWord6: true });
    expect(words(first).at(-1)).toBe('valley');
    const last = first.at(-1)!;
    const bareAdvance = GEORGIA_SUM(last.text.trimEnd()) * 11 / 2048;
    expect(last.x + bareAdvance).toBeCloseTo(233.3, 3);
  });

  it('keeps natural-width fitting outside compatibility mode 15', async () => {
    const [first] = await renderLines(paragraph(CONTROL, 'both'), 233.3, { compatibilityMode: 14 });
    expect(words(first).at(-1)).toBe('the');
  });
});

function GEORGIA_SUM(text: string): number {
  let units = 0;
  for (const character of text) units += GEORGIA[character] ?? 0;
  return units;
}
