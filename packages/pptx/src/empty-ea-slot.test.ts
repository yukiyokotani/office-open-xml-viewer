import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { paragraphInputRuns, renderTextBody } from './renderer.js';
import { powerPointFaceMetrics, powerPointResourceFaceMetrics, powerPointSymbolCoverage } from './powerpoint-line-metrics.js';
import type { Paragraph, TextBody } from './types.js';

// Issue #1689: an empty East Asian slot. Expected values are independent
// readings of PowerPoint 16.113 tagged PDF exports (emptyea, emptyea2,
// emptyea3, emptyea4 and fallback controls). Faces come from the embedded
// font name tables of the exported glyphs; baselines are export positions,
// quantized by the exporter to 0.72 pt (1/100 in) steps.

const SCALE = 1 / 12700; // 1 pt = 1 canvas unit
const U = 0.72;
const RC = { themeMajorFont: null, themeMinorFont: null, dpr: 1 };

interface Faces { latin: string; cs?: string; ea?: string; lang?: string; italic?: boolean }

function textRun(text: string, f: Faces, ea: string | undefined, size = 22): TextRunData {
  return {
    type: 'text', text, bold: false, italic: f.italic ?? true, underline: false, strikethrough: false,
    fontSize: size, color: '000000', fontFamily: f.latin, fontFamilyEa: ea, fontFamilyCs: f.cs,
    lang: f.lang ?? 'en-US',
  } as TextRunData;
}

function para(runs: Paragraph['runs']): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null,
    spaceLine: null, lvl: 0, bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true, runs,
  } as Paragraph;
}

/** The faces each East Asian-slot character of `text` is drawn with. */
function faces(text: string, f: Faces): Record<string, string | undefined> {
  const out: Record<string, string | undefined> = {};
  const { input } = paragraphInputRuns(para([textRun(text, f, f.ea)]), 22, '#000', SCALE, false, true, 1, undefined, RC);
  for (const item of input) {
    if (item.type !== 'text') continue;
    for (const ch of item.text) if (!/[\s3A-Za-z]/u.test(ch)) out[ch] = item.style.faceFamily;
  }
  return out;
}

function fontOf(text: string, f: Faces, ch: string): string {
  const { input } = paragraphInputRuns(para([textRun(text, f, f.ea)]), 22, '#000', SCALE, false, true, 1, undefined, RC);
  const item = input.find((i) => i.type === 'text' && i.text.includes(ch));
  return item?.type === 'text' ? item.style.font : '';
}

/** The emptyea probe body: five lines in three paragraphs, target on line 3. */
function probeBaselines(f: Faces, target: string, control: boolean, targetSize = 22): number[] {
  const ea = control ? (f.ea ?? 'MS Gothic') : undefined;
  const br = { type: 'break' } as Paragraph['runs'][number];
  const body = {
    verticalAnchor: 't', defaultFontSize: 22, defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
    paragraphs: [
      para([textRun('Alpha before.', f, undefined), br, textRun('Latin before.', f, undefined)]),
      para([textRun('Reference', f, undefined), textRun(` ${target}3`, f, ea, targetSize),
        textRun(' ends.', f, undefined), br, textRun('Latin after.', f, undefined)]),
      para([textRun('Source ends.', f, undefined)]),
    ],
  } as unknown as TextBody;
  const ys: number[] = [];
  let font = '';
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    fillStyle: '', strokeStyle: '', direction: 'ltr', textAlign: 'left', textBaseline: 'alphabetic',
    measureText: (text: string) => ({ width: [...text].length * 10, actualBoundingBoxAscent: 7, actualBoundingBoxDescent: 2 }),
    fillText: (_text: string, _x: number, y: number) => { if (!ys.some((v) => Math.abs(v - y) < 1e-6)) ys.push(y); },
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
  } as unknown as CanvasRenderingContext2D;
  renderTextBody(ctx, body, 0, 0, 408, 400, SCALE);
  return ys;
}

/** Test minus control baselines of the five probe lines. */
function deltas(f: Faces, target: string, targetSize = 22): number[] {
  const t = probeBaselines(f, target, false, targetSize);
  const c = probeBaselines(f, target, true, targetSize);
  expect(t).toHaveLength(5);
  return t.map((y, i) => y - c[i]);
}

describe('empty East Asian slot face resolution (#1689)', () => {
  it('draws neutral symbols in the named cs face, not the Latin face (emptyea E01/E04/E07/E16)', () => {
    expect(faces(' §3', { latin: 'Calibri', cs: 'Tahoma' })['§']).toBe('Tahoma');
    expect(faces(' §3', { latin: 'Arial', cs: 'Times New Roman' })['§']).toBe('Times New Roman');
    expect(faces(' §3', { latin: 'Cambria', cs: 'Microsoft Sans Serif' })['§']).toBe('Microsoft Sans Serif');
    expect(faces(' § ° ± × ÷3', { latin: 'Calibri', cs: 'Tahoma' }))
      .toEqual({ '§': 'Tahoma', '°': 'Tahoma', '±': 'Tahoma', '×': 'Tahoma', '÷': 'Tahoma' });
    // E33: a Latin face without § does not matter; E09: an explicit ea wins.
    expect(faces(' §3', { latin: 'SimSun-ExtB', cs: 'Tahoma' })['§']).toBe('Tahoma');
    expect(faces(' §3', { latin: 'Calibri', cs: 'Tahoma', ea: 'MS Gothic' })['§']).toBe('MS Gothic');
  });

  it('draws CJK in a cs face that covers it, whatever the language (emptyea2 G02-G04, G11)', () => {
    for (const cs of ['MS Mincho', 'DengXian', 'Arial Unicode MS']) {
      for (const lang of ['en-US', 'ja-JP']) {
        const got = faces(' §◆■、漢か3', { latin: 'Calibri', cs, lang });
        expect(Object.values(got).every((face) => face === cs), `${cs} ${lang}`).toBe(true);
      }
    }
  });

  it('falls back for CJK by the selected face, never the language (emptyea, emptyea2 G05/G12/G13, emptyea3)', () => {
    // A selected face with a sans PANOSE serif style: MS Gothic.
    for (const [latin, cs] of [['Calibri', 'Tahoma'], ['Calibri', 'Lucida Sans Unicode'], ['Arial', undefined],
      ['Calibri', undefined], ['Meiryo', 'Tahoma']] as const) {
      for (const lang of ['en-US', 'ja-JP', 'zh-CN', 'ko-KR']) {
        const got = faces(' 、漢か3', { latin, cs, lang });
        expect(got, `${latin}/${cs} ${lang}`).toEqual({ '、': 'MS Gothic', 漢: 'MS Gothic', か: 'MS Gothic' });
      }
    }
    // Serif styles, including SimSun-ExtB's style 1 and both ExtB faces: MS Mincho.
    for (const [latin, cs] of [['Arial', 'Times New Roman'], ['Arial', 'Courier New'], ['Arial', 'SimSun-ExtB'],
      ['Arial', 'MingLiU-ExtB'], ['SimSun-ExtB', undefined], ['MingLiU-ExtB', undefined]] as const) {
      for (const lang of ['en-US', 'ja-JP', 'zh-CN', 'ko-KR']) {
        const got = faces(' 、漢か3', { latin, cs, lang });
        expect(got, `${latin}/${cs} ${lang}`).toEqual({ '、': 'MS Mincho', 漢: 'MS Mincho', か: 'MS Mincho' });
      }
    }
  });

  it('keeps a covering Far-East face\'s own script chain in either slot (emptyea4 K03/K04)', () => {
    for (const f of [{ latin: 'MS Mincho' }, { latin: 'Arial', cs: 'MS Mincho' }]) {
      expect(faces(' §、漢か简한3', f)).toEqual({
        '§': 'MS Mincho', '、': 'MS Mincho', 漢: 'MS Mincho', か: 'MS Mincho',
        简: 'Microsoft JhengHei', 한: 'Malgun Gothic',
      });
    }
  });

  it('leaves the second fallback after a face without basic CJK to the platform (decision B, emptyea4 K01/K02)', () => {
    // PowerPoint drew 简 / 한 in JhengHei / Malgun Gothic after SimSun-ExtB but
    // in PMingLiU / Batang after MingLiU-ExtB. No drawing face is claimed.
    for (const latin of ['SimSun-ExtB', 'MingLiU-ExtB']) {
      const got = faces(' 漢简한3', { latin });
      expect(got).toEqual({ 漢: 'MS Mincho', 简: undefined, 한: undefined });
    }
  });

  it('attributes CJK by the selected cut\'s glyph coverage, not its family classification', () => {
    // Deng's Unicode cmap maps Han, but not U+D55C; the painting chain
    // passes through PMingLiU to Batang. An unknown earlier face could paint
    // either glyph and must prevent attribution to a later known resource.
    expect(faces('漢한', { latin: 'Arial', cs: 'DengXian', italic: false }))
      .toEqual({ 漢: 'DengXian', 한: 'Batang' });
    expect(faces('漢', { latin: 'Uncatalogued Face', italic: false })).toEqual({ 漢: undefined });
  });

  it('attributes covered and missing symbols per glyph to the drawing resource', () => {
    // The installed cmap distinguishes a covered ■ from a missing ◆ in
    // Tahoma, and a missing § from a covered § in the two ExtB resources.
    expect(faces(' §■3', { latin: 'Arial', cs: 'SimSun-ExtB' }))
      .toEqual({ '§': 'Calibri', '■': 'Cambria Math' });
    expect(faces(' §■3', { latin: 'MingLiU-ExtB' }))
      .toEqual({ '§': 'MingLiU-ExtB', '■': 'Cambria Math' });
    expect(faces(' ■3', { latin: 'Calibri', cs: 'Tahoma' })['■']).toBe('Tahoma');
    expect(faces(' ■3', { latin: 'Calibri' })['■']).toBe('Cambria Math');
    // ◆ is absent in the installed Cambria Math cmap. The known next face
    // in this painting stack is MS Mincho; the PDF's different resource is
    // outside this catalogue's compatibility claim.
    expect(faces(' ◆3', { latin: 'Calibri', cs: 'Courier New' })['◆']).toBe('MS Mincho');
    expect(powerPointSymbolCoverage('Cambria Math', false, true, 0x25C6)).toBe(false);
    // Unknown S might cover the glyph: do not claim a later catalogue face.
    expect(faces(' ■3', { latin: 'Uncatalogued Face' })['■']).toBeUndefined();
  });

  it('stacks the selected face, the symbol fallback and the CJK fallback in Office order', () => {
    // emptyea2 G01/G06, emptyea3 H01: § the cs face lacks → Calibri; ◆ ■ →
    // Cambria Math; CJK → the PANOSE tier face. Coverage picks within the stack.
    expect(fontOf(' §◆■、漢か3', { latin: 'Calibri', cs: 'SimSun-ExtB' }, '§'))
      .toMatch(/^italic \d+px "SimSun-ExtB", "Calibri", "Cambria Math", "MS Mincho"/u);
    expect(fontOf(' §◆■、漢か3', { latin: 'Calibri', cs: 'Tahoma' }, '§'))
      .toMatch(/^italic \d+px "Tahoma", "Calibri", "Cambria Math", "MS Gothic"/u);
  });
});

describe('selected-resource line metrics (#1689)', () => {
  it('sizes a synthetic italic by the upright resource (fallback / emptyea embedded OS/2)', () => {
    // The PDFs embed the regular resources under a 0.3333 shear.
    expect(powerPointFaceMetrics('MS Gothic', false, true)?.share).toBeCloseTo(220 / 256, 12);
    expect(powerPointFaceMetrics('Tahoma', false, true)?.share).toBeCloseTo(2049 / 2472, 12);
    expect(powerPointFaceMetrics('Microsoft Sans Serif', false, true)?.share).toBeCloseTo(1888 / 2318, 12);
    // A real italic resource keeps its own tables (Times New Roman Italic).
    expect(powerPointFaceMetrics('Times New Roman', false, true)?.share).toBeCloseTo(1825 / 2268, 12);
  });

  it('honors USE_TYPO_METRICS on older tables for a known symbol fallback', () => {
    // The installed OS/2 v3 table and the exported resource both set bit 7.
    const metric = powerPointFaceMetrics('Cambria Math', false, true);
    expect(metric?.share).toBeCloseTo(1946 / 2401, 12);
    expect(metric?.glyph).toEqual({ ascent: 1946 / 2048, descent: 455 / 2048 });
    const { input } = paragraphInputRuns(para([textRun('■', { latin: 'Calibri' }, undefined)]),
      22, '#000', SCALE, false, true, 1, undefined, RC);
    const segment = input.find((i) => i.type === 'text');
    expect(segment?.type === 'text' && segment.style.lineMetric).toBe(metric);
  });

  it('moves only the target line, by the measured amount (emptyea, emptyea2, emptyea3)', () => {
    const cases: Array<[string, Faces, string, number]> = [
      ['E01', { latin: 'Calibri', cs: 'Tahoma' }, '§', -0.72],
      ['E02', { latin: 'Arial', cs: 'Tahoma' }, '§', 0],
      ['E04', { latin: 'Arial', cs: 'Times New Roman', ea: 'MS Mincho' }, '§', 0],
      ['E05', { latin: 'Cambria', cs: 'Times New Roman', ea: 'MS Mincho' }, '§', 0],
      ['E06', { latin: 'Cambria', cs: 'Times New Roman', ea: 'Meiryo' }, '§', 2.16],
      ['E07', { latin: 'Cambria', cs: 'Microsoft Sans Serif', ea: 'Meiryo' }, '§', 2.16],
      ['E27', { latin: 'Calibri', cs: 'Tahoma', italic: false }, '§', -0.72],
      ['G01', { latin: 'Calibri', cs: 'Tahoma' }, '§◆■、漢か', 0],
      ['G02', { latin: 'Calibri', cs: 'MS Mincho' }, '§◆■、漢か', 0],
      ['G03', { latin: 'Calibri', cs: 'DengXian' }, '§◆■、漢か', -0.72],
      ['G04', { latin: 'Calibri', cs: 'Arial Unicode MS' }, '§◆■、漢か', -0.72],
      ['G05', { latin: 'Calibri', cs: 'Lucida Sans Unicode' }, '§◆■、漢か', -1.44],
      ['G06', { latin: 'Calibri', cs: 'SimSun-ExtB' }, '§◆■、漢か', 0],
      ['G07', { latin: 'Calibri' }, '§◆■、漢か', 0],
      ['H05', { latin: 'Arial', cs: 'Courier New' }, '§◆■、漢か', -1.44],
      ['H01', { latin: 'Arial', cs: 'SimSun-ExtB' }, '§◆■、漢か', 0],
      // K01's first row has no secondary CJK miss: its § drawing resource
      // changes the line metric even though the Latin face lacks that symbol.
      ['K01 first row', { latin: 'SimSun-ExtB' }, '§、漢か', -2.16],
    ];
    for (const [id, f, target, pdf] of cases) {
      const d = deltas(f, target);
      for (const i of [0, 1, 3, 4]) expect(d[i], `${id} line ${i + 1}`).toBeCloseTo(0, 9);
      // Continuous layout within one export step of the exported delta.
      expect(Math.abs(d[2] - pdf), `${id} line 3: ${d[2]}`).toBeLessThan(U);
    }
  });
});

describe('embedded selected-resource metrics (#1689)', () => {
  it('derives the share from the resource OS/2 tables', () => {
    const win = powerPointResourceFaceMetrics({
      unitsPerEm: 256, hheaAscent: 220, hheaDescent: -36, hheaLineGap: 0,
      winAscent: 220, winDescent: 36, hasEastAsianCmap: true,
    });
    expect(win?.share).toBeCloseTo(220 / 256, 12);
    expect(win?.glyph).toEqual({ ascent: 220 / 256, descent: 36 / 256 });
    // USE_TYPO_METRICS: (typo ascender + line gap) over the typo box (Cambria Math).
    const typo = powerPointResourceFaceMetrics({
      unitsPerEm: 2048, hheaAscent: 0, hheaDescent: 0, hheaLineGap: 0, useTypoMetrics: true,
      typoAscent: 1593, typoDescent: -455, typoLineGap: 353, winAscent: 6383, winDescent: 5045,
      hasEastAsianCmap: false,
    });
    expect(typo?.share).toBeCloseTo((1593 + 353) / (1593 + 353 + 455), 12);
    expect(powerPointResourceFaceMetrics({
      unitsPerEm: 0, hheaAscent: 0, hheaDescent: 0, hheaLineGap: 0, hasEastAsianCmap: false,
    })).toBeUndefined();
  });

  it('does not substitute installed cmap facts for a same-name embedded resource', () => {
    const rc = { ...RC, embeddedFontAliases: new Map([['calibri', '__deck_calibri']]),
      embeddedFontAuthoredFamilies: new Map([['__deck_calibri', 'calibri']]) };
    const { input } = paragraphInputRuns(para([textRun('§', { latin: 'Calibri' }, undefined)]),
      22, '#000', SCALE, false, true, 1, undefined, rc);
    const segment = input.find((i) => i.type === 'text');
    expect(segment?.type === 'text' && segment.style.faceFamily).toBeUndefined();
    expect(segment?.type === 'text' && segment.style.lineMetric).toBeUndefined();
  });
});

describe('fallback controls: line scope (#1689 fallback.win.pdf)', () => {
  /** Eight 18 pt Calibri lines in three paragraphs; `line` carries the target. */
  function fallbackBaselines(line: number, target: string, ea: string | undefined): number[] {
    const f: Faces = { latin: 'Calibri', cs: 'Calibri' };
    const br = { type: 'break' } as Paragraph['runs'][number];
    const texts = [1, 2, 3, 4, 5, 6, 7, 8].map((n) => `L0${n} Alpha beta gamma.`);
    const lineRuns = (n: number): Paragraph['runs'] => n === line
      ? [textRun(`L0${n} Alpha`, f, undefined, 18), textRun(target, f, ea, 18), textRun(' ends.', f, undefined, 18)]
      : [textRun(texts[n - 1], f, undefined, 18)];
    const group = (ns: number[]) => para(ns.flatMap((n, i) => (i ? [br, ...lineRuns(n)] : lineRuns(n))));
    const body = {
      verticalAnchor: 't', defaultFontSize: 18, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
      paragraphs: [group([1, 2, 3]), group([4, 5, 6]), group([7, 8])],
    } as unknown as TextBody;
    const ys: number[] = [];
    const ctx = {
      font: '', fillStyle: '', strokeStyle: '', direction: 'ltr', textAlign: 'left', textBaseline: 'alphabetic',
      measureText: (text: string) => ({ width: [...text].length * 8, actualBoundingBoxAscent: 7, actualBoundingBoxDescent: 2 }),
      fillText: (_t: string, _x: number, y: number) => { if (!ys.some((v) => Math.abs(v - y) < 1e-6)) ys.push(y); },
      fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
      translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
      moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    } as unknown as CanvasRenderingContext2D;
    renderTextBody(ctx, body, 0, 0, 408, 360, SCALE);
    expect(ys).toHaveLength(8);
    return ys;
  }

  it('moves only the line with the other face (F01: explicit MS Gothic italic, +0.72 on L05)', () => {
    const c = fallbackBaselines(5, ' §', 'Calibri');
    const t = fallbackBaselines(5, ' §', 'MS Gothic');
    t.forEach((y, i) => {
      if (i === 4) expect(Math.abs(y - c[i] - 0.72)).toBeLessThan(U);
      else expect(y - c[i], `L0${i + 1}`).toBeCloseTo(0, 9);
    });
    expect(t[4]).toBeGreaterThan(c[4]);
  });

  it('keeps every baseline of an empty-ea citation after Latin lines (F21: §3 in the cs face)', () => {
    const c = fallbackBaselines(7, ' §3', 'Calibri');
    const t = fallbackBaselines(7, ' §3', undefined);
    t.forEach((y, i) => expect(y - c[i], `L0${i + 1}`).toBeCloseTo(0, 9));
  });
});
