import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// Word-observed SEA (Thai/Lao/Khmer) line-fit behavior — issue #991, adjudicated
// against the Word PDF of the purpose-built calibration fixture (21-paragraph
// overflow sweep + 8 no-space-run placements; record on the issue):
//
// 1. SEA lines wrap at natural fit for every tested overflow ≥ +1pt across
//    5/9/13 inter-phrase spaces (negative-overflow controls fit). Independent
//    Latin Word controls also wrap at natural width; the former global 25%
//    trailing-space allowance was not a sound cross-script policy.
//
// 2. Dictionary boundaries are SECONDARY break opportunities. A no-space SEA
//    chunk that does not fit the remaining width of a non-empty line moves to
//    the next line WHOLE when it fits a full line by itself — Word never
//    splits it mid-chunk to fill the current line (invariant across remaining
//    widths 305–348pt, across run splits into multiple w:r, and with/without a
//    leading tab). Only a chunk wider than a full line breaks at dictionary
//    boundaries (calibration part II-D; both controls broke at exactly our
//    boundaries, so the dictionaries agree).
//
// The linear stub advances FONT_PX per code point, making the arithmetic
// explicit. Thai combining marks count as full advances here — irrelevant, as
// every expected number below is computed in the same model.

const FONT_PX = 12;

function makeLinearCanvas(): HTMLCanvasElement {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => Number(/([\d.]+)px/.exec(font)?.[1] ?? FONT_PX);
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    fontKerning: 'auto',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0, style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return canvas as unknown as HTMLCanvasElement;
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'serif', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(runs: DocRun[], alignment: DocParagraph['alignment']): BodyElement {
  const p: DocParagraph = {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: FONT_PX, defaultFontFamily: 'serif',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function textPara(text: string, alignment: DocParagraph['alignment']): BodyElement {
  return para([{ type: 'text', ...textRun(text) } as DocRun], alignment);
}

function section(pageWidth: number): SectionProps {
  return {
    pageWidth, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    docGridCharSpace: undefined,
  } as SectionProps;
}

function doc(el: BodyElement, sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body: [el],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function renderLines(el: BodyElement, pageWidth: number): Promise<string[]> {
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(el, section(pageWidth)), makeLinearCanvas(), 0, {
    dpr: 1,
    width: pageWidth,
    onTextRun: (r) => { if (r.text && r.text.trim()) runs.push(r); },
  });
  const byY = new Map<number, DocxTextRunInfo[]>();
  for (const r of runs) {
    const key = Math.round(r.y);
    let arr = byY.get(key);
    if (!arr) { arr = []; byY.set(key, arr); }
    arr.push(r);
  }
  return [...byY.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, rs]) => rs.slice().sort((p, q) => p.x - q.x).map((r) => r.text).join(''));
}

const norm = (l: string) => l.replace(/\s+/g, '');

// ── Rule 1 arithmetic ────────────────────────────────────────────────────────
// 'มาก ' token = 4 cp = 48px (word 36 + trailing space 12); three tokens = 144.
// Final word 'ดังนี้' = 6 cp = 72 (no interior dictionary boundary). Natural end
// = 216. Column 210 ⇒ overflow 6px, under the Latin-style 25% budget
// (0.25 × 36 = 9px) — a Latin line would ADMIT; the SEA line must WRAP.
const THAI_FINAL = 'มาก มาก มาก ดังนี้';

// ── Rule 2 arithmetic ────────────────────────────────────────────────────────
// Lead token 'มาก ' = 48. Chunk 'น้อยน้อยน้อยน้อย' = 16 cp = 192, dictionary
// boundaries at 4/8/12 (48px words). Column 200: 48+192 = 240 overflows; the
// old greedy fill split the chunk at offset 12 (144 ≤ 152 remaining). The
// chunk alone fits a full line (192 ≤ 200) ⇒ Word moves it WHOLE.
const THAI_CHUNK = 'มาก น้อยน้อยน้อยน้อย';

describe('issue #991 — SEA justified fit (Word calibration-fixture rules)', () => {
  it('Rule 1: wraps the paragraph-final Thai word on a thaiDistribute closing line (zero space-shrink)', async () => {
    const lines = await renderLines(textPara(THAI_FINAL, 'thaiDistribute'), 210);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มากมากมาก');
    expect(norm(lines[1])).toBe('ดังนี้');
  });

  it('Rule 1: the suppression is script-scoped, not jc-scoped (left-aligned Thai wraps too)', async () => {
    const lines = await renderLines(textPara(THAI_FINAL, 'left'), 210);
    expect(lines.length).toBe(2);
    expect(norm(lines[1])).toBe('ดังนี้');
  });

  it('Rule 1 control: content that genuinely fits stays on one line', async () => {
    const lines = await renderLines(textPara(THAI_FINAL, 'thaiDistribute'), 216);
    expect(lines.length).toBe(1);
  });

  it('Rule 1: a SEA candidate on a Latin line also gets zero shrink (candidate-inclusive)', async () => {
    // 'AAAA ' ×3 = 180, Thai final word 'ดังนี้' = 72 ⇒ natural end 252; column
    // 246 ⇒ overflow 6 ≤ 0.25×36 — a Latin candidate would be admitted, but
    // admitting the Thai word makes the line SEA, so the same zero-shrink fit
    // applies and it wraps (mixed-script GT uncollected; conservative wrap).
    const lines = await renderLines(textPara('AAAA AAAA AAAA ดังนี้', 'left'), 246);
    expect(lines.length).toBe(2);
    expect(norm(lines[1])).toBe('ดังนี้');
  });

  it('Rule 1/2 guard: grapheme-fill scripts (Myanmar) keep the per-cluster greedy fill', async () => {
    // Myanmar is SEA-marked but grapheme-fill (#961): both #991 rules must NOT
    // apply. Lead 'မာ ' = 3 cp = 36; run 'စု'×7 = 14 cp = 168 ('စ' U+1005 +
    // nonspacing 'ု' U+102F = ONE grapheme cluster; boundaries every 2 cp —
    // unlike U+102C, a UAX#29 SpacingMark exception that clusters alone).
    // Column 200: the run fits a full line (168 ≤ 200) but not the remainder
    // (164) — a dictionary chunk would move whole; the grapheme-fill run must
    // instead fill the line cluster-by-cluster (6 clusters = 144 ≤ 164 < 168).
    const lines = await renderLines(textPara('မာ စုစုစုစုစုစုစု', 'left'), 200);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('မာစုစုစုစုစုစု');
    expect(norm(lines[1])).toBe('စု');
  });

  it('Rule 2 guard: a segment mixing dictionary and grapheme-fill SEA is not moved whole', async () => {
    // 'น้อยน้อยစုစုစုစุ'-style mixed runs are NOT dictionary-SEA (per-codepoint
    // scan), so the chunk pre-flush must not fire even though the run fits a
    // full line: the greedy SEA fill splits it at the merged offset set.
    // Lead 'มาก ' = 48; mixed run 'น้อยน้อย' + 'စုစုစုစု' = 16 cp = 192 ≤ 200,
    // remaining = 152 ⇒ greedy split keeps a prefix on line 1.
    const lines = await renderLines(textPara('มาก น้อยน้อยစုစုစုစု', 'left'), 200);
    expect(lines.length).toBe(2);
    expect(norm(lines[0]).length, 'a prefix of the mixed run stays on line 1 (greedy)').toBeGreaterThan(3);
    expect(norm(lines.join(''))).toBe('มากน้อยน้อยစုစုစုစု');
  });

  it('Rule 1 guard: a Latin line also wraps at natural width', async () => {
    // 'AAAA ' ×3 = 180, final 'AAAA' = 48 ⇒ natural end 228, column 225.
    const lines = await renderLines(textPara('AAAA AAAA AAAA AAAA', 'left'), 225);
    expect(lines.length).toBe(2);
  });

  it('Rule 2: a no-space chunk that fits a full line moves whole instead of splitting', async () => {
    const lines = await renderLines(textPara(THAI_CHUNK, 'thaiDistribute'), 200);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มาก');
    expect(norm(lines[1])).toBe('น้อยน้อยน้อยน้อย');
  });

  it('Rule 2: same movement under a non-justified alignment', async () => {
    const lines = await renderLines(textPara(THAI_CHUNK, 'left'), 200);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มาก');
    expect(norm(lines[1])).toBe('น้อยน้อยน้อยน้อย');
  });

  it('Rule 2: the chunk spans w:r boundaries (a run split is not a break opportunity)', async () => {
    const el = para(
      [
        { type: 'text', ...textRun('มาก ') } as DocRun,
        { type: 'text', ...textRun('น้อยน้อย') } as DocRun,
        { type: 'text', ...textRun('น้อยน้อย') } as DocRun,
      ],
      'thaiDistribute',
    );
    const lines = await renderLines(el, 200);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มาก');
    expect(norm(lines[1])).toBe('น้อยน้อยน้อยน้อย');
  });

  it('Rule 2 guard: an over-full chunk still fills the line greedily at dictionary boundaries', async () => {
    // Column 180 < chunk 192 ⇒ the chunk cannot fit any line whole; Word breaks
    // it at dictionary boundaries (calibration part II-D), which is the greedy
    // fill: remaining after the lead = 132 ⇒ split at offset 8 (96 ≤ 132 < 144).
    const lines = await renderLines(textPara(THAI_CHUNK, 'thaiDistribute'), 180);
    expect(lines.length).toBe(2);
    expect(norm(lines[0])).toBe('มากน้อยน้อย');
    expect(norm(lines[1])).toBe('น้อยน้อย');
  });

  it('Rule 2 control: a chunk that fits the remaining width stays on the line', async () => {
    const lines = await renderLines(textPara('มาก น้อยน้อย', 'thaiDistribute'), 200);
    expect(lines.length).toBe(1);
  });
});
