import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { wordTrackChangeAuthorColor } from './layout/track-changes.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// Characterization of the docx track-changes decoration draw path
// (ECMA-376 §17.13.5 `<w:ins>` / `<w:del>`): inserted runs gain an
// author-coloured underline layered on any authored run decorations.
// Deleted (`<w:del>`) and moved-away (`<w:moveFrom>`) runs are projected out
// of the laid-out final state upstream (line-layout.ts), so the canvas shows
// the accepted document; their review markup is consumer-owned via the parsed
// revision metadata. The recording ctx mirrors the underline draw-path tests
// so stroke geometry / colour can be asserted without a real canvas.

interface StrokeOp { op: string; args: (number | string)[]; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; ops: StrokeOp[] } {
  let font = '10px serif';
  const ops: StrokeOp[] = [];
  let strokeStyle = '#000';
  let fillStyle = '#000';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(v: string) { strokeStyle = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      const lowLine = s === '_';
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: lowLine ? 0 : p * 0.8,
        actualBoundingBoxDescent: lowLine ? p * 0.05 : p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo(x: number, y: number) { ops.push({ op: 'moveTo', args: [x, y] }); },
    lineTo(x: number, y: number) { ops.push({ op: 'lineTo', args: [x, y] }); },
    stroke() { ops.push({ op: 'stroke', args: [strokeStyle] }); },
    fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {},
    clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(t: string) { ops.push({ op: 'fillText', args: [String(t), fillStyle] }); },
    strokeText() {},
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, ops };
}

function para(run: Partial<DocParagraph['runs'][number]>): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: 'abc', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 40, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
      ...run,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 40, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[]): DocxDocumentModel {
  const section = {
    pageWidth: 400, pageHeight: 600,
    marginTop: 5, marginRight: 5, marginBottom: 5, marginLeft: 5,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(...runs: Partial<DocParagraph['runs'][number]>[]) {
  const { canvas, ops } = makeRecordingCanvas();
  const body = runs.map((run) => para(run) as unknown as BodyElement);
  await renderDocumentToCanvas(doc(body), canvas, 0, { dpr: 1, width: 400 });
  return ops;
}

function strokeColors(ops: StrokeOp[]): string[] {
  return ops
    .filter((op) => op.op === 'stroke')
    .map((op) => String(op.args[0]).toLowerCase());
}

describe('docx track-changes decorations (§17.13.5) draw path', () => {
  it('underlines an inserted run in the deterministic author palette color', async () => {
    const ops = await render({ revision: { kind: 'insertion', author: 'Alice' } });
    expect(strokeColors(ops)).toContain(wordTrackChangeAuthorColor('Alice').toLowerCase());
  });

  it('projects deleted runs out of the final-state layout (no text, no decoration)', async () => {
    // Upstream's accepted-final-state projection (line-layout.ts §17.13.5)
    // drops `<w:del>` runs before segment building, so deleted text neither
    // paints nor carries decorations; review UI reads the parsed metadata.
    const ops = await render(
      { text: 'kept' },
      { text: 'deleted', revision: { kind: 'deletion', author: 'Alice' } },
    );
    const drawnText = ops
      .filter((op) => op.op === 'fillText')
      .map((op) => String(op.args[0]))
      .join('');
    expect(drawnText).toContain('kept');
    expect(drawnText).not.toContain('deleted');
    expect(strokeColors(ops)).toHaveLength(0);
  });

  it('draws no decoration for a run without a revision', async () => {
    const ops = await render({});
    expect(strokeColors(ops)).toHaveLength(0);
  });

  it('gives different authors different palette colors', async () => {
    const ops = await render(
      { revision: { kind: 'insertion', author: 'Alice' } },
      { revision: { kind: 'insertion', author: 'Carol' } },
    );
    const colors = strokeColors(ops);
    expect(colors).toContain(wordTrackChangeAuthorColor('Alice').toLowerCase());
    expect(colors).toContain(wordTrackChangeAuthorColor('Carol').toLowerCase());
    expect(wordTrackChangeAuthorColor('Alice')).not.toBe(wordTrackChangeAuthorColor('Carol'));
  });

  it('stacks the author-coloured revision underline on an authored underline', async () => {
    const ops = await render({
      underline: true, underlineStyle: 'dotted', underlineColor: '112233',
      revision: { kind: 'insertion', author: 'Alice' },
    });
    const colors = strokeColors(ops);
    expect(colors).toContain('#112233');
    expect(colors).toContain(wordTrackChangeAuthorColor('Alice').toLowerCase());
  });

  it('keeps the run text color unchanged by the revision decoration color', async () => {
    const ops = await render({ revision: { kind: 'insertion', author: 'Alice' } });
    const textFills = ops
      .filter((op) => op.op === 'fillText')
      .map((op) => String(op.args[1]).toLowerCase());
    expect(textFills.length).toBeGreaterThan(0);
    expect(textFills).not.toContain(wordTrackChangeAuthorColor('Alice').toLowerCase());
  });
});
