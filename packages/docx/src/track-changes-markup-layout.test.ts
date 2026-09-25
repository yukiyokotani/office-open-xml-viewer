import { describe, it, expect } from 'vitest';
import type {
  BodyElement, DocParagraph, DocxTextRun, DocxDocumentModel, RunRevision, SectionProps,
} from './types';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { DeepReadonly, DocumentLayout, TextPlacement } from './layout/types.js';
import { WORD_TRACK_CHANGE_AUTHOR_COLORS } from './layout/paint-compatibility.js';

// ECMA-376 §17.13.5 markup view rendering (`word-track-change-decoration` +
// `word-track-change-author-palette` + `word-track-change-bar`): insertions
// and moved-in text are underlined, deletions and moved-away text struck
// through, both in the stable per-author colour (first appearance in document
// run order indexes the eight-colour palette), and every line containing
// revision content gets a vertical change bar in the left page margin. The
// default final-view variant carries none of this.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

type DocRun = DocParagraph['runs'][number];
function textRun(text: string, revision?: RunRevision): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
    ...(revision ? { revision } : {}),
  } as unknown as DocxTextRun;
  return { type: 'text', ...run } as DocRun;
}
function para(...runs: DocRun[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

function doc(body: BodyElement[]): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 600, pageHeight: 400,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function revisionDoc(): DocxDocumentModel {
  return doc([
    para(
      textRun('kept '),
      textRun('added', { kind: 'insertion', author: 'Alice' }),
      textRun('gone', { kind: 'deletion', author: 'Bob' }),
      textRun('moved-away', { kind: 'moveFrom', author: 'Bob' }),
      textRun('moved-in', { kind: 'moveTo', author: 'Bob' }),
    ),
    para(textRun('plain paragraph')),
  ]);
}

function layoutOf(model: DocxDocumentModel, showTrackedChanges?: boolean): DeepReadonly<DocumentLayout> {
  return layoutDocument(
    model,
    createLayoutServices(model, { measureContext: makeCtx() }),
    { currentDateMs: 0, ...(showTrackedChanges === undefined ? {} : { showTrackedChanges }) },
  ) as DeepReadonly<DocumentLayout>;
}

function textPlacements(layout: DeepReadonly<DocumentLayout>): DeepReadonly<TextPlacement>[] {
  return layout.pages.flatMap((page) =>
    page.layers.body.flatMap((node) => node.kind === 'paragraph'
      ? node.lines.flatMap((line) =>
          line.placements.filter((placement) => placement.kind === 'text'))
      : [])) as DeepReadonly<TextPlacement>[];
}

function placementByText(layout: DeepReadonly<DocumentLayout>, text: string): DeepReadonly<TextPlacement> {
  const placement = textPlacements(layout).find((candidate) => candidate.text === text);
  if (!placement) throw new Error(`No text placement "${text}"`);
  return placement;
}

const [ALICE_COLOR, BOB_COLOR] = WORD_TRACK_CHANGE_AUTHOR_COLORS;

describe('markup-view revision decorations (§17.13.5)', () => {
  it('underlines insertions and moveTo, strikes deletions and moveFrom, in first-appearance author colours', () => {
    const layout = layoutOf(revisionDoc(), true);
    const decorationsOf = (text: string) =>
      placementByText(layout, text).decorations.map(({ kind, color }) => ({ kind, color }));
    expect(decorationsOf('added')).toEqual([{ kind: 'underline', color: ALICE_COLOR }]);
    expect(decorationsOf('gone')).toEqual([{ kind: 'strikethrough', color: BOB_COLOR }]);
    expect(decorationsOf('moved-away')).toEqual([{ kind: 'strikethrough', color: BOB_COLOR }]);
    expect(decorationsOf('moved-in')).toEqual([{ kind: 'underline', color: BOB_COLOR }]);
    expect(decorationsOf('kept ')).toEqual([]);
  });

  it('an authorless revision still receives a stable palette colour', () => {
    const layout = layoutOf(doc([para(textRun('anon', { kind: 'insertion' }))]), true);
    expect(placementByText(layout, 'anon').decorations).toEqual([
      expect.objectContaining({ kind: 'underline', color: ALICE_COLOR }),
    ]);
  });

  it('the default final view carries no revision decorations', () => {
    const layout = layoutOf(revisionDoc());
    for (const placement of textPlacements(layout)) {
      expect(placement.decorations).toEqual([]);
    }
  });
});

describe('markup-view margin change bars (word-track-change-bar)', () => {
  it('emits one margin change bar per line containing revision text in the markup view', () => {
    const layout = layoutOf(revisionDoc(), true);
    const page = layout.pages[0]!;
    const bars = page.changeBars ?? [];
    // The revision paragraph occupies one line; the plain paragraph none.
    expect(bars).toHaveLength(1);
    const bar = bars[0]!;
    // Centered in the 20pt left margin at the fixed 0.75pt convention width.
    expect(bar.bounds.widthPt).toBeCloseTo(0.75, 6);
    expect(bar.bounds.xPt).toBeCloseTo(20 / 2 - 0.75 / 2, 6);
    // Spans the revision line's vertical extent.
    const revisionLineBounds = (() => {
      for (const node of page.layers.body) {
        if (node.kind !== 'paragraph') continue;
        for (const line of node.lines) {
          if (line.placements.some((placement) =>
            placement.kind === 'text' && placement.revision !== undefined)) {
            return line.bounds;
          }
        }
      }
      throw new Error('No revision line');
    })();
    expect(bar.bounds.yPt).toBeCloseTo(revisionLineBounds.yPt, 6);
    expect(bar.bounds.heightPt).toBeCloseTo(revisionLineBounds.heightPt, 6);
  });

  it('the default final view attaches no change bars', () => {
    const layout = layoutOf(revisionDoc());
    expect(layout.pages[0]!.changeBars).toBeUndefined();
  });
});
