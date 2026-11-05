import { getDefaultBidiEngine, setBidiEngine } from '@silurus/ooxml-core';
import type { TextShapeRequest } from './layout/text.js';
import { expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { createLayoutServicesRuntimeView } from './layout/runtime-state.js';
import { layoutDocument } from './document-layout.js';
import type { BodyElement, DocxDocumentModel, SectionProps } from './types.js';
import type { ParagraphLayout } from './layout/types.js';

const content = 'ابتXray Yoke ابت.';
function model(lateFrame = false, firstIndent = 0, height = 100): DocxDocumentModel {
  const paragraph = { type: 'paragraph', alignment: 'left', bidi: true,
    indentLeft: 0, indentRight: 0, indentFirst: firstIndent, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: null, numbering: null, tabStops: [], defaultFontSize: 10,
    runs: [{ type: 'text', text: content, fontSize: 10, fontSizeCs: 10, kerning: 1,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontFamily: 'Fixture', fontFamilyCs: 'Fixture', color: null,
      background: null, vertAlign: null, isLink: false, hyperlink: null }],
  } as unknown as Extract<BodyElement, { type: 'paragraph' }>;
  const body: BodyElement[] = [paragraph];
  if (lateFrame) body.push({ ...paragraph, runs: [], framePr: {
    dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
    hRule: 'exact', hSpace: 0, vSpace: 0, w: 10, h: 10,
  } });
  return { body, section: { pageWidth: 100, pageHeight: height,
    marginTop: 10, marginBottom: 10, marginLeft: 10, marginRight: 10,
    headerDistance: 0, footerDistance: 0, columns: null } as SectionProps,
    headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}
function context() {
  return { font: '10px serif', direction: 'inherit', textAlign: 'start', letterSpacing: '0px', fontKerning: 'none',
    measureText(this: { fontKerning: string; direction?: string }, text: string) {
      // Independent nonzero SPACE+Y pair; the native cut must not disable it or
      // merge the two words into one retained paint operation.
      const width = text.length * 5 - (this.fontKerning === 'normal' && text.includes(' Y') ? 2 : 0);
      const suffix = text.length - text.trimEnd().length;
      const visibleWidth = width - suffix * 5;
      const leftInk = this.direction === 'rtl' && visibleWidth > 0 ? suffix * 5 : 0;
      return { width, actualBoundingBoxLeft: -leftInk,
        actualBoundingBoxRight: visibleWidth > 0 ? leftInk + visibleWidth : 0,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: this.direction === 'rtl' && suffix > 0 && visibleWidth > 0 ? 7 : 2,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}
function acquire(document: DocxDocumentModel) {
  const ctx = context();
  const services = createLayoutServices(document, { measureContext: ctx });
  const layout = layoutDocument(document, services);
  return { ctx, layout, paragraphs: layout.pages.flatMap(page => page.layers.body)
    .filter((node): node is ParagraphLayout => node.kind === 'paragraph') };
}
it('acquires strong native cuts before wrapping without losing a same-level SPACE+Y pair or source ownership', () => {
  const { paragraphs, ctx } = acquire(model());
  const lines = paragraphs.flatMap(paragraph => paragraph.lines);
  const pieces = lines.flatMap(line => line.placements).filter(piece => piece.kind === 'text');
  const visible = pieces.flatMap(piece => piece.paintOps).filter(op => op.text.trim());
  expect(visible.some(op => /[ابت]/u.test(op.text) && /[A-Za-z]/u.test(op.text))).toBe(false);
  expect(visible.filter(op => /[ابت]/u.test(op.text)).every(op => op.direction === 'rtl')).toBe(true);
  expect(visible.filter(op => /[A-Za-z]/u.test(op.text) || op.text === '.').every(op => op.direction === 'ltr')).toBe(true);
  expect(visible.some(op => op.text === '.')).toBe(true);
  expect(visible.some(op => op.text.includes('Xray Yoke'))).toBe(false);
  const logical = [...pieces].sort((a, b) => a.range.start - b.range.start);
  let cursor = 0;
  for (const piece of logical) {
    expect(piece.range.start).toBe(cursor);
    expect(piece.range.end - piece.range.start).toBe(piece.text.length);
    expect(content.slice(piece.range.start, piece.range.end)).toBe(piece.text);
    cursor = piece.range.end;
  }
  expect(cursor).toBe(content.length);
  expect(lines).toHaveLength(2);
  expect(lines[0]!.placements.filter(piece => piece.kind === 'text').reduce((sum, piece) => sum + piece.advancePt, 0)).toBe(63);
  // Full-source SPACE12 belonged to the even Latin unit before wrapping.
  // L1 resets this actual line tail to paragraph level1: it becomes a charged
  // zero-ink cell at the physical LEFT, not part of the visible LTR operation.
  const firstPieces = lines[0]!.placements.filter(piece => piece.kind === 'text');
  const tail = firstPieces.find(piece => piece.range.start === 12 && piece.range.end === 13);
  expect(tail?.text).toBe(' ');
  expect(tail?.advancePt).toBe(5);
  expect(tail?.origin.xPt).toBe(Math.min(...firstPieces.map(piece => piece.origin.xPt)));
  expect(tail?.paintOps.every(op => !op.text.trim())).toBe(true);

  expect([ctx.direction, ctx.textAlign, ctx.letterSpacing]).toEqual(['inherit', 'start', '0px']);
});
it('keeps the original native packet when a later frame makes future body bands variable', () => {
  const { paragraphs } = acquire(model(true));
  const operations = paragraphs[0]!.lines.flatMap(line => line.placements)
    .filter(piece => piece.kind === 'text').flatMap(piece => piece.paintOps);
  expect(operations.some(op => /[ابت]/u.test(op.text) && /[A-Za-z]/u.test(op.text))).toBe(true);
});

it('declines a direction-aware custom provider without the projectable-source opt-in before new requests', () => {
  const document = model();
  const services = createLayoutServices(document, { measureContext: context() });
  const { acquireProjectableNativeUnits: _capability, ...ordinary } = services.text;
  let requests = 0;
  const legacy = createLayoutServicesRuntimeView(services, { text: Object.freeze({ ...ordinary,
    shape(request: Readonly<TextShapeRequest>) { requests++; expect(request.baseDirection).toBeUndefined(); return ordinary.shape(request); },
    shapeDirected() { throw new Error('Unrequested directed provider call'); },
  }) });
  const layout = layoutDocument(document, legacy);
  const operations = layout.pages.flatMap(page => page.layers.body)
    .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
    .flatMap(paragraph => paragraph.lines).flatMap(line => line.placements)
    .filter(piece => piece.kind === 'text').flatMap(piece => piece.paintOps);
  expect(requests).toBeGreaterThan(0);
  expect(operations.some(op => /[ابت]/u.test(op.text) && /[A-Za-z]/u.test(op.text))).toBe(true);
});
it('keeps the original first-indent eligibility and source topology across continuations', () => {
  const ordinaryWord = 'ابتABCDEFGHIJKL'; // 75pt: fits column80, not original first-indent band70.
  const texts = [Array(4).fill(content.slice(0, -1)).join(' ') + '.',
    Array(4).fill(content.slice(0, -1)).join(' ') + ' ' + ordinaryWord + ' ابت.'];
  for (const [index, text] of texts.entries()) {
    const document = model(false, 10, 44);
    const paragraph = document.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
    paragraph.runs = [{ ...paragraph.runs[0]!, type: 'text', text } as typeof paragraph.runs[number]];
    const { layout, paragraphs } = acquire(document);
    expect(layout.pages.length).toBeGreaterThan(2);
    const pieces = paragraphs.flatMap(p => p.lines).flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text').sort((a, b) => a.range.start - b.range.start);
    let cursor = 0;
    for (const piece of pieces) {
      expect(piece.range.start).toBe(cursor);
      expect(text.slice(piece.range.start, piece.range.end)).toBe(piece.text);
      expect(piece.range.end - piece.range.start).toBe(piece.text.length);
      if (index === 0) expect(piece.paintOps.some(op => /[ابت]/u.test(op.text) && /[A-Za-z]/u.test(op.text))).toBe(false);
      cursor = piece.range.end;
    }
    expect(cursor).toBe(text.length);
    if (index === 1) expect(pieces.flatMap(piece => piece.paintOps).some(op => op.text === ordinaryWord)).toBe(true);
  }
});
it('does not reacquire mixed-direction words whose existing native operations are already uniform', () => {
  for (const text of ['Xray ابت Yoke ']) {
    const document = model();
    const paragraph = document.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
    paragraph.runs = [{ ...paragraph.runs[0]!, type: 'text', text } as typeof paragraph.runs[number]];
    const services = createLayoutServices(document, { measureContext: context() });
    let bundles = 0;
    const observed = createLayoutServicesRuntimeView(services, { text: Object.freeze({ ...services.text,
      acquireProjectableNativeUnits(requests: readonly Readonly<TextShapeRequest & { baseDirection: 'ltr' | 'rtl' }>[]) { bundles++; return services.text.acquireProjectableNativeUnits!(requests); },
    }) });
    const layout = layoutDocument(document, observed);
    expect(layout.pages.length).toBeGreaterThan(0);
    expect(bundles).toBe(0);
  }
});

it('does not acquire native bundles under a forwarding engine without the pinned native certificate', () => {
  const native = getDefaultBidiEngine();
  const document = model();
  const services = createLayoutServices(document, { measureContext: context() });
  let bundles = 0;
  const observed = createLayoutServicesRuntimeView(services, { text: Object.freeze({ ...services.text,
    acquireProjectableNativeUnits(requests: readonly Readonly<TextShapeRequest & { baseDirection: 'ltr' | 'rtl' }>[]) {
      bundles++; return services.text.acquireProjectableNativeUnits!(requests);
    },
  }) });
  try {
    setBidiEngine({ computeLevels: (...args) => native.computeLevels(...args),
      reorderVisual: (...args) => native.reorderVisual(...args), getMirror: code => native.getMirror(code) });
    const layout = layoutDocument(document, observed);
    const operations = layout.pages.flatMap(page => page.layers.body)
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
      .flatMap(paragraph => paragraph.lines).flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text').flatMap(piece => piece.paintOps);
    expect(bundles).toBe(0);
    expect(operations.some(op => /[ابت]/u.test(op.text) && /[A-Za-z]/u.test(op.text))).toBe(true);
  } finally { setBidiEngine(native); }
});

it('uses owned word-only ink for a projected RTL suffix on cold and warm shape caches', () => {
  const document = model();
  const paragraph = document.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
  paragraph.runs = [{ ...paragraph.runs[0]!, type: 'text', text: 'ابت ابتXray Yoke ابت.', underline: true } as typeof paragraph.runs[number]];
  const services = createLayoutServices(document, { measureContext: context() });
  for (let pass = 0; pass < 2; pass++) {
    const layout = layoutDocument(document, services);
    const pieces = layout.pages.flatMap(page => page.layers.body)
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
      .flatMap(paragraph => paragraph.lines).flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text');
    const word = pieces.find(piece => piece.range.start === 0 && piece.range.end === 3);
    const space = pieces.find(piece => piece.range.start === 3 && piece.range.end === 4);
    const operation = word?.paintOps.find(op => op.range.start === 0 && op.range.end === 3);
    expect(operation?.text).toBe('ابت');
    expect(operation?.direction).toBe('rtl');
    // The actual painted request is word-only: the fake native backend gives
    // it [0,15], while full RTL word+SPACE has shifted ink [5,20].
    expect(operation?.inkBounds?.xMinPt).toBe(0);
    expect(operation?.inkBounds?.xMaxPt).toBe(15);
    expect(operation?.blockAxisInkBounds).toEqual({ startPt: -8, endPt: 2 });
    const ownerLine = layout.pages.flatMap(page => page.layers.body)
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
      .flatMap(paragraph => paragraph.lines).find(line => line.placements.includes(word!));
    const underlines = ownerLine!.placements.filter(piece => piece.kind === 'text')
      .flatMap(piece => piece.decorations).filter(decoration => decoration.kind === 'underline'
        && decoration.from.xPt <= word!.origin.xPt
        && decoration.to.xPt >= word!.origin.xPt + word!.advancePt);
    expect(underlines.length).toBeGreaterThan(0);
    expect(underlines.map(underline => underline.from.yPt)).toEqual([(word?.origin.yPt ?? 0) + 7]);
    expect(space?.text).toBe(' ');
    expect(space?.advancePt).toBe(5);
    expect(space?.paintOps.every(op => !op.text.trim())).toBe(true);
  }
});

it('preserves source coverage when the engine changes between paragraph page acquisitions', () => {
  const document = model(false, 10, 44);
  const paragraph = document.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
  const text = Array(6).fill(content.slice(0, -1)).join(' ') + '.';
  paragraph.runs = [{ ...paragraph.runs[0]!, type: 'text', text } as typeof paragraph.runs[number]];
  const services = createLayoutServices(document, { measureContext: context() });
  const native = getDefaultBidiEngine();
  let bundles = 0;
  const observed = createLayoutServicesRuntimeView(services, { text: Object.freeze({ ...services.text,
    acquireProjectableNativeUnits(requests: Parameters<NonNullable<typeof services.text.acquireProjectableNativeUnits>>[0]) {
      const result = services.text.acquireProjectableNativeUnits!(requests);
      if (++bundles === 2) {
        setBidiEngine({ computeLevels: (...args) => native.computeLevels(...args),
          reorderVisual: (...args) => native.reorderVisual(...args), getMirror: code => native.getMirror(code) });
      }
      return result;
    },
  }) });
  try {
    const layout = layoutDocument(document, observed);
    expect(bundles).toBe(2);
    const firstPageOperations = layout.pages[0]!.layers.body
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
      .flatMap(paragraph => paragraph.lines).flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text').flatMap(piece => piece.paintOps);
    expect(firstPageOperations.some(op => /[A-Za-z]/u.test(op.text))).toBe(true);
    expect(firstPageOperations.some(op => /[ابت]/u.test(op.text) && /[A-Za-z]/u.test(op.text))).toBe(false);
    const pieces = layout.pages.flatMap(page => page.layers.body)
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
      .flatMap(paragraph => paragraph.lines).flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text').sort((a, b) => a.range.start - b.range.start);
    let cursor = 0;
    for (const piece of pieces) {
      expect(piece.range.start).toBe(cursor);
      expect(text.slice(piece.range.start, piece.range.end)).toBe(piece.text);
      cursor = piece.range.end;
    }
    expect(cursor).toBe(text.length);
  } finally { setBidiEngine(native); }
});

it('declines new native units when a strong level seam crosses an original Prepend grapheme', () => {
  const document = model();
  const paragraph = document.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
  const text = 'ب\u0D4Eب.';
  paragraph.runs = [{ ...paragraph.runs[0]!, type: 'text', text } as typeof paragraph.runs[number]];
  const services = createLayoutServices(document, { measureContext: context() });
  let bundles = 0;
  const observed = createLayoutServicesRuntimeView(services, { text: Object.freeze({ ...services.text,
    acquireProjectableNativeUnits(requests: Parameters<NonNullable<typeof services.text.acquireProjectableNativeUnits>>[0]) {
      bundles++;
      return services.text.acquireProjectableNativeUnits!(requests);
    },
  }) });
  const layout = layoutDocument(document, observed);
  expect(bundles).toBe(0);
  const pieces = layout.pages.flatMap(page => page.layers.body)
    .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
    .flatMap(paragraph => paragraph.lines).flatMap(line => line.placements)
    .filter(piece => piece.kind === 'text').sort((left, right) => left.range.start - right.range.start);
  let cursor = 0;
  for (const piece of pieces) {
    expect(piece.range.start).toBe(cursor);
    expect(text.slice(piece.range.start, piece.range.end)).toBe(piece.text);
    cursor = piece.range.end;
  }
  expect(cursor).toBe(text.length);
});

it('rejects a missing current directed receipt after a valid native preflight bundle', () => {
  const document = model();
  const services = createLayoutServices(document, { measureContext: context() });
  let bundles = 0;
  const observed = createLayoutServicesRuntimeView(services, { text: Object.freeze({ ...services.text,
    acquireProjectableNativeUnits(requests: Parameters<NonNullable<typeof services.text.acquireProjectableNativeUnits>>[0]) {
      bundles++;
      return services.text.acquireProjectableNativeUnits!(requests);
    },
    shapeDirected(request: Parameters<NonNullable<typeof services.text.shapeDirected>>[0]) {
      return { ...services.text.shapeDirected!(request), nativeBaseDirection: undefined };
    },
  }) });
  expect(() => layoutDocument(document, observed)).toThrow(/directed.*receipt/i);
  expect(bundles).toBe(1);
});

it('preserves source coverage when an ordinary continuation later acquires native units', () => {
  const document = model(false, 10, 44);
  const paragraph = document.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
  const text = Array(6).fill(content.slice(0, -1)).join(' ') + '.';
  paragraph.runs = [{ ...paragraph.runs[0]!, type: 'text', text } as typeof paragraph.runs[number]];
  const services = createLayoutServices(document, { measureContext: context() });
  const native = getDefaultBidiEngine();
  let initialSourceRequests = 0, switches = 0, bundles = 0;
  const observed = createLayoutServicesRuntimeView(services, { text: Object.freeze({ ...services.text,
    shape(request: Readonly<TextShapeRequest>) {
      // The initial ordinary word request owns the complete source context.
      // Observe its source-start call before the original service's shape cache.
      if (request.measure === false && request.baseDirection === undefined
        && request.text.length > 0 && request.substituteContext?.text === text
        && request.substituteContext.offset === 0 && ++initialSourceRequests === 2) {
        switches++;
        setBidiEngine(native);
      }
      return services.text.shape(request);
    },
    acquireProjectableNativeUnits(requests: Parameters<NonNullable<typeof services.text.acquireProjectableNativeUnits>>[0]) {
      const result = services.text.acquireProjectableNativeUnits!(requests);
      if (result) bundles++;
      return result;
    },
  }) });
  setBidiEngine({ computeLevels: (...args) => native.computeLevels(...args),
    reorderVisual: (...args) => native.reorderVisual(...args), getMirror: code => native.getMirror(code) });
  try {
    const layout = layoutDocument(document, observed);
    expect(initialSourceRequests).toBeGreaterThanOrEqual(2);
    expect(switches).toBe(1);
    expect(bundles).toBeGreaterThan(0);
    const operations = (pages: typeof layout.pages) => pages.flatMap(page => page.layers.body)
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
      .flatMap(paragraph => paragraph.lines).flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text').flatMap(piece => piece.paintOps);
    const mixed = (text: string) => /[ابت]/u.test(text) && /[A-Za-z]/u.test(text);
    expect(operations(layout.pages.slice(0, 1)).some(op => mixed(op.text))).toBe(true);
    const later = operations(layout.pages.slice(1));
    expect(later.some(op => /[A-Za-z]/u.test(op.text))).toBe(true);
    expect(later.some(op => mixed(op.text))).toBe(false);
    const pieces = layout.pages.flatMap(page => page.layers.body)
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
      .flatMap(paragraph => paragraph.lines).flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text').sort((a, b) => a.range.start - b.range.start);
    let cursor = 0;
    for (const piece of pieces) {
      expect(piece.range.start).toBe(cursor);
      expect(piece.range.end - piece.range.start).toBe(piece.text.length);
      expect(text.slice(piece.range.start, piece.range.end)).toBe(piece.text);
      cursor = piece.range.end;
    }
    expect(cursor).toBe(text.length);
  } finally { setBidiEngine(native); }
});
