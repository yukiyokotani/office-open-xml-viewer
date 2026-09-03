import { describe, expect, it } from 'vitest';
import {
  buildSegments,
  layoutLines,
  normalizeFontFamilyUncached,
  type LineLayoutEnvironment,
} from '../line-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { DocRun, DocxDocumentModel } from '../types.js';
import type { InternalDocxDocumentModel, InternalFieldRun } from '../parser-model.js';
import type { TextLayoutService } from './text.js';
import { mathResourceKey } from './resources.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { privateResourceLookupOf } from './runtime-state.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { canvasFontString } from '@silurus/ooxml-core';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 8,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function model(overrides: Partial<DocxDocumentModel> = {}): DocxDocumentModel {
  return {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    ...overrides,
  };
}

function textRun(text: string, extra: Record<string, unknown> = {}): DocRun {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Authored Sans',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...extra,
  } as DocRun;
}

describe('production layout service integration', () => {
  it('retains the document kernel when production text services are composed by object spread', () => {
    const document = model();
    const base = createLayoutServices(document, { measureContext: measureContext() });
    const text = Object.freeze({ ...base.text });
    const composed = Object.freeze({ ...base, text });

    const layout = layoutDocument(document, composed, { currentDateMs: 10 });

    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]?.geometry).toMatchObject({ widthPt: 612, heightPt: 792 });
  });

  it('routes every normal run segmentation and measurement through the injected text service', () => {
    const base = createLayoutServices(model(), { measureContext: measureContext() });
    let calls = 0;
    const countingText: TextLayoutService = Object.freeze({
      ...base.text,
      shape(request: Parameters<TextLayoutService['shape']>[0]) {
        calls += 1;
        return base.text.shape(request);
      },
    });
    const services = Object.freeze({ ...base, text: countingText });
    const environment: LineLayoutEnvironment = { pageIndex: 0, totalPages: 1, layoutServices: services };
    const segments = buildSegments([textRun('first'), textRun('second')], environment);
    const afterSegmentation = calls;
    const lines = layoutLines(measureContext(), segments, 300, 0, 1);

    expect(afterSegmentation).toBeGreaterThanOrEqual(2);
    expect(calls).toBeGreaterThan(afterSegmentation);
    expect(lines).toHaveLength(1);
  });

  it('keeps cross-slot scalar spans in one unbreakable grapheme', () => {
    const ctx = measureContext();
    const services = createLayoutServices(model(), { measureContext: ctx });
    const segments = buildSegments([textRun('a\u0301', {
      fontFamily: 'ASCII Face',
      fontFamilyHighAnsi: 'HANSI Face',
    })], { pageIndex: 0, totalPages: 1, layoutServices: services });
    const text = segments.filter((segment) => 'text' in segment);

    expect(text.map((segment) => 'text' in segment
      ? [segment.text, segment.fontFamily, segment.joinPrev ?? false]
      : null)).toEqual([
        ['a', 'ASCII Face', false],
        ['\u0301', 'HANSI Face', true],
      ]);
    const lines = layoutLines(ctx, segments, 8, 0, 1);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments.map((segment) => 'text' in segment ? segment.text : '')).toEqual(['a', '\u0301']);
  });

  it('normalizes service-backed w:sym private encoding before measurement and paint', () => {
    const services = createLayoutServices(model(), { measureContext: measureContext() });
    const segments = buildSegments([textRun('\uF0B0', {
      fontFamily: 'Symbol',
      fontFamilyHighAnsi: 'Symbol',
      fontFamilyEastAsia: 'Symbol',
    })], { pageIndex: 0, totalPages: 1, layoutServices: services });
    const text = segments.filter((segment) => 'text' in segment);

    expect(text).toHaveLength(1);
    expect(text[0]).toMatchObject({ text: '°' });
    expect((text[0] as { fontFamily?: string | null }).fontFamily).not.toBe('Symbol');
  });

  it('carries the w:kern threshold through the text service measure adapter', () => {
    let fontKerning: CanvasFontKerning = 'auto';
    const states: CanvasFontKerning[] = [];
    const ctx = {
      font: '',
      letterSpacing: '0px',
      get fontKerning() { return fontKerning; },
      set fontKerning(value: CanvasFontKerning) { fontKerning = value; },
      measureText(text: string) {
        states.push(fontKerning);
        const width = fontKerning === 'normal' ? 40 : fontKerning === 'none' ? 30 : 20;
        return {
          width: text ? width : 0,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const services = createLayoutServices(model(), { measureContext: ctx });
    const environment: LineLayoutEnvironment = { pageIndex: 0, totalPages: 1, layoutServices: services };
    const measure = (kerning: number) => {
      const segments = buildSegments([textRun('AV', { fontSize: 10, kerning })], environment);
      return layoutLines(ctx, segments, 300, 0, 1)[0].segments[0].measuredWidth;
    };

    expect(measure(5)).toBe(40);
    expect(measure(20)).toBe(30);
    expect(states).toContain('normal');
    expect(states).toContain('none');
    expect(ctx.fontKerning).toBe('auto');
  });

  it('projects finite Canvas ink metrics for retained trim geometry', () => {
    const ctx = {
      ...measureContext(),
      measureText: () => ({
        width: 10,
        actualBoundingBoxLeft: 2,
        actualBoundingBoxRight: 11,
        actualBoundingBoxAscent: 7,
        actualBoundingBoxDescent: 3,
        fontBoundingBoxAscent: 8,
        fontBoundingBoxDescent: 4,
      } as TextMetrics),
    } as unknown as CanvasRenderingContext2D;
    const services = createLayoutServices(model(), { measureContext: ctx });

    const shaped = services.text.shape({
      text: 'X', fontSizePt: 10, weight: 400, style: 'normal', measure: true,
      fonts: { ascii: 'Authored Sans' },
    });

    expect(shaped).toMatchObject({
      advancePt: 10,
      ascentPt: 8,
      descentPt: 4,
      inkBounds: { xMinPt: -2, xMaxPt: 11, ascentPt: 7, descentPt: 3 },
    });
  });

  it('uses the byte-identical Canvas route in the service measurer and line probes', () => {
    const fonts: string[] = [];
    let font = '';
    const ctx = {
      ...measureContext(),
      get font() { return font; },
      set font(value: string) { font = value; fonts.push(value); },
    } as CanvasRenderingContext2D;
    const services = createLayoutServices(model({
      fontFamilyClasses: { 'Roman Face': 'roman' },
    }), { measureContext: ctx });
    const shaped = services.text.shape({
      text: 'AV', fontSizePt: 10, fonts: { ascii: 'Roman Face' },
    });
    const span = shaped.spans[0] as NonNullable<typeof shaped.spans[0]>;
    const expected = canvasFontString(span.fontRoute, 10, 400, 'normal');
    expect(fonts).toContain(expected);

    fonts.length = 0;
    const segments = buildSegments([textRun('AVA', { fontFamily: 'Roman Face' })], {
      pageIndex: 0, totalPages: 1, layoutServices: services,
    });
    layoutLines(ctx, segments, 300, 0, 1);
    expect(fonts.filter(Boolean)).toEqual(expect.arrayContaining([expected]));
    expect(segments[0]).toMatchObject({ fontRoute: span.fontRoute });
  });

  it('keeps hint-protected eastAsia text on non-cs formatting inside an rtl run', () => {
    const services = createLayoutServices(model(), { measureContext: measureContext() });
    const run = textRun('A国', {
      fontFamily: 'Latin Face',
      fontFamilyEastAsia: 'EA Face',
      fontFamilyCs: 'CS Face',
      fontHint: 'eastAsia',
      langEastAsia: 'zh-cn',
      rtl: true,
      cs: true,
      fontSize: 10,
      fontSizeCs: 20,
      bold: false,
      boldCs: true,
    });
    const segments = buildSegments([run], { pageIndex: 0, totalPages: 1, layoutServices: services });
    const textSegments = segments.filter((segment) => 'text' in segment && segment.text.length > 0);

    expect(textSegments.map((segment) => 'text' in segment ? {
      text: segment.text,
      fontFamily: segment.fontFamily,
      fontSize: segment.fontSize,
      bold: segment.bold,
    } : null)).toEqual([
      { text: 'A', fontFamily: 'CS Face', fontSize: 20, bold: true },
      { text: '国', fontFamily: 'EA Face', fontSize: 10, bold: false },
    ]);
    expect(textSegments.map((segment) => 'text' in segment
      ? segment.textShapeRequest?.fonts.complexScript === 'CS Face' && segment.textShapeRequest?.complexScript
      : false)).toEqual([true, false]);
  });

  it('classifies Latin runs with a resolved eastAsia axis only when useFELayout is active', () => {
    const hinted = textRun('Hinted Latin title', {
      fontFamilyEastAsia: 'EA Face',
      fontHint: 'eastAsia',
      langEastAsia: 'ja-jp',
    });
    const unhinted = textRun('Unhinted Latin title', {
      fontFamilyEastAsia: 'EA Face',
      fontHint: null,
      langEastAsia: 'ja-jp',
    });
    const off = buildSegments([hinted, unhinted], { pageIndex: 0, totalPages: 1 });
    const on = buildSegments(
      [hinted, unhinted],
      { pageIndex: 0, totalPages: 1, useFeLayout: true },
    );

    expect(off.filter((segment) => 'text' in segment)
      .every((segment) => !('metricEastAsian' in segment))).toBe(true);
    expect(on.filter((segment) => 'text' in segment)
      .every((segment) => 'metricEastAsian' in segment && segment.metricEastAsian === true)).toBe(true);
  });

  it('keeps a single pure-CJK hint-protected rtl span on non-cs formatting', () => {
    const services = createLayoutServices(model(), { measureContext: measureContext() });
    const [segment] = buildSegments([textRun('国', {
      fontFamily: 'Latin Face', fontFamilyEastAsia: 'EA Face', fontFamilyCs: 'CS Face',
      fontHint: 'eastAsia', langEastAsia: 'zh-cn', rtl: true, cs: true,
      fontSize: 10, fontSizeCs: 20, bold: false, boldCs: true,
    })], { pageIndex: 0, totalPages: 1, layoutServices: services });

    expect(segment).toMatchObject({ text: '国', fontSize: 10, bold: false, fontFamily: 'EA Face' });
    expect('text' in segment && segment.textShapeRequest?.complexScript).toBe(false);
  });

  it('classifies a CJK-hinted field result on its inherited East Asian axis', () => {
    const services = createLayoutServices(model(), { measureContext: measureContext() });
    const field: InternalFieldRun & { type: 'field' } = {
      type: 'field', fieldType: 'other', instruction: 'REF cjk', fallbackText: '国',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Latin Face', background: null,
      vertAlign: null, fontFamilyEastAsia: 'EA Face', fontHint: 'eastAsia',
      langEastAsia: 'zh-cn', rtl: true, cs: true, fontFamilyCs: 'CS Face',
      fontSizeCs: 20, boldCs: true, italicCs: true,
    };
    const [segment] = buildSegments([field as DocRun], {
      pageIndex: 0, totalPages: 1, layoutServices: services,
    });

    expect(segment).toMatchObject({ text: '国', fontSize: 10, bold: false, italic: false });
    expect('text' in segment && segment.textShapeRequest).toMatchObject({
      fontHint: 'eastAsia', eastAsiaLanguage: 'zh-cn', complexScript: false,
    });
  });

  it('applies inherited complex-script font, size, and style to an RTL field result', () => {
    const services = createLayoutServices(model(), { measureContext: measureContext() });
    const field: InternalFieldRun & { type: 'field' } = {
      type: 'field', fieldType: 'other', instruction: 'REF rtl', fallbackText: 'A',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Latin Face', background: null,
      vertAlign: null, rtl: true, cs: true, fontFamilyCs: 'CS Face',
      fontSizeCs: 20, boldCs: true, italicCs: true, langBidi: 'ar-sa',
    };
    const [segment] = buildSegments([field as DocRun], {
      pageIndex: 0, totalPages: 1, layoutServices: services,
    });

    expect(segment).toMatchObject({ text: 'A', fontSize: 20, bold: true, italic: true, rtl: true });
    expect('text' in segment && segment.textShapeRequest).toMatchObject({
      complexScript: true,
      fonts: { complexScript: 'CS Face' },
    });
  });

  it('plumbs the selected eastAsia font charset from fontTable into slot selection', () => {
    const doc = model() as InternalDocxDocumentModel;
    doc.fontFamilyCharsets = { 'EA Face': '86' };
    const services = createLayoutServices(doc, { measureContext: measureContext() });
    const shaped = services.text.shape({
      text: '\u0100',
      fontSizePt: 10,
      fontHint: 'eastAsia',
      eastAsiaLanguage: 'ja-jp',
      fonts: { ascii: 'Latin Face', highAnsi: 'Latin Face', eastAsia: 'EA Face' },
    });

    expect(shaped.spans[0]?.script).toBe('eastAsia');
  });

  it('makes local availability, geometry, diagnostics, and fingerprints truthful and worker-stable', () => {
    const fonts: string[] = [];
    const ctx = {
      ...measureContext(),
      get font() { return fonts.at(-1) ?? ''; },
      set font(value: string) { fonts.push(value); },
      measureText(text: string) {
        const width = (fonts.at(-1) ?? '').includes('__local_times_bi') ? 17 : 5;
        return {
          width: text.length * width,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as CanvasRenderingContext2D;
    const localMetrics = {
      'times new roman:700:italic': {
        family: '__local_times_bi', lineHeightRatio: 1.2,
        requestedFamily: 'Times New Roman', weight: 700, style: 'italic' as const,
      },
    };
    const shape = (services: ReturnType<typeof createLayoutServices>) => services.text.shape({
      text: 'AV', fontSizePt: 10, weight: 700, style: 'italic',
      fonts: { ascii: 'Times New Roman' },
    });
    const present = createLayoutServices(model(), { measureContext: ctx, localMetrics });
    const worker = createLayoutServices(model(), { measureContext: ctx, localMetrics });
    const absent = createLayoutServices(model(), { measureContext: ctx, localMetrics: {} });

    expect(shape(present).spans[0]?.font).toMatchObject({ source: 'local', resolvedFamily: '__local_times_bi' });
    expect(shape(present).advancePt).toBe(34);
    expect(shape(absent).spans[0]?.font).toMatchObject({
      source: 'native', resolvedFamily: 'Times New Roman',
      route: { familyList: '"Times New Roman", sans-serif', scope: 'native' },
    });
    expect(shape(absent).advancePt).toBe(10);
    expect(present.text.fingerprint).toBe(worker.text.fingerprint);
    expect(present.text.fingerprint).not.toBe(absent.text.fingerprint);
  });

  it('takes one deeply immutable local-metric snapshot at the document boundary', () => {
    const callerMetric = {
      family: '__local_authored',
      lineHeightRatio: 1.25,
      requestedFamily: 'Authored Sans',
      weight: 400,
      style: 'normal' as const,
      sourceIdentity: 'local:Authored Sans',
      synthesized: false,
    };
    const caller: Record<string, typeof callerMetric> = { 'authored sans:400:normal': callerMetric };
    const services = createLayoutServices(model(), {
      measureContext: measureContext(),
      localMetrics: caller,
    });
    const before = services.text.fingerprint;

    callerMetric.family = '__mutated';
    callerMetric.lineHeightRatio = 99;
    caller['late face:400:normal'] = { ...callerMetric };

    expect(services.text.localMetrics).toEqual({
      'authored sans:400:normal': {
        family: '__local_authored',
        lineHeightRatio: 1.25,
        requestedFamily: 'Authored Sans',
        weight: 400,
        style: 'normal',
        sourceIdentity: 'local:Authored Sans',
        synthesized: false,
      },
    });
    expect(Object.isFrozen(services.text.localMetrics)).toBe(true);
    expect(Object.isFrozen(services.text.localMetrics['authored sans:400:normal'])).toBe(true);
    expect(services.text.fingerprint).toBe(before);
  });

  it('derives generic fallback from fontTable family and pitch for each selected face', () => {
    const doc = model({
      body: [{
        type: 'paragraph',
        runs: [
          textRun('serif', { fontFamily: 'Garamond' }),
          textRun('sans', { fontFamily: 'Arial' }),
        ],
      } as DocxDocumentModel['body'][number]],
      fontFamilyClasses: {
        'Roman Face': 'roman',
        'Swiss Face': 'swiss',
        'Fixed Modern': 'modern',
        'Variable Modern': 'modern',
      },
      fontFamilyPitches: {
        'Fixed Modern': 'fixed',
        'Variable Modern': 'variable',
      },
    });
    const services = createLayoutServices(doc, { measureContext: measureContext() });
    const generic = (family: string) => services.text.shape({
      text: 'x', fontSizePt: 10, fonts: { ascii: family },
    }).spans[0]?.font.genericFamily;

    expect(generic('Roman Face')).toBe('serif');
    expect(generic('Swiss Face')).toBe('sans-serif');
    expect(generic('Fixed Modern')).toBe('monospace');
    expect(generic('Variable Modern')).toBe('sans-serif');
    expect(generic('Garamond')).toBe('serif');
    expect(generic('Arial')).toBe('sans-serif');
    expect(services.text.shape({
      text: 'x', fontSizePt: 10, fonts: { ascii: 'Roman Face' },
    }).spans[0]?.font.route).toMatchObject({
      familyList: expect.stringMatching(/^"Roman Face", .*"Noto Serif".*serif$/),
      scope: 'native',
    });
  });

  it('uses bounded consumer defaults only when no rFonts slot resolves', () => {
    const services = createLayoutServices(model(), { measureContext: measureContext() });

    expect(services.text.shape({
      text: 'Header 1', fontSizePt: 10, fonts: {},
    }).spans[0]?.font).toMatchObject({
      source: 'generic', requestedFamily: 'serif', genericFamily: 'serif',
      route: { familyList: 'serif', scope: 'generic' },
    });
    expect(services.text.shape({
      text: 'Header 1', fontSizePt: 10, fonts: { ascii: 'Arial' },
    }).spans[0]?.font).toMatchObject({
      source: 'native', requestedFamily: 'Arial', genericFamily: 'sans-serif',
    });
    expect(services.text.shape({
      text: '漢字', fontSizePt: 10, fonts: {},
    }).spans[0]?.font).toMatchObject({
      source: 'generic', requestedFamily: 'sans-serif', genericFamily: 'sans-serif',
      route: { familyList: 'sans-serif', scope: 'generic' },
    });
  });

  it('retains the full fallback route and fingerprint for a run-only face absent from fontTable', () => {
    const family = 'Times New Roman';
    const document = model({
      body: [{ type: 'paragraph', runs: [textRun('run-only', { fontFamily: family })] }],
    } as Partial<DocxDocumentModel>);
    const alternate = model({
      body: [{ type: 'paragraph', runs: [textRun('run-only', { fontFamily: 'Arial' })] }],
    } as Partial<DocxDocumentModel>);
    const services = createLayoutServices(document, { measureContext: measureContext() });
    const route = services.text.shape({
      text: 'x', fontSizePt: 10, fonts: { ascii: family },
    }).spans[0]?.font.route;

    expect(route).toMatchObject({
      familyList: normalizeFontFamilyUncached(family, {}, {}),
      scope: 'native',
    });
    expect(services.text.fingerprint).not.toBe(
      createLayoutServices(alternate, { measureContext: measureContext() }).text.fingerprint,
    );
  });

  it('retains the full fallback route and fingerprint for a theme-only face', () => {
    const family = 'Cambria';
    const services = createLayoutServices(model({ majorFont: family }), {
      measureContext: measureContext(),
    });
    const route = services.text.shape({
      text: 'x', fontSizePt: 10, fonts: {}, themeFonts: { ascii: family },
    }).spans[0]?.font.route;

    expect(route).toMatchObject({
      familyList: normalizeFontFamilyUncached(family, {}, {}),
      scope: 'native',
    });
    expect(services.text.fingerprint).not.toBe(
      createLayoutServices(model({ majorFont: 'Arial' }), {
        measureContext: measureContext(),
      }).text.fingerprint,
    );
  });

  it('inventories only successfully registered faces and labels Office replacements as substitutions', () => {
    const doc = model({
      majorFont: 'Calibri',
      embeddedFonts: [{ fontName: 'Broken Embedded', partPath: 'word/fonts/missing.odttf', fontKey: '', style: 'regular' }],
    });
    const failed = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [],
      googleFaces: [],
      useGoogleFonts: true,
    });
    const missingEmbedded = failed.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Broken Embedded' } });
    const missingGoogle = failed.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Calibri' } });
    expect(missingEmbedded.spans[0]?.font.source).toBe('native');
    expect(missingEmbedded.diagnostics).toEqual([]);
    expect(missingGoogle.spans[0]?.font.source).toBe('native');

    const carlito = { family: 'Carlito', weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    const loaded = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [],
      googleFaces: [carlito],
      useGoogleFonts: true,
    });
    const substituted = loaded.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Calibri' } });
    expect(substituted.spans[0]?.font).toMatchObject({ source: 'substitute', resolvedFamily: 'Carlito' });
    expect(substituted.diagnostics[0]?.message).toMatch(/implementation-dependent/i);
  });

  it('requires loaded status and an exact family/weight/style match for every face', () => {
    const doc = model({
      embeddedFonts: [
        { fontName: 'Partial Embedded', partPath: 'word/fonts/regular.odttf', fontKey: '', style: 'regular' },
        { fontName: 'Partial Embedded', partPath: 'word/fonts/bold.odttf', fontKey: '', style: 'bold' },
      ],
    });
    const services = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [
        { family: '"Partial Embedded"', weight: '400', style: 'normal', status: 'loaded' },
        { family: 'Partial Embedded', weight: '700', style: 'normal', status: 'error' },
        { family: 'Timed Out', weight: '400', style: 'normal', status: 'loading' },
      ] as FontFace[],
    });
    const shape = (family: string, weight: number, style: 'normal' | 'italic' = 'normal') =>
      services.text.shape({ text: 'x', fontSizePt: 10, weight, style, fonts: { ascii: family } });

    expect(shape('Partial Embedded', 400).spans[0]?.font)
      .toMatchObject({ source: 'embedded', resolvedFamily: 'Partial Embedded' });
    expect(shape('Partial Embedded', 700).spans[0]?.font.source).toBe('native');
    expect(shape('Partial Embedded', 400, 'italic').spans[0]?.font.source).toBe('native');
    expect(shape('Timed Out', 400).spans[0]?.font.source).toBe('native');
  });

  it('collects every currently representable math story, including rich text boxes and nested tables', () => {
    const math = (value: string) => ({
      type: 'math', nodes: [{ type: 'text', text: value }], display: false, fontSize: 10,
    });
    const paragraph = (value: string) => ({ type: 'paragraph', runs: [math(value)] });
    const table = (value: string) => ({
      type: 'table', rows: [{ cells: [{ content: [paragraph(value)] }] }],
    });
    const shapeWithRichTextBox = {
      type: 'shape',
      textBoxContent: [paragraph('shape-inline'), table('shape-table')],
    };
    const doc = model({
      body: [
        paragraph('body'),
        { type: 'paragraph', runs: [shapeWithRichTextBox] },
      ],
      headers: { default: { body: [table('header')] }, first: null, even: null },
      footers: { default: null, first: { body: [paragraph('footer')] }, even: null },
      footnotes: [{ id: '1', content: [table('footnote')] }],
      endnotes: [{ id: '2', content: [paragraph('endnote')] }],
    } as unknown as Partial<DocxDocumentModel>);
    const normalized = normalizeInternalDocumentModel(doc);
    const services = createLayoutServices(normalized.document, { measureContext: measureContext() });

    expect(normalized.mathOccurrences).toHaveLength(7);
    expect(normalized.mathOccurrences.map((occurrence) => occurrence.source)).toEqual(
      expect.arrayContaining([
        { story: 'textbox', storyInstance: 'body:body:1.0', path: [0, 0] },
        { story: 'textbox', storyInstance: 'body:body:1.0', path: [1, 0, 0, 0, 0] },
      ]),
    );
    for (const occurrence of normalized.mathOccurrences) {
      expect(occurrence.resourceKey).toBe(mathResourceKey(occurrence.source, 'inline'));
      expect(() => services.math.resolve(occurrence.resourceKey)).not.toThrow();
    }
  });

  it('keeps identical and private math contents out of public resource identity', () => {
    const first = { type: 'math', nodes: [{ type: 'text', text: 'PRIVATE-SENTINEL' }], display: false, fontSize: 10 };
    const second = { ...first, nodes: [{ type: 'text', text: 'PRIVATE-SENTINEL' }] };
    const doc = model({ body: [
      { type: 'paragraph', runs: [first] },
      { type: 'paragraph', runs: [second] },
    ] } as unknown as Partial<DocxDocumentModel>);
    const normalized = normalizeInternalDocumentModel(doc);
    const services = createLayoutServices(normalized.document);
    const keys = normalized.mathOccurrences.map((occurrence) => occurrence.resourceKey);

    expect(new Set(keys).size).toBe(2);
    expect(keys.join(' ')).not.toContain('PRIVATE-SENTINEL');
    expect(services.math.fingerprint).not.toContain('PRIVATE-SENTINEL');
  });

  it('resolves repeated math ASTs by structural keys after cloning without mutating input', () => {
    const mathRun = () => ({
      type: 'math', nodes: [{ type: 'text', text: 'same' }], display: false, fontSize: 10,
    });
    const paragraphs = [
      { type: 'paragraph', runs: [mathRun()] },
      { type: 'paragraph', runs: [mathRun()] },
    ] as unknown as Array<{ type: 'paragraph'; runs: DocRun[] }>;
    const doc = model({ body: paragraphs } as unknown as Partial<DocxDocumentModel>);
    const before = structuredClone(doc);
    const normalized = normalizeInternalDocumentModel(doc);
    const services = createLayoutServices(normalized.document);
    const normalizedParagraphs = normalized.document.body as unknown as typeof paragraphs;
    const environment = { pageIndex: 0, totalPages: 1, layoutServices: services };
    const first = buildSegments([...normalizedParagraphs[0].runs], environment)[0];
    const second = buildSegments([...normalizedParagraphs[1].runs], environment)[0];

    expect(doc).toEqual(before);
    expect(normalizedParagraphs[0]).not.toBe(paragraphs[0]);
    expect('mathResourceKey' in first && 'mathResourceKey' in second
      ? first.mathResourceKey === second.mathResourceKey
      : true).toBe(false);
  });

  it('requires runtime math handles to match available metadata exactly', () => {
    const doc = model({ body: [
      { type: 'paragraph', runs: [{ type: 'math', nodes: [], display: false, fontSize: 10 }] },
      { type: 'paragraph', runs: [{ type: 'math', nodes: [], display: true, fontSize: 10 }] },
    ] } as unknown as Partial<DocxDocumentModel>);
    const [availableOccurrence, unavailableOccurrence] = layoutSourceStore(doc).mathOccurrences;
    const available = { resourceKey: mathResourceKey(availableOccurrence.source, 'inline'), widthEm: 1, ascentEm: 0.8, descentEm: 0.2, diagnostics: [] };
    const unavailable = { resourceKey: mathResourceKey(unavailableOccurrence.source, 'display'), widthEm: 0, ascentEm: 0, descentEm: 0, available: false as const, diagnostics: [] };
    const drawable = {} as CanvasImageSource;

    expect(() => createLayoutServices(doc, { mathResources: [available], mathDrawables: new Map() }))
      .toThrow(/math metadata.*missing/i);
    expect(() => createLayoutServices(doc, {
      mathResources: [unavailable],
      mathDrawables: new Map([[unavailable.resourceKey, drawable]]),
    })).toThrow(/math metadata.*missing|extra/i);

    const services = createLayoutServices(doc, {
      mathResources: [available, unavailable],
      mathDrawables: new Map([[available.resourceKey, drawable]]),
    });
    expect(privateResourceLookupOf(services)?.keys).toEqual([available.resourceKey]);
  });

  it('gives main and worker factories identical fingerprints for identical successful snapshots', () => {
    const embedded = { family: 'Embedded', weight: '700', style: 'italic', status: 'loaded' } as FontFace;
    const options = {
      measureContext: measureContext(),
      embeddedFaces: [embedded],
      googleFaces: [] as FontFace[],
      localMetrics: { authored: { family: '__local_authored', lineHeightRatio: 1.25 } },
    };
    const main = createLayoutServices(model(), options);
    const worker = createLayoutServices(model(), { ...options, measureContext: measureContext() });

    expect(main.text.fingerprint).toBe(worker.text.fingerprint);
    expect(main.images.fingerprint).toBe(worker.images.fingerprint);
    expect(main.math.fingerprint).toBe(worker.math.fingerprint);
  });
});
