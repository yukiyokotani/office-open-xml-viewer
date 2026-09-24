import { describe, expect, it } from 'vitest';
import type { FontPreloadEntry } from '@silurus/ooxml-core';
import {
  PPTX_GOOGLE_FONTS,
  PptxFontPreloadAccumulator,
  pptxFontPreloadNames,
  pptxSlideCjkFallback,
  pptxSlideOfficeFontRequests,
} from './google-fonts';
import type { Presentation, Slide } from './types';

describe('PPTX exact Office face requests', () => {
  it('does not pin a theme font when no run or inherited style names a face', () => {
    const slide = {
      elements: [{ type: 'shape', textBody: { paragraphs: [{
        defFontFamily: null, bullet: { type: 'none' },
        runs: [{ type: 'text', text: 'unstyled', fontFamily: null }],
      }] } }],
    } as unknown as Slide;
    expect(pptxSlideOfficeFontRequests(slide, 'Calibri Light', 'Calibri')).toEqual([]);
  });

  it('resolves theme, paragraph, and run style while excluding unrelated faces', () => {
    const slide = {
      elements: [{ type: 'shape', textBody: {
        defaultBold: false, defaultItalic: false,
        paragraphs: [{
          defFontFamily: '+mn-lt', defBold: false, defItalic: false,
          bullet: { type: 'none' },
          runs: [
            { type: 'text', text: 'regular', fontFamily: null, bold: null, italic: null },
            { type: 'text', text: 'bold', fontFamily: null, bold: true, italic: null },
            { type: 'text', text: 'other', fontFamily: 'Arial', bold: true, italic: true },
          ],
        }],
      } }],
    } as unknown as Slide;
    expect(pptxSlideOfficeFontRequests(slide, 'Cambria', 'Calibri')).toEqual([
      { family: 'Calibri', weight: 400, style: 'normal' },
      { family: 'Calibri', weight: 700, style: 'normal' },
    ]);
  });

  it('requests the face actually used by character and auto-number markers', () => {
    const slide = {
      elements: [{ type: 'shape', textBody: { paragraphs: [
        { defBold: true, defItalic: true, bullet: { type: 'char', char: '•', fontFamily: 'Calibri' }, runs: [] },
        { defBold: true, defItalic: true, bullet: { type: 'autoNum', numType: 'arabicPeriod', fontFamily: null },
          runs: [{ type: 'text', text: 'Item', fontFamily: 'Calibri', bold: true, italic: true }] },
      ] } }],
    } as unknown as Slide;
    expect(pptxSlideOfficeFontRequests(slide, null, 'Calibri')).toEqual(expect.arrayContaining([
      { family: 'Calibri', weight: 400, style: 'normal' },
      { family: 'Calibri', weight: 700, style: 'italic' },
    ]));
  });
});

// Verbatim snapshot of the PPTX Office-font substitute map BEFORE the shared
// registry consolidation (Phase 3 C7), excluding the SCRIPT_GOOGLE_FONTS spread
// (unchanged, shared already). Frozen as the oracle so the consolidated map's
// effective entries can only ADD keys, never drop or alter one.
const PPTX_GOOGLE_FONTS_OLD: Record<string, FontPreloadEntry> = {
  'calibri':           { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'calibri light':     { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'cambria':           { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'cambria math':      { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'nunito sans':       { url: 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'nunito':            { url: 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'open sans':         { url: 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'roboto':            { url: 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'lato':              { url: 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'montserrat':        { url: 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'poppins':           { url: 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'raleway':           { url: 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'playfair display':  { url: 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'sakkal majalla':      { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic':  { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic':   { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting':  { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Sans Arabic' },
  'noto naskh arabic':   { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic':    { url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Sans Arabic' },
};

describe('PPTX_GOOGLE_FONTS — shared registry consolidation (oracle)', () => {
  it('preserves valid pre-consolidation entries byte-for-byte', () => {
    for (const [key, entry] of Object.entries(PPTX_GOOGLE_FONTS_OLD)) {
      if (key === 'calibri light' || key === 'cambria math') continue;
      expect(PPTX_GOOGLE_FONTS[key], `entry "${key}"`).toEqual(entry);
    }
    expect(PPTX_GOOGLE_FONTS['calibri light']).toBeUndefined();
    expect(PPTX_GOOGLE_FONTS['cambria math']).toBeUndefined();
  });

  it('adds the safe, documented Ubuntu and Franklin-family substitutes', () => {
    // pptx already carried the full web-font + Office-substitute set. The shared
    // registry additionally contributes "ubuntu" (a generic Google web font, no
    // format affinity): a slide that requests Ubuntu now measures glyphs with
    // the real face instead of a narrower system sans. Purely additive.
    const oldKeys = new Set(Object.keys(PPTX_GOOGLE_FONTS_OLD));
    const added = Object.keys(PPTX_GOOGLE_FONTS).filter(
      (k) => !oldKeys.has(k) && !k.startsWith('noto '),
    );
    expect(new Set(added)).toEqual(new Set([
      'ubuntu',
      'franklin gothic book',
      'franklin gothic medium',
    ]));
    expect(PPTX_GOOGLE_FONTS['ubuntu'].url).toMatch(/family=Ubuntu(?:[:&]|$)/);
    expect(PPTX_GOOGLE_FONTS['ubuntu'].loadFamily).toBeUndefined();
    expect(PPTX_GOOGLE_FONTS['franklin gothic medium']).toMatchObject({
      loadFamily: 'Libre Franklin',
    });
  });

  it('includes slide-local paragraph and run families, not only the first theme fonts', () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'shape',
        textBody: {
          paragraphs: [{
            defFontFamily: 'Franklin Gothic Medium',
            runs: [{
              type: 'text',
              text: 'Title',
              fontFamily: null,
              fontFamilyEa: 'Yu Gothic',
              fontFamilySym: null,
            }],
          }],
        },
      }],
    } as unknown as Slide;
    const accumulator = new PptxFontPreloadAccumulator('Aptos Display', 'Aptos');
    accumulator.addSlide(slide);

    expect(accumulator.names()).toEqual(expect.arrayContaining([
      'Aptos Display',
      'Aptos',
      'Franklin Gothic Medium',
      'Yu Gothic',
    ]));
  });
});

describe('PptxFontPreloadAccumulator', () => {
  it('preserves full-presentation shape, table, and chart text semantics incrementally', () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [
        { type: 'shape', textBody: { paragraphs: [{ runs: [{ type: 'text', text: '日本語' }] }] } },
        { type: 'table', rows: [{ cells: [{ textBody: { paragraphs: [{ runs: [{ type: 'text', text: 'العربية' }] }] } }] }] },
        {
          type: 'chart',
          chart: {
            title: 'Заголовок',
            categories: ['หมวด'],
            series: [{ name: 'סדרה' }],
          },
        },
      ],
    } as unknown as Slide;
    const pres: Presentation = {
      slideWidth: 1,
      slideHeight: 1,
      slides: [slide],
      defaultTextColor: null,
      majorFont: 'Yu Gothic',
      minorFont: 'Aptos',
    };
    const incremental = new PptxFontPreloadAccumulator(pres.majorFont, pres.minorFont);
    incremental.addSlide(slide);
    expect(incremental.names()).toEqual(pptxFontPreloadNames(pres));
    expect(incremental.names()).toEqual([
      'Yu Gothic', 'Aptos',
      'Noto Sans JP', 'Noto Serif JP',
      'Noto Sans', 'Noto Serif',
      'Noto Naskh Arabic', 'Noto Sans Arabic',
      'Noto Sans Thai',
      'Noto Sans Hebrew', 'Noto Serif Hebrew',
    ]);
  });
});


it('keeps the union of stable per-slide preferences during progressive preflight', () => {
  const slideWith = (text: string) => ({ elements: [{ type: 'shape', textBody: {
    paragraphs: [{ runs: [{ type: 'text', text }] }],
  } }] } as Slide);
  const han = slideWith('漢字');
  const kana = slideWith('漢字かな');
  const korean = slideWith('漢字한글');
  const fonts = new PptxFontPreloadAccumulator('Calibri', 'Calibri', undefined, undefined, 'sc');
  fonts.addSlide(han);
  expect(fonts.names()).toContain('Noto Sans SC');
  const next = fonts.withSlide(kana);
  expect(next.names()).toContain('Noto Sans SC');
  expect(next.names()).toContain('Noto Sans JP');
  expect(fonts.names()).not.toContain('Noto Sans JP');
  expect(pptxSlideCjkFallback(han, 'Calibri', 'Calibri', 'sc')).toBe('sc');
  expect(pptxSlideCjkFallback(kana, 'Calibri', 'Calibri', 'sc')).toBe('jp');
  expect(pptxSlideCjkFallback(korean, 'Calibri', 'Calibri', 'sc')).toBe('kr');
});
