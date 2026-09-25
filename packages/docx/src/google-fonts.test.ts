import { describe, expect, it } from 'vitest';
import type { FontPreloadEntry } from '@silurus/ooxml-core';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames, docxOfficeFontFallbackRequests } from './google-fonts.js';
import type { DocxDocumentModel } from './types.js';

// Verbatim snapshot of the DOCX Office-font substitute map BEFORE the shared
// registry consolidation (PR #-, Phase 3 C7), excluding the SCRIPT_GOOGLE_FONTS
// spread (unchanged, shared already). Frozen as the oracle so the consolidated
// map's effective entries can only ADD keys, never drop or alter one.
const DOCX_GOOGLE_FONTS_OLD: Record<string, FontPreloadEntry> = {
  'calibri':           { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'cambria':           { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'nunito sans':       { url: 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'nunito':            { url: 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'open sans':         { url: 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'roboto':            { url: 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'lato':              { url: 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'montserrat':        { url: 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'poppins':           { url: 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'raleway':           { url: 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'playfair display':  { url: 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'ubuntu':            { url: 'https://fonts.googleapis.com/css2?family=Ubuntu:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'sakkal majalla':      { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic':  { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic':   { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting':  { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Sans Arabic' },
  'noto naskh arabic':   { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic':    { url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Sans Arabic' },
};

/** Build a minimal model whose body is a single paragraph with one text run. */
function docWith(text: string, major = 'Calibri', minor = 'Calibri'): DocxDocumentModel {
  return {
    section: {} as DocxDocumentModel['section'],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    majorFont: major,
    minorFont: minor,
    body: [
      {
        type: 'paragraph',
        runs: [{ type: 'text', text } as never],
      } as never,
    ],
  } as DocxDocumentModel;
}

describe('docxFontPreloadNames — script-aware preload', () => {
  it('pure-Latin doc preloads ONLY the theme fonts (no CJK / script faces)', () => {
    const names = docxFontPreloadNames(docWith('Hello, world.'));
    expect(names).toEqual(['Calibri', 'Calibri']);
    // The expensive CJK faces must NOT be queued for a Latin document.
    expect(names).not.toContain('Noto Sans JP');
    expect(names).not.toContain('Noto Sans KR');
    expect(names).not.toContain('Noto Naskh Arabic');
  });

  it('Japanese doc preloads the JP Noto faces', () => {
    const names = docxFontPreloadNames(docWith('こんにちは世界'));
    expect(names).toContain('Noto Sans JP');
    expect(names).toContain('Noto Serif JP');
    expect(names).not.toContain('Noto Sans KR');
  });

  it('Han with a Korean theme font uses the kr lang hint', () => {
    const names = docxFontPreloadNames(docWith('漢字', 'Malgun Gothic', 'Malgun Gothic'));
    expect(names).toContain('Noto Sans KR');
    expect(names).not.toContain('Noto Sans JP');
  });

  it('is deterministic — same model yields the same set (main == worker)', () => {
    const doc = docWith('日本語 العربية');
    expect(docxFontPreloadNames(doc)).toEqual(docxFontPreloadNames(doc));
  });
});

it('collects the distinct authored font tuples used by rendered text', () => {
  const doc = docWith('body', 'Aptos', 'Aptos');
  (doc.body[0] as { runs: object[] }).runs = [
    { type: 'text', text: 'regular', fontFamily: 'Calibri', bold: false, italic: false },
    { type: 'text', text: 'emphasis', fontFamily: 'calibri', bold: true, italic: true },
    { type: 'text', text: 'light', fontFamily: 'Calibri Light', bold: false, italic: false },
  ];
  const requests = docxOfficeFontFallbackRequests(doc);
  expect(requests).toContainEqual({ family: 'Calibri', weight: 400, style: 'normal' });
  expect(requests).toContainEqual({ family: 'calibri', weight: 700, style: 'italic' });
  expect(requests).toContainEqual({ family: 'Calibri Light', weight: 400, style: 'normal' });
  expect(requests.filter((request) => request.family.toLowerCase() === 'calibri')).toHaveLength(2);
});

it('loads the themed bold face when a run inherits its family', () => {
  const doc = docWith('body');
  (doc.body[0] as { runs: object[] }).runs = [
    { type: 'text', text: 'heading', bold: true },
  ];
  expect(docxOfficeFontFallbackRequests(doc)).toContainEqual({
    family: 'Calibri', weight: 700, style: 'normal',
  });
});

it('loads the inherited Latin face when only an East Asian face is authored', () => {
  const doc = docWith('Latin text', 'Aptos', 'Calibri');
  (doc.body[0] as { runs: object[] }).runs = [
    { type: 'text', text: 'Latin text', fontFamilyEastAsia: 'Meiryo', bold: true },
  ];
  const requests = docxOfficeFontFallbackRequests(doc);
  expect(requests).toContainEqual({ family: 'Calibri', weight: 400, style: 'normal' });
  expect(requests).toContainEqual({ family: 'Calibri', weight: 700, style: 'normal' });
  expect(requests).toContainEqual({ family: 'Meiryo', weight: 700, style: 'normal' });
});

it('does not load the theme face for an explicitly authored Latin face', () => {
  const doc = docWith('Latin text', 'Aptos', 'Calibri');
  (doc.body[0] as { runs: object[] }).runs = [
    { type: 'text', text: 'Latin text', fontFamily: 'Arial', fontFamilyEastAsia: 'Meiryo', bold: true },
  ];
  const requests = docxOfficeFontFallbackRequests(doc);
  expect(requests).not.toContainEqual({ family: 'Calibri', weight: 700, style: 'normal' });
  expect(requests).toContainEqual({ family: 'Arial', weight: 700, style: 'normal' });
  expect(requests).toContainEqual({ family: 'Meiryo', weight: 700, style: 'normal' });
});

describe('DOCX_GOOGLE_FONTS — theme typeface coverage', () => {
  // Templates whose theme minorFont is Ubuntu (e.g. sample-11) emit runs with
  // family "Ubuntu". Without an explicit mapping the preloader skips the name
  // and the renderer falls back to a system sans whose metrics are narrower
  // than Ubuntu's — table cells sized against the Ubuntu width then fail to
  // wrap where Word would. The map entry resolves Ubuntu against Google Fonts
  // so the canvas measures glyphs with the actual face.
  it('resolves Ubuntu to a Google Fonts Ubuntu stylesheet', () => {
    const entry = DOCX_GOOGLE_FONTS['ubuntu'];
    expect(entry).toBeDefined();
    expect(entry.url).toMatch(/^https:\/\/fonts\.googleapis\.com\/css2\?/);
    expect(entry.url).toMatch(/family=Ubuntu(?:[:&]|$)/);
    // No loadFamily override — Google Fonts ships the same family name, so
    // the renderer's canvas font stack can use "Ubuntu" directly.
    expect(entry.loadFamily).toBeUndefined();
  });

  it('includes Ubuntu in the preload list when the theme minorFont is Ubuntu', () => {
    const names = docxFontPreloadNames(docWith('City or Town', 'Calibri', 'Ubuntu'));
    expect(names).toContain('Ubuntu');
  });
});

describe('DOCX_GOOGLE_FONTS — shared registry consolidation (oracle)', () => {
  it('preserves every pre-consolidation entry byte-for-byte', () => {
    for (const [key, entry] of Object.entries(DOCX_GOOGLE_FONTS_OLD)) {
      expect(DOCX_GOOGLE_FONTS[key], `entry "${key}"`).toEqual(entry);
    }
  });

  it('adds only supported text-face alternatives', () => {
    // Consolidating everything into the shared GOOGLE_FONT_SUBSTITUTES pulls the
    // The Franklin Gothic faces share Libre Franklin. Calibri Light needs a
    // distinct light face, and Cambria Math needs mathematical font support.
    const oldScriptKeys = new Set(Object.keys(DOCX_GOOGLE_FONTS_OLD));
    const added = Object.keys(DOCX_GOOGLE_FONTS).filter(
      (k) => !oldScriptKeys.has(k) && !k.startsWith('noto '),
    );
    expect(new Set(added)).toEqual(new Set([
      'franklin gothic book',
      'franklin gothic medium',
    ]));
    expect(DOCX_GOOGLE_FONTS['calibri light']).toBeUndefined();
    expect(DOCX_GOOGLE_FONTS['cambria math']).toBeUndefined();
    expect(DOCX_GOOGLE_FONTS['franklin gothic medium']).toMatchObject({
      loadFamily: 'Libre Franklin',
    });
  });
});


it('preloads the configured fallback only for ambiguous Han', () => {
  expect(docxFontPreloadNames(docWith('漢字'), 'sc')).toContain('Noto Sans SC');
  expect(docxFontPreloadNames(docWith('漢字'), 'sc')).not.toContain('Noto Sans JP');
  expect(docxFontPreloadNames(docWith('Hello'), 'sc')).toEqual(['Calibri', 'Calibri']);
  expect(docxFontPreloadNames(docWith('漢字', 'Meiryo'), 'sc')).toContain('Noto Sans JP');
  expect(docxFontPreloadNames(docWith('漢字かな'), 'sc')).toContain('Noto Sans JP');
});


it('preloads an explicitly authored East Asian run language ahead of host preference', () => {
  const doc = docWith('漢字');
  Object.assign((doc.body[0] as { runs: object[] }).runs[0], { langEastAsia: 'ja-JP' });
  const names = docxFontPreloadNames(doc, 'sc');
  expect(names).toContain('Noto Sans JP');
  expect(names).not.toContain('Noto Sans SC');
});


it('preloads the explicit Chinese region even when the same run contains kana', () => {
  const doc = docWith('漢字かな');
  Object.assign((doc.body[0] as { runs: object[] }).runs[0], { langEastAsia: 'zh-CN' });
  expect(docxFontPreloadNames(doc, 'tc')).toContain('Noto Sans SC');
  expect(docxFontPreloadNames(doc, 'tc')).toContain('Noto Sans JP');
});
