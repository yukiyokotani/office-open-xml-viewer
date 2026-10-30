import { describe, expect, it } from 'vitest';
import { googleFonts } from '@silurus/ooxml-core/internal/font-sources';
import {
  completeDocxFontPresence,
  declaredFontSubstitutions,
  docxFontDemandCandidates,
  docxFontFamilyPlan,
  docxOfficeFontFallbackRequests,
  type DocxOfficeFontPresence,
} from './font-plan.js';
import type { DocxDocumentModel } from './types.js';

const GOOGLE = [googleFonts()];
const SUBSTITUTES = declaredFontSubstitutions(GOOGLE);

/** Demanded names before the installed check: authored, then fallback. */
function docxFontPreloadNames(doc: DocxDocumentModel, fallback?: Parameters<typeof docxFontDemandCandidates>[2]): string[] {
  const candidates = docxFontDemandCandidates(doc, SUBSTITUTES, fallback);
  return [...candidates.authored, ...candidates.fallback];
}

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
    expect(names).toEqual(['Calibri']);
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
  expect(requests.filter((request) => request.family.toLowerCase() === 'calibri' && !request.presenceOnly))
    .toHaveLength(2);
  // One family presence probe per authored family, independent of its tuples.
  expect(requests.filter((request) => request.presenceOnly).map((request) => request.family))
    .toEqual(['Aptos', 'Calibri', 'Calibri Light']);
});

it('also requests the catalogued regular face of a styled tuple the catalog lacks', () => {
  // LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY: main and worker share this list,
  // so both load the regular face a synthesized bold run is compared with.
  const doc = docWith('body', 'Arial', 'Arial');
  (doc.body[0] as { runs: object[] }).runs = [
    { type: 'text', text: '見出し', fontFamily: 'MS Mincho', bold: true, italic: false },
  ];
  const requests = docxOfficeFontFallbackRequests(doc);
  expect(requests).toContainEqual({ family: 'MS Mincho', weight: 700, style: 'normal' });
  expect(requests).toContainEqual({ family: 'MS Mincho', weight: 400, style: 'normal' });
});

it('demands every rendered authored family, whatever a source serves', () => {
  const doc = docWith('body');
  doc.majorFont = undefined;
  doc.minorFont = undefined;
  (doc.body[0] as { runs: object[] }).runs = [
    { type: 'text', text: 'a', fontFamily: 'Ubuntu' },
    { type: 'text', text: 'b', fontFamily: 'Arial', fontFamilyEastAsia: 'Roboto' },
    { type: 'text', text: 'c', fontFamily: 'Arial', fontFamilyCs: 'Franklin Gothic Book' },
  ];
  expect(docxFontPreloadNames(doc)).toEqual(['Ubuntu', 'Arial', 'Roboto', 'Franklin Gothic Book']);
});

it('checks substitute family presence independent of the styles used', () => {
  const doc = docWith('body');
  doc.majorFont = undefined;
  doc.minorFont = undefined;
  (doc.body[0] as { runs: object[] }).runs = [
    { type: 'text', text: 'Bold only', fontFamily: 'Franklin Gothic Book', bold: true },
  ];
  // The bold tuple, plus the catalogued regular face the uncatalogued bold
  // is compared with (LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY), are style
  // requests; presence is asked for the family.
  const requests = docxOfficeFontFallbackRequests(doc, { fontDemand: true });
  expect(requests.filter((request) => request.family === 'Franklin Gothic Book')).toEqual([
    { family: 'Franklin Gothic Book', weight: 700, style: 'normal' },
    { family: 'Franklin Gothic Book', weight: 400, style: 'normal' },
    { family: 'Franklin Gothic Book', presenceOnly: true },
  ]);
  // Without web fonts the family is still probed once (layout's installed
  // families), so presence never depends on the styles the document uses.
  expect(docxOfficeFontFallbackRequests(doc).filter((request) => request.presenceOnly))
    .toEqual([{ family: 'Franklin Gothic Book', presenceOnly: true }]);

  for (const officeFonts of [
    { routes: {}, installed: ['franklin gothic book'] },
    { routes: { 'franklin gothic book:700:normal': {} } },
  ]) {
    const plan = docxFontFamilyPlan(doc, GOOGLE, undefined, officeFonts);
    expect(plan.families).not.toContain('Franklin Gothic Book');
    expect(plan.installedSubstituteFamilies).toEqual(['franklin gothic book']);
  }
  const absent = docxFontFamilyPlan(doc, GOOGLE, undefined, { routes: {} });
  expect(absent.families).toContain('Franklin Gothic Book');
  // An installed family without a declared substitute is not demanded either.
  expect(docxFontFamilyPlan(doc, [], undefined, { routes: {}, installed: ['franklin gothic book'] }).families)
    .toEqual([]);
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

it('demands a theme minorFont such as Ubuntu', () => {
  expect(docxFontPreloadNames(docWith('City or Town', 'Calibri', 'Ubuntu'))).toContain('Ubuntu');
});

it('demands a family with Arabic-scoped substitutes, admitting the substitute only for Arabic text', () => {
  const doc = docWith('body', 'Aptos', 'Aptos');
  (doc.body[0] as { runs: object[] }).runs = [
    { type: 'text', text: 'Latin only', fontFamily: 'Sakkal Majalla' },
  ];
  // The authored family is always demanded (a source may serve it by name);
  // only its Arabic-scoped substitute waits for Arabic text.
  expect(docxFontPreloadNames(doc)).toContain('Sakkal Majalla');
  expect(docxFontDemandCandidates(doc, SUBSTITUTES).substituteScripts['sakkal majalla']).toEqual([]);
  (doc.body[0] as { runs: object[] }).runs.push({ type: 'text', text: 'مرحبا', fontFamily: 'Sakkal Majalla' });
  expect(docxFontDemandCandidates(doc, SUBSTITUTES).substituteScripts['sakkal majalla']).toEqual(['arabic']);
});

it('preloads the configured fallback only for ambiguous Han', () => {
  expect(docxFontPreloadNames(docWith('漢字'), 'sc')).toContain('Noto Sans SC');
  expect(docxFontPreloadNames(docWith('漢字'), 'sc')).not.toContain('Noto Sans JP');
  expect(docxFontPreloadNames(docWith('Hello'), 'sc')).toEqual(['Calibri']);
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

it('keeps default native preflight at its existing slot projection', () => {
  const doc = docWith('body');
  (doc.body[0] as { runs: object[] }).runs = [{ type: 'text', text: 'body',
    fontFamily: 'Arial', fontSlots: { direct: { ascii: 'Arial', highAnsi: 'Cambria',
      eastAsia: 'Arial', complexScript: 'Arial' }, themePresent: {} } }];
  expect(docxOfficeFontFallbackRequests(doc).some((request) => request.family === 'Cambria')).toBe(false);
  expect(docxOfficeFontFallbackRequests(doc, { fontDemand: true }))
    .toContainEqual({ family: 'Cambria', weight: 400, style: 'normal' });
  expect(docxFontPreloadNames(doc)).toContain('Cambria');
});

it('demands only families whose presence was probed, probing past the office-fallback cap', async () => {
  // 33 authored families; the capped office-fallback probe checked the first
  // 32, and "Face 33" is installed on the host.
  const doc = docWith('body');
  doc.majorFont = undefined;
  doc.minorFont = undefined;
  const names = Array.from({ length: 33 }, (_, index) => `Face ${index + 1}`);
  (doc.body[0] as { runs: object[] }).runs = names.map((family) => ({ type: 'text', text: 'x', fontFamily: family }));
  const capped: DocxOfficeFontPresence = { routes: {}, presenceChecked: names.slice(0, 32).map((family) => family.toLowerCase()) };
  // Unproven: never demanded, so no source face can shadow it.
  expect(docxFontFamilyPlan(doc, GOOGLE, undefined, capped).families).toEqual(names.slice(0, 32));
  const probed: string[][] = [];
  const completed = await completeDocxFontPresence(doc, GOOGLE, undefined, capped, async (families) => {
    probed.push([...families]);
    return { installed: new Set(['face 33']), checked: new Set(families.map((family) => family.toLowerCase())) };
  });
  expect(probed).toEqual([['Face 33']]);
  const plan = docxFontFamilyPlan(doc, GOOGLE, undefined, completed);
  expect(plan.families).toEqual(names.slice(0, 32));
  expect(completed.installed).toEqual(['face 33']);
});
