/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { createLayoutServices } from './layout-runtime.js';
import { normalizeInternalDocumentModel, paragraphMarkShapeInput } from './parser-model.js';
import { renderDocumentToCanvas } from './renderer.js';
import { fontFiles, googleFonts } from '@silurus/ooxml-core/internal/font-sources';
import { declaredFontSubstitutions, docxFontDemandCandidates } from './font-plan.js';
import { sourceFontLayoutOptions } from './testing/font-sources.js';
import { LineMeasurementAdapter } from './line-breaker/measurement-adapter.js';
import type { LayoutTextSeg } from './line-breaker/model.js';
import { createLayoutServicesRuntimeView } from './layout/runtime-state.js';
import { assertTextShapeRunContext, sliceTextShapeRequest, textScopeScanStats } from './layout/text.js';
import type { DocxDocumentModel } from './types.js';

// Font-source substitutes (googleFonts(): Calibri → Carlito, Cambria → Caladea)
// are fetched only for families in the document's demand, and the layout font
// inventory routes only demanded names to a loaded substitute. Each case parses a synthetic DOCX
// with NO theme part (so theme names cannot mask the gap) and asserts that the
// family production shaping requests is both collected and painted with the
// loaded substitute.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();

function docx(body: string, numbering?: string): Uint8Array {
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + (numbering ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : '')
      + '</Types>')],
    ['_rels/.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>')],
    ['word/_rels/document.xml.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + (numbering ? '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : '')
      + '</Relationships>')],
    ['word/document.xml', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:document xmlns:w="${W}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr></w:body></w:document>')],
  ]);
  if (numbering) {
    parts.set('word/numbering.xml', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${numbering}</w:abstractNum>`
      + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'));
  }
  return storeZip(parts);
}

const arialRun = (text: string) =>
  `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/></w:rPr><w:t>${text}</w:t></w:r>`;
const numberedParagraph = `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${arialRun('Item')}</w:p>`;

function parse(bytes: Uint8Array): DocxDocumentModel {
  const archive = new DocxArchive(bytes);
  try {
    return normalizeInternalDocumentModel(
      JSON.parse(new TextDecoder().decode(archive.parse())),
    ).document;
  } finally {
    archive.free();
  }
}

const GENERIC_FAMILIES = new Set(['serif', 'sans-serif', 'monospace']);
const ARABIC_LETTER = /\p{Script=Arabic}/u;

function cssFamilies(font: string): string[] {
  const list = /\d+(?:\.\d+)?px\s+(.*)$/u.exec(font)?.[1] ?? '';
  return list.split(',').map((entry) => entry.trim().replace(/^"(.*)"$/u, '$1'));
}

/** Emulated Canvas font selection: the first family in the CSS list that is
 * available (loaded web face, installed face or generic) and has a glyph for
 * the character. Every available face covers Latin; Arabic is covered by the
 * Arabic Noto faces, an installed Sakkal Majalla and the generics. Faces differ in advance and font box, so
 * a leaked face changes measurement as well as paint. */
function selectedFace(font: string, text: string, available: ReadonlySet<string>): string {
  const character = [...text].find((c) => c.trim()) ?? 'x';
  const arabic = ARABIC_LETTER.test(character);
  return cssFamilies(font).find((family) => (GENERIC_FAMILIES.has(family) || available.has(family))
    && (!arabic || GENERIC_FAMILIES.has(family) || /Arabic|Sakkal/u.test(family))) ?? 'serif';
}

const NASKH_METRICS = { advance: 0.9, ascent: 1.4, descent: 0.6 };
const DEFAULT_METRICS = { advance: 0.5, ascent: 0.8, descent: 0.2 };

function recordingCanvas(available: ReadonlySet<string>) {
  let font = '10px serif';
  const calls: { text: string; font: string; face: string }[] = [];
  const measurements: { text: string; face: string; width: number }[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const noop = () => {};
  const ctx = new Proxy({
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const metrics = selectedFace(font, text, available) === 'Noto Naskh Arabic'
        ? NASKH_METRICS : DEFAULT_METRICS;
      const width = [...text].length * px() * metrics.advance;
      measurements.push({ text, face: selectedFace(font, text, available), width });
      return {
        width,
        fontBoundingBoxAscent: px() * metrics.ascent,
        fontBoundingBoxDescent: px() * metrics.descent,
        actualBoundingBoxAscent: px() * metrics.ascent,
        actualBoundingBoxDescent: px() * metrics.descent,
      } as TextMetrics;
    },
    fillText(text: string) { calls.push({ text, font, face: selectedFace(font, text, available) }); },
    createLinearGradient: () => ({ addColorStop: noop }),
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
  } as Record<string | symbol, unknown>, {
    get(target, key) {
      if (key in target) return target[key];
      return noop;
    },
    set(target, key, value) {
      target[key] = value;
      return true;
    },
  });
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, measurements };
}

const loaded = (family: string) =>
  ({ family, weight: '400', style: 'normal', status: 'loaded' }) as FontFace;

const WEB_FACES = ['Carlito', 'Caladea', 'Noto Naskh Arabic'];

/** The face Canvas actually selects for each painted text, plus the layout
 * service used, so a test can also inspect shaped measurement. */
async function paintedFamilies(
  model: DocxDocumentModel,
  texts: readonly string[],
  options: Readonly<{ installedSubstituteFamilies?: readonly string[]; installedFaces?: readonly string[] }> = {},
): Promise<(string | undefined)[]> {
  const available = new Set([...WEB_FACES, ...(options.installedFaces ?? [])]);
  const { canvas, calls } = recordingCanvas(available);
  await renderDocumentToCanvas(model, canvas, 0, {
    dpr: 1,
    width: 612,
    layoutServices: createLayoutServices(model, {
      ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded), { installed: options.installedSubstituteFamilies }),
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    }),
  });
  return texts.map((text) => calls.find((entry) => entry.text.includes(text))?.face);
}

async function paintedFamily(model: DocxDocumentModel, text: string): Promise<string | undefined> {
  return (await paintedFamilies(model, [text]))[0];
}

const GOOGLE_SUBSTITUTES = declaredFontSubstitutions([googleFonts()]);
/** Demanded names before the installed check. */
function docxFontPreloadNames(model: DocxDocumentModel): string[] {
  const candidates = docxFontDemandCandidates(model, GOOGLE_SUBSTITUTES);
  return [...candidates.authored, ...candidates.fallback];
}

/** Whether demand admits the Arabic-scoped substitute for Sakkal Majalla. */
function arabicSubstituteAdmitted(model: DocxDocumentModel): boolean {
  const candidates = docxFontDemandCandidates(model, GOOGLE_SUBSTITUTES);
  return candidates.authored.includes('Sakkal Majalla')
    && (candidates.substituteScripts['sakkal majalla'] ?? []).includes('arabic');
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('Google Fonts preload collects every rendered substitute family', () => {
  it('collects and paints a directly authored Cambria body run with Caladea', async () => {
    const model = parse(docx(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/></w:rPr><w:t>SerifBody</w:t></w:r></w:p>',
    ));
    expect(model.majorFont ?? null).toBeNull();
    expect(docxFontPreloadNames(model)).toContain('Cambria');
    expect(docxFontPreloadNames(model)).not.toContain('Calibri');
    expect(await paintedFamily(model, 'SerifBody')).toBe('Caladea');
  });

  it('keeps an installed authored family even when its web substitute is already loaded', async () => {
    const model = parse(docx(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/></w:rPr><w:t>InstalledSerif</w:t></w:r></w:p>',
    ));
    expect(await paintedFamilies(model, ['InstalledSerif'], {
      installedSubstituteFamilies: ['cambria'], installedFaces: ['Cambria'],
    })).toEqual(['Cambria']);
  });

  it('collects Calibri from a table nested in a text box story', async () => {
    const model = parse(docx(
      `<w:p>${arialRun('Anchor')}<w:r><w:pict>`
      + '<v:shape id="box" type="#_x0000_t202" style="position:relative;width:300pt;height:120pt" filled="f" stroked="f">'
      + '<v:textbox><w:txbxContent>'
      + '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
      + '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
      + '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr><w:t>CellText</w:t></w:r></w:p>'
      + '</w:tc></w:tr></w:tbl><w:p/>'
      + '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>',
    ));
    expect(docxFontPreloadNames(model)).toContain('Calibri');
    expect(await paintedFamily(model, 'CellText')).toBe('Carlito');
  });

  it('collects the highAnsi Calibri slot of a bullet whose ascii slot is Arial', async () => {
    const model = parse(docx(numberedParagraph, '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
      + '<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:suff w:val="space"/>'
      + '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Calibri"/></w:rPr></w:lvl>'));
    expect(docxFontPreloadNames(model)).toContain('Calibri');
    expect(await paintedFamily(model, '•')).toBe('Carlito');
  });

  it('collects the complex-script Calibri slot of a cs number whose ascii slot is Arial', async () => {
    const model = parse(docx(numberedParagraph, '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
      + '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="space"/>'
      + '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Calibri"/><w:cs/></w:rPr></w:lvl>'));
    expect(docxFontPreloadNames(model)).toContain('Calibri');
    expect(await paintedFamily(model, '1.')).toBe('Carlito');
  });
});

describe('script-scoped Arabic visual substitutes', () => {
  const sakkal = (text: string, rtl = false) => '<w:p><w:r><w:rPr>'
    + '<w:rFonts w:ascii="Sakkal Majalla" w:hAnsi="Sakkal Majalla" w:cs="Sakkal Majalla"/>'
    + `${rtl ? '<w:rtl/>' : ''}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

  const run = (text: string, ascii: string, hAnsi: string) => '<w:p><w:r><w:rPr>'
    + `<w:rFonts w:ascii="${ascii}" w:hAnsi="${hAnsi}" w:cs="${ascii}"/>`
    + `</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  /** Every paint call that draws part of `text`, with the face Canvas selects. */
  const draws = async (body: string, text: string, useGoogleFonts = true, installed: string[] = []) => {
    const { canvas, calls } = recordingCanvas(new Set([...WEB_FACES, ...installed]));
    const model = parse(docx(body));
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 612,
      layoutServices: createLayoutServices(model, {
        ...(useGoogleFonts ? sourceFontLayoutOptions(model, WEB_FACES.map(loaded)) : {}),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      }),
    });
    return calls.filter((call) => [...call.text].some((c) => text.includes(c)))
      .map((call) => [call.text, call.face, cssFamilies(call.font)[0]]);
  };

  it('keeps paragraph marks and non-Arabic runs independent of body Arabic proof', async () => {
    const univers = '<w:rFonts w:ascii="Univers Next Arabic" w:hAnsi="Univers Next Arabic" w:cs="Arial"/>';
    const mark = `<w:pPr><w:rPr>${univers}<w:sz w:val="22"/></w:rPr></w:pPr>`;
    const model = parse(docx(
      `<w:p>${mark}<w:r><w:rPr>${univers}</w:rPr><w:t>مرحبا</w:t></w:r></w:p>`
      + `<w:p>${mark}</w:p>`
      + `<w:p>${mark}<w:r><w:rPr>${univers}<w:sz w:val="44"/></w:rPr>`
      + '<w:t xml:space="preserve">   Caption</w:t></w:r>'
      + '<w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla" w:hAnsi="Sakkal Majalla"/>'
      + '</w:rPr><w:t>ما</w:t></w:r></w:p>',
    ));
    const faces = [...WEB_FACES, 'Noto Sans Arabic', 'Sakkal Majalla'];
    const { canvas, calls, measurements } = recordingCanvas(new Set(faces));
    const options = { measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      installedSubstituteFamilies: ['sakkal majalla'] };
    const services = createLayoutServices(model, {
      ...options, ...sourceFontLayoutOptions(model, faces.map(loaded), { installed: ['sakkal majalla'] }),
    });
    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 612, layoutServices: services });
    expect(calls.some(({ text, face }) => text.includes('مرحبا') && face === 'Noto Sans Arabic')).toBe(true);
    expect(calls.some(({ text, face }) => text.includes('ما') && face === 'Sakkal Majalla')).toBe(true);
    expect(calls.filter(({ text }) => text.includes('Caption')).map(({ face }) => face)).toEqual(['sans-serif']);
    const neutral = measurements.filter(({ text }) => text === 'x' || /^ +$/u.test(text));
    expect(neutral.length).toBeGreaterThan(0);
    expect(neutral.every(({ face }) => !face.includes('Arabic'))).toBe(true);
    const withoutGoogle = createLayoutServices(model, options);
    // §17.3.1.29 marks are independent shape inputs. A paragraph/body Arabic
    // substitute must not change the empty mark's line box or borrow its proof.
    for (const paragraph of model.body) {
      if (paragraph.type !== 'paragraph') continue;
      const input = paragraphMarkShapeInput(paragraph)!;
      const request = { ...input, text: 'x' };
      const expected = withoutGoogle.text.shape(request);
      const actual = services.text.shape(request);
      expect([actual.advancePt, actual.ascentPt, actual.descentPt])
        .toEqual([expected.advancePt, expected.ascentPt, expected.descentPt]);
    }
  });

  it('measures emphasis clusters at their retained run offsets', async () => {
    const model = parse(docx('<w:p><w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla"'
      + ' w:hAnsi="Sakkal Majalla"/><w:sz w:val="20"/><w:em w:val="dot"/>'
      + '</w:rPr><w:t xml:space="preserve">ما ١٢a</w:t></w:r></w:p>'));
    const { canvas, calls, measurements } = recordingCanvas(new Set(WEB_FACES));
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 612,
      layoutServices: createLayoutServices(model, {
        ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      }),
    });
    expect(calls.some((entry) => entry.text.includes('٢') && entry.face === 'Noto Naskh Arabic')).toBe(true);
    expect(calls.some((entry) => entry.text === '•')).toBe(true);
    expect(measurements).toContainEqual({ text: '٢', face: 'Noto Naskh Arabic', width: 9 });
    expect(measurements.filter((entry) => entry.text.includes('٢'))
      .every((entry) => entry.face === 'Noto Naskh Arabic')).toBe(true);
  });

  it.each(['embedded', 'local', 'installed', 'css'] as const)(
    'keeps scalar slots when an authored %s face wins over the substitute', async (source) => {
      const model = parse(docx(run('م\u0301رحبا', 'Sakkal Majalla', 'Arial')));
      const alias = source === 'embedded' ? 'Embedded Sakkal' : 'Sakkal Majalla';
      const { canvas, calls, measurements } = recordingCanvas(new Set([...WEB_FACES, alias, 'Arial']));
      if (source === 'css') vi.stubGlobal('self', { fonts: new Set([loaded('Sakkal Majalla')]) });
      const services = createLayoutServices(model, {
        ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
        ...(source === 'embedded' ? { embeddedRoutes: [{ requestedFamily: 'Sakkal Majalla',
          resolvedFamily: alias, weight: 400, style: 'normal' as const, resourceIdentity: 'embedded:test' }] } : {}),
        // Arial is present here (its own resource metric), so the scalar-slot
        // assertions below concern substitute scoping, not absent-face metrics.
        localMetrics: {
          arial: {
            requestedFamily: 'Arial', family: 'Arial', weight: 400, style: 'normal' as const,
            sourceIdentity: 'local:arial', lineHeightRatio: 1.15, designAscentRatio: 0.905, designDescentRatio: 0.212,
          },
          ...(source === 'local' ? { 'sakkal majalla': {
            requestedFamily: 'Sakkal Majalla', family: alias, weight: 400, style: 'normal' as const,
            sourceIdentity: 'local:test', lineHeightRatio: 1,
          } } : {}),
        },
        ...(source === 'installed'
          ? sourceFontLayoutOptions(model, WEB_FACES.map(loaded), { installed: ['sakkal majalla'] }) : {}),
      });
      if (source === 'css') vi.unstubAllGlobals();
      await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 612, layoutServices: services });
      expect(calls.filter((entry) => entry.text.includes('\u0301')).map((entry) => entry.face)).toEqual(['Arial']);
      expect(measurements.filter((entry) => entry.text.includes('\u0301'))
        .every((entry) => entry.face === 'Arial')).toBe(true);
      expect(calls.some((entry) => entry.text.includes('م') && entry.face === alias)).toBe(true);
    },
  );

  it('does not lend base-run Arabic proof to an independent ruby guide', async () => {
    const model = parse(docx('<w:p><w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla"'
      + ' w:hAnsi="Sakkal Majalla"/></w:rPr><w:ruby><w:rubyPr><w:hps w:val="10"/></w:rubyPr>'
      + '<w:rt><w:r><w:t>١٢</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>ما</w:t></w:r></w:rubyBase>'
      + '</w:ruby></w:r></w:p>'));
    const { canvas, calls, measurements } = recordingCanvas(new Set(WEB_FACES));
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 612,
      layoutServices: createLayoutServices(model, {
        ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      }),
    });
    expect(calls.some((entry) => entry.text === 'ما' && entry.face === 'Noto Naskh Arabic')).toBe(true);
    // The Arabic-script guide is pair-kerned, so it paints as one contextual
    // span (WORD_ARABIC_PAIR_KERNING).
    expect(calls.filter((entry) => /[١٢]/u.test(entry.text)).map((entry) => [entry.text, entry.face]))
      .toEqual([['١٢', 'serif']]);
    const guideMeasurements = measurements.filter((entry) => /[١٢]/u.test(entry.text));
    expect(guideMeasurements.length).toBeGreaterThan(0);
    expect(guideMeasurements.every((entry) => entry.face === 'serif')).toBe(true);
  });

  it('rejects a substring measurement that bypasses context projection', () => {
    const model = parse(docx(sakkal('ما ١٢a')));
    const { canvas } = recordingCanvas(new Set(WEB_FACES));
    const services = createLayoutServices(model, {
      ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    });
    const word = sliceTextShapeRequest({ text: 'ما ١٢a', fontSizePt: 10,
      fonts: { ascii: 'Sakkal Majalla', highAnsi: 'Sakkal Majalla' } }, 3, 5);
    // All retained requests carry context. A future caller spreading one and
    // replacing its text must fail, rather than silently measure another face.
    expect(() => services.text.shape({ ...word, text: word.text.slice(1), measure: true }))
      .toThrow(/context/iu);
    expect(services.text.shape({ ...sliceTextShapeRequest(word, 1, 2), measure: true }).advancePt).toBe(9);
    // A fragment can agree with its own truncated context and still lose the
    // owning run. Acquisition's guard must also reject that false identity.
    expect(() => assertTextShapeRunContext({ text: word.text,
      substituteContext: { text: word.text, offset: 0 } }, 'ما ١٢a')).toThrow(/full run context/iu);
    const adapter = new LineMeasurementAdapter(canvas.getContext('2d') as CanvasRenderingContext2D,
      1, () => '10px serif');
    expect(() => adapter.measureSegment({ text: '٢', textShapeRequest: word,
      textLayoutService: services.text } as LayoutTextSeg)).toThrow(/range context/iu);
  });

  it.each([
    ['ما ١٢a', false, false, false],
    ['ما ١٢a', false, false, true],
    ['م\u0301رحبا', true, true, false],
    ['ما ١٢', true, false, false],
    ['ما ' + '١٢'.repeat(60) + 'a', false, false, false],
  ])('retains the same scoped face in measurement and paint: %s', async (text, rtl, hint, table) => {
    const paragraph = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla"'
      + ' w:hAnsi="Sakkal Majalla" w:eastAsia="Sakkal Majalla" w:cs="Sakkal Majalla"'
      + `${hint ? ' w:hint="eastAsia"' : ''}/>${rtl ? '<w:rtl/>' : ''}<w:sz w:val="20"/>`
      + `</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const body = table ? '<w:tbl><w:tblPr><w:tblLayout w:type="autofit"/></w:tblPr>'
      + `<w:tr><w:tc>${paragraph}</w:tc></w:tr></w:tbl>` : paragraph;
    const model = parse(docx(body));
    const { canvas, calls, measurements } = recordingCanvas(new Set(WEB_FACES));
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 612,
      layoutServices: createLayoutServices(model, {
        ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      }),
    });
    const probe = hint ? '\u0301' : '١';
    const measured = measurements.filter((entry) => entry.text.includes(probe));
    expect(measured.length).toBeGreaterThan(0);
    expect(measured.every((entry) => entry.face === 'Noto Naskh Arabic')).toBe(true);
    const painted = calls.filter((entry) => entry.text.includes(probe));
    expect(painted.length).toBeGreaterThan(0);
    expect(painted.every((entry) => entry.face === 'Noto Naskh Arabic')).toBe(true);
    if (!hint && text.length < 10) {
      expect(measurements).toContainEqual({ text: '١٢', face: 'Noto Naskh Arabic', width: 18 });
    }
  });

  it.each([false, true])('keeps a Latin-attached Arabic mark excluded during auto-fit and final measurement (split Leader: %s)', async (splitLeader) => {
    // §17.3.2.26 selects hAnsi for é and ascii for U+064E. Even when
    // intrinsic merging joins the ascii suffix to Leader, the mark's Latin
    // base in the full run must keep it outside the Arabic substitute scope.
    const run = (text: string) => '<w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla"'
      + ' w:hAnsi="Sakkal Majalla"/><w:sz w:val="20"/></w:rPr>'
      + `<w:t xml:space="preserve">${text}</w:t></w:r>`;
    const paragraph = '<w:p>' + (splitLeader
      ? run('مرحباé\u064e ') + run('Leader') : run('مرحباé\u064e Leader')) + '</w:p>';
    const model = parse(docx('<w:tbl><w:tblPr><w:tblLayout w:type="autofit"/></w:tblPr>'
      + `<w:tr><w:tc>${paragraph}</w:tc></w:tr></w:tbl>`));
    const { canvas, calls, measurements } = recordingCanvas(new Set(WEB_FACES));
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 612,
      layoutServices: createLayoutServices(model, {
        ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      }),
    });
    const measured = measurements.filter(({ text }) => text.includes('\u064e'));
    const painted = calls.filter(({ text }) => text.includes('\u064e'));
    expect(measured.length).toBeGreaterThan(0);
    expect(painted.length).toBeGreaterThan(0);
    expect(painted.every(({ face }) => face === 'serif')).toBe(true);
    expect(measured.every(({ face }) => face === 'serif')).toBe(true);
    // At the authored 10pt size, intrinsic and final shaping must agree on
    // the emulated serif advance. A same-run merge still shapes the suffix
    // together; a different-run seam must retain each run's context.
    expect(measured).toContainEqual({ text: '\u064e ', face: 'serif', width: 10 });
    if (!splitLeader) {
      expect(measured).toContainEqual({ text: '\u064e Leader', face: 'serif', width: 40 });
      expect(measurements).toContainEqual({ text: 'Leader', face: 'serif', width: 30 });
    }
    expect(calls.some(({ text, face }) => text.includes('مرحبا')
      && face === 'Noto Naskh Arabic')).toBe(true);
  });

  it("preserves each run's external Arabic proof during table intrinsic measurement", async () => {
    // U+08A0 proves Arabic through a different slot/face from the digits.
    // Adjacent digit spans in two runs therefore depend on different contexts.
    const run = (text: string) => '<w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla"'
      + ' w:hAnsi="Univers Next Arabic"/><w:sz w:val="20"/></w:rPr>'
      + `<w:t xml:space="preserve">${text}</w:t></w:r>`;
    const model = parse(docx('<w:tbl><w:tblPr><w:tblLayout w:type="autofit"/></w:tblPr>'
      + `<w:tr><w:tc><w:p>${run('\u08A0 ١٢ ')}${run('١٢ \u08A0')}</w:p></w:tc></w:tr></w:tbl>`));
    const { canvas, calls, measurements } = recordingCanvas(new Set(WEB_FACES));
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 612,
      layoutServices: createLayoutServices(model, {
        ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      }),
    });
    const measured = measurements.filter((entry) => entry.text.includes('١'));
    const painted = calls.filter((entry) => entry.text.includes('١'));
    expect(measured.length).toBeGreaterThan(0);
    expect(painted.length).toBeGreaterThan(0);
    expect(measured.every((entry) => entry.face === 'Noto Naskh Arabic')).toBe(true);
    expect(painted.every((entry) => entry.face === 'Noto Naskh Arabic')).toBe(true);
  });

  it('scans each alternating parent/child scope configuration once per run', async () => {
    const text = 'ما a\u0301 '.repeat(1000);
    const model = parse(docx('<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"'
      + ' w:eastAsia="Arial" w:cs="Sakkal Majalla" w:hint="eastAsia"/><w:rtl/>'
      + `</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`));
    const { canvas, calls } = recordingCanvas(new Set([...WEB_FACES, 'Arial']));
    const services = createLayoutServices(model, {
      ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    });
    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 612, layoutServices: services });
    expect(calls.some((call) => call.text.includes('ما') && call.face === 'Noto Naskh Arabic')).toBe(true);
    expect(calls.some((call) => call.text.includes('\u0301') && call.face === 'Arial')).toBe(true);
    // Parent tokens force cs; resolved Latin children do not. Repeated words
    // must reuse both full-run descriptors, rather than evicting each other.
    expect(textScopeScanStats(services.text)).toEqual({ scans: 2, utf16Units: 2 * text.length });
  }, 30_000);

  it.each(['', '<w:caps/>', '<w:smallCaps/>'])(
    'retains the full case-transformed run around a Latin base and Arabic mark: %s', async (transform) => {
      const original = 'مرحباé\u064e';
      const fullText = transform ? original.toUpperCase() : original;
      const model = parse(docx('<w:p><w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla"'
        + ` w:hAnsi="Sakkal Majalla"/>${transform}</w:rPr><w:t>${original}</w:t></w:r></w:p>`));
      const { canvas, calls, measurements } = recordingCanvas(new Set(WEB_FACES));
      const services = createLayoutServices(model, {
        ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      });
      const seen: string[] = [];
      const view = createLayoutServicesRuntimeView(services, { text: {
        ...services.text,
        shape(request) {
          if (request.substituteContext) {
            // Checking only substring == context.slice(offset) would accept
            // a fresh, truncated context. Compare with the independent full run.
            expect(request.substituteContext.text).toBe(fullText);
            seen.push(request.text);
          }
          return services.text.shape(request);
        },
      } });
      await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 612, layoutServices: view });
      expect(seen.some((text) => text.includes('\u064e'))).toBe(true);
      expect(calls.filter((call) => call.text.includes('\u064e')).map((call) => call.face)).toEqual(['serif']);
      expect(measurements.filter((entry) => entry.text.includes('\u064e'))
        .every((entry) => entry.face === 'serif')).toBe(true);
    },
  );

  it.each([
    ['Sakkal Majalla', '<w:smallCaps/>', 'مرحباßé\u064e\t١٢', 'مرحباSSÉ\u064e\t١٢'],
    // Wingdings 0x21 is U+1F589 (astral), 0x76 U+2756; neither is emoji-capable.
    ['Wingdings', '', '\uf021\t\uf076', '\u{1F589}\t\u2756'],
  ])('projects expanded display scalars and tab pieces into one full %s run', async (family, transform, original, convertedText) => {
    const model = parse(docx('<w:p><w:r><w:rPr>'
      + `<w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>${transform}`
      + `</w:rPr><w:t>${original}</w:t></w:r></w:p>`));
    const { canvas, calls } = recordingCanvas(new Set(WEB_FACES));
    const services = createLayoutServices(model, { ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D });
    // A symbol run converts only what the measured face lacks; a context that
    // cannot judge it converts the 0.89 subset (its BMP pencil, with VS15).
    const fullText = family === 'Wingdings'
      ? services.text.symbolText!({ text: original, fonts: { ascii: family, highAnsi: family }, weight: 400, style: 'normal' }).map((part) => part.text).join('')
      : convertedText;
    expect([convertedText, '\u270F\uFE0E\t\u2756']).toContain(fullText);
    const ranges: string[] = [];
    const view = createLayoutServicesRuntimeView(services, { text: { ...services.text, shape(request) {
      if (request.substituteContext) {
        assertTextShapeRunContext(request, fullText);
        ranges.push(request.text);
      }
      return services.text.shape(request);
    } } });
    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 612, layoutServices: view });
    expect(ranges.some((text) => text.includes(family === 'Wingdings' ? '\u2756' : '١٢'))).toBe(true);
    expect(calls.some((call) => call.text.includes(family === 'Wingdings' ? '\u2756' : 'É'))).toBe(true);
    if (family === 'Sakkal Majalla') {
      expect(calls.filter((call) => call.text.includes('١٢') || call.text.includes('\u064e'))
        .every((call) => call.face === 'serif')).toBe(true);
    }
  });

  it('does not merge hidden Arabic proof into the next visible run', async () => {
    const model = parse(docx('<w:p><w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla"'
      + ' w:hAnsi="Sakkal Majalla"/><w:vanish/></w:rPr><w:t>ما</w:t></w:r>'
      + '<w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla" w:hAnsi="Sakkal Majalla"/>'
      + '</w:rPr><w:t>١٢</w:t></w:r></w:p>'));
    expect(await paintedFamilies(model, ['١٢'])).toEqual(['serif']);
  });

  it('shares scoped word shapes across 4,000 offsets and different run contexts', async () => {
    const context = 'ما ' + '١٢ '.repeat(4000);
    const model = parse(docx(sakkal(context)));
    const { canvas, calls, measurements } = recordingCanvas(new Set(WEB_FACES));
    const services = createLayoutServices(model, {
      ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    });
    const fonts = { ascii: 'Sakkal Majalla', highAnsi: 'Sakkal Majalla', complexScript: 'Sakkal Majalla' };
    const shape = (text: string, context: string, offset: number) => services.text.shape({
      text, fonts, fontSizePt: 10, measure: false, substituteContext: { text: context, offset },
    });
    const first = shape('١٢', context, 3);
    for (let index = 1; index < 4000; index += 1) {
      expect(shape('١٢', context, 3 + index * 3)).toBe(first);
    }
    expect(shape('١٢', 'مرحبا ١٢', 6)).toBe(first);
    expect(first.spans[0]?.font.resolvedFamily).toBe('Noto Naskh Arabic');
    // Identical text outside proven Arabic must never alias that cached shape.
    const unproven = shape('١٢', 'Leader ١٢', 7);
    expect(unproven).not.toBe(first);
    expect(unproven.spans[0]?.font.resolvedFamily).not.toBe('Noto Naskh Arabic');
    expect(textScopeScanStats(services.text)).toEqual({ scans: 3,
      utf16Units: context.length + 'مرحبا ١٢'.length + 'Leader ١٢'.length });
    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 612, layoutServices: services });
    expect(calls.some((entry) => entry.text.includes('١٢'))).toBe(true);
    expect(calls.every((entry) => entry.face === 'Noto Naskh Arabic')).toBe(true);
    const digitMeasurements = measurements.filter((entry) => entry.text.includes('١٢'));
    expect(digitMeasurements.length).toBeGreaterThan(0);
    expect(digitMeasurements.every((entry) => entry.face === 'Noto Naskh Arabic')).toBe(true);
  // Scope work and shape sharing above are deterministic resource bounds;
  // allow the full paint fixture to finish on hosts running parallel sessions.
  }, 60_000);

  it('merges joined Arabic across slots when both resolve to the same substitute face', async () => {
    // U+08A0 is a highAnsi scalar; the two requested families both resolve to
    // the Noto Naskh Arabic substitute face.
    const text = '\u0644\u08A0\u0627';
    expect((await draws(run(text, 'Sakkal Majalla', 'Traditional Arabic'), text))
      .map(([drawnText, face]) => [drawnText, face])).toEqual([[text, 'Noto Naskh Arabic']]);
  });

  it('keeps transparent controls inside joined Arabic across slot boundaries', async () => {
    for (const control of ['\u200F', '\uFEFF', '\u200D', '\u200C']) {
      const text = `\u0644${control}\u0627`;
      expect((await draws(sakkal(text), text)).map(([drawnText, face]) => [drawnText, face]))
        .toEqual([[text, 'Noto Naskh Arabic']]);
    }
  });

  it('never lets Arabic digits, comma or tatweel alone enable the substitute', async () => {
    for (const [text, probe] of [
      ['Leader\u0661\u0662', '\u0661'], ['\u0661\u0662', '\u0661'], ['\u060C', '\u060C'], ['\u0640', '\u0640'],
    ] as const) {
      const drawnFaces = (await draws(sakkal(text), probe)).map(([, face]) => face);
      expect(drawnFaces.length).toBeGreaterThan(0);
      expect(drawnFaces).not.toContain('Noto Naskh Arabic');
    }
    // Inside proven Arabic text they continue the scope.
    const proven = '\u0645\u0640\u0627 \u0661\u0662';
    const provenDraws = await draws(sakkal(proven), proven);
    expect(provenDraws.length).toBeGreaterThan(0);
    expect(provenDraws.every(([, face]) => face === 'Noto Naskh Arabic')).toBe(true);
  });

  it('keeps main\'s per-scalar slots for general text without a scoped family', async () => {
    // ECMA-376 §17.3.2.26 per code point: the combining acute is a highAnsi
    // scalar and keeps the hAnsi face, as on main.
    const text = 'a\u0301';
    const painted = await draws(run(text, 'Courier New', 'Arial'), text, false, ['Courier New', 'Arial']);
    expect(painted.map(([drawnText, , first]) => [drawnText, first]))
      .toEqual([['a', 'Courier New'], ['\u0301', 'Arial']]);
  });

  it('neither preloads nor paints Noto Naskh Arabic for Latin-only Sakkal Majalla text', async () => {
    const model = parse(docx(sakkal('Leader')));
    expect(arabicSubstituteAdmitted(model)).toBe(false);
    // Naskh is loaded (another run's Arabic text would load it too), but the
    // missing authored face must fall back past it for Latin.
    expect(await paintedFamily(model, 'Leader')).toBe('serif');
    const { canvas } = recordingCanvas(new Set(WEB_FACES));
    const services = createLayoutServices(model, {
      ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    });
    const slots = { ascii: 'Sakkal Majalla', highAnsi: 'Sakkal Majalla', complexScript: 'Sakkal Majalla' };
    const latin = services.text.shape({ text: 'Leader', fontSizePt: 10, fonts: slots, measure: true });
    expect(latin.advancePt).toBeCloseTo(6 * 10 * DEFAULT_METRICS.advance, 6);
    expect(latin.ascentPt).toBeCloseTo(10 * DEFAULT_METRICS.ascent, 6);
    const arabic = services.text.shape({ text: 'مرحبا', fontSizePt: 10, fonts: slots, measure: true });
    // This service was created for Latin-only text: no Arabic substitute was
    // requested/loaded for the authored family, so registry presence alone
    // must not enable its slot override or Noto fallback.
    expect(arabic.ascentPt).toBeCloseTo(10 * DEFAULT_METRICS.ascent, 6);
  });

  it('paints only the Arabic characters of a mixed ascii-slot run with the substitute', async () => {
    const model = parse(docx(sakkal('مرحبا Leader')));
    expect(arabicSubstituteAdmitted(model)).toBe(true);
    expect(await paintedFamilies(model, ['مرحبا', 'Leader']))
      .toEqual(['Noto Naskh Arabic', 'serif']);
  });

  it('uses the shared scope rule for Latin punctuation in complex-script spans', async () => {
    // ASCII punctuation is other in the shared rule, even in the cs slot.
    const model = parse(docx(sakkal('مرحبا! بكم', true)));
    const [family] = await paintedFamilies(model, ['!']);
    expect(family).toBe('serif');
  });

  it('draws vocalised Arabic once with the substitute, never splitting marks from letters', async () => {
    const vocalised = '\u0645\u064E\u0631\u0652\u062D\u064E\u0628\u064B\u0627'; // مَرْحَبًا
    const tatweel = '\u0645\u0640\u0640\u0627';
    for (const text of [vocalised, tatweel, '\u064E']) {
      const { canvas, calls } = recordingCanvas(new Set(WEB_FACES));
      const model = parse(docx(sakkal(text)));
      await renderDocumentToCanvas(model, canvas, 0, {
        dpr: 1, width: 612,
        layoutServices: createLayoutServices(model, {
          ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
          measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
        }),
      });
      const drawn = calls.filter((call) => [...call.text].some((c) => text.includes(c)));
      expect(drawn.map((call) => [call.text, call.face])).toEqual([[text, 'Noto Naskh Arabic']]);
    }
  });

  it('decides slot and face per grapheme cluster across the slot boundary', async () => {
    const drawn = async (text: string) => {
      const { canvas, calls } = recordingCanvas(new Set(WEB_FACES));
      const model = parse(docx(sakkal(text)));
      await renderDocumentToCanvas(model, canvas, 0, {
        dpr: 1, width: 612,
        layoutServices: createLayoutServices(model, {
          ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
          measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
        }),
      });
      return calls.filter((call) => [...call.text].some((c) => text.includes(c)))
        .map((call) => [call.text, call.face]).sort();
    };
    const arabic = '\u0645\u0631\u062D\u0628\u0627'; // مرحبا
    // A generic combining mark on an Arabic letter stays with it in Naskh.
    expect(await drawn('\u0645\u0301\u0631\u062D\u0628\u0627'))
      .toEqual([['\u0645\u0301\u0631\u062D\u0628\u0627', 'Noto Naskh Arabic']]);
    // An Arabic mark on a Latin base or on Latin punctuation takes its base's
    // (out-of-scope) face, never Naskh. Outside the Arabic scope each scalar
    // keeps its own ECMA-376 §17.3.2.26 slot, as on main.
    const facesByScalar = (draws: string[][]) => Object.fromEntries(
      draws.flatMap(([text, face]) => [...text!].map((scalar) => [scalar, face])),
    );
    for (const [base, mark] of [['\u00E9', '\u064E'], ['\u2019', '\u064E'], ['e', '\u0301']]) {
      const draws = await drawn(`${arabic}${base}${mark}`);
      expect(draws).toContainEqual([arabic, 'Noto Naskh Arabic']);
      expect(facesByScalar(draws)).toMatchObject({ [base!]: 'serif', [mark!]: 'serif' });
    }
    // Joiner contexts are shaped and painted as one string.
    for (const joiner of ['\u200D', '\u200C']) {
      const text = `\u0644${joiner}\u0627`;
      expect(await drawn(text)).toEqual([[text, 'Noto Naskh Arabic']]);
    }
  });

  it('splits mixed Latin and vocalised Arabic only at the script boundary', async () => {
    const vocalised = '\u0645\u064E\u0631\u0652\u062D\u064E\u0628\u064B\u0627';
    const model = parse(docx(sakkal(`${vocalised}Leader`)));
    expect(await paintedFamilies(model, [vocalised, 'Leader'])).toEqual(['Noto Naskh Arabic', 'serif']);
    const { canvas } = recordingCanvas(new Set(WEB_FACES));
    const services = createLayoutServices(model, {
      ...sourceFontLayoutOptions(model, WEB_FACES.map(loaded)),
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    });
    const shaped = services.text.shape({
      text: `${vocalised}Leader`, fontSizePt: 10, measure: true,
      fonts: { ascii: 'Sakkal Majalla', highAnsi: 'Sakkal Majalla', complexScript: 'Sakkal Majalla' },
    });
    expect(shaped.spans.map((span) => span.text)).toEqual([vocalised, 'Leader']);
  });

  it('never lets an invisible control enable the substitute for a complex-script span', async () => {
    const model = parse(docx(sakkal('Leader\uFEFF', true)));
    expect(arabicSubstituteAdmitted(model)).toBe(false);
    expect(await paintedFamily(model, 'Leader')).toBe('serif');
  });

  it('never routes an installed authored family to its substitute', async () => {
    const model = parse(docx(sakkal('مرحبا Leader')));
    expect(await paintedFamilies(model, ['مرحبا'], {
      installedSubstituteFamilies: ['sakkal majalla'],
      installedFaces: ['Sakkal Majalla'],
    })).toEqual(['Sakkal Majalla']);
  });
});

describe('substitution relationships are keyed per authored family', () => {
  it('serves a directly supplied family for its own Latin text although another family substitutes to it', async () => {
    // A face supplied for A, with the Arabic-scoped pairs A → B and C → A.
    // Being C's target must not scope A out of A's own route.
    const sources = [fontFiles({
      faces: [{ family: 'Alpha Face', url: 'https://app.example/alpha.woff2' }],
      substitutes: [
        { for: 'Alpha Face', use: 'Beta Face', script: 'arabic' },
        { for: 'Gamma Face', use: 'Alpha Face', script: 'arabic' },
      ],
    })];
    const model = parse(docx('<w:p><w:r><w:rPr><w:rFonts w:ascii="Alpha Face" w:hAnsi="Alpha Face"'
      + ' w:cs="Alpha Face"/></w:rPr><w:t>Latin text</w:t></w:r></w:p>'));
    const { canvas, calls } = recordingCanvas(new Set(['Alpha Face']));
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 612,
      layoutServices: createLayoutServices(model, {
        ...sourceFontLayoutOptions(model, [loaded('Alpha Face')], { sources }),
        measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
      }),
    });
    const painted = calls.filter((call) => call.text.includes('Latin'));
    expect(painted.length).toBeGreaterThan(0);
    expect(painted.map((call) => call.face)).toEqual(painted.map(() => 'Alpha Face'));
    // Its own served family leads, through the document's private family.
    expect(painted.every((call) => /^ooxml-font-source-test-\d+$/.test(cssFamilies(call.font)[0] ?? '')
      && cssFamilies(call.font)[1] === 'Alpha Face')).toBe(true);
  });
});
