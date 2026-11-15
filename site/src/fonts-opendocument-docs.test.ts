import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { announcements } from './lib/announcements';
import { sitemapPaths } from './pages/sitemap.xml';

const source = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8');
const json = (path: string) => JSON.parse(source(path)) as { exports: Record<string, unknown>; peerDependencies?: Record<string, string>; version: string };

const fontFactories = ['googleFonts', 'fontCssApi', 'fontCss', 'fontFiles', 'fontResolver'] as const;
const odfReaders = [
  ['odp', 'createOdp', 'PptxViewer', 'opendocument_odp_bg.wasm'],
  ['odt', 'createOdt', 'DocxViewer', 'opendocument_odt_bg.wasm'],
  ['ods', 'createOds', 'XlsxViewer', 'opendocument_ods_bg.wasm'],
] as const;

describe('font sources guide', () => {
  const page = source('./pages/fonts.astro');

  it('is published and linked like the other guides', () => {
    expect(page).toContain('canonicalPath="/fonts/"');
    expect(sitemapPaths).toContain('/fonts/');
    for (const path of ['./components/Nav.astro', './components/SiteFooter.astro']) {
      expect(source(path)).toContain('href="/fonts"');
    }
    expect(source('./pages/errors.astro')).toContain('href="/fonts#limits"');
  });

  it('documents only factories and substitution lists that every public entry exports', () => {
    const root = source('../../src/index.ts');
    const baselines = ['docx', 'xlsx', 'pptx'].map((format) => source(`../../packages/${format}/api/public-api-baseline.d.ts`));
    for (const factory of fontFactories) {
      expect(page).toContain(`${factory}(`);
      expect(root).toMatch(new RegExp(`\\b${factory},`));
      for (const baseline of baselines) expect(baseline).toContain(`export function ${factory}(`);
    }
    for (const list of ['officeFontSubstitutions', 'googleFontsSubstitutions']) {
      expect(page).toContain(list);
      expect(root).toContain(list);
      for (const baseline of baselines) expect(baseline).toContain(`export const ${list}:`);
    }
    for (const type of ['FontResolveFn', 'FontDemand', 'FontResolveContext', 'FontFaceDeclaration']) {
      expect(page).toContain(type);
      expect(root).toContain(`type ${type},`);
    }
  });

  it('answers fallback families in the resolver examples', () => {
    const request = 'JSON.stringify({ families: demand.families, fallbackFamilies: demand.fallbackFamilies })';
    expect(page).toContain(request);
    expect(JSON.stringify(announcements[0])).toContain(request.replace(/"/g, '\\"'));
  });

  it('keeps the limits table and migration table identical to the migration guide', () => {
    const guide = source('../../docs/migration-0.90.md');
    const rows = [...page.matchAll(/\['([^']+)', '([^']+)'\],/g)].map(([, name, value]) => [name, value]);
    expect(rows.length).toBe(6);
    const canonical = guide.replace(/`/g, '');
    for (const [name, value] of rows) expect(canonical).toContain(`| ${name} | ${value} |`);
    expect(guide).toContain('| `useGoogleFonts: true` | `fonts: [googleFonts()]` |');
    expect(page).toContain('<td><code>useGoogleFonts: true</code></td><td><code>fonts: [googleFonts()]</code></td>');
    expect(guide).toContain('two phases, each with its own 15 s deadline');
    expect(page).toContain('two phases, each with its own 15 s deadline');
    expect(page).toContain('https://github.com/yukiyokotani/office-open-xml-viewer/blob/main/docs/migration-0.90.md');
  });
});

describe('OpenDocument guide', () => {
  const page = source('./pages/opendocument.astro');
  const pkg = json('../../packages/opendocument/package.json');
  const odfGuide = source('../../docs/opendocument.md');

  it('is published and linked like the legacy Office guide', () => {
    expect(page).toContain('canonicalPath="/opendocument/"');
    expect(sitemapPaths).toContain('/opendocument/');
    for (const path of ['./components/Nav.astro', './components/SiteFooter.astro', './components/Capabilities.astro', './pages/production.astro']) {
      expect(source(path)).toContain('href="/opendocument"');
    }
    expect(source('../../README.md')).toContain('https://ooxml.silurus.dev/opendocument/');
  });

  it('documents only entries and factories the add-on package exports', () => {
    expect(page).toContain('@silurus/ooxml-opendocument');
    expect(pkg.peerDependencies?.['@silurus/ooxml']).toBe(pkg.version);
    for (const [entry, factory, viewer, wasm] of odfReaders) {
      expect(Object.keys(pkg.exports)).toContain(`./${entry}`);
      const entrySource = source(`../../packages/opendocument/src/${entry}.ts`);
      expect(entrySource).toContain(`export const ${entry}:`);
      expect(entrySource).toContain(`export function ${factory}(`);
      expect(entrySource).toContain(wasm);
      expect(page).toContain(`import { ${entry} } from '@silurus/ooxml-opendocument/${entry}'`);
      expect(page).toContain(`{ ${entry} }`);
      expect(page).toContain(factory);
      expect(page).toContain(viewer);
      expect(page).toContain(`${entry}-source-module.mjs`);
      // Asset names come from the canonical OpenDocument reference.
      expect(odfGuide).toContain(`${entry}-source-module.mjs`);
      expect(odfGuide).toContain(wasm);
      expect(page).toContain(wasm);
    }
    for (const contract of ['companion-runtime-mismatch', 'unsupported-encryption']) {
      expect(odfGuide).toContain(contract);
      expect(page).toContain(contract);
    }
  });
});

describe('v0.90.0 release note', () => {
  const announcement = announcements.find(({ slug }) => slug === 'v090-opendocument-and-font-sources');

  it('leads the announcements and links both guides', () => {
    expect(announcements[0]).toBe(announcement);
    expect(announcement).toMatchObject({ label: 'Release note', version: 'v0.90.0' });
    expect(announcement?.sections[0]).toMatchObject({ title: 'In short', kind: 'summary' });
    expect(announcement?.sections.at(-1)?.title).toBe('Upgrading');
    const hrefs = announcement?.sections.flatMap(({ links }) => links?.map(({ href }) => href) ?? []) ?? [];
    expect(hrefs).toContain('/fonts#migration');
    expect(hrefs).toContain('/opendocument');
    const text = JSON.stringify(announcement);
    for (const phrase of ['@silurus/ooxml-opendocument', '{ odp }', 'createOdp', 'fonts: [googleFonts()]', 'useGoogleFonts', 'connect-src', 'font-source-limit', 'companion-runtime-mismatch', 'Wingdings']) {
      expect(text).toContain(phrase);
    }
    expect(announcement?.summary).not.toContain('`');
  });
});
