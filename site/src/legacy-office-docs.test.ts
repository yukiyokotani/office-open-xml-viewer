import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { announcements } from './lib/announcements';
import { sitemapPaths } from './pages/sitemap.xml';

const source = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8');

describe('experimental legacy Office add-on guidance', () => {
  it('publishes an opt-in guide without promoting legacy formats into the main navigation', () => {
    const page = source('./pages/legacy-office.astro');
    expect(page).toContain('canonicalPath="/legacy-office/"');
    expect(sitemapPaths).toContain('/legacy-office/');
    for (const path of ['./components/SiteFooter.astro', './components/Capabilities.astro', './pages/production.astro']) {
      expect(source(path)).toContain('href="/legacy-office"');
    }
    expect(source('./components/Nav.astro')).not.toContain('/legacy-office');
    for (const [entry, reader, viewer] of [
      ['doc', 'legacyDoc', 'DocxViewer'],
      ['xls', 'legacyXls', 'XlsxViewer'],
      ['ppt', 'legacyPpt', 'PptxViewer'],
    ]) {
      expect(page).toContain(`@silurus/ooxml-legacy-office/${entry}`);
      expect(page).toContain(`import { ${reader} } from '@silurus/ooxml-legacy-office/${entry}'`);
      expect(page).toContain(`{ ${reader} }`);
      expect(page).toContain(viewer);
    }
    expect(page).not.toContain('modelSources');
    for (const contract of ['experimental', 'versions must match', 'openPptxPresentation', 'session.close()', 'createLegacyPpt', 'legacy-binary-format', 'unsupported-content', 'source-subset', 'Markdown', 'WASM', 'moduleUrl', 'wasmUrl']) {
      expect(page).toContain(contract);
    }
    expect(source('./pages/bundle-size.astro')).toContain('Legacy Office add-on');
    expect(source('./pages/bundle-size.astro')).toContain('OpenDocument add-on');
    expect(source('../../README.md')).toContain('https://ooxml.silurus.dev/legacy-office/');
  });

  it('announces the optional subset as the v0.89.0 release note with an explicit migration boundary', () => {
    const announcement = announcements.find(({ slug }) => slug === 'experimental-legacy-office');
    expect(announcement).toMatchObject({ label: 'Release note', version: 'v0.89.0', date: '2026-10-19' });
    expect(announcement?.sections[0].kind).toBe('summary');
    const text = JSON.stringify(announcement);
    for (const phrase of ['experimental', '@silurus/ooxml-legacy-office', '{ legacyDoc }', 'createLegacyDoc', 'No migration is required', 'versions must match', 'reject', 'OOXML viewer']) {
      expect(text).toContain(phrase);
    }
    expect(text).not.toContain('modelSources');
    expect(text).not.toMatch(/\b(?:KiB|MiB|gzip)\b|\d+%/);
  });
});
