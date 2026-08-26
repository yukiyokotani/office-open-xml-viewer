import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  canonicalPageUrl,
  formatSeo,
  siteMetadata,
} from './lib/seo';
import { GET as getRobots } from './pages/robots.txt';
import { GET as getSitemap, sitemapPaths } from './pages/sitemap.xml';

const source = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8');

describe('public-site SEO foundation', () => {
  it('keeps the established OOXML Viewer term while adding the JavaScript Office viewer intent', () => {
    expect(siteMetadata.title).toBe('OOXML Viewer | JavaScript Office File Viewer');
    expect(siteMetadata.description).toContain('JavaScript and TypeScript Office file viewer library');

    const homepage = source('./pages/index.astro');
    expect(homepage).toContain('title={siteMetadata.title}');
    expect(homepage).toContain('canonicalPath="/"');
    expect(homepage).toContain('JavaScript Office file viewer library');
    expect(homepage).toContain("'@type': 'WebSite'");
    expect(source('./layouts/Base.astro')).toContain('property="og:site_name"');
  });

  it('gives each format page a distinct JavaScript viewer-library search intent', () => {
    expect(formatSeo).toEqual({
      docx: {
        title: 'JavaScript DOCX Viewer Library | @silurus/ooxml',
        description: expect.stringContaining('DOCX'),
        heading: 'Render DOCX files in the browser with JavaScript.',
      },
      xlsx: {
        title: 'JavaScript XLSX Viewer Library | @silurus/ooxml',
        description: expect.stringContaining('XLSX'),
        heading: 'Render XLSX files in the browser with JavaScript.',
      },
      pptx: {
        title: 'JavaScript PPTX Viewer Library | @silurus/ooxml',
        description: expect.stringContaining('PPTX'),
        heading: 'Render PPTX files in the browser with JavaScript.',
      },
    });
  });

  it('normalizes public canonicals to the trailing-slash URLs served by GitHub Pages', () => {
    expect(canonicalPageUrl('/')).toBe('https://ooxml.silurus.dev/');
    expect(canonicalPageUrl('/docx')).toBe('https://ooxml.silurus.dev/docx/');
    expect(canonicalPageUrl('frameworks/react/')).toBe(
      'https://ooxml.silurus.dev/frameworks/react/',
    );
  });

  it.each([
    ['./pages/try.astro', 'canonicalPath="/try/"'],
    ['./pages/errors.astro', 'canonicalPath="/errors/"'],
    ['./pages/production.astro', 'canonicalPath="/production/"'],
    ['./pages/announcements/index.astro', 'canonicalPath="/announcements/"'],
    ['./pages/announcements/[slug].astro', 'canonicalPath={`/announcements/${announcement.slug}/`}'],
    ['./layouts/FormatPage.astro', 'canonicalPath={`/${format}/`}'],
    ['./layouts/ApiPage.astro', 'canonicalPath={`/api/${format}/`}'],
    ['./components/FrameworkGuide.astro', 'canonicalPath={canonicalPath}'],
  ])('declares a canonical for %s', (path, canonicalDeclaration) => {
    expect(source(path)).toContain(canonicalDeclaration);
  });

  it('publishes only canonical public pages in the sitemap', async () => {
    const response = getSitemap();
    const xml = await response.text();

    expect(response.headers.get('content-type')).toContain('application/xml');
    for (const path of sitemapPaths) {
      expect(xml).toContain(`<loc>${canonicalPageUrl(path)}</loc>`);
    }
    expect(xml).not.toContain('/preview/');
  });

  it('advertises the sitemap in robots.txt', async () => {
    const response = getRobots();
    const robots = await response.text();

    expect(response.headers.get('content-type')).toContain('text/plain');
    expect(robots).toContain('User-agent: *');
    expect(robots).toContain('Allow: /');
    expect(robots).toContain(`Sitemap: ${canonicalPageUrl('/sitemap.xml')}`);
  });

  it.each(['lower', 'showcase'])('keeps the %s visual preview out of search results', (page) => {
    expect(source(`./pages/preview/${page}.astro`)).toContain(
      '<meta name="robots" content="noindex, nofollow" />',
    );
  });
});
