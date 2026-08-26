import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { sitemapPaths } from './pages/sitemap.xml';

const source = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8');

describe('documentation information architecture', () => {
  it.each(['docx', 'xlsx', 'pptx'])('keeps /%s focused on live examples', (format) => {
    const page = source(`./pages/${format}.astro`);
    expect(page).toContain('<DemoBlock');
    expect(page).not.toContain('ApiReference');
    expect(source('./layouts/FormatPage.astro')).toContain('href={`/api/${format}`}');
  });

  it.each(['docx', 'xlsx', 'pptx'])('publishes /api/%s as a separate API reference', (format) => {
    expect(source(`./pages/api/${format}.astro`)).toContain('<ApiPage');
    expect([...sitemapPaths]).toContain(`/api/${format}/`);
  });

  it('keeps API pages navigable without mixing in design guidance', () => {
    const layout = source('./layouts/ApiPage.astro');
    const reference = source('./components/ApiReference.astro');
    expect(layout).toContain('class="api-toc"');
    expect(layout).toContain('href="/production"');
    expect(reference).toContain('Options &amp; methods');
    expect(reference).not.toContain('Choose a rendering mode');
    expect(reference).not.toContain('Load once for one Viewer');
    expect(reference).not.toContain('ChartEx, 3-D charts and Region Maps');
  });

  it('puts shared decisions on one production page', () => {
    const page = source('./pages/production.astro');
    expect(page).toContain('id="rendering-mode"');
    expect(page).toContain('id="ownership"');
    expect(page).toContain('id="optional-renderers"');
    expect(sitemapPaths).toContain('/production/');
  });
});
