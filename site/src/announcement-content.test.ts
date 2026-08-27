import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { announcements } from './lib/announcements';

const articlePage = readFileSync(new URL('./pages/announcements/[slug].astro', import.meta.url), 'utf8');
const bundleSizePage = readFileSync(new URL('./pages/bundle-size.astro', import.meta.url), 'utf8');
const apiReference = readFileSync(new URL('./components/ApiReference.astro', import.meta.url), 'utf8');
const apiReferenceData = readFileSync(new URL('./lib/api-reference.ts', import.meta.url), 'utf8');
const siteFooter = readFileSync(new URL('./components/SiteFooter.astro', import.meta.url), 'utf8');
const capabilities = readFileSync(new URL('./components/Capabilities.astro', import.meta.url), 'utf8');
const readme = readFileSync(new URL('../../README.md', import.meta.url), 'utf8');

describe('v0.82 review announcement', () => {
  const announcement = announcements.find((item) => item.slug === 'v082-review-comments');

  it('states the read-only cross-format boundary and compatibility', () => {
    expect(announcement).toMatchObject({
      label: 'Release note',
      version: 'v0.82.0',
      title: 'Comments and tracked changes in v0.82.0',
    });
    expect(announcement?.sections[0]).toMatchObject({ title: 'Comments in context', kind: 'summary' });

    const text = announcement?.sections.flatMap((section) => [
      section.title,
      ...section.paragraphs,
      ...(section.examples?.map(({ code }) => code) ?? []),
    ]).join('\n') ?? '';

    for (const format of ['DOCX', 'XLSX', 'PPTX']) expect(text).toContain(format);
    expect(text).toContain('read-only');
    expect(text).toContain('comments: true');
    expect(text).toContain('stable CSS classes and custom properties');
    expect(text).toContain('list virtualization');
    expect(text).toContain('insertions, deletions and moves');
    expect(text).toContain('there is no built-in tracked-change markup view');
    expect(text).not.toContain('select the tracked-change presentation');
    expect(text).toContain('No existing option is removed or renamed');
    expect(text).not.toMatch(/\b(?:KB|KiB|gzip)\b/);
  });
});

describe('v0.81 ChartEx migration guide', () => {
  const announcement = announcements.find((item) => item.slug === 'v081-chartex-opt-in');

  it('makes the required migration decision explicit', () => {
    expect(announcement).toMatchObject({
      label: 'Release note',
      version: 'v0.81.0',
      title: 'Migrating to v0.81.0',
    });
    expect(announcement?.sections.map(({ title }) => title)).toEqual(['ChartEx support', 'Migration']);
    expect(announcement?.sections[0]).toMatchObject({ kind: 'summary' });

    const text = announcement?.sections.flatMap((section) => [
      section.title,
      ...(section.modules ?? []),
      ...section.paragraphs,
      ...(section.bullets ?? []),
      ...(section.examples?.map(({ code }) => code) ?? []),
    ]).join('\n') ?? '';

    expect(text).toContain('Classic charts remain built in and require no application changes');
    expect(text).toContain("@silurus/ooxml/chart-ex");
    expect(text).toContain('chartEx');
    expect(text).toContain('waterfall');
    expect(announcement?.summary).toContain('expands Microsoft ChartEx rendering');
    expect(announcement?.summary).not.toContain('adds Microsoft ChartEx rendering');
  });

  it('keeps volatile measurements and implementation detail out of the announcement', () => {
    const text = announcement?.sections.flatMap((section) => [
      ...section.paragraphs,
      ...(section.bullets ?? []),
    ]).join('\n') ?? '';

    expect(text).not.toMatch(/\b(?:KB|KiB|gzip)\b/);
    expect(text).not.toContain('structured-clone');
    expect(text).not.toContain('self-contained render-worker');
    expect(text).not.toContain('shared chart model');
  });
});

describe('stable documentation boundaries', () => {
  it('keeps the current bundle measurements on one stable page', () => {
    expect(bundleSizePage).toContain('Current production assets in v0.82.1');
    expect(bundleSizePage).toContain('DOCX static JavaScript');
    expect(bundleSizePage).toContain('XLSX static JavaScript');
    expect(bundleSizePage).toContain('PPTX static JavaScript');
    expect(bundleSizePage).toContain('ChartEx');
    expect(bundleSizePage).toContain('Parser WASM');
    expect(bundleSizePage).toContain('complete synchronous import graph');
    expect(bundleSizePage).not.toContain('change from v0.81.0');
    expect(siteFooter).toContain('href="/bundle-size"');
    expect(readme).toContain('https://ooxml.silurus.dev/bundle-size/');
    expect(readme).not.toContain('For v0.79.0, the complete npm package');
  });

  it('describes ChartEx as opt-in for every document host', () => {
    expect(capabilities.match(/optional ChartEx/g)).toHaveLength(3);
  });

  it('does not link API details directly to release notes', () => {
    expect(apiReference).not.toContain('href="/announcements/');
    expect(apiReferenceData).not.toContain("detailsHref: '/announcements/");
  });
});

describe('v0.80 worker rendering announcement', () => {
  const announcement = announcements.find((item) => item.slug === 'v080-worker-rendering');

  it('presents worker rendering as a released compatible minor version', () => {
    expect(announcement).toBeDefined();
    expect(announcement).toMatchObject({ label: 'Release note', version: 'v0.80.0' });
    expect(announcement?.summary).toContain('extends the existing');
    expect(announcement?.summary).not.toContain('adds one worker rendering mode');
    expect(announcement?.sections[0]).toMatchObject({ title: 'In short', kind: 'summary' });
    expect(announcement?.sections[0]?.paragraphs.join(' ')).toContain('introduced in v0.59.0');
    expect(announcement?.sections[0]?.paragraphs.join(' ')).toContain('Main-thread mode remains the default');
  });

  it('states the main-thread boundary and production trade-offs', () => {
    const text = announcement?.sections.flatMap((section) => [
      section.title,
      ...section.paragraphs,
      ...(section.bullets ?? []),
      ...(section.examples?.map(({ code }) => code) ?? []),
    ]).join('\n') ?? '';

    expect(text).toContain('Choose the mode that fits your app');
    expect(text).toContain('Use main mode for smaller documents');
    expect(text).toContain('Use worker mode when larger or more complex documents');
    expect(text).toContain("mode: 'worker'");
    expect(text).toContain('published render worker is therefore self-contained');
    expect(text).toContain('larger self-contained worker asset');
    expect(text).toContain('transfers a rendered bitmap for each frame');
    expect(text).toContain('equations, 3-D charts and Region Maps');
    expect(text).toContain('structured-clone boundary');
    expect(text).toContain('opaque asset');
    expect(text).toContain('size-bounded surface');
    expect(text).toContain('cached at 64 px/em');
    expect(text).toContain('pixel-identical');
    expect(text).toContain('automatically uses main mode');
  });
});

describe('v0.79 chart rendering announcement', () => {
  const announcement = announcements.find((item) => item.slug === 'v079-chart-rendering-addons');

  it('publishes the release and states the opt-in decision before implementation detail', () => {
    expect(announcement).toBeDefined();
    expect(announcement?.label).toBe('Release note');
    expect(announcement?.sections[0]).toMatchObject({ title: 'In short', kind: 'summary' });
    const summary = announcement?.sections[0]?.paragraphs.join(' ') ?? '';
    expect(summary).toContain('optional renderer modules');
    expect(summary).toContain('No migration');
  });

  it('documents module boundaries, Region Map provenance and fidelity scope', () => {
    const text = announcement?.sections.flatMap((section) => [
      ...(section.modules ?? []),
      ...section.paragraphs,
      ...(section.bullets ?? []),
      ...(section.examples?.map(({ code }) => code) ?? []),
    ]).join('\n') ?? '';
    expect(text).toContain('@silurus/ooxml/three-d');
    expect(text).toContain('@silurus/ooxml/region-map');
    expect(text).toContain('Natural Earth');
    expect(text).toContain('country-level world maps');
    expect(text).toContain('main thread');
    expect(text).not.toContain('interactive orbit');
    expect(text).not.toContain('copied from Excel or Bing');
  });

  it('uses one local renderer-produced image with explicit provenance', () => {
    expect(announcement?.image).toMatchObject({
      src: '/announcements/chart-rendering-v079.webp',
    });
    expect(announcement?.image?.alt).toContain('synthetic');
    expect(announcement?.image?.caption).toContain('Natural Earth');
    expect(articlePage).toContain('announcement.image');
    expect(articlePage).toContain('<figcaption>');
  });
});

describe('resource-governance announcement', () => {
  const announcement = announcements.find((item) => item.slug === 'v075-resource-governance');

  it('starts with a direct migration decision', () => {
    expect(announcement).toBeDefined();
    expect(announcement?.sections[0]).toMatchObject({ title: 'In short', kind: 'summary' });
    expect(announcement?.sections[0]?.paragraphs.join(' ')).toContain('do not need to change');
    expect(announcement?.sections[0]?.bullets?.join(' ')).toContain('maxZipEntryBytes');
  });

  it('documents defaults, typed failures, metrics and the WASM boundary', () => {
    const text = announcement?.sections.flatMap((section) => [
      ...section.paragraphs,
      ...(section.bullets ?? []),
    ]).join('\n') ?? '';

    expect(text).toContain('128 MiB');
    expect(text).toContain('256 MiB');
    expect(text).toContain('OoxmlResourceLimitError');
    expect(text).toContain('getResourceMetrics()');
    expect(text).toContain('cannot be recovered reliably after the trap');
  });

  it('provides executable-shaped examples and renders them as highlighted code', () => {
    const examples = announcement?.sections.flatMap((section) => section.examples ?? []) ?? [];
    expect(examples.map((example) => example.title)).toEqual(expect.arrayContaining([
      'Show a specific preview error',
      'Collect metrics without console output',
      'Before',
      'After',
    ]));
    expect(articlePage).toContain('<Code code={example.code} lang="ts" themes={codeThemes}');
  });

  it('keeps code examples inside the article column on mobile', () => {
    expect(articlePage).toContain('aside, .article-body { min-width: 0; overflow-wrap: anywhere; }');
    expect(articlePage).toContain('grid-template-columns: minmax(0, 1fr);');
    expect(articlePage).toContain('overflow-x: auto;');
  });
});
