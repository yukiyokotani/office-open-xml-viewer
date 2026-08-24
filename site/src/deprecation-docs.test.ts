import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { apiReference } from './lib/api-reference.js';
import { announcements } from './lib/announcements.js';

const deprecatedPage = new URL('./pages/deprecations.astro', import.meta.url);
const announcementPage = readFileSync(
  new URL('./pages/announcements/[slug].astro', import.meta.url),
  'utf8',
);
const selectionContextMigration = readFileSync(
  new URL('../../docs/migration-selection-context-0.77.md', import.meta.url),
  'utf8',
);

describe('public migration documentation', () => {
  it('keeps release migration history in announcements instead of an API page', () => {
    expect(existsSync(deprecatedPage)).toBe(false);
    const migration = announcements.find(({ slug }) => slug === 'v077-migration-guide');
    expect(migration).toBeDefined();
    expect(migration?.sections.every(({ modules, rationale }) =>
      Boolean(modules?.length && rationale))).toBe(true);
    const text = JSON.stringify(migration);
    for (const api of [
      'onSelectionChange',
      'PptxElementSelectionContext',
      'readDocxSelectionContext()',
      'Viewer load() and onError',
      'showTrackChanges',
      'XlsxRenderViewportOptions',
      'OoxmlErrorStage',
    ]) {
      expect(text).toContain(api);
    }
    expect(text).toContain('Why these changes ship together');
    expect(announcementPage).not.toContain('Who needs to act');
    expect(announcementPage).toContain('Affected modules');
    expect(announcementPage).not.toContain('Applies to ·');
    expect(announcementPage).toContain('Table of contents');
  });

  it('documents the selection type migrations in the 0.77 announcement', () => {
    const text = JSON.stringify(announcements.find(({ slug }) => slug === 'v077-migration-guide'));
    expect(text).toContain('PptxElementContext');
    expect(text).toContain('DocxSelectionContext');
    expect(text).toContain('XlsxSelectionContext');
    expect(selectionContextMigration).toContain("context?.kind === 'text'");
    expect(selectionContextMigration).toContain("context?.kind === 'range'");
    expect(selectionContextMigration).toContain('DocxTextSelectionContext');
    expect(selectionContextMigration).toContain('XlsxRangeSelectionContext');
  });

  it('documents maxZipEntryBytes directly without linking API details to a release note', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const option = apiClass.options?.find(({ name }) => name === 'maxZipEntryBytes');
        expect(option?.desc, apiClass.name).toContain('scheduled for removal in a future breaking release');
        expect(option?.detailsHref, apiClass.name).toBeUndefined();
      }
    }
  });
});
