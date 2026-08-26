import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const productionPage = readFileSync(
  new URL('./pages/production.astro', import.meta.url),
  'utf8',
);

describe('format API loading-mode documentation', () => {
  it('documents acquisition ownership and execution mode as separate choices', () => {
    expect(productionPage).toContain('id="rendering-mode"');
    expect(productionPage).toContain('Parsing always runs in a Worker');
    expect(productionPage).toContain('id="ownership"');
    expect(productionPage).toContain('one Viewer own one document');
    expect(productionPage).toContain('several views must share one parse');
    expect(productionPage.indexOf('id="rendering-mode"'))
      .toBeLessThan(productionPage.indexOf('id="ownership"'));
    expect(productionPage).toContain('Destroy every view before the shared document');
  });
});
