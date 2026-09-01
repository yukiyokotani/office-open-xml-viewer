import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { advancedChartRenderers, fullRenderers } from './webview/advancedChartRenderers';

const HERE = dirname(fileURLToPath(import.meta.url));

describe('VS Code webview advanced chart renderers', () => {
  it('bundles every optional chart renderer used by the three Office viewers', () => {
    expect(Object.keys(advancedChartRenderers).sort()).toEqual([
      'chartEx',
      'regionMap',
      'threeD',
    ]);
    expect(typeof advancedChartRenderers.chartEx.render).toBe('function');
    expect(typeof advancedChartRenderers.threeD.render).toBe('function');
    expect(typeof advancedChartRenderers.regionMap.render).toBe('function');
    expect(Object.keys(fullRenderers).sort()).toEqual([
      'chartEx',
      'regionMap',
      'threeD',
      'tiff',
    ]);
    expect(typeof fullRenderers.tiff.render).toBe('function');
  });

  it('injects the same renderer set into DOCX, XLSX, and PPTX viewers', () => {
    const bootstrap = readFileSync(resolve(HERE, 'webview/bootstrap.ts'), 'utf8');
    expect(bootstrap.match(/\.\.\.fullRenderers/g)).toHaveLength(3);
  });
});
