import { describe, expect, it } from 'vitest';
import {
  apiReference,
  formatRenderModeGuidance,
  optionalChartRenderers,
} from './api-reference.js';

describe('official-site API reference', () => {
  it('documents both optional chart renderer entries and their shared contracts', () => {
    expect(optionalChartRenderers.map(({ entry, exportName, contract }) => ({ entry, exportName, contract })))
      .toEqual([
        {
          entry: '@silurus/ooxml/chart-ex',
          exportName: 'chartEx',
          contract: 'ChartExRenderer',
        },
        {
          entry: '@silurus/ooxml/three-d',
          exportName: 'threeD',
          contract: 'ChartThreeDRenderer',
        },
        {
          entry: '@silurus/ooxml/region-map',
          exportName: 'regionMap',
          contract: 'ChartRegionMapRenderer',
        },
      ]);
  });

  it('documents optional chart injection on every format Viewer and engine', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const options = apiClass.options ?? [];
        expect(options.find(({ name }) => name === 'threeD')?.type, apiClass.name)
          .toBe('ChartThreeDRenderer');
        expect(options.find(({ name }) => name === 'regionMap')?.type, apiClass.name)
          .toBe('ChartRegionMapRenderer');
        expect(options.find(({ name }) => name === 'chartEx')?.type, apiClass.name)
          .toBe('ChartExRenderer');
      }
    }
  });

  it('gives every format a current main/worker choice and feature-parity guide', () => {
    for (const [format, classes] of Object.entries(apiReference)) {
      const guidance = formatRenderModeGuidance[format as keyof typeof formatRenderModeGuidance];
      expect(guidance, format).toContain('both modes');
      expect(guidance, format).toContain('Worker mode');
      expect(guidance, format).toContain('ChartEx');
      for (const apiClass of classes) {
        const mode = apiClass.options?.find(({ name }) => name === 'mode');
        expect(mode?.def, apiClass.name).toBe("'main'");
        expect(mode?.desc, apiClass.name).toContain("Use 'main'");
        expect(mode?.desc, apiClass.name).toContain("Use 'worker'");
        expect(mode?.desc, apiClass.name).toContain('larger');
        expect(mode?.desc, apiClass.name).toMatch(/built-in/i);
        expect(mode?.desc, apiClass.name).toContain('ChartEx');
      }
    }
    expect(formatRenderModeGuidance.docx).toContain('automatically use main mode');
    expect(formatRenderModeGuidance.docx).toContain("document's mode");
  });

  it('keeps worker bitmap descriptions synchronized with optional ChartEx rendering', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        for (const method of apiClass.methods.filter(({ sig }) => sig.includes('ToBitmap('))) {
          expect(method.desc, `${apiClass.name}: ${method.sig}`).toContain('ChartEx');
        }
      }
    }
  });

  it('documents the shared resource controls on every browser API class', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const optionNames = apiClass.options?.map(({ name }) => name) ?? [];
        expect(optionNames, apiClass.name).toEqual(expect.arrayContaining([
          'resourceLimits',
          'onResourceMetrics',
          'debug',
        ]));
      }
    }
  });

  it('keeps every semantic emphasis synchronized with its description', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        for (const item of [...(apiClass.options ?? []), ...apiClass.methods]) {
          if (item.emphasis) {
            expect(item.desc, `${apiClass.name}: ${'name' in item ? item.name : item.sig}`)
              .toContain(item.emphasis);
          }
        }
      }
    }
  });

  it('documents the Viewer error-delivery contract and typed resource failures', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes.filter(({ name }) => name.endsWith('Viewer'))) {
        const onError = apiClass.options?.find(({ name }) => name === 'onError');
        expect(onError, apiClass.name).toBeDefined();
        expect(onError?.desc, apiClass.name).toContain('load(), navigation, and other awaitable operations reject');
        expect(onError?.desc, apiClass.name).toContain('the same failure is never delivered twice');
        expect(onError?.desc, apiClass.name).toContain('OoxmlResourceLimitError');
        expect(onError?.desc, apiClass.name).toContain('OoxmlDecodedImageLimitError');
        expect(onError?.desc, apiClass.name).toContain('message text is not a stable discriminator');
        expect(onError?.detailsHref, apiClass.name).toBe('/errors#delivery');
      }
    }
  });

  it('links resource-limit options to their typed error fields', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const resourceLimits = apiClass.options?.find(({ name }) => name === 'resourceLimits');
        expect(resourceLimits?.detailsHref, apiClass.name).toBe('/errors#ooxml-resource-limit-error');
      }
    }
  });

  it('documents password on every self-loading Viewer and engine', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes) {
        const password = apiClass.options?.find(({ name }) => name === 'password');
        expect(password?.type, apiClass.name).toBe('string');
        expect(password?.desc, apiClass.name).toContain('borrowed');
      }
    }
  });

  it('documents the common native context-menu handoff on every Viewer', () => {
    for (const classes of Object.values(apiReference)) {
      for (const apiClass of classes.filter(({ name }) => name.endsWith('Viewer'))) {
        const option = apiClass.options?.find(({ name }) => name === 'onContextMenu');
        expect(option?.type, apiClass.name).toContain('ViewerContextMenuEvent<');
        expect(option?.desc, apiClass.name).toContain('originalEvent.preventDefault()');
        expect(option?.desc, apiClass.name).toContain('getContext()');
        expect(option?.desc, apiClass.name).toContain('native browser behavior unchanged');
      }
    }
  });
});
