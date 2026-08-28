import { describe, expect, it } from 'vitest';
import {
  apiReference,
  formatRenderModeGuidance,
  optionalRenderers,
} from './api-reference.js';

describe('official-site API reference', () => {
  it('documents every optional renderer entry and its shared contract', () => {
    expect(optionalRenderers.map(({ entry, exportName, contract }) => ({ entry, exportName, contract })))
      .toEqual([
        {
          entry: '@silurus/ooxml/math',
          exportName: 'math',
          contract: 'MathRenderer',
        },
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

  it('documents the complete progressive DOCX contract on every loading surface', () => {
    for (const apiClass of apiReference.docx) {
      const options = apiClass.options ?? [];
      const progressive = options.find(({ name }) => name === 'progressiveLayout');
      expect(progressive?.def, apiClass.name).toBe('false');
      expect(progressive?.desc, apiClass.name).toContain('pages available so far');
      expect(progressive?.detailsHref, apiClass.name).toBe('/docx#progressive-layout');
      expect(options.map(({ name }) => name), apiClass.name).toEqual(expect.arrayContaining([
        'showTrackedChanges',
        'currentDate',
        'sliceLayout',
        'onLayoutProgress',
        'onLayoutPartial',
        'onLayoutComplete',
      ]));
      expect(options.find(({ name }) => name === 'onLayoutPartial')?.type, apiClass.name)
        .toContain('availableUnits: number; totalUnits?: number; exact: boolean');
      expect(options.find(({ name }) => name === 'onLayoutComplete')?.type, apiClass.name)
        .toBe('(error?: unknown) => void');
      expect(options.find(({ name }) => name === 'currentDate')?.desc, apiClass.name)
        .toContain('layout variant');

      const pageCount = apiClass.methods.find(({ sig }) => sig === 'get pageCount(): number');
      expect(pageCount?.desc, apiClass.name).toContain('available so far');
      expect(apiClass.methods.find(({ sig }) => sig === 'get layoutComplete(): boolean')?.desc, apiClass.name)
        .toContain('remains false');
      expect(apiClass.methods.find(({ sig }) => sig === 'waitUntilLayoutComplete(): Promise<void>')?.desc, apiClass.name)
        .toContain('authoritative full layout');
    }

    const viewerCallback = apiReference.docx[0].options?.find(({ name }) => name === 'onPageChange');
    expect(viewerCallback?.desc).toContain('page-count publication changes total');
    const scrollCallback = apiReference.docx[2].options?.find(({ name }) => name === 'onVisiblePageChange');
    expect(scrollCallback?.desc).toContain('even if the same page remains visible');
  });

  it('documents a symmetric PPTX lifecycle with a stable final slide extent', () => {
    for (const apiClass of apiReference.pptx) {
      const options = apiClass.options ?? [];
      const progressive = options.find(({ name }) => name === 'progressiveLayout');
      expect(progressive?.def, apiClass.name).toBe('false');
      expect(progressive?.desc, apiClass.name).toContain('slideCount');
      expect(progressive?.detailsHref, apiClass.name).toBe('/pptx#progressive-layout');
      expect(options.map(({ name }) => name), apiClass.name).toEqual(expect.arrayContaining([
        'onLayoutProgress',
        'onLayoutPartial',
        'onLayoutComplete',
      ]));
      expect(options.find(({ name }) => name === 'onLayoutPartial')?.type, apiClass.name)
        .toContain('availableUnits: number; totalUnits?: number; exact: boolean');
      expect(apiClass.methods.find(({ sig }) => sig === 'get availableSlideCount(): number'), apiClass.name)
        .toBeDefined();
      expect(apiClass.methods.find(({ sig }) => sig === 'get layoutComplete(): boolean')?.desc, apiClass.name)
        .toContain('remains false');
      expect(apiClass.methods.find(({ sig }) => sig === 'waitUntilLayoutComplete(): Promise<void>'), apiClass.name)
        .toBeDefined();
    }

    const viewerCallback = apiReference.pptx[0].options?.find(({ name }) => name === 'onSlideChange');
    expect(viewerCallback?.type).toContain('layoutComplete: boolean');
    const scrollCallback = apiReference.pptx[2].options?.find(({ name }) => name === 'onVisibleSlideChange');
    expect(scrollCallback?.desc).toContain('same slide remains visible');
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
