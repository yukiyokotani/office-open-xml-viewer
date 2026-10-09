import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import type { DocParagraph, DocxDocumentModel } from '../types.js';
import { imageResourceKey, mathResourceKey } from './source-key.js';
import {
  chartPaintResourceKey,
} from './production-paint-resources.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { documentImageMetadataRecords } from './resources.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';

function chartModel(): ChartModel {
  return {
    chartType: 'line', title: null, categories: [], series: [], varyColors: false,
    showDataLabels: false, valMin: null, valMax: null, catAxisTitle: null,
    valAxisTitle: null, catAxisHidden: false, valAxisHidden: false,
    catAxisLineHidden: false, valAxisLineHidden: false, plotAreaBg: null,
    chartBg: null, showLegend: false, legendPos: null, catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'cross', catAxisMajorTickMark: 'cross',
    titleFontSizeHpt: null, titleFontColor: null, titleFontFace: null,
    catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null, subtotalIndices: [],
  } as ChartModel;
}

function paragraph(runs: DocParagraph['runs']): DocParagraph {
  return {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
    numbering: {
      numId: 1, level: 0, format: 'bullet', text: '', indentLeft: 0, tab: 18,
      suff: 'tab', picBulletImagePath: 'word/media/bullet.gif',
      picBulletMimeType: 'image/gif', picBulletWidthPt: 6, picBulletHeightPt: 7,
    },
    runs,
  } as DocParagraph;
}

const section = {
  pageWidth: 612, pageHeight: 792,
  marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
  headerDistance: 36, footerDistance: 36,
} as DocxDocumentModel['section'];

function documentModel(): DocxDocumentModel {
  return {
    section,
    headers: {
      default: {
        body: [paragraph([{
          type: 'image', imagePath: 'word/media/header.png', mimeType: 'image/png',
          widthPt: 12, heightPt: 8,
        }])],
      },
    },
    footers: {},
    body: [paragraph([{
      type: 'image', imagePath: 'word/media/image.png', svgImagePath: 'word/media/image.svg',
      mimeType: 'image/png', widthPt: 40, heightPt: 30,
      srcRect: { l: .1, t: .2, r: .1, b: 0 }, rotation: 15, flipH: true,
      flipV: true, alpha: .5, colorReplaceFrom: 'FFFFFF',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    }, {
      type: 'chart', chart: chartModel(), widthPt: 120, heightPt: 80, anchor: false,
    }, {
      type: 'math', nodes: [], display: false, fontSize: 10,
    }, {
      type: 'shape', widthPt: 50, heightPt: 30, anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: false, zOrder: 0, subpaths: [],
      fill: {
        fillType: 'image', imagePath: 'word/media/shape-fill.png', mimeType: 'image/png',
        svgImagePath: 'word/media/shape-fill.svg',
        srcRect: { l: .2, t: 0, r: .1, b: .05 },
        alpha: .6, duotone: { clr1: '112233', clr2: 'DDEEFF' },
      }, stroke: null, textBlocks: [{
        text: '', fontSizePt: 10, alignment: 'left', imagePath: 'word/media/textbox.png',
        mimeType: 'image/png', imageWidthPt: 9, imageHeightPt: 5,
      }],
    }])],
  } as DocxDocumentModel;
}

describe('production paint resources', () => {
  it('preserves picture-bullet carrier transforms without changing its explicit box', () => {
    const owner = paragraph([]);
    const transform = {
      srcRect: { l: .25, t: .125, r: 0, b: .25 },
      rotation: 90, flipH: true, flipV: false,
    };
    owner.numbering!.picBulletTransform = transform;
    const doc = { section, headers: {}, footers: {}, body: [owner] } as DocxDocumentModel;
    const registry = layoutSourceStore(doc).paintResources;
    const key = imageResourceKey(
      { story: 'body', storyInstance: 'body', path: [0] }, 'word/media/bullet.gif',
    );
    expect(registry.resolve(key, 'picture-bullet')).toMatchObject({
      intrinsicSize: { widthPt: 6, heightPt: 7 },
      srcRect: { l: .25, t: .125, r: 0, b: .25 },
      rotation: 90, flipH: true, flipV: false,
    });
    transform.srcRect.l = .5;
    expect(registry.resolve(key, 'picture-bullet')).toMatchObject({
      srcRect: { l: .25 },
    });
  });

  it('builds clone-safe descriptors with the same structural keys as retained layout', () => {
    const registry = layoutSourceStore(documentModel()).paintResources;
    const body = { story: 'body' as const, storyInstance: 'body', path: [0] };
    const imageKey = imageResourceKey({ ...body, path: [0, 0] }, 'word/media/image.png');
    const chartKey = chartPaintResourceKey({ ...body, path: [0, 1] });
    const mathKey = mathResourceKey({ ...body, path: [0, 2] }, 'inline');
    const bulletKey = imageResourceKey(body, 'word/media/bullet.gif');
    const textBoxKey = imageResourceKey({
      story: 'textbox', storyInstance: 'body:body:0.3', path: [0, 0],
    }, 'word/media/textbox.png');
    const shapeFillKey = imageResourceKey(
      { ...body, path: [0, 3] }, 'word/media/shape-fill.png',
    );
    const headerImageKey = imageResourceKey({
      story: 'header', storyInstance: 'default', path: [0, 0],
    }, 'word/media/header.png');
    const headerBulletKey = imageResourceKey({
      story: 'header', storyInstance: 'default', path: [0],
    }, 'word/media/bullet.gif');

    expect(registry.keys).toEqual([
      bulletKey, headerBulletKey, headerImageKey, imageKey, shapeFillKey, textBoxKey, chartKey, mathKey,
    ].sort((left, right) => left.localeCompare(right)));
    expect(registry.resolve(imageKey, 'image')).toMatchObject({
      partPath: 'word/media/image.png', svgImagePath: 'word/media/image.svg',
      intrinsicSize: { widthPt: 40, heightPt: 30 },
      srcRect: { l: .1, t: .2, r: .1, b: 0 }, rotation: 15,
      flipH: true, flipV: true, alpha: .5, colorReplaceFrom: 'FFFFFF',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    });
    expect(registry.resolve(chartKey, 'chart').model.chartType).toBe('line');
    expect(registry.resolve(shapeFillKey, 'image')).toMatchObject({
      partPath: 'word/media/shape-fill.png',
      intrinsicSize: { widthPt: 50, heightPt: 30 },
      svgImagePath: 'word/media/shape-fill.svg',
      srcRect: { l: .2, t: 0, r: .1, b: .05 },
      alpha: .6, duotone: { clr1: '112233', clr2: 'DDEEFF' },
    });
    expect(registry.resolve(bulletKey, 'picture-bullet')).toMatchObject({
      partPath: 'word/media/bullet.gif', intrinsicSize: { widthPt: 6, heightPt: 7 },
    });
    expect(structuredClone(registry.descriptors)).toEqual(registry.descriptors);
  });

  it('addresses a valid image by authored order after an unavailable drawing', () => {
    const doc = {
      section,
      headers: {},
      footers: {},
      body: [paragraph([{
        type: 'unavailableDrawing',
        resourceKind: 'image',
        widthPt: 12,
        heightPt: 8,
      } as never, {
        type: 'image',
        imagePath: 'word/media/available.png',
        mimeType: 'image/png',
        widthPt: 24,
        heightPt: 16,
      }])],
    } as DocxDocumentModel;
    const normalized = normalizeInternalDocumentModel(doc);
    const projections = normalized.bodyModelGateway.acquisitionInputs;
    const metadata = documentImageMetadataRecords(
      normalized.document,
      undefined,
      projections,
    );
    const registry = layoutSourceStore(normalized.document).paintResources;
    const authoredKey = imageResourceKey(
      { story: 'body', storyInstance: 'body', path: [0, 1] },
      'word/media/available.png',
    );

    expect(registry.resolve(authoredKey, 'image')).toMatchObject({
      partPath: 'word/media/available.png',
      intrinsicSize: { widthPt: 24, heightPt: 16 },
    });
  });

  it('projects image and chart resources from a complete rich text box nested table', () => {
    const nestedParagraph = paragraph([{
      type: 'image', imagePath: 'word/media/rich.png', mimeType: 'image/png',
      widthPt: 17, heightPt: 13,
    }, {
      type: 'chart', chart: chartModel(), widthPt: 60, heightPt: 40, anchor: false,
    }]);
    nestedParagraph.numbering = null;
    const shape = {
      type: 'shape', widthPt: 50, heightPt: 30, anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: false, zOrder: 0, subpaths: [],
      fill: null, stroke: null,
      textBlocks: [{ text: 'compatibility sentinel', fontSizePt: 10, alignment: 'left' }],
      textBoxContent: [{
        type: 'table', colWidths: [80], rows: [{ cells: [{
          content: [nestedParagraph], colSpan: 1, vMerge: false, vAlign: 'top',
        }] }],
      }],
    };
    const store = layoutSourceStore({
      section, headers: {}, footers: {}, body: [paragraph([shape as never])],
    } as DocxDocumentModel);
    const storyInstance = 'body:body:0.0';
    const imageSource = {
      story: 'textbox' as const, storyInstance, path: [0, 0, 0, 0, 0],
    };
    const chartSource = { ...imageSource, path: [0, 0, 0, 0, 1] };
    const imageKey = imageResourceKey(imageSource, 'word/media/rich.png');
    const chartKey = chartPaintResourceKey(chartSource);
    const canonicalParagraph = store.blocks.resolve({
      story: 'body', storyInstance: 'body', path: [0],
    });
    if (canonicalParagraph.type !== 'paragraph') throw new Error('Expected paragraph');
    const canonicalShape = canonicalParagraph.runs[0];

    expect(canonicalShape?.type).toBe('shape');
    expect(canonicalShape).not.toHaveProperty('textBlocks');
    expect(canonicalShape).not.toHaveProperty('textBoxContent');
    expect(store.paintResources.resolve(imageKey, 'image')).toMatchObject({
      intrinsicSize: { widthPt: 17, heightPt: 13 },
    });
    expect(store.paintResources.resolve(chartKey, 'chart').model.chartType).toBe('line');
  });
});
