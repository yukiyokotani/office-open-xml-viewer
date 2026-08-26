import { join } from 'node:path';

import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import { getSlideMutationId } from '../../src/adapters/pptx-json-adapter';
import { ELEMENT_ORIGINS } from '../../src/domain/element-origin';
import type { ElementRef } from '../../src/domain/mutation';
import { AddElementMutation } from '../../src/mutations/add-element';
import { RemoveElementMutation } from '../../src/mutations/remove-element';
import { toOfficeCliBatch } from '../../src/transport/officecli/officecli-translator';
import { shape } from '../fixtures/presentation';
import {
  addShape,
  addSlide,
  assertLiveOfficeCli,
  createDeck,
  createLiveWorkspace,
  destroyLiveWorkspace,
  elementIdOfPath,
  flushDeck,
  getNode,
  normalizeShapeFormat,
  parseDeck,
  refForElementId,
  runBatch,
  tryGetNode,
} from './harness';

const BASE_SHAPE_PROPS = {
  x: '914400emu',
  y: '457200emu',
  width: '1828800emu',
  height: '914400emu',
} as const;

function newElementRef(presentation: Presentation, elementId: string): ElementRef {
  return {
    origin: ELEMENT_ORIGINS.SLIDE,
    slideId: getSlideMutationId(presentation.slides[0]),
    elementId,
  };
}

function shapeIdsInTree(pptxPath: string): string[] {
  // Preset-less text shapes read back as type "textbox" and placeholders as
  // their own type, but every spTree shape shares the `shape[@id=N]` path
  // family, which is what slide-tree ordering assertions care about.
  return getNode(pptxPath, '/slide[1]')
    .children
    .filter((child) => /\/shape\[@id=\d+\]$/.test(child.path))
    .map((child) => elementIdOfPath(child.path));
}

describe('AddElementMutation × OfficeCLI 真实执行', () => {
  let dir: string;
  const openedDecks: string[] = [];

  beforeAll(() => {
    assertLiveOfficeCli();
    dir = createLiveWorkspace('add-element');
  });

  afterAll(() => destroyLiveWorkspace(dir, openedDecks));

  function newDeck(name: string): string {
    const pptxPath = join(dir, name);
    openedDecks.push(pptxPath);
    createDeck(pptxPath);
    return pptxPath;
  }

  it('AddElement 生成的 add 命令能以指定 id、几何、变换与文本真实创建 shape', () => {
    const pptxPath = newDeck('create.pptx');
    addSlide(pptxPath);
    addShape(pptxPath, '/slide[1]', { text: 'existing', ...BASE_SHAPE_PROPS });
    flushDeck(pptxPath);
    const presentation = parseDeck(pptxPath);

    const element = shape('7777', '新增的形状文本', {
      geometry: 'ellipse',
      x: 123457,
      y: 765431,
      width: 2000003,
      height: 1000001,
      rotation: 15.5,
      flipH: true,
      name: 'Restored Shape',
    });
    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-add-1',
      mutations: [new AddElementMutation({
        target: newElementRef(presentation, '7777'),
        element,
        presentationElementIndex: 1,
      })],
    }));

    const node = getNode(pptxPath, '/slide[1]/shape[@id=7777]');
    expect(node.text).toBe('新增的形状文本');
    expect(node.format.geometry).toBe('ellipse');
    expect(node.format.name).toBe('Restored Shape');
    expect(normalizeShapeFormat(node.format)).toEqual({
      x: 123457,
      y: 765431,
      width: 2000003,
      height: 1000001,
      rotation: 15.5,
      flipH: true,
      flipV: false,
    });
  });

  it('zorder 属性能把 shape 插入 slide 树的指定位置（纯 slide 元素 deck）', () => {
    const pptxPath = newDeck('zorder.pptx');
    addSlide(pptxPath);
    const firstPath = addShape(pptxPath, '/slide[1]', { text: 'first', ...BASE_SHAPE_PROPS });
    const lastPath = addShape(pptxPath, '/slide[1]', { text: 'last', ...BASE_SHAPE_PROPS });
    flushDeck(pptxPath);
    const presentation = parseDeck(pptxPath);

    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-add-zorder-1',
      mutations: [new AddElementMutation({
        target: newElementRef(presentation, '8888'),
        element: shape('8888', 'middle', { ...shapeOverrides() }),
        presentationElementIndex: 1,
      })],
    }));

    expect(shapeIdsInTree(pptxPath)).toEqual([
      elementIdOfPath(firstPath),
      '8888',
      elementIdOfPath(lastPath),
    ]);
  });

  it('zorder 属性在带 layout placeholder 的 deck 上不发生序号偏移', () => {
    const pptxPath = newDeck('placeholder.pptx');
    addSlide(pptxPath, { layout: 'Title and Content', title: '占位符标题' });
    const extraPath = addShape(pptxPath, '/slide[1]', { text: 'extra', ...BASE_SHAPE_PROPS });
    flushDeck(pptxPath);
    const presentation = parseDeck(pptxPath);

    // 前置条件：标题占位符落在 slide spTree 内、且位于额外 shape 之前
    //（两者均为 origin:'slide'）。
    expect(presentation.slides[0].elementSources).toEqual([
      { origin: 'slide' },
      { origin: 'slide' },
    ]);
    const placeholderId = (presentation.slides[0].elements[0] as { id?: string }).id;

    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-add-placeholder-1',
      mutations: [new AddElementMutation({
        target: newElementRef(presentation, '9999'),
        element: shape('9999', 'between', { ...shapeOverrides() }),
        presentationElementIndex: 1,
      })],
    }));

    expect(shapeIdsInTree(pptxPath)).toEqual([
      placeholderId,
      '9999',
      elementIdOfPath(extraPath),
    ]);
  });

  it('remove 后由 inverse() 生成的 AddElement 能真实恢复 shape（含填充、描边、阴影样式），且乐观状态与重新解析的 PPTX 一致', () => {
    const pptxPath = newDeck('undo.pptx');
    addSlide(pptxPath);
    addShape(pptxPath, '/slide[1]', { text: 'keep', ...BASE_SHAPE_PROPS });
    const victimPath = addShape(pptxPath, '/slide[1]', {
      text: 'victim 第一行\n第二行',
      x: '123457emu',
      y: '765431emu',
      width: '2000003emu',
      height: '1000001emu',
      rotation: '30.5',
      flipH: 'true',
      preset: 'roundRect',
      adj: 'adj:val 40000',
      fill: '112233',
      opacity: '0.5',
      line: '445566:1.5',
      lineOpacity: '0.25',
      lineDash: 'lgDashDot',
      lineCap: 'round',
      cmpd: 'dbl',
      headEnd: 'triangle',
      tailEnd: 'oval',
      shadow: '808080-4-45-3-40',
    });
    flushDeck(pptxPath);
    const original = parseDeck(pptxPath);
    const ref = refForElementId(original, elementIdOfPath(victimPath));
    // Precondition: the authored corner-radius handle survives parsing, so
    // this undo also exercises the documented adj degradation.
    expect(original.slides[0].elements[1]).toMatchObject({
      geometry: 'roundRect',
      adj: 40000,
    });

    // Undo relies on the pre-removal snapshot: inverse() must be captured
    // before the removal mutates any state, exactly as the history stack does.
    const remove = new RemoveElementMutation({ target: ref });
    const undo = remove.inverse(original);
    if (!undo) throw new TypeError('A removed shape must remain restorable');
    const afterRemove = remove.apply(original).presentation;

    runBatch(pptxPath, toOfficeCliBatch(original, {
      id: 'live-undo-remove',
      mutations: [remove],
    }));
    expect(tryGetNode(pptxPath, victimPath)).toBeUndefined();

    runBatch(pptxPath, toOfficeCliBatch(afterRemove, {
      id: 'live-undo-restore',
      mutations: [undo],
    }));

    const optimistic = undo.apply(afterRemove).presentation;
    const authoritative = parseDeck(pptxPath);
    // The adjust handle degrades to the preset default by design; everything
    // else (fill/opacity/line/shadow/text) restores faithfully, and the
    // optimistic state agrees with the file about the degradation.
    expect(authoritative.slides[0].elements[1]).toMatchObject({
      geometry: 'roundRect',
      adj: null,
    });
    expect(authoritative.slides[0].elements).toEqual(optimistic.slides[0].elements);
    expect(authoritative.slides[0].elementSources)
      .toEqual(optimistic.slides[0].elementSources);
  });
});

function shapeOverrides() {
  return {
    x: 914400,
    y: 457200,
    width: 1828800,
    height: 914400,
  };
}
