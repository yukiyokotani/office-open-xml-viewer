import { join } from 'node:path';

import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import type { Presentation } from '@maxgent/ooxml/pptx';

import { RemoveElementMutation } from '../../src/mutations/remove-element';
import { toOfficeCliBatch } from '../../src/transport/officecli/officecli-translator';
import {
  addShape,
  addSlide,
  addSlideElement,
  assertLiveOfficeCli,
  createDeck,
  createLiveWorkspace,
  destroyLiveWorkspace,
  elementIdOfPath,
  flushDeck,
  getNode,
  parseDeck,
  refForElementId,
  runBatch,
  tryGetNode,
} from './harness';

describe('RemoveElementMutation × OfficeCLI 真实执行', () => {
  let dir: string;
  let pptxPath: string;
  let presentation: Presentation;
  let victimShapePath: string;
  let keptShapePath: string;
  let picturePptxPath: string;
  let picturePresentation: Presentation;
  let picturePath: string;

  beforeAll(() => {
    assertLiveOfficeCli();
    dir = createLiveWorkspace('remove-element');
    pptxPath = join(dir, 'deck.pptx');
    createDeck(pptxPath);
    addSlide(pptxPath);
    victimShapePath = addShape(pptxPath, '/slide[1]', {
      text: 'victim',
      x: '914400emu',
      y: '457200emu',
      width: '1828800emu',
      height: '914400emu',
    });
    keptShapePath = addShape(pptxPath, '/slide[1]', {
      text: 'kept',
      x: '914400emu',
      y: '1828800emu',
      width: '1828800emu',
      height: '914400emu',
    });
    flushDeck(pptxPath);
    presentation = parseDeck(pptxPath);

    picturePptxPath = join(dir, 'picture-deck.pptx');
    createDeck(picturePptxPath);
    addSlide(picturePptxPath);
    picturePath = addSlideElement(picturePptxPath, '/slide[1]', 'picture', {
      src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
      x: '0',
      y: '0',
      width: '914400emu',
      height: '914400emu',
    });
    flushDeck(picturePptxPath);
    picturePresentation = parseDeck(picturePptxPath);
  });

  afterAll(() => destroyLiveWorkspace(dir, [pptxPath, picturePptxPath]));

  it('RemoveElement 生成的 remove 命令能真实删除目标 shape 且不影响同页其他元素', () => {
    const ref = refForElementId(presentation, elementIdOfPath(victimShapePath));

    runBatch(pptxPath, toOfficeCliBatch(presentation, {
      id: 'live-remove-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    }));

    expect(tryGetNode(pptxPath, victimShapePath)).toBeUndefined();
    const kept = getNode(pptxPath, keptShapePath);
    expect(kept.text).toBe('kept');

    // The on-disk model must agree: exactly one shape survives.
    const reparsed = parseDeck(pptxPath);
    expect(reparsed.slides[0].elements).toHaveLength(1);
    expect((reparsed.slides[0].elements[0] as { id?: string }).id)
      .toBe(elementIdOfPath(keptShapePath));
  });

  it('RemoveElement deletes an ordinary picture from its frontend type and id', () => {
    const ref = refForElementId(picturePresentation, elementIdOfPath(picturePath));
    const batch = toOfficeCliBatch(picturePresentation, {
      id: 'live-remove-picture-1',
      mutations: [new RemoveElementMutation({ target: ref })],
    });
    expect(batch.commands).toEqual([{ command: 'remove', path: picturePath }]);

    runBatch(picturePptxPath, batch);

    expect(tryGetNode(picturePptxPath, picturePath)).toBeUndefined();
    expect(parseDeck(picturePptxPath).slides[0].elements).toEqual([]);
  });
});
