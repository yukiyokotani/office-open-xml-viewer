import assert from 'node:assert/strict';
import test from 'node:test';
import { stableCanvasRender } from './stable-canvas-render.mjs';

test('waits for fonts around the discovery paint and captures after a repaint', async () => {
  const events = [];
  let readyRead = 0;
  const fonts = {
    get ready() {
      const phase = ++readyRead;
      return Promise.resolve().then(() => events.push(`fonts:${phase}`));
    },
  };

  await stableCanvasRender({
    fonts,
    render: async () => events.push('render'),
    nextFrame: async () => events.push('frame'),
  });

  assert.deepEqual(events, [
    'fonts:1',
    'render',
    'fonts:2',
    'render',
    'frame',
    'frame',
  ]);
});

test('keeps the same repaint and frame contract when FontFaceSet is unavailable', async () => {
  const events = [];
  await stableCanvasRender({
    fonts: null,
    render: async () => events.push('render'),
    nextFrame: async () => events.push('frame'),
  });
  assert.deepEqual(events, ['render', 'render', 'frame', 'frame']);
});
