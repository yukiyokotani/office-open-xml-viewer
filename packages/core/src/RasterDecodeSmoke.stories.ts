import type { Meta, StoryObj } from '@storybook/html-vite';
import { decodeRasterOrMetafile } from './image/raster-or-metafile.js';

const meta: Meta = {
  title: 'Internal/Raster decode smoke',
  parameters: { docs: { disable: true } },
};

export default meta;
type Story = StoryObj;

export const Harness: Story = {
  render: () => {
    Object.assign(globalThis, { __ooxmlDecodeRasterOrMetafile: decodeRasterOrMetafile });
    const ready = document.createElement('div');
    ready.dataset.rasterDecodeReady = 'true';
    return ready;
  },
};
