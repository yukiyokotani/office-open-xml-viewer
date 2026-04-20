import type { Meta, StoryObj } from '@storybook/html';
import { buildViewerUI } from './XlsxViewer.stories';

type Args = { scale: number };

const meta: Meta<Args> = {
  title: 'XlsxViewer/Examples',
  argTypes: {
    scale: {
      control: { type: 'range', min: 0.25, max: 2, step: 0.05 },
      description: 'Cell/header scale (1 = normal size)',
    },
  },
  args: { scale: 1 },
};
export default meta;
type Story = StoryObj<Args>;

export const Demo: Story = {
  name: 'Demo — single viewer (demo.xlsx)',
  render(args) {
    const { root } = buildViewerUI(args, `${import.meta.env.BASE_URL}xlsx/demo/sample-1.xlsx`);
    return root;
  },
};
