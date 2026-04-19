import type { Meta, StoryObj } from '@storybook/html';
import { buildViewerUI } from './PptxViewer.stories';

type Args = { width: number };

const meta: Meta<Args> = {
  title: 'PptxViewer/Demo',
  argTypes: {
    width: {
      control: { type: 'range', min: 400, max: 1600, step: 40 },
      description: 'Canvas render width (px)',
    },
  },
  args: { width: 960 },
};
export default meta;
type Story = StoryObj<Args>;

export const Demo: Story = {
  name: 'demo.pptx',
  render(args) {
    const { root } = buildViewerUI(args, `${import.meta.env.BASE_URL}pptx/demo/sample-1.pptx`);
    return root;
  },
};
