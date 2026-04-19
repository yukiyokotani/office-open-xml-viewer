import type { Meta, StoryObj } from '@storybook/html';
import { buildViewerUI } from './DocxViewer.stories';

type Args = { width: number };

const meta: Meta<Args> = {
  title: 'DocxViewer/Demo',
  argTypes: {
    width: {
      control: { type: 'range', min: 400, max: 1200, step: 40 },
      description: 'Canvas render width (px)',
    },
  },
  args: { width: 700 },
};
export default meta;
type Story = StoryObj<Args>;

export const Demo: Story = {
  name: 'demo.docx',
  render(args) {
    const { root } = buildViewerUI(args, `${import.meta.env.BASE_URL}docx/demo/sample-1.docx`);
    return root;
  },
};
