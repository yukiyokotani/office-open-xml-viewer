import type { Preview } from '@storybook/html';

const preview: Preview = {
  parameters: {
    backgrounds: {
      default: 'gray',
      values: [
        { name: 'gray', value: '#888888' },
        { name: 'white', value: '#ffffff' },
      ],
    },
  },
};
export default preview;
