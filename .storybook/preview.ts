import type { Preview } from '@storybook/html-vite';

const preview: Preview = {
  parameters: {
    controls: {
      matchers: {
        color: /(background|color)$/i,
        date: /Date$/i,
      },
    },
    a11y: {
      test: 'todo',
    },
    options: {
      storySort: {
        order: [
          'PptxViewer', ['*', 'Examples', 'PrivateExamples'],
          'DocxViewer', ['*', 'Examples', 'PrivateExamples'],
          'XlsxViewer', ['*', 'Examples', 'PrivateExamples'],
        ],
      },
    },
  },
};

export default preview;
