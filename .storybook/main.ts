import type { StorybookConfig } from '@storybook/html-vite';

const config: StorybookConfig = {
  stories: [
    './**/*.mdx',
    '../packages/*/src/**/*.mdx',
    '../packages/*/src/**/*.stories.@(js|jsx|mjs|ts|tsx)',
  ],
  addons: [
    '@storybook/addon-a11y',
    '@storybook/addon-docs',
    '@chromatic-com/storybook',
  ],
  framework: '@storybook/html-vite',
  staticDirs: [
    { from: '../packages/pptx/public', to: '/pptx' },
    { from: '../packages/xlsx/public', to: '/xlsx' },
    { from: '../packages/docx/public', to: '/docx' },
  ],
  async viteFinal(config) {
    const { default: wasm } = await import('vite-plugin-wasm');
    return {
      ...config,
      base: '/',
      plugins: [...(config.plugins ?? []), wasm()],
      worker: { format: 'es' as const, plugins: () => [wasm()] },
    };
  },
};
export default config;
