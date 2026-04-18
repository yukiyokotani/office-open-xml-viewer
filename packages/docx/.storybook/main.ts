import type { StorybookConfig } from '@storybook/html-vite';

const config: StorybookConfig = {
  stories: [
    '../src/**/*.mdx',
    '../src/**/*.stories.@(js|jsx|mjs|ts|tsx)',
  ],
  addons: [
    '@storybook/addon-docs',
  ],
  framework: '@storybook/html-vite',
  staticDirs: ['../public'],
};
export default config;
