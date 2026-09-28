/** The webview does not expose DOCX's opt-in bundled-font setting. Its IIFE
 * build would otherwise absorb all four optional font data URLs even though
 * the setting is never selected. Keep the published library's lazy asset
 * module out of this consumer without changing the DOCX package build. */
export const optionalDocxFontStub = {
  name: 'optional-docx-font-stub',
  setup(build) {
    build.onResolve({ filter: /^(?:\.\/)?(?:assets\/carlito\/urls\.js|urls-[\w-]+\.js)$/ }, (args) => {
      if (!args.importer.replaceAll('\\', '/').includes('/packages/docx/')) return null;
      return { path: args.path, namespace: 'stub-optional-docx-fonts' };
    });
    build.onLoad({ filter: /.*/, namespace: 'stub-optional-docx-fonts' }, () => ({
      contents: 'export const CARLITO_URLS = undefined;',
      loader: 'js',
    }));
  },
};
