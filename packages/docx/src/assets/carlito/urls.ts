// Vite resolves these static module-relative URLs as packaged assets. Keeping
// them behind the optional loader avoids static font imports in hosts that do
// not select the offline Calibri substitute (for example, the VS Code webview).
export const CARLITO_URLS = Object.freeze({
  regular: new URL('./Carlito-Regular.ttf', import.meta.url).href,
  bold: new URL('./Carlito-Bold.ttf', import.meta.url).href,
  italic: new URL('./Carlito-Italic.ttf', import.meta.url).href,
  boldItalic: new URL('./Carlito-BoldItalic.ttf', import.meta.url).href,
});
