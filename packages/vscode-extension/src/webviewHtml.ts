import * as vscode from 'vscode';

/**
 * Generate the HTML for the webview panel.
 * The webview script (dist/webview.js) is allowed via the content security policy,
 * and receives the file bytes via a `ooxml-init` message posted from the extension host.
 */
export function getWebviewHtml(
  webview: vscode.Webview,
  extensionUri: vscode.Uri,
  fileType: 'docx' | 'xlsx' | 'pptx',
): string {
  const scriptUri = webview.asWebviewUri(
    vscode.Uri.joinPath(extensionUri, 'dist', 'webview.js'),
  );

  const nonce = getNonce();

  return /* html */ `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta http-equiv="Content-Security-Policy"
    content="default-src 'none';
             img-src ${webview.cspSource} data: blob:;
             media-src ${webview.cspSource} blob:;
             font-src ${webview.cspSource};
             script-src 'nonce-${nonce}' 'wasm-unsafe-eval';
             worker-src data: blob:;
             style-src 'unsafe-inline';
             connect-src ${webview.cspSource} data: blob:;" />
  <title>OOXML Viewer</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    html, body {
      width: 100%;
      height: 100%;
      background: var(--vscode-editor-background);
      color: var(--vscode-foreground);
      font-family: var(--vscode-font-family, sans-serif);
    }
    /* xlsx fills the whole viewport; docx/pptx scroll inside #viewer-root. */
    body.layout-xlsx { overflow: hidden; }
    body.layout-stack { overflow: auto; }
    #viewer-root {
      width: 100%;
      min-height: 100%;
      display: flex;
      justify-content: center;
      align-items: flex-start;
      padding: 16px;
    }
    body.layout-xlsx #viewer-root { padding: 0; height: 100%; }
    #viewer-container { max-width: 100%; width: 100%; }
    body.layout-stack #viewer-container {
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    #status {
      position: fixed;
      inset: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      pointer-events: none;
      z-index: 10;
    }
    #status[data-state="error"] {
      position: static;
      pointer-events: auto;
      color: var(--vscode-errorForeground, #f44747);
      font-size: 13px;
      padding: 8px;
      justify-content: flex-start;
    }
    .spinner {
      width: 28px;
      height: 28px;
      border: 3px solid color-mix(in srgb, var(--vscode-foreground) 20%, transparent);
      border-top-color: var(--vscode-progressBar-background, var(--vscode-foreground));
      border-radius: 50%;
      animation: spin 0.9s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    /* docx / pptx scroll-stack styling */
    .page-stack {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 16px;
      width: 100%;
    }
    .page-wrapper {
      position: relative;
      width: 100%;
      margin: 0 auto;
    }
    .page-canvas {
      display: block;
      width: 100%;
      background: #fff;
      box-shadow: 0 1px 4px rgba(0, 0, 0, 0.35);
    }
    .text-layer {
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      overflow: hidden;
      pointer-events: none;
      user-select: text;
      -webkit-user-select: text;
    }
  </style>
</head>
<body class="${fileType === 'xlsx' ? 'layout-xlsx' : 'layout-stack'}">
  <div id="viewer-root">
    <div id="viewer-container">
      <div id="status"><div class="spinner"></div></div>
    </div>
  </div>
  <script nonce="${nonce}">
    window.__OOXML_FILE_TYPE__ = ${JSON.stringify(fileType)};
  </script>
  <script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
}

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}
