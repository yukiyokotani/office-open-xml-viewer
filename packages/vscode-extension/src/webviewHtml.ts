import * as vscode from 'vscode';
import * as path from 'path';

/**
 * Generate the HTML for the webview panel.
 * The webview script (dist/webview.js) is allowed via the content security policy,
 * and receives the file bytes via a `ooxml-init` message posted from the extension host.
 */
export function getWebviewHtml(
  webview: vscode.Webview,
  extensionUri: vscode.Uri,
  fileType: 'xlsx' | 'docx' | 'pptx',
): string {
  const scriptUri = webview.asWebviewUri(
    vscode.Uri.joinPath(extensionUri, 'dist', 'webview.js'),
  );
  const wasmUri = webview.asWebviewUri(
    vscode.Uri.joinPath(extensionUri, 'dist', `${fileType}_parser_bg.wasm`),
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
             script-src 'nonce-${nonce}';
             style-src 'unsafe-inline';
             connect-src ${webview.cspSource};" />
  <title>OOXML Viewer</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    html, body { width: 100%; height: 100%; overflow: hidden; }
    body { background: var(--vscode-editor-background, #1e1e1e); }
    #viewer-root {
      width: 100%;
      height: 100%;
      overflow: auto;
      display: flex;
      justify-content: center;
      align-items: flex-start;
      padding: 16px;
    }
    #viewer-container {
      max-width: 100%;
    }
    #status {
      color: var(--vscode-foreground, #ccc);
      font-family: var(--vscode-font-family, sans-serif);
      font-size: 13px;
      padding: 8px;
    }
    /* Slide/page navigation bar */
    #nav-bar {
      display: none;
      align-items: center;
      gap: 8px;
      padding: 4px 0 8px;
      font-family: var(--vscode-font-family, sans-serif);
      font-size: 13px;
      color: var(--vscode-foreground, #ccc);
    }
    #nav-bar.visible { display: flex; }
    #nav-bar button {
      background: var(--vscode-button-background, #0e639c);
      color: var(--vscode-button-foreground, #fff);
      border: none;
      border-radius: 2px;
      padding: 2px 10px;
      cursor: pointer;
      font-size: 12px;
    }
    #nav-bar button:disabled {
      opacity: 0.4;
      cursor: default;
    }
  </style>
</head>
<body>
  <div id="viewer-root">
    <div id="viewer-container">
      <div id="nav-bar">
        <button id="prev-btn">← Prev</button>
        <span id="page-info"></span>
        <button id="next-btn">Next →</button>
      </div>
      <div id="status">Loading…</div>
    </div>
  </div>
  <script nonce="${nonce}">
    window.__OOXML_FILE_TYPE__ = ${JSON.stringify(fileType)};
    window.__OOXML_WASM_URL__ = ${JSON.stringify(wasmUri.toString())};
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
