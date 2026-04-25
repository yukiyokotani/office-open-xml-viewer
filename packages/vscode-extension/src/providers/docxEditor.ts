import * as vscode from 'vscode';
import { getWebviewHtml } from '../webviewHtml';

export class DocxEditorProvider implements vscode.CustomReadonlyEditorProvider {
  static readonly viewType = 'ooxmlViewer.docxEditor';

  static register(context: vscode.ExtensionContext): vscode.Disposable {
    return vscode.window.registerCustomEditorProvider(
      DocxEditorProvider.viewType,
      new DocxEditorProvider(context),
      { supportsMultipleEditorsPerDocument: true },
    );
  }

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(
    uri: vscode.Uri,
  ): Promise<vscode.CustomDocument> {
    return { uri, dispose: () => undefined };
  }

  async resolveCustomEditor(
    document: vscode.CustomDocument,
    webviewPanel: vscode.WebviewPanel,
  ): Promise<void> {
    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.joinPath(this.context.extensionUri, 'dist'),
      ],
    };

    webviewPanel.webview.html = getWebviewHtml(
      webviewPanel.webview,
      this.context.extensionUri,
      'docx',
    );

    const bytes = await vscode.workspace.fs.readFile(document.uri);

    webviewPanel.webview.onDidReceiveMessage(async (msg) => {
      if (msg.type === 'webview-ready') {
        await webviewPanel.webview.postMessage({
          type: 'ooxml-init',
          fileType: 'docx',
          data: Array.from(bytes),
        });
      } else if (msg.type === 'copy') {
        vscode.env.clipboard.writeText(msg.text);
      }
    });
  }
}
