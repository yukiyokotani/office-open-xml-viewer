import * as vscode from 'vscode';
import { getWebviewHtml } from '../webviewHtml';
import { shouldUseGoogleFonts } from '../config';
import {
  type SelectionContextHandle,
  SelectionContextRegistry,
  selectionDocumentIdentity,
} from '../selectionContextRegistry';

/** Live webview panel + everything needed to rebuild its HTML, tracked so the
 *  HTML can be regenerated when the effective Google Fonts flag changes
 *  (CSP differs → full reload needed). */
interface OpenView {
  panel: vscode.WebviewPanel;
  extensionUri: vscode.Uri;
  fileType: 'docx' | 'xlsx' | 'pptx';
  selection: SelectionContextHandle;
  selectionSession: number;
  /**
   * The webview document that was sent `ooxml-init`: the HTML session and the
   * instance id that document generated. VS Code destroys a hidden editor's
   * document and loads the same HTML again when the tab is shown, so one
   * session can have several documents, and each needs its own init.
   */
  initialized: { session: number; instance: string } | null;
}

/** Registry of every open OOXML webview across all three editor providers, so a
 *  single config/trust change can refresh them all at once. */
const openViews = new Set<OpenView>();

/**
 * Rebuild the HTML of every open OOXML webview.
 *
 * Called when `ooxmlViewer.useGoogleFonts` changes or workspace trust is granted:
 * the CSP is baked into the document `<head>`, so toggling Google Fonts requires a
 * full HTML reload rather than a re-render message. Re-setting `webview.html`
 * re-runs the bootstrap, which re-requests the bytes and re-instantiates the viewer
 * with the new flag.
 */
export function refreshAllWebviews(): void {
  for (const view of openViews) {
    renderWebview(view);
  }
}

/** Open the in-preview find popup in the active OOXML custom editor. */
export async function showFindInActiveWebview(): Promise<void> {
  const active = [...openViews].find((view) => view.panel.active);
  if (!active) return;
  await active.panel.webview.postMessage({ type: 'show-find' });
}

/** Set the webview HTML for a panel using the current effective flag. */
function renderWebview(view: OpenView): void {
  view.selectionSession++;
  view.initialized = null;
  view.selection.update(null);
  view.selection.updateView(null);
  view.panel.webview.html = getWebviewHtml(
    view.panel.webview,
    view.extensionUri,
    view.fileType,
    shouldUseGoogleFonts(),
    view.selectionSession,
  );
}

/**
 * Shared implementation for the docx / xlsx / pptx readonly custom editors. The
 * three formats differ only by their `viewType` and `fileType` tag; everything
 * else (CSP-gated HTML, byte hand-off, Google-Fonts wiring) is common.
 */
export abstract class BaseEditorProvider
  implements vscode.CustomReadonlyEditorProvider
{
  protected abstract readonly fileType: 'docx' | 'xlsx' | 'pptx';

  constructor(
    protected readonly context: vscode.ExtensionContext,
    private readonly selectionContexts: SelectionContextRegistry,
  ) {}

  async openCustomDocument(uri: vscode.Uri): Promise<vscode.CustomDocument> {
    return { uri, dispose: () => undefined };
  }

  async resolveCustomEditor(
    document: vscode.CustomDocument,
    webviewPanel: vscode.WebviewPanel,
  ): Promise<void> {
    // Only the bundled assets are webview-loadable. The document's bytes are
    // read here and posted to the webview, so neither its directory nor the
    // document itself is a resource root: `connect-src`/`img-src`
    // `${cspSource}` therefore cannot reach sibling files.
    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, 'dist')],
    };

    const selection = this.selectionContexts.track(
      webviewPanel,
      selectionDocumentIdentity(this.fileType, document.uri),
    );
    const view: OpenView = {
      panel: webviewPanel,
      extensionUri: this.context.extensionUri,
      fileType: this.fileType,
      selection,
      selectionSession: 0,
      initialized: null,
    };
    openViews.add(view);
    let disposed = false;
    webviewPanel.onDidDispose(() => {
      disposed = true;
      openViews.delete(view);
      selection.dispose();
    });

    renderWebview(view);

    webviewPanel.webview.onDidReceiveMessage(async (msg) => {
      if (msg.type === 'webview-ready') {
        if (msg.selectionSession !== view.selectionSession) return;
        const session = view.selectionSession;
        const instance = typeof msg.instance === 'string' ? msg.instance : '';
        // One init per webview document: a repeated ready from the document
        // that already has (or is being sent) its init is ignored. A document
        // re-created for the same HTML has a new instance id and is served.
        if (view.initialized?.session === session && view.initialized.instance === instance) return;
        const target = { session, instance };
        view.initialized = target;
        let data: Uint8Array | undefined;
        let error: string | undefined;
        try {
          // `readFile` may return a Node Buffer (often a view into a shared
          // pool). The webview serializer JSON-encodes a Buffer as
          // `{type:"Buffer",data:[...]}` instead of transferring it, so copy
          // into a plain, exact-length Uint8Array that it transfers as bytes.
          data = new Uint8Array(await vscode.workspace.fs.readFile(document.uri));
        } catch (err) {
          error = err instanceof Error ? err.message : String(err);
        }
        // A reload (new session), a re-created document (new instance) or a
        // close during the read supersedes it.
        if (disposed || view.initialized !== target) return;
        const posted = await webviewPanel.webview.postMessage({
          type: 'ooxml-init',
          fileType: this.fileType,
          data,
          error,
          useGoogleFonts: shouldUseGoogleFonts(),
          selectionSession: session,
          instance,
        });
        if (!posted && view.initialized === target) {
          view.initialized = null;
        }
      } else if (msg.type === 'selection-context') {
        if (
          msg.selectionSession !== view.selectionSession ||
          msg.fileType !== this.fileType
        ) return;
        selection.update(msg.context ?? null);
      } else if (msg.type === 'viewer-location') {
        if (
          msg.selectionSession !== view.selectionSession ||
          msg.fileType !== this.fileType
        ) return;
        selection.updateView(msg.location ?? null);
      }
    });
  }
}
