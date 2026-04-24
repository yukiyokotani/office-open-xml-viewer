import * as vscode from 'vscode';
import { XlsxEditorProvider } from './providers/xlsxEditor';
import { DocxEditorProvider } from './providers/docxEditor';
import { PptxEditorProvider } from './providers/pptxEditor';

export function activate(context: vscode.ExtensionContext): void {
  context.subscriptions.push(
    XlsxEditorProvider.register(context),
    DocxEditorProvider.register(context),
    PptxEditorProvider.register(context),
  );
}

export function deactivate(): void {
  // nothing to clean up
}
