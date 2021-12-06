import * as vscode from 'vscode';
import { readFile, readFileSync } from 'fs';
import { docxToString } from './docxToString';

interface Test extends vscode.CustomDocument {
  test: string;
}

export class DocxViewerProvider implements vscode.CustomReadonlyEditorProvider<Test> {

  openCustomDocument(uri: vscode.Uri, openContext: vscode.CustomDocumentOpenContext): Test {
    console.log(uri);
    return {
      test: 'testing...',
      dispose: () => { },
      uri: uri,
    };
  }

  resolveCustomEditor(document: Test, webviewPanel: vscode.WebviewPanel, token: vscode.CancellationToken): Thenable<void> {
    console.log(document.uri.path);
    docxToString(document.uri.fsPath).then(text => {
      console.log(text);
      webviewPanel.webview.html = `<!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Document</title>
      </head>
      <body>
          <p style="white-space: pre-wrap;">${text}</p>
      </body>
      </html>`;
    });

    return Promise.resolve();
  }

  // register the provider for the custom editor
  public static register(context: vscode.ExtensionContext) {
    return vscode.window.registerCustomEditorProvider('docx.docx', new DocxViewerProvider());
  }

}