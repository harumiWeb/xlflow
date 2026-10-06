import * as assert from "assert";
import * as vscode from "vscode";
import type { XlflowLanguageClientManager } from "../../src/client";
import { UserFormEditorProvider } from "../../src/userFormEditor/provider";
import type { DocumentEditQueue } from "../../src/userFormEditor/edits";
import type { DesignerError, HostMessage, WebviewMessage } from "../../src/userFormEditor/protocol";

// Use the real provider and VS Code events with a controllable panel transport
// and deferred queue result, so visibility cannot race the test driver.
export async function runUserFormProviderAssertions(extensionUri: vscode.Uri): Promise<void> {
  const incoming = new vscode.EventEmitter<WebviewMessage>();
  const view = new vscode.EventEmitter<vscode.WebviewPanelOnDidChangeViewStateEvent>();
  const disposed = new vscode.EventEmitter<void>();
  const connection = new vscode.EventEmitter<void>();
  const cancel = new vscode.CancellationTokenSource();
  const messages: HostMessage[] = [];
  let visible = true;
  let delivered = true;
  const panel = {
    get visible() {
      return visible;
    },
    onDidChangeViewState: view.event,
    onDidDispose: disposed.event,
    webview: {
      options: {},
      html: "",
      cspSource: "https://test.invalid",
      asWebviewUri: (uri: vscode.Uri) => uri,
      onDidReceiveMessage: incoming.event,
      postMessage: async (message: HostMessage) => {
        if (delivered) messages.push(message);
        return delivered;
      },
    },
  } as unknown as vscode.WebviewPanel;
  const client = {
    onDidChangePreviewConnection: connection.event,
  } as unknown as XlflowLanguageClientManager;
  const provider = new UserFormEditorProvider(extensionUri, client);
  let finish!: (error: DesignerError | undefined) => void;
  let current!: () => boolean;
  (provider as unknown as { edits: Pick<DocumentEditQueue, "run"> }).edits = {
    run: async (_document, _version, _operations, _client, permitted) => {
      current = permitted;
      return new Promise<DesignerError | undefined>((resolve) => {
        finish = resolve;
      });
    },
  };
  const document = await vscode.workspace.openTextDocument({
    language: "yaml",
    content: "form: {name: Main, caption: Original}\ncontrols: []\n",
  });
  const source = document.getText();
  const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0));
  const show = async () => {
    visible = true;
    view.fire({ webviewPanel: panel });
    await tick();
  };
  const edit = (requestId: number) =>
    incoming.fire({
      type: "edit",
      requestId,
      version: document.version,
      operations: [{ type: "setFormProperty", field: "caption", value: "Draft" }],
    });
  const replies = () => messages.filter((m) => m.type === "editResult");
  try {
    await provider.resolveCustomTextEditor(document, panel, cancel.token);
    incoming.fire({ type: "ready" });
    await tick();
    const errors: (DesignerError | undefined)[] = [
      { code: "staleDocument", message: "Hidden edit is stale" },
      {
        code: "invalid",
        message: "Invalid caption",
        diagnostics: [
          { operationIndex: 0, field: "caption", code: "invalid_caption", message: "Rejected" },
        ],
      },
      undefined,
    ];
    for (const [index, error] of errors.entries()) {
      const requestId = index + 1;
      const before = replies().length;
      edit(requestId);
      visible = false;
      assert.strictEqual(current(), false, "hidden panels cannot apply a late source edit");
      finish(error);
      await tick();
      assert.strictEqual(replies().length, before, "hidden replies are retained, not posted");
      await show();
      assert.deepStrictEqual(replies().at(-1), { type: "editResult", requestId, error });
      assert.strictEqual(replies().length, before + 1);
      await show();
      assert.strictEqual(replies().length, before + 1, "delivered replies are not replayed twice");
      assert.strictEqual(document.getText(), source);
    }
    // A transport rejection must remain queued until a subsequent visible flush.
    edit(4);
    delivered = false;
    finish(undefined);
    await tick();
    const beforeRetry = replies().length;
    delivered = true;
    await show();
    assert.strictEqual(replies().length, beforeRetry + 1);
    assert.deepStrictEqual(replies().at(-1), {
      type: "editResult",
      requestId: 4,
      error: undefined,
    });

    // Recreated Webviews restart IDs: neither queued nor still-running old
    // replies may acknowledge a new context's request with the same ID.
    edit(5);
    visible = false;
    finish(undefined);
    await tick();
    edit(6);
    const beforeReset = replies().length;
    visible = true;
    incoming.fire({ type: "ready" });
    assert.strictEqual(current(), false, "a new context invalidates old in-flight edits");
    finish(undefined);
    await tick();
    await show();
    assert.strictEqual(replies().length, beforeReset);
    edit(1);
    finish(undefined);
    await tick();
    assert.strictEqual(
      replies().length,
      beforeReset + 1,
      "new context edits still receive replies",
    );
  } finally {
    disposed.fire();
    incoming.dispose();
    view.dispose();
    disposed.dispose();
    connection.dispose();
    cancel.dispose();
  }
}
