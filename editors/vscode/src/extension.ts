import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "./client";
import { registerCommands } from "./commands";
import { createChannels } from "./logging";
import { XlflowTestController } from "./testing";

let clientManager: XlflowLanguageClientManager | undefined;
let testController: XlflowTestController | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const channels = createChannels();
  clientManager = new XlflowLanguageClientManager(channels);
  testController = new XlflowTestController(channels);

  context.subscriptions.push(channels.output, channels.trace, clientManager, testController);
  registerCommands(context, clientManager, channels);
  context.subscriptions.push(
    vscode.workspace.onDidChangeWorkspaceFolders(() => {
      void testController?.refresh();
    }),
    vscode.workspace.onDidChangeTextDocument((event) => {
      clientManager?.scheduleSuggest(event.document);
    }),
    vscode.workspace.onDidChangeConfiguration(async (event) => {
      if (event.affectsConfiguration("xlflow.path") || event.affectsConfiguration("xlflow.lsp")) {
        await clientManager?.restart();
      }
    }),
  );

  await clientManager.start();
  void testController.refresh();
}

export async function deactivate(): Promise<void> {
  const manager = clientManager;
  const tests = testController;
  clientManager = undefined;
  testController = undefined;
  tests?.dispose();
  await manager?.stop();
}
