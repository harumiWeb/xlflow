import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "./client";
import { registerCommands } from "./commands";
import { createChannels } from "./logging";

let clientManager: XlflowLanguageClientManager | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const channels = createChannels();
  clientManager = new XlflowLanguageClientManager(channels);

  context.subscriptions.push(channels.output, channels.trace, clientManager);
  registerCommands(context, clientManager, channels);
  context.subscriptions.push(
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
}

export async function deactivate(): Promise<void> {
  const manager = clientManager;
  clientManager = undefined;
  await manager?.stop();
}
