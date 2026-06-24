import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "./client";
import { registerCommands } from "./commands";
import { createChannels } from "./logging";
import { SessionManager } from "./session";
import { XlflowTestController } from "./testing";

let clientManager: XlflowLanguageClientManager | undefined;
let testController: XlflowTestController | undefined;
let sessionManager: SessionManager | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const channels = createChannels();
  clientManager = new XlflowLanguageClientManager(channels);
  testController = new XlflowTestController(channels);
  sessionManager = new SessionManager(channels);

  context.subscriptions.push(
    channels.output,
    channels.trace,
    clientManager,
    testController,
    sessionManager,
  );
  registerCommands(context, clientManager, channels, sessionManager);
  context.subscriptions.push(
    vscode.workspace.onDidChangeWorkspaceFolders(() => {
      void testController?.refresh();
      void sessionManager?.refreshStatus();
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
  void sessionManager.refreshStatus();
}

export async function deactivate(): Promise<void> {
  const manager = clientManager;
  const tests = testController;
  const sessions = sessionManager;
  clientManager = undefined;
  testController = undefined;
  sessionManager = undefined;
  tests?.dispose();
  sessions?.dispose();
  await manager?.stop();
}
