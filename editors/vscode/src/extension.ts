import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "./client";
import { registerCommands } from "./commands";
import { createChannels } from "./logging";
import { XlflowProjectStateService } from "./projectState";
import { SessionManager } from "./session";
import { XlflowSidebar } from "./sidebar";
import { XlflowTestController } from "./testing";

let clientManager: XlflowLanguageClientManager | undefined;
let testController: XlflowTestController | undefined;
let sessionManager: SessionManager | undefined;
let projectState: XlflowProjectStateService | undefined;
let sidebar: XlflowSidebar | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const channels = createChannels();
  clientManager = new XlflowLanguageClientManager(channels);
  testController = new XlflowTestController(channels);
  sessionManager = new SessionManager(channels);
  projectState = new XlflowProjectStateService();
  sidebar = new XlflowSidebar(projectState, sessionManager, channels);

  context.subscriptions.push(
    channels.output,
    channels.trace,
    clientManager,
    testController,
    sessionManager,
    projectState,
    sidebar,
  );
  const refreshSelectedProject = async (): Promise<void> => {
    await projectState?.refresh();
    await sessionManager?.refreshStatus();
    await testController?.refreshAuto();
    await Promise.all([
      sidebar?.refreshModules(),
      sidebar?.refreshUserForms(),
      sidebar?.refreshTests(),
    ]);
  };

  registerCommands(context, clientManager, channels, sessionManager, projectState, {
    refreshAll: refreshSelectedProject,
    refreshProject: () => {
      sidebar?.refreshProjectViews();
    },
    refreshModules: async () => {
      await sidebar?.refreshModules();
    },
    refreshUserForms: async () => {
      await sidebar?.refreshUserForms();
    },
    refreshTests: async () => {
      await testController?.refreshAuto();
      await sidebar?.refreshTests();
    },
  });

  const configWatcher = vscode.workspace.createFileSystemWatcher("**/xlflow.toml");
  context.subscriptions.push(
    vscode.workspace.onDidChangeWorkspaceFolders(() => {
      void refreshSelectedProject();
    }),
    vscode.window.onDidChangeActiveTextEditor(() => {
      void refreshSelectedProject();
    }),
    configWatcher,
    configWatcher.onDidCreate(() => {
      void refreshSelectedProject();
    }),
    configWatcher.onDidChange(() => {
      void refreshSelectedProject();
    }),
    configWatcher.onDidDelete(() => {
      void refreshSelectedProject();
    }),
    vscode.workspace.onDidChangeTextDocument((event) => {
      clientManager?.scheduleSuggest(event.document);
    }),
    vscode.workspace.onDidChangeConfiguration(async (event) => {
      if (event.affectsConfiguration("xlflow.path") || event.affectsConfiguration("xlflow.lsp")) {
        await clientManager?.restart();
      }
      if (event.affectsConfiguration("xlflow.testing.autoDiscover")) {
        await testController?.refreshAuto();
      }
    }),
  );

  try {
    await clientManager.start();
  } catch (error) {
    channels.output.error(`xlflow language server startup failed: ${String(error)}`);
    vscode.window.showWarningMessage(
      "xlflow language server failed to start. Command palette actions remain available; check xlflow.path or run xlflow: Check Environment.",
    );
  }
  await projectState.refresh();
  void testController.refreshAuto();
  void sessionManager.refreshStatus();
}

export async function deactivate(): Promise<void> {
  const manager = clientManager;
  const tests = testController;
  const sessions = sessionManager;
  const states = projectState;
  const bars = sidebar;
  clientManager = undefined;
  testController = undefined;
  sessionManager = undefined;
  projectState = undefined;
  sidebar = undefined;
  bars?.dispose();
  states?.dispose();
  tests?.dispose();
  sessions?.dispose();
  await manager?.stop();
}
