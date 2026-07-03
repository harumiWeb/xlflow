import * as vscode from "vscode";
import { showProjectCliUnavailableNotice, XlflowCliAvailabilityService } from "./cliAvailability";
import { XlflowLanguageClientManager } from "./client";
import { registerCommands } from "./commands";
import { createChannels } from "./logging";
import { selectedWorkspaceFolder, XlflowProjectStateService } from "./projectState";
import { SessionManager } from "./session";
import { XlflowSidebar } from "./sidebar";
import { XlflowUpdateService } from "./updateCheck";
import { XlflowTestController } from "./testing";
import { setXlflowCliAvailabilityService } from "./xlflow";

let clientManager: XlflowLanguageClientManager | undefined;
let testController: XlflowTestController | undefined;
let sessionManager: SessionManager | undefined;
let projectState: XlflowProjectStateService | undefined;
let sidebar: XlflowSidebar | undefined;
let cliAvailability: XlflowCliAvailabilityService | undefined;
let updateService: XlflowUpdateService | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const channels = createChannels();
  cliAvailability = new XlflowCliAvailabilityService();
  setXlflowCliAvailabilityService(cliAvailability);
  clientManager = new XlflowLanguageClientManager(channels, cliAvailability);
  testController = new XlflowTestController(channels);
  sessionManager = new SessionManager(channels);
  projectState = new XlflowProjectStateService();
  updateService = new XlflowUpdateService(context);
  sidebar = new XlflowSidebar(
    projectState,
    sessionManager,
    cliAvailability,
    updateService,
    channels,
  );

  context.subscriptions.push(
    channels.output,
    channels.trace,
    cliAvailability,
    clientManager,
    testController,
    sessionManager,
    projectState,
    updateService,
    sidebar,
  );
  let lastSelectedWorkspaceKey = selectedWorkspaceKey();

  const refreshProjectStatus = async (options: { restartLsp?: boolean } = {}): Promise<void> => {
    const state = await projectState?.refresh();
    if (options.restartLsp === true) {
      await clientManager?.restartIfWorkspaceChanged();
    }
    await sessionManager?.refreshStatus();
    sidebar?.refreshProjectViews();
    const availability = cliAvailability?.current();
    if (state?.kind === "ready" && availability !== undefined) {
      await showProjectCliUnavailableNotice(context, state.workspaceFolder, availability);
    }
  };
  const refreshProjectDetails = async (): Promise<void> => {
    await testController?.refreshAuto();
    await Promise.all([
      sidebar?.refreshModules(),
      sidebar?.refreshUserForms(),
      sidebar?.refreshTests(),
    ]);
  };
  const refreshSelectedProject = async (
    options: { restartLsp?: boolean; details?: boolean } = {},
  ): Promise<void> => {
    await refreshProjectStatus({ restartLsp: options.restartLsp });
    if (options.details !== false) {
      await refreshProjectDetails();
    }
  };

  registerCommands(
    context,
    clientManager,
    cliAvailability,
    updateService,
    channels,
    sessionManager,
    projectState,
    {
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
    },
  );

  const configWatcher = vscode.workspace.createFileSystemWatcher("**/xlflow.toml");
  context.subscriptions.push(
    vscode.workspace.onDidChangeWorkspaceFolders(() => {
      lastSelectedWorkspaceKey = selectedWorkspaceKey();
      void refreshSelectedProject({ restartLsp: true });
    }),
    vscode.window.onDidChangeActiveTextEditor(() => {
      const key = selectedWorkspaceKey();
      if (key === lastSelectedWorkspaceKey) {
        return;
      }
      lastSelectedWorkspaceKey = key;
      void refreshSelectedProject({ restartLsp: true });
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
      const pathChanged = event.affectsConfiguration("xlflow.path");
      const lspChanged = event.affectsConfiguration("xlflow.lsp");
      if (pathChanged) {
        await cliAvailability?.refresh();
        await updateService?.checkAutomatic(cliAvailability?.current());
      }
      if (pathChanged || lspChanged) {
        await clientManager?.restart();
      }
      if (pathChanged) {
        await refreshSelectedProject();
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
      vscode.l10n.t(
        "xlflow language server failed to start. Command palette actions remain available; check xlflow.path or run xlflow: Check Environment.",
      ),
    );
  }
  await cliAvailability.refresh();
  await updateService.checkAutomatic(cliAvailability.current());
  await refreshSelectedProject({ restartLsp: false });
}

export async function deactivate(): Promise<void> {
  const manager = clientManager;
  const tests = testController;
  const sessions = sessionManager;
  const states = projectState;
  const bars = sidebar;
  const availability = cliAvailability;
  const updates = updateService;
  clientManager = undefined;
  testController = undefined;
  sessionManager = undefined;
  projectState = undefined;
  sidebar = undefined;
  cliAvailability = undefined;
  updateService = undefined;
  setXlflowCliAvailabilityService(undefined);
  bars?.dispose();
  states?.dispose();
  tests?.dispose();
  sessions?.dispose();
  updates?.dispose();
  availability?.dispose();
  await manager?.stop();
}

function selectedWorkspaceKey(): string | undefined {
  return selectedWorkspaceFolder()?.uri.toString();
}
