import * as vscode from "vscode";
import {
  registerDocumentationCodeActions,
  registerLineSuppressionCodeActions,
} from "./codeActions";
import { showProjectCliUnavailableNotice, XlflowCliAvailabilityService } from "./cliAvailability";
import { XlflowLanguageClientManager } from "./client";
import { registerCommands } from "./commands";
import { checkVbaLanguageAssociation } from "./languageAssociation";
import { createChannels } from "./logging";
import { readConfig } from "./config";
import { selectedWorkspaceFolder, XlflowProjectStateService } from "./projectState";
import { SessionManager } from "./session";
import { XlflowSidebar } from "./sidebar";
import { StartupTelemetry } from "./startupTelemetry";
import { XlflowUpdateService } from "./updateCheck";
import { XlflowTestController } from "./testing";
import { XlflowCapabilitiesService } from "./capabilities";
import { XlflowRulesRegistryService } from "./rulesRegistry";
import { setXlflowCapabilitiesService, setXlflowCliAvailabilityService } from "./xlflow";

let clientManager: XlflowLanguageClientManager | undefined;
let testController: XlflowTestController | undefined;
let sessionManager: SessionManager | undefined;
let projectState: XlflowProjectStateService | undefined;
let sidebar: XlflowSidebar | undefined;
let cliAvailability: XlflowCliAvailabilityService | undefined;
let updateService: XlflowUpdateService | undefined;
let capabilitiesService: XlflowCapabilitiesService | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const activationStarted = performance.now();
  const channels = createChannels();
  const startup = new StartupTelemetry(
    readConfig().lspPerformanceLogging,
    (record) => channels.output.info(`performance ${JSON.stringify(record)}`),
    undefined,
    undefined,
    activationStarted,
  );
  startup.mark("extensionActivationStart");
  cliAvailability = new XlflowCliAvailabilityService();
  setXlflowCliAvailabilityService(cliAvailability);
  clientManager = new XlflowLanguageClientManager(channels, startup);
  testController = new XlflowTestController(channels);
  sessionManager = new SessionManager(channels);
  capabilitiesService = new XlflowCapabilitiesService(channels, {
    currentBusyOperation: () => sessionManager?.currentBusyOperation(),
    operationStarted: (operation) => sessionManager?.setManagedOperation(operation),
    operationFinished: () => sessionManager?.setManagedOperation(undefined),
    refreshStatus: async () => {
      await sessionManager?.refreshStatus();
    },
  });
  const rulesRegistry = new XlflowRulesRegistryService();
  setXlflowCapabilitiesService(capabilitiesService);
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
    registerLineSuppressionCodeActions(rulesRegistry),
    registerDocumentationCodeActions(),
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
      sidebar?.refreshFormulas(),
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
      refreshFormulas: async () => {
        await sidebar?.refreshFormulas();
      },
    },
  );

  const configWatcher = vscode.workspace.createFileSystemWatcher("**/xlflow.toml");
  const formulasManifestWatcher = vscode.workspace.createFileSystemWatcher(
    "**/formulas/manifest.json",
  );
  const formulasJsonlWatcher = vscode.workspace.createFileSystemWatcher("**/formulas/**/*.jsonl");
  const refreshFormulas = () => {
    void sidebar?.refreshFormulas();
  };
  let configRefreshPromise: Promise<void> = Promise.resolve();
  const refreshAfterConfigChange = () => {
    configRefreshPromise = configRefreshPromise.then(async () => {
      // The language server loads xlflow.toml at startup, so a project
      // configuration change needs a full restart rather than only a view refresh.
      try {
        await clientManager?.restart();
      } catch (error) {
        channels.output.error(`xlflow configuration refresh failed: ${String(error)}`);
      }
      try {
        await refreshSelectedProject();
      } catch (error) {
        channels.output.error(`xlflow project refresh failed: ${String(error)}`);
      }
    });
  };
  context.subscriptions.push(
    vscode.workspace.onDidChangeWorkspaceFolders(() => {
      lastSelectedWorkspaceKey = selectedWorkspaceKey();
      void refreshSelectedProject({ restartLsp: true });
    }),
    vscode.window.onDidChangeActiveTextEditor(() => {
      const key = selectedWorkspaceKey();
      if (key === lastSelectedWorkspaceKey) {
        void checkVbaLanguageAssociation(context);
        return;
      }
      lastSelectedWorkspaceKey = key;
      void checkVbaLanguageAssociation(context);
      void refreshSelectedProject({ restartLsp: true });
    }),
    configWatcher,
    formulasManifestWatcher,
    formulasJsonlWatcher,
    configWatcher.onDidCreate(refreshAfterConfigChange),
    configWatcher.onDidChange(refreshAfterConfigChange),
    configWatcher.onDidDelete(refreshAfterConfigChange),
    formulasManifestWatcher.onDidCreate(refreshFormulas),
    formulasManifestWatcher.onDidChange(refreshFormulas),
    formulasManifestWatcher.onDidDelete(refreshFormulas),
    formulasJsonlWatcher.onDidCreate(refreshFormulas),
    formulasJsonlWatcher.onDidChange(refreshFormulas),
    formulasJsonlWatcher.onDidDelete(refreshFormulas),
    vscode.workspace.onDidChangeTextDocument((event) => {
      clientManager?.scheduleSuggest(event.document);
    }),
    vscode.workspace.onDidChangeConfiguration(async (event) => {
      const pathChanged = event.affectsConfiguration("xlflow.path");
      const lspChanged = event.affectsConfiguration("xlflow.lsp");
      if (pathChanged) {
        capabilitiesService?.invalidate();
        rulesRegistry.invalidate();
        await cliAvailability?.refresh();
        void rulesRegistry.load().catch(() => undefined);
        void capabilitiesService?.load();
        await updateService?.checkAutomatic(cliAvailability?.current());
      }
      if (pathChanged || lspChanged) {
        try {
          await clientManager?.restart();
        } catch (error) {
          channels.output.error(`xlflow language server restart failed: ${String(error)}`);
        }
      }
      if (pathChanged) {
        await refreshSelectedProject();
      }
      if (event.affectsConfiguration("xlflow.testing.autoDiscover")) {
        await testController?.refreshAuto();
      }
    }),
  );

  const availabilityPromise = (async () => {
    startup.mark("cliAvailabilityStart");
    const availability = await cliAvailability!.refresh();
    startup.mark("cliAvailabilityComplete", { outcome: availability.ok ? "ok" : "error" });
    return availability;
  })();
  const projectRefreshPromise = refreshSelectedProject({ restartLsp: false });
  const projectRefreshState = projectRefreshPromise.then(
    () => projectState?.current(),
    () => projectState?.current(),
  );
  void availabilityPromise
    .then(async (availability) => {
      await updateService?.checkAutomatic(availability);
      const state = await projectRefreshState;
      if (state?.kind === "ready") {
        await showProjectCliUnavailableNotice(context, state.workspaceFolder, availability);
      }
    })
    .catch((error) =>
      channels.output.error(`xlflow CLI availability refresh failed: ${String(error)}`),
    );
  void rulesRegistry.load().catch((error) => {
    channels.output.error(`xlflow rules refresh failed: ${String(error)}`);
  });
  void capabilitiesService.load().catch((error) => {
    channels.output.error(`xlflow capabilities refresh failed: ${String(error)}`);
  });
  void projectRefreshPromise.catch((error) => {
    channels.output.error(`xlflow project refresh failed: ${String(error)}`);
  });
  void checkVbaLanguageAssociation(context).catch((error) => {
    channels.output.error(`xlflow VBA language association check failed: ${String(error)}`);
  });

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
  capabilitiesService = undefined;
  setXlflowCliAvailabilityService(undefined);
  setXlflowCapabilitiesService(undefined);
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
