import * as path from "path";
import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "./client";
import { readConfig } from "./config";
import { XlflowChannels } from "./logging";
import { XlflowProjectStateService } from "./projectState";
import { SessionManager } from "./session";
import { resolveWorkspaceRoot, runXlflowCommand, runXlflowJsonCommand } from "./xlflow";

type RunProcedureArgs = {
  uri: string;
  name: string;
  moduleName?: string;
  qualifiedName?: string;
  kind?: "sub" | "test";
  line?: number;
};

interface XlflowMutationEnvelope {
  status?: string;
  error?: {
    message?: string;
    code?: string;
  };
  source?: {
    renamed?: string[];
    removed?: string[];
  };
}

interface CommandRefreshHooks {
  refreshAll(): Promise<void>;
  refreshProject(): Promise<void> | void;
  refreshModules(): Promise<void>;
  refreshUserForms(): Promise<void>;
  refreshTests(): Promise<void>;
}

export function registerCommands(
  context: vscode.ExtensionContext,
  clientManager: XlflowLanguageClientManager,
  channels: XlflowChannels,
  sessionManager: SessionManager,
  projectState: XlflowProjectStateService,
  hooks: CommandRefreshHooks,
): void {
  context.subscriptions.push(
    vscode.commands.registerCommand("xlflow.restartLanguageServer", async () => {
      await clientManager.restart();
    }),
    vscode.commands.registerCommand("xlflow.checkEnvironment", async () => {
      await runXlflowCommand(["lsp", "--check"], "xlflow environment check", channels.output, {
        requireWorkspace: false,
        uiLabel: vscode.l10n.t("xlflow environment check"),
      });
    }),
    vscode.commands.registerCommand("xlflow.newProject", async () => {
      const workbook = await vscode.window.showInputBox({
        title: vscode.l10n.t("xlflow: New Project"),
        prompt: vscode.l10n.t(
          "Workbook filename or project name. Leave empty to use xlflow's default.",
        ),
        placeHolder: "Book.xlsm",
        value: "Book.xlsm",
      });
      if (workbook === undefined) {
        return;
      }
      const args = ["new"];
      if (workbook.trim() !== "") {
        args.push(workbook.trim());
      }
      const code = await runXlflowCommand(args, "xlflow new", channels.output, {
        requireWorkspace: true,
        uiLabel: vscode.l10n.t("xlflow new"),
      });
      if (code === 0) {
        await hooks.refreshAll();
      }
    }),
    vscode.commands.registerCommand("xlflow.initProject", async () => {
      const workbooks = await vscode.window.showOpenDialog({
        title: vscode.l10n.t("xlflow: Initialize Project"),
        canSelectFiles: true,
        canSelectFolders: false,
        canSelectMany: false,
        filters: {
          [vscode.l10n.t("Excel workbooks")]: ["xlsm", "xlsb", "xlsx", "xls"],
          [vscode.l10n.t("All files")]: ["*"],
        },
      });
      const workbook = workbooks?.[0];
      if (workbook === undefined) {
        return;
      }
      const code = await runXlflowCommand(
        ["init", workbook.fsPath],
        "xlflow init",
        channels.output,
        {
          requireWorkspace: true,
          uiLabel: vscode.l10n.t("xlflow init"),
        },
      );
      if (code === 0) {
        await hooks.refreshAll();
      }
    }),
    vscode.commands.registerCommand("xlflow.skillInstall", async () => {
      await installAgentSkill(channels);
    }),
    vscode.commands.registerCommand("xlflow.moduleInstall", async () => {
      await installHelperModules(channels);
    }),
    vscode.commands.registerCommand("xlflow.newModule", async () => {
      await newModule(channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.newStandardModule", async () => {
      await newModuleOfType("standard", channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.newClassModule", async () => {
      await newModuleOfType("class", channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.newUserForm", async () => {
      await newUserForm(channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.pull", async () => {
      const code = await runXlflowCommand(["pull"], "xlflow pull", channels.output, {
        requireWorkspace: true,
        uiLabel: vscode.l10n.t("xlflow pull"),
      });
      if (code === 0) {
        await Promise.all([
          hooks.refreshProject(),
          hooks.refreshModules(),
          hooks.refreshUserForms(),
          hooks.refreshTests(),
        ]);
      }
    }),
    vscode.commands.registerCommand("xlflow.push", async () => {
      const pushLabel = vscode.l10n.t("Push");
      const confirmed = await vscode.window.showWarningMessage(
        vscode.l10n.t("Push sources to workbook?"),
        { modal: true },
        pushLabel,
      );
      if (confirmed !== pushLabel) {
        return;
      }
      const code = await runXlflowCommand(["push"], "xlflow push", channels.output, {
        requireWorkspace: true,
        uiLabel: vscode.l10n.t("xlflow push"),
      });
      if (code === 0) {
        await Promise.all([
          hooks.refreshProject(),
          hooks.refreshModules(),
          hooks.refreshUserForms(),
          hooks.refreshTests(),
        ]);
      }
    }),
    vscode.commands.registerCommand("xlflow.runMacro", async () => {
      await runXlflowCommand(["run"], "xlflow run", channels.output, {
        requireWorkspace: true,
        uiLabel: vscode.l10n.t("xlflow run"),
      });
    }),
    vscode.commands.registerCommand("xlflow.runProcedure", async (args: unknown) => {
      await runProcedure(args, channels);
    }),
    vscode.commands.registerCommand("xlflow.runTestProcedure", async (args: unknown) => {
      await runTestProcedureFromCodeLens(args, channels);
    }),
    vscode.commands.registerCommand("xlflow.test", async () => {
      const code = await runXlflowCommand(["test"], "xlflow test", channels.output, {
        requireWorkspace: true,
        uiLabel: vscode.l10n.t("xlflow test"),
      });
      if (code === 0) {
        await hooks.refreshTests();
      }
    }),
    vscode.commands.registerCommand("xlflow.lintWorkspace", async () => {
      await runXlflowCommand(["lint"], "xlflow lint", channels.output, {
        requireWorkspace: true,
        uiLabel: vscode.l10n.t("xlflow lint"),
      });
    }),
    vscode.commands.registerCommand("xlflow.formatDocument", async () => {
      const editor = vscode.window.activeTextEditor;
      if (editor === undefined) {
        vscode.window.showWarningMessage(
          vscode.l10n.t("xlflow format document requires an active editor."),
        );
        return;
      }
      await vscode.commands.executeCommand("editor.action.formatDocument");
    }),
    vscode.commands.registerCommand("xlflow.formatProject", async () => {
      await runXlflowCommand(["fmt", "--write"], "xlflow fmt", channels.output, {
        requireWorkspace: true,
        uiLabel: vscode.l10n.t("xlflow fmt"),
      });
    }),
    vscode.commands.registerCommand("xlflow.saveWorkbook", async () => {
      await runXlflowCommand(["save"], "xlflow save", channels.output, {
        requireWorkspace: true,
        uiLabel: vscode.l10n.t("xlflow save"),
      });
    }),
    vscode.commands.registerCommand("xlflow.sessionStart", async () => {
      await sessionManager.start();
      hooks.refreshProject();
    }),
    vscode.commands.registerCommand("xlflow.sessionStatus", async () => {
      await sessionManager.showStatus();
    }),
    vscode.commands.registerCommand("xlflow.sessionStop", async () => {
      await sessionManager.stop();
      hooks.refreshProject();
    }),
    vscode.commands.registerCommand("xlflow.sessionRestart", async () => {
      await sessionManager.restart();
      hooks.refreshProject();
    }),
    vscode.commands.registerCommand("xlflow.sessionActions", async () => {
      await sessionManager.showActions();
    }),
    vscode.commands.registerCommand("xlflow.openOutput", () => {
      sessionManager.openOutput();
    }),
    vscode.commands.registerCommand("xlflow.refreshProject", async () => {
      await projectState.refresh();
      await sessionManager.refreshStatus();
      hooks.refreshProject();
    }),
    vscode.commands.registerCommand("xlflow.refreshModules", async () => {
      await hooks.refreshModules();
    }),
    vscode.commands.registerCommand("xlflow.collapseModules", async () => {
      await vscode.commands.executeCommand("workbench.actions.treeView.xlflow.modules.collapseAll");
    }),
    vscode.commands.registerCommand("xlflow.refreshUserForms", async () => {
      await hooks.refreshUserForms();
    }),
    vscode.commands.registerCommand("xlflow.collapseUserForms", async () => {
      await vscode.commands.executeCommand(
        "workbench.actions.treeView.xlflow.userForms.collapseAll",
      );
    }),
    vscode.commands.registerCommand("xlflow.refreshTests", async () => {
      await hooks.refreshTests();
    }),
    vscode.commands.registerCommand("xlflow.runAllTests", async () => {
      await vscode.commands.executeCommand("xlflow.test");
    }),
    vscode.commands.registerCommand("xlflow.runDoctor", async () => {
      await sessionManager.runDoctor();
    }),
    vscode.commands.registerCommand("xlflow.sessionToggle", async () => {
      if (sessionManager.currentSnapshot().state === "active") {
        await sessionManager.stop();
      } else {
        await sessionManager.start();
      }
      hooks.refreshProject();
    }),
    vscode.commands.registerCommand("xlflow.setupActions", async () => {
      await showSetupActions();
    }),
    vscode.commands.registerCommand("xlflow.openDocumentation", async () => {
      await vscode.env.openExternal(vscode.Uri.parse("https://harumiweb.github.io/xlflow/"));
    }),
    vscode.commands.registerCommand("xlflow.openModule", async (value: unknown) => {
      const uri = treeUri(value);
      if (uri !== undefined) {
        await vscode.window.showTextDocument(uri);
      }
    }),
    vscode.commands.registerCommand("xlflow.renameModule", async (value: unknown) => {
      await renameModule(value, channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.deleteModule", async (value: unknown) => {
      await deleteModule(value, channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.revealSourceFile", async (value: unknown) => {
      const uri = treeUri(value);
      if (uri !== undefined) {
        await vscode.commands.executeCommand("revealInExplorer", uri);
      }
    }),
    vscode.commands.registerCommand("xlflow.copyModuleName", async (value: unknown) => {
      await copyText(vscode.l10n.t("Module name"), treeName(value));
    }),
    vscode.commands.registerCommand("xlflow.copyRelativePath", async (value: unknown) => {
      const uri = treeUri(value);
      await copyText(
        vscode.l10n.t("Relative path"),
        uri === undefined ? undefined : relativePathForUri(uri),
      );
    }),
    vscode.commands.registerCommand("xlflow.copyProcedureName", async (value: unknown) => {
      await copyText(vscode.l10n.t("Procedure name"), treeName(value));
    }),
    vscode.commands.registerCommand("xlflow.copyQualifiedName", async (value: unknown) => {
      await copyText(vscode.l10n.t("Qualified name"), treeQualifiedName(value));
    }),
    vscode.commands.registerCommand("xlflow.renameUserForm", async (value: unknown) => {
      await renameUserForm(value, channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.deleteUserForm", async (value: unknown) => {
      await deleteUserForm(value, channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.revealUserFormSource", async (value: unknown) => {
      const uri = userFormSourceUri(value);
      if (uri !== undefined) {
        await vscode.commands.executeCommand("revealInExplorer", uri);
      }
    }),
    vscode.commands.registerCommand("xlflow.copyUserFormName", async (value: unknown) => {
      await copyText(vscode.l10n.t("UserForm name"), treeName(value));
    }),
    vscode.commands.registerCommand("xlflow.copyUserFormRelativePath", async (value: unknown) => {
      await copyText(
        vscode.l10n.t("Relative path"),
        userFormRelativePath(value) ?? relativePathFromUri(userFormSourceUri(value)),
      );
    }),
    vscode.commands.registerCommand("xlflow.openProcedure", async (value: unknown) => {
      const uri = treeUri(value);
      if (uri === undefined) {
        return;
      }
      const line = treeLine(value);
      const document = await vscode.workspace.openTextDocument(uri);
      const editor = await vscode.window.showTextDocument(document);
      if (line !== undefined) {
        const position = new vscode.Position(Math.max(0, line - 1), 0);
        editor.selection = new vscode.Selection(position, position);
        editor.revealRange(
          new vscode.Range(position, position),
          vscode.TextEditorRevealType.InCenter,
        );
      }
    }),
    vscode.commands.registerCommand("xlflow.openUserFormArtifact", async (value: unknown) => {
      const uri = treeUri(value);
      if (uri !== undefined) {
        await vscode.window.showTextDocument(uri);
      }
    }),
  );
}

async function showSetupActions(): Promise<void> {
  const action = await vscode.window.showQuickPick(
    [
      { label: vscode.l10n.t("New Project"), command: "xlflow.newProject" },
      { label: vscode.l10n.t("Init Existing Workbook"), command: "xlflow.initProject" },
      { label: vscode.l10n.t("Run Doctor"), command: "xlflow.runDoctor" },
      { label: vscode.l10n.t("Open Documentation"), command: "xlflow.openDocumentation" },
    ],
    {
      title: vscode.l10n.t("xlflow Project Setup"),
      placeHolder: vscode.l10n.t("Select a setup action"),
    },
  );
  if (action !== undefined) {
    await vscode.commands.executeCommand(action.command);
  }
}

async function installAgentSkill(channels: XlflowChannels): Promise<void> {
  const workspaceFolder = await resolveWorkspaceRoot({ prompt: true });
  if (workspaceFolder === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow skill install requires an open workspace folder."),
    );
    return;
  }
  const provider = await vscode.window.showQuickPick(
    [
      { label: "codex", description: vscode.l10n.t("Install for Codex") },
      { label: "claude", description: vscode.l10n.t("Install for Claude Code") },
      { label: "cursor", description: vscode.l10n.t("Install for Cursor") },
      { label: "gemini", description: vscode.l10n.t("Install for Gemini CLI") },
      { label: "agents", description: vscode.l10n.t("Install shared .agents instructions") },
    ],
    {
      title: vscode.l10n.t("xlflow: Install Agent Skill"),
      placeHolder: vscode.l10n.t("Select the agent provider target"),
    },
  );
  if (provider === undefined) {
    return;
  }

  const overwrite = await vscode.window.showQuickPick(
    [
      {
        label: vscode.l10n.t("Install without overwrite"),
        description: vscode.l10n.t("Fail if the xlflow skill already exists"),
        force: false,
      },
      {
        label: vscode.l10n.t("Overwrite existing installation"),
        description: vscode.l10n.t("Pass --force to replace an existing xlflow skill"),
        force: true,
      },
    ],
    {
      title: vscode.l10n.t("xlflow: Install Agent Skill"),
      placeHolder: vscode.l10n.t("Choose overwrite behavior"),
    },
  );
  if (overwrite === undefined) {
    return;
  }

  const args = ["skill", "install", "--agent", provider.label];
  if (overwrite.force) {
    args.push("--force");
  }
  await runXlflowCommand(args, `xlflow skill install --agent ${provider.label}`, channels.output, {
    requireWorkspace: true,
    uiLabel: vscode.l10n.t("xlflow skill install for {provider}", { provider: provider.label }),
    workspaceFolder,
  });
}

async function installHelperModules(channels: XlflowChannels): Promise<void> {
  const workspaceFolder = await resolveWorkspaceRoot({ prompt: true });
  if (workspaceFolder === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow module install requires an open workspace folder."),
    );
    return;
  }
  const mode = await vscode.window.showQuickPick(
    [
      {
        label: vscode.l10n.t("Install to source only"),
        description: vscode.l10n.t("Run xlflow module install"),
        args: ["module", "install"],
      },
      {
        label: vscode.l10n.t("Install and push to workbook"),
        description: vscode.l10n.t("Run xlflow module install --push"),
        args: ["module", "install", "--push"],
      },
    ],
    {
      title: vscode.l10n.t("xlflow: Install Helper Modules"),
      placeHolder: vscode.l10n.t("Choose how to install bundled helper modules"),
    },
  );
  if (mode === undefined) {
    return;
  }
  await runXlflowCommand(mode.args, `xlflow ${mode.args.join(" ")}`, channels.output, {
    requireWorkspace: true,
    uiLabel: mode.args.includes("--push")
      ? vscode.l10n.t("xlflow module install --push")
      : vscode.l10n.t("xlflow module install"),
    workspaceFolder,
  });
}

async function newModule(channels: XlflowChannels, hooks: CommandRefreshHooks): Promise<void> {
  const moduleType = await vscode.window.showQuickPick(
    [
      {
        label: vscode.l10n.t("Standard Module"),
        description: vscode.l10n.t("Run xlflow module new --type standard"),
        moduleType: "standard" as const,
        placeholder: "InvoiceProcessor",
      },
      {
        label: vscode.l10n.t("Class Module"),
        description: vscode.l10n.t("Run xlflow module new --type class"),
        moduleType: "class" as const,
        placeholder: "InvoiceService",
      },
    ],
    {
      title: vscode.l10n.t("xlflow: New Module"),
      placeHolder: vscode.l10n.t("Select the module type"),
    },
  );
  if (moduleType === undefined) {
    return;
  }
  await newModuleOfType(moduleType.moduleType, channels, hooks, moduleType.placeholder);
}

async function newModuleOfType(
  moduleType: "standard" | "class",
  channels: XlflowChannels,
  hooks: CommandRefreshHooks,
  placeholder?: string,
): Promise<void> {
  const workspaceFolder = await resolveWorkspaceRoot({ prompt: true });
  if (workspaceFolder === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow module new requires an open workspace folder."),
    );
    return;
  }
  const title =
    moduleType === "standard"
      ? vscode.l10n.t("xlflow: New Standard Module")
      : vscode.l10n.t("xlflow: New Class Module");
  const name = await promptComponentName(title, placeholder ?? "InvoiceProcessor");
  if (name === undefined) {
    return;
  }
  const args = ["module", "new", name, "--type", moduleType];
  const code = await runXlflowCommand(args, `xlflow module new ${name}`, channels.output, {
    requireWorkspace: true,
    uiLabel: vscode.l10n.t("xlflow module new {name}", { name }),
    workspaceFolder,
  });
  if (code === 0) {
    await hooks.refreshModules();
  }
}

async function newUserForm(channels: XlflowChannels, hooks: CommandRefreshHooks): Promise<void> {
  const workspaceFolder = await resolveWorkspaceRoot({ prompt: true });
  if (workspaceFolder === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow form new requires an open workspace folder."),
    );
    return;
  }
  const name = await promptComponentName(vscode.l10n.t("xlflow: New UserForm"), "CustomerForm");
  if (name === undefined) {
    return;
  }
  const code = await runXlflowCommand(
    ["form", "new", name],
    `xlflow form new ${name}`,
    channels.output,
    {
      requireWorkspace: true,
      uiLabel: vscode.l10n.t("xlflow form new {name}", { name }),
      workspaceFolder,
    },
  );
  if (code === 0) {
    await hooks.refreshUserForms();
  }
}

async function renameUserForm(
  value: unknown,
  channels: XlflowChannels,
  hooks: CommandRefreshHooks,
): Promise<void> {
  const name = treeName(value);
  const uri = userFormSourceUri(value);
  if (name === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow rename UserForm received invalid UserForm arguments."),
    );
    return;
  }
  const newName = await promptComponentName(vscode.l10n.t("xlflow: Rename UserForm"), name);
  if (newName === undefined || newName.trim() === name) {
    return;
  }

  const workspaceFolder = await workspaceFolderForUserForm(value, uri);
  const wasOpen =
    uri !== undefined &&
    vscode.workspace.textDocuments.some((document) => document.uri.toString() === uri.toString());
  const result = await runModuleMutation(
    ["rename", name, newName.trim()],
    `xlflow module rename ${name} ${newName.trim()}`,
    channels.output,
    workspaceFolder,
  );
  if (result.exitCode !== 0) {
    showMutationFailure("rename", name, result, "UserForm");
    return;
  }

  await Promise.all([hooks.refreshUserForms(), hooks.refreshModules(), hooks.refreshTests()]);
  if (wasOpen && workspaceFolder !== undefined) {
    const renamedUri = renamedModuleUri(workspaceFolder, result.json, newName.trim());
    if (renamedUri !== undefined) {
      await vscode.window.showTextDocument(renamedUri);
    }
  }
}

async function deleteUserForm(
  value: unknown,
  channels: XlflowChannels,
  hooks: CommandRefreshHooks,
): Promise<void> {
  const name = treeName(value);
  const uri = userFormSourceUri(value);
  if (name === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow delete UserForm received invalid UserForm arguments."),
    );
    return;
  }

  const deleteLabel = vscode.l10n.t("Delete");
  const confirmed = await vscode.window.showWarningMessage(
    vscode.l10n.t(
      'Delete UserForm "{name}" from the xlflow project?\n\nThis may remove the .frm, .frx, sidecar code, and designer spec artifacts. The workbook will be updated on the next xlflow push.',
      { name },
    ),
    { modal: true },
    deleteLabel,
  );
  if (confirmed !== deleteLabel) {
    return;
  }

  const workspaceFolder = await workspaceFolderForUserForm(value, uri);
  const result = await runModuleMutation(
    ["remove", name],
    `xlflow module remove ${name}`,
    channels.output,
    workspaceFolder,
  );
  if (result.exitCode !== 0) {
    showMutationFailure("delete", name, result, "UserForm");
    return;
  }

  await Promise.all([hooks.refreshUserForms(), hooks.refreshModules(), hooks.refreshTests()]);
}

async function renameModule(
  value: unknown,
  channels: XlflowChannels,
  hooks: CommandRefreshHooks,
): Promise<void> {
  const moduleName = treeName(value);
  const uri = treeUri(value);
  if (moduleName === undefined || uri === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow rename module received invalid module arguments."),
    );
    return;
  }
  if (treeModuleKind(value) === "document") {
    vscode.window.showErrorMessage(
      vscode.l10n.t('Failed to rename module "{moduleName}": document modules cannot be renamed.', {
        moduleName,
      }),
    );
    return;
  }

  const newName = await promptComponentName(vscode.l10n.t("xlflow: Rename Module"), moduleName);
  if (newName === undefined || newName.trim() === moduleName) {
    return;
  }

  const workspaceFolder = await workspaceFolderForUri(uri);
  const wasOpen = vscode.workspace.textDocuments.some(
    (document) => document.uri.toString() === uri.toString(),
  );
  const result = await runModuleMutation(
    ["rename", moduleName, newName.trim()],
    `xlflow module rename ${moduleName} ${newName.trim()}`,
    channels.output,
    workspaceFolder,
  );
  if (result.exitCode !== 0) {
    showMutationFailure("rename", moduleName, result);
    return;
  }

  await Promise.all([hooks.refreshModules(), hooks.refreshTests()]);
  if (wasOpen && workspaceFolder !== undefined) {
    const renamedUri = renamedModuleUri(workspaceFolder, result.json, newName.trim());
    if (renamedUri !== undefined) {
      await vscode.window.showTextDocument(renamedUri);
    }
  }
}

async function deleteModule(
  value: unknown,
  channels: XlflowChannels,
  hooks: CommandRefreshHooks,
): Promise<void> {
  const moduleName = treeName(value);
  const uri = treeUri(value);
  if (moduleName === undefined || uri === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow delete module received invalid module arguments."),
    );
    return;
  }
  if (treeModuleKind(value) === "document") {
    vscode.window.showErrorMessage(
      vscode.l10n.t('Failed to delete module "{moduleName}": document modules cannot be removed.', {
        moduleName,
      }),
    );
    return;
  }

  const deleteLabel = vscode.l10n.t("Delete");
  const confirmed = await vscode.window.showWarningMessage(
    vscode.l10n.t(
      'Delete module "{moduleName}" from the xlflow project?\n\nThis removes the source file. The workbook will be updated on the next xlflow push.',
      { moduleName },
    ),
    { modal: true },
    deleteLabel,
  );
  if (confirmed !== deleteLabel) {
    return;
  }

  const workspaceFolder = await workspaceFolderForUri(uri);
  const result = await runModuleMutation(
    ["remove", moduleName],
    `xlflow module remove ${moduleName}`,
    channels.output,
    workspaceFolder,
  );
  if (result.exitCode !== 0) {
    showMutationFailure("delete", moduleName, result);
    return;
  }

  await Promise.all([hooks.refreshModules(), hooks.refreshTests()]);
}

async function promptComponentName(
  title: string,
  placeHolder: string,
): Promise<string | undefined> {
  return vscode.window.showInputBox({
    title,
    prompt: vscode.l10n.t("Enter a VBA component name without a file extension or path."),
    placeHolder,
    validateInput: validateComponentNameInput,
  });
}

function validateComponentNameInput(value: string): string | undefined {
  const name = value.trim();
  if (name.length === 0) {
    return vscode.l10n.t("Name is required.");
  }
  if (/[\\/]/.test(name) || name.includes("..")) {
    return vscode.l10n.t("Use a component name, not a path.");
  }
  if (/\.(bas|cls|frm)$/i.test(name)) {
    return vscode.l10n.t("Do not include a file extension.");
  }
  return undefined;
}

async function runProcedure(value: unknown, channels: XlflowChannels): Promise<void> {
  const args = normalizeRunProcedureArgs(value);
  if (args === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow received invalid run procedure arguments."),
    );
    return;
  }
  const uri = vscode.Uri.parse(args.uri);
  if (!(await saveDirtyDocumentIfNeeded(uri))) {
    return;
  }
  const workspaceFolder = await workspaceFolderForUri(uri);
  const target = readNonEmpty(args.qualifiedName) ?? args.name;
  await runXlflowCommand(["run", target], `xlflow run ${target}`, channels.output, {
    requireWorkspace: true,
    uiLabel: vscode.l10n.t("xlflow run {target}", { target }),
    workspaceFolder,
  });
}

async function runTestProcedureFromCodeLens(
  value: unknown,
  channels: XlflowChannels,
): Promise<void> {
  const args = normalizeRunProcedureArgs(value);
  if (args === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow CodeLens received invalid test arguments."),
    );
    return;
  }
  const moduleName = readNonEmpty(args.moduleName);
  if (moduleName === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow CodeLens received invalid test arguments."),
    );
    return;
  }
  const uri = vscode.Uri.parse(args.uri);
  if (!(await saveDirtyDocumentIfNeeded(uri))) {
    return;
  }
  const workspaceFolder = await workspaceFolderForUri(uri);
  await runXlflowCommand(
    ["test", "--module", moduleName, "--filter", args.name],
    `xlflow test ${moduleName}.${args.name}`,
    channels.output,
    {
      requireWorkspace: true,
      uiLabel: vscode.l10n.t("xlflow test {target}", { target: `${moduleName}.${args.name}` }),
      workspaceFolder,
    },
  );
}

function normalizeRunProcedureArgs(value: unknown): RunProcedureArgs | undefined {
  if (isRunProcedureArgs(value)) {
    return value;
  }
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  const obj = value as Record<string, unknown>;
  const uri = obj.uri;
  const name = readNonEmpty(obj.name);
  if (!(uri instanceof vscode.Uri) || name === undefined) {
    const test = obj.test;
    if (typeof test !== "object" || test === null) {
      return undefined;
    }
    const testObj = test as Record<string, unknown>;
    const testName = readNonEmpty(testObj.name);
    const testModule = readNonEmpty(testObj.module);
    const testUri = obj.uri;
    if (!(testUri instanceof vscode.Uri) || testName === undefined) {
      return undefined;
    }
    const qualifiedName = readNonEmpty(testObj.qualified_name) ?? `${testModule}.${testName}`;
    const line = typeof testObj.line === "number" ? testObj.line : undefined;
    return {
      uri: testUri.toString(),
      name: testName,
      moduleName: testModule,
      qualifiedName,
      kind: "test",
      line,
    };
  }
  const qualifiedName = readNonEmpty(obj.qualifiedName) ?? name;
  const moduleName = readNonEmpty(obj.moduleName);
  const kind = obj.test === true ? "test" : "sub";
  const line = typeof obj.line === "number" ? obj.line : undefined;
  return { uri: uri.toString(), name, moduleName, qualifiedName, kind, line };
}

function isRunProcedureArgs(value: unknown): value is RunProcedureArgs {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const obj = value as Record<string, unknown>;
  const kind = obj.kind;
  return (
    typeof obj.uri === "string" &&
    obj.uri.trim().length > 0 &&
    typeof obj.name === "string" &&
    obj.name.trim().length > 0 &&
    (obj.moduleName === undefined || typeof obj.moduleName === "string") &&
    (obj.qualifiedName === undefined || typeof obj.qualifiedName === "string") &&
    (kind === undefined || kind === "sub" || kind === "test") &&
    (obj.line === undefined || typeof obj.line === "number")
  );
}

async function saveDirtyDocumentIfNeeded(uri: vscode.Uri): Promise<boolean> {
  if (!readConfig().runSaveBeforeRun) {
    return true;
  }
  const document = vscode.workspace.textDocuments.find(
    (candidate) => candidate.uri.toString() === uri.toString(),
  );
  if (document === undefined || !document.isDirty) {
    return true;
  }
  const saved = await document.save();
  if (!saved) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow run was cancelled because the VBA document was not saved."),
    );
  }
  return saved;
}

async function workspaceFolderForUri(uri: vscode.Uri): Promise<vscode.WorkspaceFolder | undefined> {
  return (
    vscode.workspace.getWorkspaceFolder(uri) ??
    (await resolveWorkspaceRoot({ prompt: true, fallbackToFirst: true }))
  );
}

async function workspaceFolderForUserForm(
  value: unknown,
  sourceUri: vscode.Uri | undefined,
): Promise<vscode.WorkspaceFolder | undefined> {
  if (sourceUri !== undefined) {
    return workspaceFolderForUri(sourceUri);
  }
  const workspaceUri = treeWorkspaceUri(value);
  if (workspaceUri !== undefined) {
    const folder = vscode.workspace.getWorkspaceFolder(workspaceUri);
    if (folder !== undefined) {
      return folder;
    }
  }
  return resolveWorkspaceRoot({ prompt: true, fallbackToFirst: true });
}

async function runModuleMutation(
  args: string[],
  label: string,
  outputChannel: vscode.OutputChannel,
  workspaceFolder: vscode.WorkspaceFolder | undefined,
): Promise<{ exitCode: number; stderr: string; json?: XlflowMutationEnvelope }> {
  return runXlflowJsonCommand<XlflowMutationEnvelope>(
    ["--json", "module", ...args],
    label,
    outputChannel,
    { requireWorkspace: true, workspaceFolder },
  );
}

function renamedModuleUri(
  workspaceFolder: vscode.WorkspaceFolder,
  envelope: XlflowMutationEnvelope | undefined,
  newName: string,
): vscode.Uri | undefined {
  const renamed = envelope?.source?.renamed ?? [];
  const preferred =
    renamed.find((candidate) => basenameWithoutExtension(candidate) === newName) ?? renamed[0];
  if (preferred === undefined) {
    return undefined;
  }
  if (path.isAbsolute(preferred)) {
    return vscode.Uri.file(preferred);
  }
  return vscode.Uri.joinPath(workspaceFolder.uri, ...preferred.replace(/\\/g, "/").split("/"));
}

function basenameWithoutExtension(filePath: string): string {
  const base = filePath.replace(/\\/g, "/").split("/").pop() ?? filePath;
  const dot = base.lastIndexOf(".");
  return dot === -1 ? base : base.slice(0, dot);
}

function showMutationFailure(
  operation: "rename" | "delete",
  moduleName: string,
  result: { exitCode: number; stderr: string; json?: XlflowMutationEnvelope },
  targetKind = "module",
): void {
  const message =
    readNonEmpty(result.json?.error?.message) ??
    readNonEmpty(result.stderr.split(/\r?\n/).find((line) => line.trim().length > 0)) ??
    vscode.l10n.t("xlflow exited with code {exitCode}.", { exitCode: result.exitCode });
  vscode.window.showErrorMessage(
    vscode.l10n.t('Failed to {operation} {targetKind} "{moduleName}": {message}', {
      operation: mutationOperationLabel(operation),
      targetKind: mutationTargetKindLabel(targetKind),
      moduleName,
      message,
    }),
  );
}

function mutationOperationLabel(operation: "rename" | "delete"): string {
  return operation === "rename" ? vscode.l10n.t("rename") : vscode.l10n.t("delete");
}

function mutationTargetKindLabel(targetKind: string): string {
  return targetKind === "UserForm" ? vscode.l10n.t("UserForm") : vscode.l10n.t("module");
}

async function copyText(localizedLabel: string, value: string | undefined): Promise<void> {
  if (value === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("xlflow could not determine the {label}.", {
        label: localizedLabel,
      }),
    );
    return;
  }
  await vscode.env.clipboard.writeText(value);
  vscode.window.showInformationMessage(vscode.l10n.t("{label} copied.", { label: localizedLabel }));
}

function relativePathForUri(uri: vscode.Uri): string {
  const folder = vscode.workspace.getWorkspaceFolder(uri);
  if (folder === undefined) {
    return vscode.workspace.asRelativePath(uri, false).replace(/\\/g, "/");
  }
  return path.relative(folder.uri.fsPath, uri.fsPath).replace(/\\/g, "/");
}

function relativePathFromUri(uri: vscode.Uri | undefined): string | undefined {
  return uri === undefined ? undefined : relativePathForUri(uri);
}

function readNonEmpty(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}

function treeName(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  return readNonEmpty((value as Record<string, unknown>).name);
}

function treeQualifiedName(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  return readNonEmpty((value as Record<string, unknown>).qualifiedName);
}

function userFormRelativePath(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  const obj = value as Record<string, unknown>;
  return readNonEmpty(obj.primaryRelativePath) ?? readNonEmpty(obj.relativePath);
}

function userFormSourceUri(value: unknown): vscode.Uri | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  const obj = value as Record<string, unknown>;
  const primaryUri = obj.primaryUri;
  if (primaryUri instanceof vscode.Uri) {
    return primaryUri;
  }
  return treeUri(value);
}

function treeWorkspaceUri(value: unknown): vscode.Uri | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  const workspaceUri = (value as Record<string, unknown>).workspaceUri;
  return workspaceUri instanceof vscode.Uri ? workspaceUri : undefined;
}

function treeModuleKind(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  return readNonEmpty((value as Record<string, unknown>).moduleKind);
}

function treeUri(value: unknown): vscode.Uri | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  const uri = (value as Record<string, unknown>).uri;
  return uri instanceof vscode.Uri ? uri : undefined;
}

function treeLine(value: unknown): number | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  const line = (value as Record<string, unknown>).line;
  return typeof line === "number" ? line : undefined;
}
