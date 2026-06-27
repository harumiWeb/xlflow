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
      });
    }),
    vscode.commands.registerCommand("xlflow.newProject", async () => {
      const workbook = await vscode.window.showInputBox({
        title: "xlflow: New Project",
        prompt: "Workbook filename or project name. Leave empty to use xlflow's default.",
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
      });
      if (code === 0) {
        await hooks.refreshAll();
      }
    }),
    vscode.commands.registerCommand("xlflow.initProject", async () => {
      const workbooks = await vscode.window.showOpenDialog({
        title: "xlflow: Initialize Project",
        canSelectFiles: true,
        canSelectFolders: false,
        canSelectMany: false,
        filters: {
          "Excel workbooks": ["xlsm", "xlsb", "xlsx", "xls"],
          "All files": ["*"],
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
      const confirmed = await vscode.window.showWarningMessage(
        "Push sources to workbook?",
        { modal: true },
        "Push",
      );
      if (confirmed !== "Push") {
        return;
      }
      const code = await runXlflowCommand(["push"], "xlflow push", channels.output, {
        requireWorkspace: true,
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
      await runXlflowCommand(["run"], "xlflow run", channels.output, { requireWorkspace: true });
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
      });
      if (code === 0) {
        await hooks.refreshTests();
      }
    }),
    vscode.commands.registerCommand("xlflow.lintWorkspace", async () => {
      await runXlflowCommand(["lint"], "xlflow lint", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.formatDocument", async () => {
      const editor = vscode.window.activeTextEditor;
      if (editor === undefined) {
        vscode.window.showWarningMessage("xlflow format document requires an active editor.");
        return;
      }
      await vscode.commands.executeCommand("editor.action.formatDocument");
    }),
    vscode.commands.registerCommand("xlflow.formatProject", async () => {
      await runXlflowCommand(["fmt", "--write"], "xlflow fmt", channels.output, {
        requireWorkspace: true,
      });
    }),
    vscode.commands.registerCommand("xlflow.saveWorkbook", async () => {
      await runXlflowCommand(["save"], "xlflow save", channels.output, { requireWorkspace: true });
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
      await copyText("Module name", treeName(value));
    }),
    vscode.commands.registerCommand("xlflow.copyRelativePath", async (value: unknown) => {
      const uri = treeUri(value);
      await copyText("Relative path", uri === undefined ? undefined : relativePathForUri(uri));
    }),
    vscode.commands.registerCommand("xlflow.copyProcedureName", async (value: unknown) => {
      await copyText("Procedure name", treeName(value));
    }),
    vscode.commands.registerCommand("xlflow.copyQualifiedName", async (value: unknown) => {
      await copyText("Qualified name", treeQualifiedName(value));
    }),
    vscode.commands.registerCommand("xlflow.renameUserForm", async (value: unknown) => {
      await renameUserForm(value, channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.deleteUserForm", async (value: unknown) => {
      await deleteUserForm(value, channels, hooks);
    }),
    vscode.commands.registerCommand("xlflow.revealUserFormSource", async (value: unknown) => {
      const uri = treeUri(value);
      if (uri !== undefined) {
        await vscode.commands.executeCommand("revealInExplorer", uri);
      }
    }),
    vscode.commands.registerCommand("xlflow.copyUserFormName", async (value: unknown) => {
      await copyText("UserForm name", treeName(value));
    }),
    vscode.commands.registerCommand("xlflow.copyUserFormRelativePath", async (value: unknown) => {
      await copyText("Relative path", treeRelativePath(value) ?? relativePathFromTreeUri(value));
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
      { label: "New Project", command: "xlflow.newProject" },
      { label: "Init Existing Workbook", command: "xlflow.initProject" },
      { label: "Run Doctor", command: "xlflow.runDoctor" },
      { label: "Open Documentation", command: "xlflow.openDocumentation" },
    ],
    { title: "xlflow Project Setup", placeHolder: "Select a setup action" },
  );
  if (action !== undefined) {
    await vscode.commands.executeCommand(action.command);
  }
}

async function installAgentSkill(channels: XlflowChannels): Promise<void> {
  const workspaceFolder = await resolveWorkspaceRoot({ prompt: true });
  if (workspaceFolder === undefined) {
    vscode.window.showWarningMessage("xlflow skill install requires an open workspace folder.");
    return;
  }
  const provider = await vscode.window.showQuickPick(
    [
      { label: "codex", description: "Install for Codex" },
      { label: "claude", description: "Install for Claude Code" },
      { label: "cursor", description: "Install for Cursor" },
      { label: "gemini", description: "Install for Gemini CLI" },
      { label: "agents", description: "Install shared .agents instructions" },
    ],
    {
      title: "xlflow: Install Agent Skill",
      placeHolder: "Select the agent provider target",
    },
  );
  if (provider === undefined) {
    return;
  }

  const overwrite = await vscode.window.showQuickPick(
    [
      {
        label: "Install without overwrite",
        description: "Fail if the xlflow skill already exists",
        force: false,
      },
      {
        label: "Overwrite existing installation",
        description: "Pass --force to replace an existing xlflow skill",
        force: true,
      },
    ],
    {
      title: "xlflow: Install Agent Skill",
      placeHolder: "Choose overwrite behavior",
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
    workspaceFolder,
  });
}

async function installHelperModules(channels: XlflowChannels): Promise<void> {
  const workspaceFolder = await resolveWorkspaceRoot({ prompt: true });
  if (workspaceFolder === undefined) {
    vscode.window.showWarningMessage("xlflow module install requires an open workspace folder.");
    return;
  }
  const mode = await vscode.window.showQuickPick(
    [
      {
        label: "Install to source only",
        description: "Run xlflow module install",
        args: ["module", "install"],
      },
      {
        label: "Install and push to workbook",
        description: "Run xlflow module install --push",
        args: ["module", "install", "--push"],
      },
    ],
    {
      title: "xlflow: Install Helper Modules",
      placeHolder: "Choose how to install bundled helper modules",
    },
  );
  if (mode === undefined) {
    return;
  }
  await runXlflowCommand(mode.args, `xlflow ${mode.args.join(" ")}`, channels.output, {
    requireWorkspace: true,
    workspaceFolder,
  });
}

async function newModule(channels: XlflowChannels, hooks: CommandRefreshHooks): Promise<void> {
  const moduleType = await vscode.window.showQuickPick(
    [
      {
        label: "Standard Module",
        description: "Run xlflow module new --type standard",
        moduleType: "standard" as const,
        placeholder: "InvoiceProcessor",
      },
      {
        label: "Class Module",
        description: "Run xlflow module new --type class",
        moduleType: "class" as const,
        placeholder: "InvoiceService",
      },
    ],
    {
      title: "xlflow: New Module",
      placeHolder: "Select the module type",
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
    vscode.window.showWarningMessage("xlflow module new requires an open workspace folder.");
    return;
  }
  const title =
    moduleType === "standard" ? "xlflow: New Standard Module" : "xlflow: New Class Module";
  const name = await promptComponentName(title, placeholder ?? "InvoiceProcessor");
  if (name === undefined) {
    return;
  }
  const args = ["module", "new", name, "--type", moduleType];
  const code = await runXlflowCommand(args, `xlflow module new ${name}`, channels.output, {
    requireWorkspace: true,
    workspaceFolder,
  });
  if (code === 0) {
    await hooks.refreshModules();
  }
}

async function newUserForm(channels: XlflowChannels, hooks: CommandRefreshHooks): Promise<void> {
  const workspaceFolder = await resolveWorkspaceRoot({ prompt: true });
  if (workspaceFolder === undefined) {
    vscode.window.showWarningMessage("xlflow form new requires an open workspace folder.");
    return;
  }
  const name = await promptComponentName("xlflow: New UserForm", "CustomerForm");
  if (name === undefined) {
    return;
  }
  const code = await runXlflowCommand(
    ["form", "new", name],
    `xlflow form new ${name}`,
    channels.output,
    {
      requireWorkspace: true,
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
  const uri = treeUri(value);
  if (name === undefined) {
    vscode.window.showWarningMessage("xlflow rename UserForm received invalid UserForm arguments.");
    return;
  }
  const newName = await promptComponentName("xlflow: Rename UserForm", name);
  if (newName === undefined || newName.trim() === name) {
    return;
  }

  const workspaceFolder =
    uri === undefined
      ? await resolveWorkspaceRoot({ prompt: true, fallbackToFirst: true })
      : await workspaceFolderForUri(uri);
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
  const uri = treeUri(value);
  if (name === undefined) {
    vscode.window.showWarningMessage("xlflow delete UserForm received invalid UserForm arguments.");
    return;
  }

  const confirmed = await vscode.window.showWarningMessage(
    `Delete UserForm "${name}" from the xlflow project?\n\nThis may remove the .frm, .frx, sidecar code, and designer spec artifacts. The workbook will be updated on the next xlflow push.`,
    { modal: true },
    "Delete",
  );
  if (confirmed !== "Delete") {
    return;
  }

  const workspaceFolder =
    uri === undefined
      ? await resolveWorkspaceRoot({ prompt: true, fallbackToFirst: true })
      : await workspaceFolderForUri(uri);
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
    vscode.window.showWarningMessage("xlflow rename module received invalid module arguments.");
    return;
  }
  if (treeModuleKind(value) === "document") {
    vscode.window.showErrorMessage(
      `Failed to rename module "${moduleName}": document modules cannot be renamed.`,
    );
    return;
  }

  const newName = await promptComponentName("xlflow: Rename Module", moduleName);
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
    vscode.window.showWarningMessage("xlflow delete module received invalid module arguments.");
    return;
  }
  if (treeModuleKind(value) === "document") {
    vscode.window.showErrorMessage(
      `Failed to delete module "${moduleName}": document modules cannot be removed.`,
    );
    return;
  }

  const confirmed = await vscode.window.showWarningMessage(
    `Delete module "${moduleName}" from the xlflow project?\n\nThis removes the source file. The workbook will be updated on the next xlflow push.`,
    { modal: true },
    "Delete",
  );
  if (confirmed !== "Delete") {
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
    prompt: "Enter a VBA component name without a file extension or path.",
    placeHolder,
    validateInput: validateComponentNameInput,
  });
}

function validateComponentNameInput(value: string): string | undefined {
  const name = value.trim();
  if (name.length === 0) {
    return "Name is required.";
  }
  if (/[\\/]/.test(name) || name.includes("..")) {
    return "Use a component name, not a path.";
  }
  if (/\.(bas|cls|frm)$/i.test(name)) {
    return "Do not include a file extension.";
  }
  return undefined;
}

async function runProcedure(value: unknown, channels: XlflowChannels): Promise<void> {
  const args = normalizeRunProcedureArgs(value);
  if (args === undefined) {
    vscode.window.showWarningMessage("xlflow received invalid run procedure arguments.");
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
    workspaceFolder,
  });
}

async function runTestProcedureFromCodeLens(
  value: unknown,
  channels: XlflowChannels,
): Promise<void> {
  const args = normalizeRunProcedureArgs(value);
  if (args === undefined) {
    vscode.window.showWarningMessage("xlflow CodeLens received invalid test arguments.");
    return;
  }
  const moduleName = readNonEmpty(args.moduleName);
  if (moduleName === undefined) {
    vscode.window.showWarningMessage("xlflow CodeLens received invalid test arguments.");
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
    { requireWorkspace: true, workspaceFolder },
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
      "xlflow run was cancelled because the VBA document was not saved.",
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
    `xlflow exited with code ${result.exitCode}.`;
  vscode.window.showErrorMessage(
    `Failed to ${operation} ${targetKind} "${moduleName}": ${message}`,
  );
}

async function copyText(label: string, value: string | undefined): Promise<void> {
  if (value === undefined) {
    vscode.window.showWarningMessage(`xlflow could not determine the ${label.toLowerCase()}.`);
    return;
  }
  await vscode.env.clipboard.writeText(value);
  vscode.window.showInformationMessage(`${label} copied.`);
}

function relativePathForUri(uri: vscode.Uri): string {
  const folder = vscode.workspace.getWorkspaceFolder(uri);
  if (folder === undefined) {
    return vscode.workspace.asRelativePath(uri, false).replace(/\\/g, "/");
  }
  return path.relative(folder.uri.fsPath, uri.fsPath).replace(/\\/g, "/");
}

function relativePathFromTreeUri(value: unknown): string | undefined {
  const uri = treeUri(value);
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

function treeRelativePath(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  return readNonEmpty((value as Record<string, unknown>).relativePath);
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
