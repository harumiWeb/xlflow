import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "./client";
import { readConfig } from "./config";
import { XlflowChannels } from "./logging";
import { SessionManager } from "./session";
import { resolveWorkspaceRoot, runXlflowCommand } from "./xlflow";

type RunProcedureArgs = {
  uri: string;
  name: string;
  moduleName?: string;
  qualifiedName?: string;
  kind?: "sub" | "test";
  line?: number;
};

export function registerCommands(
  context: vscode.ExtensionContext,
  clientManager: XlflowLanguageClientManager,
  channels: XlflowChannels,
  sessionManager: SessionManager,
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
      await runXlflowCommand(args, "xlflow new", channels.output, { requireWorkspace: true });
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
      await runXlflowCommand(["init", workbook.fsPath], "xlflow init", channels.output, {
        requireWorkspace: true,
      });
    }),
    vscode.commands.registerCommand("xlflow.skillInstall", async () => {
      await installAgentSkill(channels);
    }),
    vscode.commands.registerCommand("xlflow.moduleInstall", async () => {
      await installHelperModules(channels);
    }),
    vscode.commands.registerCommand("xlflow.pull", async () => {
      await runXlflowCommand(["pull"], "xlflow pull", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.push", async () => {
      await runXlflowCommand(["push"], "xlflow push", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.runMacro", async () => {
      await runXlflowCommand(["run"], "xlflow run", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.runProcedure", async (args: unknown) => {
      await runProcedureFromCodeLens(args, channels);
    }),
    vscode.commands.registerCommand("xlflow.runTestProcedure", async (args: unknown) => {
      await runTestProcedureFromCodeLens(args, channels);
    }),
    vscode.commands.registerCommand("xlflow.test", async () => {
      await runXlflowCommand(["test"], "xlflow test", channels.output, { requireWorkspace: true });
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
    }),
    vscode.commands.registerCommand("xlflow.sessionStatus", async () => {
      await sessionManager.showStatus();
    }),
    vscode.commands.registerCommand("xlflow.sessionStop", async () => {
      await sessionManager.stop();
    }),
    vscode.commands.registerCommand("xlflow.sessionRestart", async () => {
      await sessionManager.restart();
    }),
    vscode.commands.registerCommand("xlflow.sessionActions", async () => {
      await sessionManager.showActions();
    }),
    vscode.commands.registerCommand("xlflow.openOutput", () => {
      sessionManager.openOutput();
    }),
  );
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

async function runProcedureFromCodeLens(value: unknown, channels: XlflowChannels): Promise<void> {
  if (!isRunProcedureArgs(value)) {
    vscode.window.showWarningMessage("xlflow CodeLens received invalid run arguments.");
    return;
  }
  const uri = vscode.Uri.parse(value.uri);
  if (!(await saveDirtyDocumentIfNeeded(uri))) {
    return;
  }
  const workspaceFolder = await workspaceFolderForUri(uri);
  const target = readNonEmpty(value.qualifiedName) ?? value.name;
  await runXlflowCommand(["run", target], `xlflow run ${target}`, channels.output, {
    requireWorkspace: true,
    workspaceFolder,
  });
}

async function runTestProcedureFromCodeLens(
  value: unknown,
  channels: XlflowChannels,
): Promise<void> {
  if (!isRunProcedureArgs(value)) {
    vscode.window.showWarningMessage("xlflow CodeLens received invalid test arguments.");
    return;
  }
  const moduleName = readNonEmpty(value.moduleName);
  if (moduleName === undefined) {
    vscode.window.showWarningMessage("xlflow CodeLens received invalid test arguments.");
    return;
  }
  const uri = vscode.Uri.parse(value.uri);
  if (!(await saveDirtyDocumentIfNeeded(uri))) {
    return;
  }
  const workspaceFolder = await workspaceFolderForUri(uri);
  await runXlflowCommand(
    ["test", "--module", moduleName, "--filter", value.name],
    `xlflow test ${moduleName}.${value.name}`,
    channels.output,
    { requireWorkspace: true, workspaceFolder },
  );
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

function readNonEmpty(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}
