import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "./client";
import { XlflowChannels } from "./logging";
import { runXlflowCommand } from "./xlflow";

export function registerCommands(
  context: vscode.ExtensionContext,
  clientManager: XlflowLanguageClientManager,
  channels: XlflowChannels,
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
    vscode.commands.registerCommand("xlflow.pull", async () => {
      await runXlflowCommand(["pull"], "xlflow pull", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.push", async () => {
      await runXlflowCommand(["push"], "xlflow push", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.runMacro", async () => {
      await runXlflowCommand(["run"], "xlflow run", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.test", async () => {
      await runXlflowCommand(["test"], "xlflow test", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.saveWorkbook", async () => {
      await runXlflowCommand(["save"], "xlflow save", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.sessionStart", async () => {
      await runXlflowCommand(["session", "start"], "xlflow session start", channels.output, {
        requireWorkspace: true,
      });
    }),
    vscode.commands.registerCommand("xlflow.sessionStatus", async () => {
      await runXlflowCommand(["session", "status"], "xlflow session status", channels.output, {
        requireWorkspace: true,
      });
    }),
    vscode.commands.registerCommand("xlflow.sessionStop", async () => {
      await runXlflowCommand(["session", "stop"], "xlflow session stop", channels.output, {
        requireWorkspace: true,
      });
    }),
  );
}
