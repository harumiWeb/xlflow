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
    vscode.commands.registerCommand("xlflow.pull", async () => {
      await runXlflowCommand(["pull"], "xlflow pull", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.push", async () => {
      await runXlflowCommand(["push"], "xlflow push", channels.output, { requireWorkspace: true });
    }),
    vscode.commands.registerCommand("xlflow.runMacro", async () => {
      await runXlflowCommand(["run"], "xlflow run", channels.output, { requireWorkspace: true });
    }),
  );
}
