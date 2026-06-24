import * as childProcess from "child_process";
import * as vscode from "vscode";
import { readConfig } from "./config";
import { appendProcessOutput } from "./logging";

export interface WorkspaceRootOptions {
  prompt: boolean;
  fallbackToFirst?: boolean;
}

export async function resolveWorkspaceRoot(
  options: WorkspaceRootOptions,
): Promise<vscode.WorkspaceFolder | undefined> {
  const folders = vscode.workspace.workspaceFolders ?? [];
  if (folders.length === 0) {
    return undefined;
  }
  if (folders.length === 1) {
    return folders[0];
  }

  const activeDocument = vscode.window.activeTextEditor?.document;
  if (activeDocument?.uri.scheme === "file") {
    const containingFolder = vscode.workspace.getWorkspaceFolder(activeDocument.uri);
    if (containingFolder !== undefined) {
      return containingFolder;
    }
  }

  if (options.prompt) {
    return vscode.window.showWorkspaceFolderPick({
      placeHolder: "Select the workspace folder for xlflow commands",
    });
  }

  return options.fallbackToFirst ? folders[0] : undefined;
}

export async function runXlflowCommand(
  args: string[],
  label: string,
  outputChannel: vscode.OutputChannel,
  options: { requireWorkspace: boolean },
): Promise<number> {
  const folder = await resolveWorkspaceRoot({
    prompt: options.requireWorkspace,
    fallbackToFirst: !options.requireWorkspace,
  });
  if (options.requireWorkspace && folder === undefined) {
    vscode.window.showWarningMessage(`${label} requires an open workspace folder.`);
    return -1;
  }

  const config = readConfig();
  const cwd = folder?.uri.fsPath;
  outputChannel.show(true);
  outputChannel.appendLine(
    `> ${config.path} ${args.join(" ")}${cwd === undefined ? "" : ` (cwd: ${cwd})`}`,
  );

  return new Promise((resolve) => {
    const child = childProcess.spawn(config.path, args, {
      cwd,
      windowsHide: true,
    });

    child.stdout.on("data", (data: Buffer) => appendProcessOutput(outputChannel, "stdout", data));
    child.stderr.on("data", (data: Buffer) => appendProcessOutput(outputChannel, "stderr", data));
    child.on("error", (error) => {
      outputChannel.appendLine(`[error] ${error.message}`);
      vscode.window.showErrorMessage(`${label} failed: ${error.message}`);
      resolve(-1);
    });
    child.on("close", (code) => {
      const exitCode = code ?? -1;
      outputChannel.appendLine(`${label} exited with code ${exitCode}`);
      if (exitCode === 0) {
        vscode.window.showInformationMessage(`${label} completed.`);
      } else {
        vscode.window.showErrorMessage(`${label} failed with exit code ${exitCode}.`);
      }
      resolve(exitCode);
    });
  });
}
