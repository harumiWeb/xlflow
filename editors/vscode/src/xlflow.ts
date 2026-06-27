import * as childProcess from "child_process";
import * as vscode from "vscode";
import { XlflowCliAvailabilityService } from "./cliAvailability";
import { readConfig } from "./config";
import { appendProcessOutput } from "./logging";

let cliAvailabilityService: XlflowCliAvailabilityService | undefined;

export function setXlflowCliAvailabilityService(
  service: XlflowCliAvailabilityService | undefined,
): void {
  cliAvailabilityService = service;
}

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
  if (activeDocument !== undefined) {
    const containingFolder = vscode.workspace.getWorkspaceFolder(activeDocument.uri);
    if (containingFolder !== undefined) {
      return containingFolder;
    }
  }

  if (options.prompt) {
    return vscode.window.showWorkspaceFolderPick({
      placeHolder: vscode.l10n.t("Select the workspace folder for xlflow commands"),
    });
  }

  return options.fallbackToFirst ? folders[0] : undefined;
}

export async function runXlflowCommand(
  args: string[],
  label: string,
  outputChannel: vscode.OutputChannel,
  options: {
    requireWorkspace: boolean;
    notify?: boolean;
    showOutput?: boolean;
    showCliUnavailable?: boolean;
    uiLabel?: string;
    workspaceFolder?: vscode.WorkspaceFolder;
  },
): Promise<number> {
  const uiLabel = options.uiLabel ?? label;
  const folder =
    options.workspaceFolder ??
    (await resolveWorkspaceRoot({
      prompt: options.requireWorkspace,
      fallbackToFirst: !options.requireWorkspace,
    }));
  if (options.requireWorkspace && folder === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("{label} requires an open workspace folder.", {
        label: uiLabel,
      }),
    );
    return -1;
  }

  if (!(await ensureCliAvailable(options.showCliUnavailable !== false))) {
    return -1;
  }

  const config = readConfig();
  const cwd = folder?.uri.fsPath;
  if (options.showOutput !== false) {
    outputChannel.show(true);
  }
  outputChannel.appendLine(
    `> ${config.path} ${args.join(" ")}${cwd === undefined ? "" : ` (cwd: ${cwd})`}`,
  );
  const notify = options.notify !== false;

  const run = new Promise<number>((resolve) => {
    let settled = false;
    const settle = (exitCode: number): void => {
      if (settled) {
        return;
      }
      settled = true;
      resolve(exitCode);
    };
    const child = childProcess.spawn(config.path, args, {
      cwd,
      windowsHide: true,
    });

    child.stdout.on("data", (data: Buffer) => appendProcessOutput(outputChannel, "stdout", data));
    child.stderr.on("data", (data: Buffer) => appendProcessOutput(outputChannel, "stderr", data));
    child.on("error", (error) => {
      outputChannel.appendLine(`[error] ${error.message}`);
      vscode.window.showErrorMessage(
        vscode.l10n.t("{label} failed: {message}", {
          label: uiLabel,
          message: error.message,
        }),
      );
      settle(-1);
    });
    child.on("close", (code) => {
      if (settled) {
        return;
      }
      const exitCode = code ?? -1;
      outputChannel.appendLine(`${label} exited with code ${exitCode}`);
      if (!notify) {
        settle(exitCode);
        return;
      }
      if (exitCode === 0) {
        vscode.window.showInformationMessage(
          vscode.l10n.t("{label} completed.", { label: uiLabel }),
        );
      } else {
        vscode.window.showErrorMessage(
          vscode.l10n.t("{label} failed with exit code {exitCode}.", {
            label: uiLabel,
            exitCode,
          }),
        );
      }
      settle(exitCode);
    });
  });
  if (!notify) {
    return run;
  }
  return vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: uiLabel,
      cancellable: false,
    },
    () => run,
  );
}

export interface XlflowJsonCommandResult<T> {
  exitCode: number;
  stdout: string;
  stderr: string;
  json?: T;
}

export async function runXlflowJsonCommand<T>(
  args: string[],
  label: string,
  outputChannel: vscode.OutputChannel,
  options: {
    requireWorkspace: boolean;
    showCliUnavailable?: boolean;
    uiLabel?: string;
    workspaceFolder?: vscode.WorkspaceFolder;
  },
): Promise<XlflowJsonCommandResult<T>> {
  const uiLabel = options.uiLabel ?? label;
  const folder =
    options.workspaceFolder ??
    (await resolveWorkspaceRoot({
      prompt: options.requireWorkspace,
      fallbackToFirst: !options.requireWorkspace,
    }));
  if (options.requireWorkspace && folder === undefined) {
    vscode.window.showWarningMessage(
      vscode.l10n.t("{label} requires an open workspace folder.", {
        label: uiLabel,
      }),
    );
    return { exitCode: -1, stdout: "", stderr: "" };
  }

  if (!(await ensureCliAvailable(options.showCliUnavailable !== false))) {
    return { exitCode: -1, stdout: "", stderr: "xlflow CLI is unavailable." };
  }

  const config = readConfig();
  const cwd = folder?.uri.fsPath;
  outputChannel.appendLine(
    `> ${config.path} ${args.join(" ")}${cwd === undefined ? "" : ` (cwd: ${cwd})`}`,
  );

  return new Promise((resolve) => {
    let settled = false;
    const settle = (result: XlflowJsonCommandResult<T>): void => {
      if (settled) {
        return;
      }
      settled = true;
      resolve(result);
    };
    const stdoutChunks: Buffer[] = [];
    const stderrChunks: Buffer[] = [];
    const child = childProcess.spawn(config.path, args, {
      cwd,
      windowsHide: true,
    });

    child.stdout.on("data", (data: Buffer) => {
      stdoutChunks.push(data);
      appendProcessOutput(outputChannel, "stdout", data);
    });
    child.stderr.on("data", (data: Buffer) => {
      stderrChunks.push(data);
      appendProcessOutput(outputChannel, "stderr", data);
    });
    child.on("error", (error) => {
      outputChannel.appendLine(`[error] ${error.message}`);
      settle({ exitCode: -1, stdout: "", stderr: error.message });
    });
    child.on("close", (code) => {
      if (settled) {
        return;
      }
      const exitCode = code ?? -1;
      const stdout = Buffer.concat(stdoutChunks).toString("utf8");
      const stderr = Buffer.concat(stderrChunks).toString("utf8");
      let json: T | undefined;
      if (stdout.trim() !== "") {
        try {
          json = JSON.parse(stdout) as T;
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          outputChannel.appendLine(`[error] Failed to parse ${label} JSON: ${message}`);
        }
      }
      outputChannel.appendLine(`${label} exited with code ${exitCode}`);
      settle({ exitCode, stdout, stderr, json });
    });
  });
}

async function ensureCliAvailable(showActions: boolean): Promise<boolean> {
  if (cliAvailabilityService === undefined) {
    return true;
  }
  return cliAvailabilityService.ensureAvailable({ showActions });
}
