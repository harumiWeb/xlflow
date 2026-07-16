import * as childProcess from "child_process";
import * as vscode from "vscode";
import type { XlflowCapabilitiesService, XlflowCapabilityOperation } from "./capabilities";
import { XlflowCliAvailabilityService } from "./cliAvailability";
import { readConfig } from "./config";
import { appendProcessOutput } from "./logging";

let cliAvailabilityService: XlflowCliAvailabilityService | undefined;
let capabilitiesService: XlflowCapabilitiesService | undefined;

export function setXlflowCliAvailabilityService(
  service: XlflowCliAvailabilityService | undefined,
): void {
  cliAvailabilityService = service;
}

export function setXlflowCapabilitiesService(service: XlflowCapabilitiesService | undefined): void {
  capabilitiesService = service;
}

export function xlflowCapabilities(): XlflowCapabilitiesService | undefined {
  return capabilitiesService;
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
    skipCoordination?: boolean;
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

  const commandArgs = ensureJsonArgs(args);
  const operation =
    options.skipCoordination === true ? undefined : await beginManagedCommand(commandArgs);
  if (operation === "blocked") {
    return -1;
  }

  const config = readConfig();
  const cwd = folder?.uri.fsPath;
  if (options.showOutput === true) {
    outputChannel.show(true);
  }
  outputChannel.appendLine(
    `> ${config.path} ${commandArgs.join(" ")}${cwd === undefined ? "" : ` (cwd: ${cwd})`}`,
  );
  const notify = options.notify !== false;
  let jsonResult: XlflowJsonCommandResult<unknown> | undefined;

  const run = new Promise<number>((resolve) => {
    let settled = false;
    const settle = (exitCode: number): void => {
      if (settled) {
        return;
      }
      settled = true;
      resolve(exitCode);
    };
    const child = childProcess.spawn(config.path, commandArgs, {
      cwd,
      windowsHide: true,
    });
    const stdoutChunks: Buffer[] = [];
    const stderrChunks: Buffer[] = [];

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
      showVBAObjectModelAccessNotice(error.message);
      showCommandFailure(
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
      const stdout = Buffer.concat(stdoutChunks).toString("utf8");
      const stderr = Buffer.concat(stderrChunks).toString("utf8");
      const combinedOutput = `${stdout}\n${stderr}`;
      let json: unknown;
      if (stdout.trim() !== "") {
        try {
          json = JSON.parse(stdout) as unknown;
        } catch {
          // Older or non-conforming CLIs retain the existing text-only fallback.
        }
      }
      jsonResult = { exitCode, stdout, stderr, json };
      outputChannel.appendLine(`${label} exited with code ${exitCode}`);
      showVBAObjectModelAccessNotice(combinedOutput);
      if (!notify) {
        settle(exitCode);
        return;
      }
      if (exitCode === 0) {
        vscode.window.showInformationMessage(
          vscode.l10n.t("{label} completed.", { label: uiLabel }),
        );
      } else if (!isWorkbookBusyEnvelope(json)) {
        showCommandFailure(
          vscode.l10n.t("{label} failed with exit code {exitCode}.", {
            label: uiLabel,
            exitCode,
          }),
        );
      }
      settle(exitCode);
    });
  });
  const completed = !notify
    ? await run
    : await vscode.window.withProgress(
        {
          location: vscode.ProgressLocation.Notification,
          title: uiLabel,
          cancellable: false,
        },
        () => run,
      );
  await finishManagedCommand(operation);
  if (completed !== 0 && jsonResult !== undefined) {
    const retryArgs = await workbookBusyRetryArgs(commandArgs, jsonResult);
    if (retryArgs !== undefined) {
      return runXlflowCommand(retryArgs, label, outputChannel, options);
    }
  }
  return completed;
}

export async function runXlflowTerminalCommand(
  args: string[],
  label: string,
  options: {
    requireWorkspace: boolean;
    showCliUnavailable?: boolean;
    uiLabel?: string;
    workspaceFolder?: vscode.WorkspaceFolder;
    terminalName?: string;
  },
): Promise<boolean> {
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
    return false;
  }

  if (!(await ensureCliAvailable(options.showCliUnavailable !== false))) {
    return false;
  }

  if (!((await capabilitiesService?.beforeTerminalCommand(args)) ?? true)) {
    return false;
  }

  const config = readConfig();
  const terminal = vscode.window.createTerminal({
    name: options.terminalName ?? "xlflow",
    cwd: folder?.uri.fsPath,
  });
  terminal.show(true);
  terminal.sendText(buildTerminalCommandLine(config.path, args), true);
  return true;
}

export function buildTerminalCommandLine(executable: string, args: string[]): string {
  return [executable, ...args].map(quoteTerminalArgument).join(" ");
}

function quoteTerminalArgument(value: string): string {
  if (/^[A-Za-z0-9_./:\\-]+$/.test(value)) {
    return value;
  }
  return `"${value.replace(/(["`$])/g, "`$1")}"`;
}

export interface XlflowJsonCommandResult<T> {
  exitCode: number;
  stdout: string;
  stderr: string;
  json?: T;
}

interface XlflowFailureEnvelope {
  error?: {
    code?: unknown;
    message?: unknown;
    details?: unknown;
  };
}

export async function workbookBusyRetryArgs(
  args: string[],
  result: XlflowJsonCommandResult<unknown>,
): Promise<string[] | undefined> {
  const failure = result.json as XlflowFailureEnvelope | undefined;
  const code = failure?.error?.code;
  if (typeof code !== "string" || !code.startsWith("workbook_busy")) {
    return undefined;
  }
  const details = recordValue(failure?.error?.details);
  const owner = recordValue(details?.owner);
  const ownerCommand = stringValue(owner?.command) ?? stringValue(details?.command);
  const operation = stringValue(owner?.operation_kind) ?? stringValue(details?.operation_kind);
  const scope = stringValue(owner?.resource_scope) ?? stringValue(details?.resource_scope);
  const retryableFromError = details?.retryable;
  const capability = await capabilitiesService?.capabilityForArgs(args);
  const retryable =
    typeof retryableFromError === "boolean"
      ? retryableFromError
      : capability?.retryable_when_busy === true;
  const retry = vscode.l10n.t("Retry");
  const wait = vscode.l10n.t("Wait up to 30 seconds");
  const message = workbookBusyMessage(ownerCommand, operation, scope, failure?.error?.message);
  const action = await vscode.window.showErrorMessage(
    message,
    ...(retryable ? [retry, wait] : [retry]),
  );
  if (action === retry) {
    return [...args];
  }
  if (action === wait) {
    return withBusyWaitArgs(args);
  }
  return undefined;
}

function workbookBusyMessage(
  command: string | undefined,
  operation: string | undefined,
  scope: string | undefined,
  fallback: unknown,
): string {
  const parts = [
    command === undefined ? undefined : vscode.l10n.t("Command: {command}", { command }),
    operation === undefined ? undefined : vscode.l10n.t("Operation: {operation}", { operation }),
    scope === undefined ? undefined : vscode.l10n.t("Scope: {scope}", { scope }),
  ].filter((part): part is string => part !== undefined);
  const prefix = vscode.l10n.t("Workbook is busy.");
  const detail = parts.length === 0 ? "" : ` ${parts.join("; ")}`;
  const message = stringValue(fallback);
  return message === undefined ? `${prefix}${detail}` : `${prefix}${detail} ${message}`;
}

export function withBusyWaitArgs(args: string[]): string[] {
  const withoutWait: string[] = [];
  for (let index = 0; index < args.length; index += 1) {
    if (args[index] === "--wait") {
      continue;
    }
    if (args[index] === "--wait-timeout") {
      index += 1;
      continue;
    }
    withoutWait.push(args[index]);
  }
  return [...withoutWait, "--wait", "--wait-timeout", "30s"];
}

function ensureJsonArgs(args: string[]): string[] {
  return args.includes("--json") ? args : ["--json", ...args];
}

function isWorkbookBusyEnvelope(value: unknown): boolean {
  const error = recordValue(value)?.error;
  const code = recordValue(error)?.code;
  return typeof code === "string" && code.startsWith("workbook_busy");
}

function recordValue(value: unknown): Record<string, unknown> | undefined {
  return typeof value === "object" && value !== null
    ? (value as Record<string, unknown>)
    : undefined;
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() !== "" ? value.trim() : undefined;
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
    skipCoordination?: boolean;
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
    return { exitCode: -1, stdout: "", stderr: vscode.l10n.t("xlflow CLI is unavailable.") };
  }

  const operation = options.skipCoordination === true ? undefined : await beginManagedCommand(args);
  if (operation === "blocked") {
    return { exitCode: -1, stdout: "", stderr: "" };
  }

  const config = readConfig();
  const cwd = folder?.uri.fsPath;
  outputChannel.appendLine(
    `> ${config.path} ${args.join(" ")}${cwd === undefined ? "" : ` (cwd: ${cwd})`}`,
  );

  const result = await new Promise<XlflowJsonCommandResult<T>>((resolve) => {
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
      showVBAObjectModelAccessNotice(error.message);
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
      showVBAObjectModelAccessNotice(`${stdout}\n${stderr}`);
      settle({ exitCode, stdout, stderr, json });
    });
  });
  await finishManagedCommand(operation);
  return result;
}

async function beginManagedCommand(
  args: string[],
): Promise<XlflowCapabilityOperation | undefined | "blocked"> {
  return capabilitiesService?.beforeManagedCommand(args);
}

async function finishManagedCommand(
  operation: XlflowCapabilityOperation | undefined | "blocked",
): Promise<void> {
  if (operation === "blocked") {
    return;
  }
  await capabilitiesService?.afterManagedCommand(operation);
}

async function ensureCliAvailable(showActions: boolean): Promise<boolean> {
  if (cliAvailabilityService === undefined) {
    return true;
  }
  return cliAvailabilityService.ensureAvailable({ showActions });
}

function showCommandFailure(message: string): void {
  const openOutput = vscode.l10n.t("Open xlflow Output");
  void vscode.window.showErrorMessage(message, openOutput).then((action) => {
    if (action === openOutput) {
      void vscode.commands.executeCommand("xlflow.openOutput");
    }
  });
}

export function containsVBAObjectModelAccessIssue(text: string): boolean {
  const lower = text.toLowerCase();
  return (
    lower.includes("vba_object_model_access_disabled") ||
    lower.includes("vbproject access is denied") ||
    lower.includes("vbide access is not available") ||
    lower.includes("vbide access unavailable") ||
    lower.includes("trust access to the vba project object model") ||
    lower.includes("get_vbproject failed") ||
    lower.includes("import_vba_components failed") ||
    text.includes("プログラミングによる Visual Basic プロジェクトへのアクセス")
  );
}

function showVBAObjectModelAccessNotice(outputText: string): void {
  if (!containsVBAObjectModelAccessIssue(outputText)) {
    return;
  }
  const openOutput = vscode.l10n.t("Open xlflow Output");
  const runDoctor = vscode.l10n.t("Run Doctor");
  void vscode.window
    .showWarningMessage(
      vscode.l10n.t(
        "Excel is blocking programmatic access to the VBA project object model. Enable Trust access to the VBA project object model in Excel Trust Center, then retry.",
      ),
      openOutput,
      runDoctor,
    )
    .then((action) => {
      if (action === openOutput) {
        void vscode.commands.executeCommand("xlflow.openOutput");
      } else if (action === runDoctor) {
        void vscode.commands.executeCommand("xlflow.runDoctor");
      }
    });
}
