import * as childProcess from "child_process";
import * as vscode from "vscode";
import { readConfig } from "./config";

export type XlflowCliAvailability =
  | { ok: true; executable: string; version?: string }
  | { ok: false; reason: "notFound" | "failed"; executable: string; message: string };

export interface AvailabilityFailureInput {
  code?: unknown;
  message?: unknown;
  signal?: unknown;
  stdout?: string;
  stderr?: string;
  timedOut?: boolean;
}

const installGuideUri = vscode.Uri.parse("https://harumiweb.github.io/xlflow/installation");
const checkTimeoutMs = 5000;
const shownProjectNoticeKeys = new Set<string>();

export class XlflowCliAvailabilityService implements vscode.Disposable {
  private readonly emitter = new vscode.EventEmitter<XlflowCliAvailability>();
  private availability: XlflowCliAvailability | undefined;
  private pending: Promise<XlflowCliAvailability> | undefined;

  readonly onDidChangeAvailability = this.emitter.event;

  current(): XlflowCliAvailability | undefined {
    return this.availability;
  }

  async refresh(): Promise<XlflowCliAvailability> {
    if (this.pending !== undefined) {
      return this.pending;
    }
    this.pending = checkXlflowAvailability()
      .then((availability) => {
        this.availability = availability;
        this.emitter.fire(availability);
        return availability;
      })
      .finally(() => {
        this.pending = undefined;
      });
    return this.pending;
  }

  async ensureAvailable(options: { showActions?: boolean } = {}): Promise<boolean> {
    const availability = await this.refresh();
    if (availability.ok) {
      return true;
    }
    if (options.showActions === false) {
      return false;
    }
    return showCliUnavailableActions(availability, () => this.refresh());
  }

  dispose(): void {
    this.emitter.dispose();
  }
}

export async function checkXlflowAvailability(): Promise<XlflowCliAvailability> {
  const executable = readConfig().path;
  try {
    const result = await execFileWithTimeout(executable, ["--version"], checkTimeoutMs);
    if (result.exitCode === 0) {
      return normalizeAvailabilitySuccess(executable, result.stdout, result.stderr);
    }
    return normalizeAvailabilityFailure(executable, {
      code: result.exitCode,
      stdout: result.stdout,
      stderr: result.stderr,
      message: vscode.l10n.t("xlflow --version exited with code {exitCode}.", {
        exitCode: result.exitCode,
      }),
    });
  } catch (error) {
    return normalizeAvailabilityFailure(executable, errorToFailureInput(error));
  }
}

export function normalizeAvailabilitySuccess(
  executable: string,
  stdout: string,
  stderr: string,
): XlflowCliAvailability {
  return {
    ok: true,
    executable,
    version: readNonEmpty(stdout) ?? readNonEmpty(stderr),
  };
}

export function normalizeAvailabilityFailure(
  executable: string,
  failure: AvailabilityFailureInput,
): XlflowCliAvailability {
  if (failure.code === "ENOENT") {
    return {
      ok: false,
      reason: "notFound",
      executable,
      message: vscode.l10n.t("xlflow executable was not found."),
    };
  }
  if (failure.timedOut === true || failure.signal === "SIGTERM") {
    return {
      ok: false,
      reason: "failed",
      executable,
      message: vscode.l10n.t("xlflow --version timed out."),
    };
  }
  const detail =
    readNonEmpty(failure.stderr) ??
    readNonEmpty(failure.stdout) ??
    (typeof failure.message === "string" ? failure.message : undefined) ??
    vscode.l10n.t("Failed to run xlflow.");
  return { ok: false, reason: "failed", executable, message: detail };
}

export function cliNotificationSuppressionKey(
  workspaceUri: vscode.Uri,
  availability: XlflowCliAvailability,
): string {
  return `xlflow.cliMissingNotice.${workspaceUri.toString()}.${availability.executable}`;
}

export async function showCliUnavailableActions(
  availability: Extract<XlflowCliAvailability, { ok: false }>,
  retry: () => Promise<XlflowCliAvailability>,
): Promise<boolean> {
  const installGuide = vscode.l10n.t("Install Guide");
  const configurePath = vscode.l10n.t("Configure Path");
  const retryLabel = vscode.l10n.t("Retry");
  const action = await vscode.window.showErrorMessage(
    cliUnavailableMessage(availability),
    installGuide,
    configurePath,
    retryLabel,
  );
  if (action === installGuide) {
    await openInstallGuide();
    return false;
  }
  if (action === configurePath) {
    await configureXlflowPath();
    return false;
  }
  if (action === retryLabel) {
    return (await retry()).ok;
  }
  return false;
}

export async function showProjectCliUnavailableNotice(
  context: vscode.ExtensionContext,
  workspaceFolder: vscode.WorkspaceFolder,
  availability: XlflowCliAvailability,
): Promise<void> {
  if (availability.ok) {
    return;
  }
  const key = cliNotificationSuppressionKey(workspaceFolder.uri, availability);
  if (shownProjectNoticeKeys.has(key) || context.workspaceState.get<boolean>(key) === true) {
    return;
  }
  shownProjectNoticeKeys.add(key);
  const installGuide = vscode.l10n.t("Install Guide");
  const configurePath = vscode.l10n.t("Configure Path");
  const retryLabel = vscode.l10n.t("Retry");
  const dontShow = vscode.l10n.t("Don't Show Again for This Workspace");
  const action = await vscode.window.showWarningMessage(
    cliUnavailableMessage(availability),
    installGuide,
    configurePath,
    retryLabel,
    dontShow,
  );
  if (action === installGuide) {
    await openInstallGuide();
  } else if (action === configurePath) {
    await configureXlflowPath();
  } else if (action === retryLabel) {
    await vscode.commands.executeCommand("xlflow.retryCliDetection");
  } else if (action === dontShow) {
    await context.workspaceState.update(key, true);
  }
}

export function cliUnavailableMessage(
  availability: Extract<XlflowCliAvailability, { ok: false }>,
): string {
  if (availability.reason === "notFound") {
    return vscode.l10n.t(
      "xlflow CLI was not found. Install xlflow or configure xlflow.path to use this extension.",
    );
  }
  return vscode.l10n.t("xlflow CLI check failed: {message}", { message: availability.message });
}

export async function openInstallGuide(): Promise<void> {
  await vscode.env.openExternal(installGuideUri);
}

export async function configureXlflowPath(): Promise<void> {
  await vscode.commands.executeCommand("workbench.action.openSettings", "xlflow.path");
}

function execFileWithTimeout(
  command: string,
  args: string[],
  timeoutMs: number,
): Promise<{ exitCode: number; stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    let settled = false;
    const stdoutChunks: Buffer[] = [];
    const stderrChunks: Buffer[] = [];
    const child = childProcess.spawn(command, args, { windowsHide: true });
    const timer = setTimeout(() => {
      if (settled) {
        return;
      }
      settled = true;
      child.kill();
      reject({ timedOut: true, signal: "SIGTERM" });
    }, timeoutMs);
    const settle = (): void => {
      clearTimeout(timer);
    };
    child.stdout.on("data", (data: Buffer) => stdoutChunks.push(data));
    child.stderr.on("data", (data: Buffer) => stderrChunks.push(data));
    child.on("error", (error) => {
      if (settled) {
        return;
      }
      settled = true;
      settle();
      reject(error);
    });
    child.on("close", (code) => {
      if (settled) {
        return;
      }
      settled = true;
      settle();
      resolve({
        exitCode: code ?? -1,
        stdout: Buffer.concat(stdoutChunks).toString("utf8"),
        stderr: Buffer.concat(stderrChunks).toString("utf8"),
      });
    });
  });
}

function errorToFailureInput(error: unknown): AvailabilityFailureInput {
  if (typeof error !== "object" || error === null) {
    return { message: String(error) };
  }
  const obj = error as Record<string, unknown>;
  return {
    code: obj.code,
    message: obj.message,
    signal: obj.signal,
    timedOut: obj.timedOut === true,
  };
}

function readNonEmpty(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}
