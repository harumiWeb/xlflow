import * as childProcess from "child_process";
import * as vscode from "vscode";
import { openInstallGuide, XlflowCliAvailability } from "./cliAvailability";
import { readConfig } from "./config";

const latestReleaseFallback = vscode.Uri.parse(
  "https://github.com/harumiWeb/xlflow/releases/latest",
);
const updateCheckTimeoutMs = 5000;
const autoCheckIntervalMs = 24 * 60 * 60 * 1000;

export interface XlflowUpdateInfo {
  currentVersion: string;
  latestVersion?: string;
  updateAvailable: boolean;
  releaseUrl?: string;
}

export type XlflowUpdateStatus =
  | { kind: "unknown" }
  | { kind: "checking" }
  | { kind: "upToDate"; info: XlflowUpdateInfo }
  | { kind: "available"; info: XlflowUpdateInfo }
  | { kind: "error"; message: string };

interface UpdateEnvelope {
  status?: string;
  error?: {
    message?: string;
  };
  update?: {
    current_version?: string;
    latest_version?: string;
    update_available?: boolean;
    release_url?: string;
  };
}

interface RawCheckResult {
  exitCode: number;
  stdout: string;
  stderr: string;
}

export class XlflowUpdateService implements vscode.Disposable {
  private readonly emitter = new vscode.EventEmitter<XlflowUpdateStatus>();
  private status: XlflowUpdateStatus = { kind: "unknown" };
  private pending: Promise<XlflowUpdateStatus> | undefined;

  readonly onDidChangeUpdate = this.emitter.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  current(): XlflowUpdateStatus {
    return this.status;
  }

  async checkAutomatic(availability: XlflowCliAvailability | undefined): Promise<void> {
    if (availability === undefined || !availability.ok) {
      this.setStatus({ kind: "unknown" });
      return;
    }
    const currentVersionHint = cliVersionSummary(availability.version);
    if (!this.shouldRunAutomaticCheck(availability, currentVersionHint)) {
      return;
    }
    const status = await this.check({
      manual: false,
      executable: availability.executable,
      currentVersionHint,
    });
    if (status.kind === "available") {
      await this.showUpdateAvailable(status.info, availability.executable);
    }
  }

  async checkManual(availability: XlflowCliAvailability | undefined): Promise<void> {
    if (availability === undefined || !availability.ok) {
      vscode.window.showWarningMessage(vscode.l10n.t("xlflow CLI is unavailable."));
      return;
    }
    const status = await this.check({
      manual: true,
      executable: availability.executable,
      currentVersionHint: cliVersionSummary(availability.version),
    });
    if (status.kind === "available") {
      await this.showUpdateAvailable(status.info, availability.executable, true);
      return;
    }
    if (status.kind === "upToDate") {
      vscode.window.showInformationMessage(
        vscode.l10n.t("xlflow is up to date. Current: {current}.", {
          current: status.info.currentVersion,
        }),
      );
      return;
    }
    if (status.kind === "error") {
      vscode.window.showWarningMessage(
        vscode.l10n.t("Failed to check for xlflow updates: {message}", {
          message: status.message,
        }),
      );
    }
  }

  dispose(): void {
    this.emitter.dispose();
  }

  private async check(options: {
    manual: boolean;
    executable: string;
    currentVersionHint?: string;
  }): Promise<XlflowUpdateStatus> {
    if (this.pending !== undefined) {
      return this.pending;
    }
    this.setStatus({ kind: "checking" });
    this.pending = checkForUpdate()
      .then((result) => normalizeUpdateResult(result))
      .then(async (status) => {
        this.setStatus(status);
        if (!options.manual) {
          await this.context.globalState.update(
            lastCheckedKey(
              options.executable,
              currentVersionFromStatus(status) ?? options.currentVersionHint,
            ),
            Date.now(),
          );
        }
        return status;
      })
      .catch(async (error) => {
        const status: XlflowUpdateStatus = {
          kind: "error",
          message: error instanceof Error ? error.message : String(error),
        };
        this.setStatus(status);
        if (!options.manual) {
          await this.context.globalState.update(
            lastCheckedKey(options.executable, options.currentVersionHint),
            Date.now(),
          );
        }
        return status;
      })
      .finally(() => {
        this.pending = undefined;
      });
    return this.pending;
  }

  private shouldRunAutomaticCheck(
    availability: Extract<XlflowCliAvailability, { ok: true }>,
    currentVersion: string | undefined,
  ) {
    const lastChecked = this.context.globalState.get<number>(
      lastCheckedKey(availability.executable, currentVersion),
    );
    return typeof lastChecked !== "number" || Date.now() - lastChecked >= autoCheckIntervalMs;
  }

  private async showUpdateAvailable(
    info: XlflowUpdateInfo,
    executable: string,
    manual = false,
  ): Promise<void> {
    const latest = readNonEmpty(info.latestVersion);
    if (latest === undefined) {
      return;
    }
    const dismissedKey = updateDismissedKey(executable, info.currentVersion, latest);
    if (!manual && this.context.globalState.get<boolean>(dismissedKey) === true) {
      return;
    }
    const openRelease = vscode.l10n.t("Open Release");
    const installGuide = vscode.l10n.t("Install Guide");
    const dontShow = vscode.l10n.t("Don't Show Again for This Version");
    const action = await vscode.window.showInformationMessage(
      vscode.l10n.t("xlflow {latest} is available. Current: {current}.", {
        latest,
        current: info.currentVersion,
      }),
      openRelease,
      installGuide,
      dontShow,
    );
    if (action === openRelease) {
      await vscode.env.openExternal(updateReleaseUri(info));
    } else if (action === installGuide) {
      await openInstallGuide();
    } else if (action === dontShow) {
      await this.context.globalState.update(dismissedKey, true);
    }
  }

  private setStatus(status: XlflowUpdateStatus): void {
    this.status = status;
    this.emitter.fire(status);
  }
}

export function normalizeUpdateResult(result: RawCheckResult): XlflowUpdateStatus {
  const envelope = parseEnvelope(result.stdout);
  if (result.exitCode !== 0 || envelope?.status === "failed") {
    return {
      kind: "error",
      message:
        readNonEmpty(envelope?.error?.message) ??
        readNonEmpty(result.stderr) ??
        vscode.l10n.t("xlflow update check exited with code {exitCode}.", {
          exitCode: result.exitCode,
        }),
    };
  }
  const update = envelope?.update;
  const currentVersion = readNonEmpty(update?.current_version);
  if (currentVersion === undefined) {
    return { kind: "error", message: vscode.l10n.t("xlflow update check returned invalid JSON.") };
  }
  const info: XlflowUpdateInfo = {
    currentVersion,
    latestVersion: readNonEmpty(update?.latest_version),
    updateAvailable: update?.update_available === true,
    releaseUrl: readNonEmpty(update?.release_url),
  };
  return info.updateAvailable ? { kind: "available", info } : { kind: "upToDate", info };
}

export function updateSummary(
  availabilityVersion: string | undefined,
  status: XlflowUpdateStatus | undefined,
): string | undefined {
  if (status?.kind === "available") {
    return `${status.info.currentVersion} -> ${status.info.latestVersion ?? "latest"}`;
  }
  return cliVersionSummary(availabilityVersion);
}

export function updateTooltip(base: string, status: XlflowUpdateStatus | undefined): string {
  if (status?.kind !== "available") {
    return base;
  }
  const releaseUrl = readNonEmpty(status.info.releaseUrl);
  const lines = [
    base,
    "",
    vscode.l10n.t("Update available: {latest}", {
      latest: status.info.latestVersion ?? "latest",
    }),
  ];
  if (releaseUrl !== undefined) {
    lines.push(releaseUrl);
  }
  return lines.join("\n");
}

export function cliVersionSummary(version: string | undefined): string | undefined {
  const text = readNonEmpty(version);
  if (text === undefined) {
    return undefined;
  }
  for (const line of text.split(/\r?\n/)) {
    const match = line.match(/^\s*Version:\s*(.+?)\s*$/);
    if (match !== null) {
      return match[1].trim();
    }
  }
  const firstUsefulLine = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line.length > 0 && !/^OK\b/i.test(line));
  const xlflowVersion = firstUsefulLine?.match(/^xlflow\s+(.+)$/i);
  return xlflowVersion?.[1]?.trim() ?? firstUsefulLine;
}

export function lastCheckedKey(executable: string, currentVersion?: string): string {
  return `xlflow.update.lastChecked.${executable}.${currentVersion ?? "unknown"}`;
}

export function updateDismissedKey(
  executable: string,
  currentVersion: string,
  latestVersion: string,
): string {
  return `xlflow.update.dismissed.${executable}.${currentVersion}.${latestVersion}`;
}

export function shouldRunAutomaticCheck(nowMs: number, lastCheckedMs: number | undefined): boolean {
  return typeof lastCheckedMs !== "number" || nowMs - lastCheckedMs >= autoCheckIntervalMs;
}

function checkForUpdate(): Promise<RawCheckResult> {
  const config = readConfig();
  return execFileWithTimeout(config.path, ["--json", "update", "check"], updateCheckTimeoutMs);
}

function currentVersionFromStatus(status: XlflowUpdateStatus): string | undefined {
  if (status.kind === "available" || status.kind === "upToDate") {
    return status.info.currentVersion;
  }
  return undefined;
}

function execFileWithTimeout(
  command: string,
  args: string[],
  timeoutMs: number,
): Promise<RawCheckResult> {
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
      reject(new Error(vscode.l10n.t("xlflow update check timed out.")));
    }, timeoutMs);
    const clear = (): void => clearTimeout(timer);
    child.stdout.on("data", (data: Buffer) => stdoutChunks.push(data));
    child.stderr.on("data", (data: Buffer) => stderrChunks.push(data));
    child.on("error", (error) => {
      if (settled) {
        return;
      }
      settled = true;
      clear();
      reject(error);
    });
    child.on("close", (code) => {
      if (settled) {
        return;
      }
      settled = true;
      clear();
      resolve({
        exitCode: code ?? -1,
        stdout: Buffer.concat(stdoutChunks).toString("utf8"),
        stderr: Buffer.concat(stderrChunks).toString("utf8"),
      });
    });
  });
}

function parseEnvelope(stdout: string): UpdateEnvelope | undefined {
  try {
    return JSON.parse(stdout) as UpdateEnvelope;
  } catch {
    return undefined;
  }
}

function updateReleaseUri(info: XlflowUpdateInfo): vscode.Uri {
  const releaseUrl = readNonEmpty(info.releaseUrl);
  if (releaseUrl === undefined) {
    return latestReleaseFallback;
  }
  try {
    return vscode.Uri.parse(releaseUrl);
  } catch {
    return latestReleaseFallback;
  }
}

function readNonEmpty(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}
