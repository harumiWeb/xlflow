import * as path from "path";
import * as vscode from "vscode";
import { XlflowChannels } from "./logging";
import { resolveWorkspaceRoot, runXlflowCommand, runXlflowJsonCommand } from "./xlflow";

export type SessionState = "unknown" | "inactive" | "starting" | "active" | "stopping" | "error";

export interface XlflowStatusEnvelope {
  status?: string;
  error?: {
    code?: string;
    message?: string;
  };
  session?: XlflowSessionPayload;
}

export interface XlflowSessionPayload {
  active?: boolean;
  workbook_path?: string;
  workbook_name?: string;
  dirty?: boolean;
  save_required?: boolean;
  running?: boolean;
  workbook_open?: boolean;
  mode?: string;
  metadata?: {
    started_at?: string;
    pid?: number;
    hwnd?: number;
    workbook_path?: string;
  } | null;
}

export interface SessionSnapshot {
  state: SessionState;
  session?: XlflowSessionPayload;
  workspaceFolder?: vscode.WorkspaceFolder;
  lastCheckedAt?: Date;
  lastError?: string;
}

type SessionAction = "start" | "stop" | "restart" | "status" | "output" | "doctor";

export class SessionManager implements vscode.Disposable {
  private readonly statusBarItem: vscode.StatusBarItem;
  private readonly emitter = new vscode.EventEmitter<SessionSnapshot>();
  private snapshot: SessionSnapshot = { state: "unknown" };
  private projectKind: "noWorkspace" | "notInitialized" | "ready" | "invalid" = "noWorkspace";

  readonly onDidChangeSnapshot = this.emitter.event;

  constructor(private readonly channels: XlflowChannels) {
    this.statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 90);
    this.statusBarItem.command = "xlflow.sessionActions";
    this.updateStatusBar();
    this.statusBarItem.show();
  }

  dispose(): void {
    this.emitter.dispose();
    this.statusBarItem.dispose();
  }

  currentSnapshot(): SessionSnapshot {
    return this.snapshot;
  }

  setProjectKind(kind: "noWorkspace" | "notInitialized" | "ready" | "invalid"): void {
    this.projectKind = kind;
    this.updateStatusBar();
  }

  async refreshStatus(options: { prompt?: boolean; showOutput?: boolean } = {}): Promise<void> {
    const folder = await resolveWorkspaceRoot({
      prompt: options.prompt === true,
      fallbackToFirst: true,
    });
    if (folder === undefined) {
      this.snapshot = {
        state: "unknown",
        lastError: vscode.l10n.t("No workspace folder is open."),
      };
      this.updateStatusBar();
      this.emitter.fire(this.snapshot);
      return;
    }

    const result = await runXlflowJsonCommand<XlflowStatusEnvelope>(
      ["--json", "session", "status"],
      "xlflow session status",
      this.channels.output,
      {
        requireWorkspace: false,
        showCliUnavailable: options.prompt === true || options.showOutput === true,
        workspaceFolder: folder,
      },
    );
    if (options.showOutput === true) {
      this.channels.output.show(true);
    }
    const lastCheckedAt = new Date();
    if (result.exitCode !== 0 || result.json === undefined || result.json.session === undefined) {
      this.snapshot = {
        state: "error",
        workspaceFolder: folder,
        lastCheckedAt,
        lastError: statusErrorMessage(result.json, result.stderr),
      };
      this.updateStatusBar();
      this.emitter.fire(this.snapshot);
      return;
    }

    const state = sessionStateFromEnvelope(result.json);
    this.snapshot = {
      state,
      session: result.json.session,
      workspaceFolder: folder,
      lastCheckedAt,
    };
    this.updateStatusBar();
    this.emitter.fire(this.snapshot);
  }

  async start(): Promise<void> {
    await this.runSessionCommand(
      "starting",
      ["session", "start"],
      "xlflow session start",
      vscode.l10n.t("xlflow session start"),
      "started",
      vscode.l10n.t("started"),
    );
  }

  async stop(): Promise<void> {
    await this.runSessionCommand(
      "stopping",
      ["session", "stop"],
      "xlflow session stop",
      vscode.l10n.t("xlflow session stop"),
      "stopped",
      vscode.l10n.t("stopped"),
    );
  }

  async restart(): Promise<void> {
    this.setTransientState("stopping");
    const stopCode = await runXlflowCommand(
      ["session", "stop"],
      "xlflow session stop",
      this.channels.output,
      {
        requireWorkspace: true,
        notify: false,
        uiLabel: vscode.l10n.t("xlflow session stop"),
      },
    );
    if (stopCode !== 0) {
      await this.handleSessionFailure();
      return;
    }
    this.setTransientState("starting");
    const startCode = await runXlflowCommand(
      ["session", "start"],
      "xlflow session start",
      this.channels.output,
      {
        requireWorkspace: true,
        notify: false,
        uiLabel: vscode.l10n.t("xlflow session start"),
      },
    );
    if (startCode !== 0) {
      await this.handleSessionFailure();
      return;
    }
    vscode.window.showInformationMessage(vscode.l10n.t("xlflow session restarted."));
    await this.refreshStatus();
  }

  async showStatus(): Promise<void> {
    this.channels.output.show(true);
    await this.refreshStatus({ prompt: true, showOutput: true });
  }

  openOutput(): void {
    this.channels.output.show(true);
  }

  async runDoctor(): Promise<void> {
    const code = await runXlflowCommand(["doctor"], "xlflow doctor", this.channels.output, {
      requireWorkspace: true,
      notify: false,
      showOutput: true,
      uiLabel: vscode.l10n.t("xlflow doctor"),
    });
    if (code === 0) {
      vscode.window.showInformationMessage(vscode.l10n.t("xlflow doctor completed."));
    } else {
      this.showOutputError(vscode.l10n.t("xlflow doctor failed. See xlflow output."));
    }
  }

  async showActions(): Promise<void> {
    const action = await vscode.window.showQuickPick(this.quickPickItems(), {
      title: vscode.l10n.t("xlflow Session"),
      placeHolder: vscode.l10n.t("Select a session action"),
    });
    if (action === undefined) {
      return;
    }
    await this.runAction(action.action);
  }

  private async runAction(action: SessionAction): Promise<void> {
    switch (action) {
      case "start":
        await this.start();
        return;
      case "stop":
        await this.stop();
        return;
      case "restart":
        await this.restart();
        return;
      case "status":
        await this.showStatus();
        return;
      case "output":
        this.openOutput();
        return;
      case "doctor":
        await this.runDoctor();
        return;
    }
  }

  private quickPickItems(): Array<vscode.QuickPickItem & { action: SessionAction }> {
    const common: Array<vscode.QuickPickItem & { action: SessionAction }> = [
      { label: vscode.l10n.t("Show Session Status"), action: "status" },
      { label: vscode.l10n.t("Open xlflow Output"), action: "output" },
    ];
    switch (this.snapshot.state) {
      case "active":
        return [
          { label: vscode.l10n.t("Stop Session"), action: "stop" },
          { label: vscode.l10n.t("Restart Session"), action: "restart" },
          ...common,
        ];
      case "error":
        return [
          { label: vscode.l10n.t("Show Session Status"), action: "status" },
          { label: vscode.l10n.t("Run Doctor"), action: "doctor" },
          { label: vscode.l10n.t("Open xlflow Output"), action: "output" },
          { label: vscode.l10n.t("Start Session"), action: "start" },
        ];
      default:
        return [
          { label: vscode.l10n.t("Start Session"), action: "start" },
          ...common,
          { label: vscode.l10n.t("Run Doctor"), action: "doctor" },
        ];
    }
  }

  private async runSessionCommand(
    transientState: SessionState,
    args: string[],
    label: string,
    uiLabel: string,
    successVerb: string,
    successVerbLabel: string,
  ): Promise<void> {
    this.setTransientState(transientState);
    const code = await runXlflowCommand(args, label, this.channels.output, {
      requireWorkspace: true,
      notify: false,
      uiLabel,
    });
    if (code !== 0) {
      await this.handleSessionFailure();
      return;
    }
    vscode.window.showInformationMessage(
      vscode.l10n.t("xlflow session {successVerb}.", {
        successVerb: successVerbLabel,
      }),
    );
    await this.refreshStatus();
  }

  private async handleSessionFailure(): Promise<void> {
    this.showOutputError(vscode.l10n.t("xlflow session failed. See xlflow output."));
    await this.refreshStatus();
    if (this.snapshot.state !== "error") {
      this.snapshot = {
        ...this.snapshot,
        state: "error",
        lastError: vscode.l10n.t("Session command failed. See xlflow output."),
      };
      this.updateStatusBar();
      this.emitter.fire(this.snapshot);
    }
  }

  private showOutputError(message: string): void {
    const openOutput = vscode.l10n.t("Open xlflow Output");
    void vscode.window.showErrorMessage(message, openOutput).then((action) => {
      if (action === openOutput) {
        this.openOutput();
      }
    });
  }

  private setTransientState(state: SessionState): void {
    this.snapshot = { ...this.snapshot, state };
    this.updateStatusBar();
    this.emitter.fire(this.snapshot);
  }

  private updateStatusBar(): void {
    this.statusBarItem.text = sessionStatusText(this.snapshot.state, this.projectKind);
    this.statusBarItem.tooltip = sessionStatusTooltip(this.snapshot);
    this.statusBarItem.command =
      this.projectKind === "ready" || this.projectKind === "invalid"
        ? "xlflow.sessionActions"
        : "xlflow.setupActions";
    this.statusBarItem.color =
      this.snapshot.state === "active" ? new vscode.ThemeColor("testing.iconPassed") : undefined;
    this.statusBarItem.backgroundColor =
      this.snapshot.state === "error"
        ? new vscode.ThemeColor("statusBarItem.warningBackground")
        : undefined;
    this.updateSessionContext();
  }

  private updateSessionContext(): void {
    const projectReady = this.projectKind === "ready";
    void vscode.commands.executeCommand(
      "setContext",
      "xlflow.sessionActive",
      projectReady && this.snapshot.state === "active",
    );
    void vscode.commands.executeCommand(
      "setContext",
      "xlflow.sessionStartEnabled",
      projectReady &&
        (this.snapshot.state === "unknown" ||
          this.snapshot.state === "inactive" ||
          this.snapshot.state === "error"),
    );
    void vscode.commands.executeCommand(
      "setContext",
      "xlflow.sessionStopEnabled",
      projectReady && this.snapshot.state === "active",
    );
    void vscode.commands.executeCommand(
      "setContext",
      "xlflow.saveRequired",
      projectReady && this.snapshot.session?.save_required === true,
    );
  }
}

export function sessionStatusText(
  state: SessionState,
  projectKind: "noWorkspace" | "notInitialized" | "ready" | "invalid" = "ready",
): string {
  if (projectKind === "noWorkspace") {
    return vscode.l10n.t("$(circle-slash) xlflow: No Workspace");
  }
  if (projectKind === "notInitialized") {
    return vscode.l10n.t("$(circle-slash) xlflow: No Project");
  }
  if (projectKind === "invalid") {
    return vscode.l10n.t("$(warning) xlflow: Project Error");
  }
  switch (state) {
    case "unknown":
      return vscode.l10n.t("$(question) xlflow: Session Unknown");
    case "inactive":
      return vscode.l10n.t("$(circle-slash) xlflow: No Session");
    case "starting":
      return vscode.l10n.t("$(sync~spin) xlflow: Starting");
    case "active":
      return vscode.l10n.t("$(check) xlflow: Session Active");
    case "stopping":
      return vscode.l10n.t("$(sync~spin) xlflow: Stopping");
    case "error":
      return vscode.l10n.t("$(warning) xlflow: Session Error");
  }
}

export function sessionStateFromEnvelope(env: XlflowStatusEnvelope): SessionState {
  if (env.status === "failed" || env.session === undefined) {
    return "error";
  }
  return env.session.active === true ? "active" : "inactive";
}

function sessionStatusTooltip(snapshot: SessionSnapshot): string {
  const lines: string[] = [];
  switch (snapshot.state) {
    case "active":
      lines.push(vscode.l10n.t("xlflow session active"));
      break;
    case "inactive":
      lines.push(vscode.l10n.t("No active xlflow session."));
      break;
    case "starting":
      lines.push(vscode.l10n.t("xlflow session starting..."));
      break;
    case "stopping":
      lines.push(vscode.l10n.t("xlflow session stopping..."));
      break;
    case "error":
      lines.push(vscode.l10n.t("xlflow session error"));
      break;
    case "unknown":
      lines.push(vscode.l10n.t("xlflow session status unknown"));
      break;
  }
  if (snapshot.workspaceFolder !== undefined) {
    lines.push(
      vscode.l10n.t("Workspace: {workspace}", { workspace: snapshot.workspaceFolder.uri.fsPath }),
    );
  }
  const workbook = workbookDisplayName(snapshot.session);
  if (workbook !== undefined) {
    lines.push(vscode.l10n.t("Workbook: {workbook}", { workbook }));
  }
  if (snapshot.session?.save_required === true) {
    lines.push(vscode.l10n.t("Save required"));
  } else if (snapshot.session?.dirty === true) {
    lines.push(vscode.l10n.t("Workbook dirty"));
  }
  const startedAt = readNonEmpty(snapshot.session?.metadata?.started_at);
  if (startedAt !== undefined) {
    lines.push(vscode.l10n.t("Started: {startedAt}", { startedAt }));
  }
  if (snapshot.lastCheckedAt !== undefined) {
    lines.push(
      vscode.l10n.t("Last check: {lastCheck}", {
        lastCheck: snapshot.lastCheckedAt.toLocaleTimeString(),
      }),
    );
  }
  if (snapshot.lastError !== undefined) {
    lines.push(vscode.l10n.t("Error: {error}", { error: snapshot.lastError }));
  }
  lines.push(vscode.l10n.t("Click for session actions."));
  return lines.join("\n");
}

function workbookDisplayName(session: XlflowSessionPayload | undefined): string | undefined {
  const explicitName = readNonEmpty(session?.workbook_name);
  if (explicitName !== undefined) {
    return explicitName;
  }
  const workbookPath = readNonEmpty(session?.workbook_path ?? session?.metadata?.workbook_path);
  return workbookPath === undefined ? undefined : path.basename(workbookPath);
}

function statusErrorMessage(env: XlflowStatusEnvelope | undefined, stderr: string): string {
  return (
    readNonEmpty(env?.error?.message) ??
    readNonEmpty(stderr) ??
    vscode.l10n.t("xlflow session status did not return a valid session payload.")
  );
}

function readNonEmpty(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}
