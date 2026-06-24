import * as path from "path";
import * as vscode from "vscode";
import { XlflowChannels } from "./logging";
import { resolveWorkspaceRoot, runXlflowCommand, runXlflowJsonCommand } from "./xlflow";

export type SessionState = "unknown" | "inactive" | "starting" | "active" | "stopping" | "error";

interface XlflowStatusEnvelope {
  status?: string;
  error?: {
    code?: string;
    message?: string;
  };
  session?: XlflowSessionPayload;
}

interface XlflowSessionPayload {
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

interface SessionSnapshot {
  state: SessionState;
  session?: XlflowSessionPayload;
  workspaceFolder?: vscode.WorkspaceFolder;
  lastCheckedAt?: Date;
  lastError?: string;
}

type SessionAction = "start" | "stop" | "restart" | "status" | "output" | "doctor";

export class SessionManager implements vscode.Disposable {
  private readonly statusBarItem: vscode.StatusBarItem;
  private snapshot: SessionSnapshot = { state: "unknown" };

  constructor(private readonly channels: XlflowChannels) {
    this.statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 90);
    this.statusBarItem.command = "xlflow.sessionActions";
    this.updateStatusBar();
    this.statusBarItem.show();
  }

  dispose(): void {
    this.statusBarItem.dispose();
  }

  async refreshStatus(options: { prompt?: boolean; showOutput?: boolean } = {}): Promise<void> {
    const folder = await resolveWorkspaceRoot({
      prompt: options.prompt === true,
      fallbackToFirst: true,
    });
    if (folder === undefined) {
      this.snapshot = {
        state: "unknown",
        lastError: "No workspace folder is open.",
      };
      this.updateStatusBar();
      return;
    }

    const result = await runXlflowJsonCommand<XlflowStatusEnvelope>(
      ["--json", "session", "status"],
      "xlflow session status",
      this.channels.output,
      { requireWorkspace: false, workspaceFolder: folder },
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
  }

  async start(): Promise<void> {
    await this.runSessionCommand(
      "starting",
      ["session", "start"],
      "xlflow session start",
      "started",
    );
  }

  async stop(): Promise<void> {
    await this.runSessionCommand("stopping", ["session", "stop"], "xlflow session stop", "stopped");
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
      },
    );
    if (startCode !== 0) {
      await this.handleSessionFailure();
      return;
    }
    vscode.window.showInformationMessage("xlflow session restarted.");
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
    });
    if (code === 0) {
      vscode.window.showInformationMessage("xlflow doctor completed.");
    } else {
      vscode.window.showErrorMessage("xlflow doctor failed. See xlflow output.");
    }
  }

  async showActions(): Promise<void> {
    const action = await vscode.window.showQuickPick(this.quickPickItems(), {
      title: "xlflow Session",
      placeHolder: "Select a session action",
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
      { label: "Show Session Status", action: "status" },
      { label: "Open xlflow Output", action: "output" },
    ];
    switch (this.snapshot.state) {
      case "active":
        return [
          { label: "Stop Session", action: "stop" },
          { label: "Restart Session", action: "restart" },
          ...common,
        ];
      case "error":
        return [
          { label: "Show Session Status", action: "status" },
          { label: "Run Doctor", action: "doctor" },
          { label: "Open xlflow Output", action: "output" },
          { label: "Start Session", action: "start" },
        ];
      default:
        return [
          { label: "Start Session", action: "start" },
          ...common,
          { label: "Run Doctor", action: "doctor" },
        ];
    }
  }

  private async runSessionCommand(
    transientState: SessionState,
    args: string[],
    label: string,
    successVerb: string,
  ): Promise<void> {
    this.setTransientState(transientState);
    const code = await runXlflowCommand(args, label, this.channels.output, {
      requireWorkspace: true,
      notify: false,
    });
    if (code !== 0) {
      await this.handleSessionFailure();
      return;
    }
    vscode.window.showInformationMessage(`xlflow session ${successVerb}.`);
    await this.refreshStatus();
  }

  private async handleSessionFailure(): Promise<void> {
    vscode.window.showErrorMessage("xlflow session failed. See xlflow output.");
    await this.refreshStatus();
    if (this.snapshot.state !== "error") {
      this.snapshot = {
        ...this.snapshot,
        state: "error",
        lastError: "Session command failed. See xlflow output.",
      };
      this.updateStatusBar();
    }
  }

  private setTransientState(state: SessionState): void {
    this.snapshot = { ...this.snapshot, state };
    this.updateStatusBar();
  }

  private updateStatusBar(): void {
    this.statusBarItem.text = sessionStatusText(this.snapshot.state);
    this.statusBarItem.tooltip = sessionStatusTooltip(this.snapshot);
    this.statusBarItem.color =
      this.snapshot.state === "active" ? new vscode.ThemeColor("testing.iconPassed") : undefined;
    this.statusBarItem.backgroundColor =
      this.snapshot.state === "error"
        ? new vscode.ThemeColor("statusBarItem.warningBackground")
        : undefined;
  }
}

export function sessionStatusText(state: SessionState): string {
  switch (state) {
    case "unknown":
      return "$(question) xlflow";
    case "inactive":
      return "$(circle-slash) xlflow";
    case "starting":
      return "$(sync~spin) xlflow";
    case "active":
      return "$(check) xlflow: Session";
    case "stopping":
      return "$(sync~spin) xlflow";
    case "error":
      return "$(warning) xlflow";
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
      lines.push("xlflow session active");
      break;
    case "inactive":
      lines.push("No active xlflow session.");
      break;
    case "starting":
      lines.push("xlflow session starting...");
      break;
    case "stopping":
      lines.push("xlflow session stopping...");
      break;
    case "error":
      lines.push("xlflow session error");
      break;
    case "unknown":
      lines.push("xlflow session status unknown");
      break;
  }
  if (snapshot.workspaceFolder !== undefined) {
    lines.push(`Workspace: ${snapshot.workspaceFolder.uri.fsPath}`);
  }
  const workbook = workbookDisplayName(snapshot.session);
  if (workbook !== undefined) {
    lines.push(`Workbook: ${workbook}`);
  }
  if (snapshot.session?.save_required === true) {
    lines.push("Save required");
  } else if (snapshot.session?.dirty === true) {
    lines.push("Workbook dirty");
  }
  const startedAt = readNonEmpty(snapshot.session?.metadata?.started_at);
  if (startedAt !== undefined) {
    lines.push(`Started: ${startedAt}`);
  }
  if (snapshot.lastCheckedAt !== undefined) {
    lines.push(`Last check: ${snapshot.lastCheckedAt.toLocaleTimeString()}`);
  }
  if (snapshot.lastError !== undefined) {
    lines.push(`Error: ${snapshot.lastError}`);
  }
  lines.push("Click for session actions.");
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
    "xlflow session status did not return a valid session payload."
  );
}

function readNonEmpty(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}
