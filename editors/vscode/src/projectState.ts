import * as vscode from "vscode";

export type XlflowProjectState =
  | { kind: "noWorkspace" }
  | { kind: "notInitialized"; workspaceFolder: vscode.WorkspaceFolder }
  | { kind: "ready"; workspaceFolder: vscode.WorkspaceFolder; configPath: vscode.Uri }
  | {
      kind: "invalid";
      workspaceFolder: vscode.WorkspaceFolder;
      configPath: vscode.Uri;
      error: string;
    };

export class XlflowProjectStateService implements vscode.Disposable {
  private readonly emitter = new vscode.EventEmitter<XlflowProjectState>();
  private state: XlflowProjectState = { kind: "noWorkspace" };

  readonly onDidChangeState = this.emitter.event;

  dispose(): void {
    this.emitter.dispose();
  }

  current(): XlflowProjectState {
    return this.state;
  }

  async refresh(): Promise<XlflowProjectState> {
    const next = await detectProjectState();
    this.state = next;
    await setProjectContext(next);
    this.emitter.fire(next);
    return next;
  }
}

export async function detectProjectState(): Promise<XlflowProjectState> {
  const workspaceFolder = selectedWorkspaceFolder();
  if (workspaceFolder === undefined) {
    return { kind: "noWorkspace" };
  }

  const configPath = vscode.Uri.joinPath(workspaceFolder.uri, "xlflow.toml");
  try {
    const stat = await vscode.workspace.fs.stat(configPath);
    if ((stat.type & vscode.FileType.File) === 0) {
      return {
        kind: "invalid",
        workspaceFolder,
        configPath,
        error: "xlflow.toml exists but is not a file.",
      };
    }
    return { kind: "ready", workspaceFolder, configPath };
  } catch (error) {
    if (isFileNotFound(error)) {
      return { kind: "notInitialized", workspaceFolder };
    }
    return {
      kind: "invalid",
      workspaceFolder,
      configPath,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

export function selectedWorkspaceFolder(): vscode.WorkspaceFolder | undefined {
  const folders = vscode.workspace.workspaceFolders ?? [];
  if (folders.length === 0) {
    return undefined;
  }
  const activeDocument = vscode.window.activeTextEditor?.document;
  if (activeDocument !== undefined) {
    const containing = vscode.workspace.getWorkspaceFolder(activeDocument.uri);
    if (containing !== undefined) {
      return containing;
    }
  }
  return folders[0];
}

export async function setProjectContext(state: XlflowProjectState): Promise<void> {
  await vscode.commands.executeCommand(
    "setContext",
    "xlflow.hasWorkspace",
    state.kind !== "noWorkspace",
  );
  await vscode.commands.executeCommand("setContext", "xlflow.projectReady", state.kind === "ready");
  await vscode.commands.executeCommand(
    "setContext",
    "xlflow.projectMissing",
    state.kind === "notInitialized",
  );
  await vscode.commands.executeCommand(
    "setContext",
    "xlflow.projectInvalid",
    state.kind === "invalid",
  );
}

export function readyWorkspaceFolder(
  state: XlflowProjectState,
): vscode.WorkspaceFolder | undefined {
  return state.kind === "ready" ? state.workspaceFolder : undefined;
}

function isFileNotFound(error: unknown): boolean {
  if (error instanceof vscode.FileSystemError) {
    return error.code === "FileNotFound";
  }
  return String(error).includes("FileNotFound");
}
