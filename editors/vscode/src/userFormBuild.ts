import * as path from "path";
import * as vscode from "vscode";
import type { XlflowChannels } from "./logging";
import type { XlflowProjectStateService } from "./projectState";
import type { SessionManager } from "./session";
import { discoverUserForms, readExcelPathFromToml, readFormsRootFromToml } from "./sidebar";
import { runXlflowJsonCommand } from "./xlflow";
import {
  collectFormBuildSources,
  executeFormBuild,
  type FormBuildOutcome,
  type FormBuildSource,
} from "./userFormBuildWorkflow";

export interface UserFormBuildHooks {
  prepareBuildSources(uris: readonly vscode.Uri[]): Promise<boolean>;
  refreshUserForms(): Promise<void>;
}

export interface FormBuildInputs {
  workbook: string;
  formsRoot: vscode.Uri;
  sources: FormBuildSource[];
  duplicates: string[][];
  excluded: number;
  codeByName: Record<string, string[]>;
}

function isMissingDirectory(error: unknown): boolean {
  if (typeof error !== "object" || error === null || !("code" in error)) return false;
  return error.code === "FileNotFound" || error.code === "ENOENT";
}

export async function discoverFormBuildInputs(
  folder: vscode.WorkspaceFolder,
): Promise<FormBuildInputs> {
  const configUri = vscode.Uri.joinPath(folder.uri, "xlflow.toml");
  const config = Buffer.from(await vscode.workspace.fs.readFile(configUri)).toString("utf8");
  const workbook = readExcelPathFromToml(config);
  if (!workbook)
    throw new Error(vscode.l10n.t("Configure [excel].path in xlflow.toml before building forms."));
  const formsRoot = vscode.Uri.file(
    path.resolve(folder.uri.fsPath, readFormsRootFromToml(config).replace(/\\/g, "/")),
  );
  let entries: [string, vscode.FileType][];
  try {
    entries = await vscode.workspace.fs.readDirectory(vscode.Uri.joinPath(formsRoot, "specs"));
  } catch (error) {
    if (!isMissingDirectory(error)) throw error;
    entries = [];
  }
  const collected = collectFormBuildSources(
    entries.filter(([, type]) => (type & vscode.FileType.File) !== 0).map(([name]) => name),
  );
  const sourceNames = new Set(collected.sources.map((source) => source.name.toLowerCase()));
  const forms = await discoverUserForms(folder, configUri);
  const codeByName: Record<string, string[]> = Object.create(null) as Record<string, string[]>;
  for (const source of collected.sources) codeByName[source.name.toLowerCase()] = [];
  let codeEntries: [string, vscode.FileType][];
  try {
    codeEntries = await vscode.workspace.fs.readDirectory(vscode.Uri.joinPath(formsRoot, "code"));
  } catch (error) {
    if (!isMissingDirectory(error)) throw error;
    codeEntries = [];
  }
  // form build always consumes sidecar code, regardless of userform.code_source.
  // Preserve the directory entry's actual spelling when targeting dirty text documents.
  for (const [file, type] of codeEntries) {
    if ((type & vscode.FileType.File) === 0 || !/\.bas$/i.test(file)) continue;
    const key = file.replace(/\.bas$/i, "").toLowerCase();
    codeByName[key]?.push(vscode.Uri.joinPath(formsRoot, "code", file).fsPath);
  }
  return {
    workbook: path.resolve(folder.uri.fsPath, workbook.replace(/\\/g, "/")),
    formsRoot,
    ...collected,
    excluded: forms.filter((form) => !sourceNames.has(form.name.toLowerCase())).length,
    codeByName,
  };
}

/** One controller covers both commands so dialogs cannot launch overlapping builds. */
export class UserFormBuildController {
  private building = false;

  constructor(
    private readonly projectState: XlflowProjectStateService,
    private readonly channels: XlflowChannels,
    private readonly sessions: SessionManager,
    private readonly hooks: UserFormBuildHooks,
  ) {}

  async build(value?: unknown, all = false): Promise<void> {
    if (this.building) {
      void vscode.window.showWarningMessage(
        vscode.l10n.t("A UserForm build is already in progress."),
      );
      return;
    }
    this.building = true;
    await vscode.commands.executeCommand("setContext", "xlflow.formBuildRunning", true);
    try {
      const node = value as { name?: unknown; workspaceUri?: unknown } | undefined;
      const state = this.projectState.current();
      // Never fall back to the active editor for an individual tree command.
      const folder = all
        ? state.kind === "ready"
          ? state.workspaceFolder
          : undefined
        : node?.workspaceUri instanceof vscode.Uri
          ? vscode.workspace.workspaceFolders?.find(
              (candidate) => candidate.uri.toString() === node.workspaceUri!.toString(),
            )
          : undefined;
      if (!folder || (!all && typeof node?.name !== "string")) {
        void vscode.window.showWarningMessage(
          vscode.l10n.t("Select a form in an initialized xlflow project."),
        );
        return;
      }
      const inputs = await discoverFormBuildInputs(folder);
      if (inputs.duplicates.length > 0) {
        throw new Error(
          vscode.l10n.t("Duplicate FormSpecs must be resolved before building:\n{files}", {
            files: inputs.duplicates.map((group) => group.join(", ")).join("\n"),
          }),
        );
      }
      const selected = all
        ? inputs.sources
        : inputs.sources.filter(
            (source) => source.name.toLowerCase() === (node!.name as string).toLowerCase(),
          );
      if (selected.length === 0) {
        void vscode.window.showWarningMessage(
          vscode.l10n.t("No persisted FormSpec is available to build. Excluded forms: {count}.", {
            count: inputs.excluded,
          }),
        );
        return;
      }
      const buildAction = vscode.l10n.t("Build Forms");
      const detail = vscode.l10n.t(
        "Workbook: {workbook}\n\nForms:\n{forms}\n\nExisting forms will be replaced (--overwrite). All workbook changes in the session will be saved after success.\nExcluded forms without a FormSpec: {excluded}.",
        {
          workbook: inputs.workbook,
          forms: selected.map((source) => source.name).join("\n"),
          excluded: all ? inputs.excluded : 0,
        },
      );
      if (
        (await vscode.window.showWarningMessage(
          vscode.l10n.t("Build UserForms in Excel?"),
          { modal: true, detail },
          buildAction,
        )) !== buildAction
      )
        return;
      const sourceUris: vscode.Uri[] = [];
      for (const source of selected) {
        sourceUris.push(vscode.Uri.joinPath(inputs.formsRoot, "specs", source.specPath));
        sourceUris.push(
          ...(inputs.codeByName[source.name.toLowerCase()] ?? []).map((file) =>
            vscode.Uri.file(file),
          ),
        );
      }
      if (!(await this.hooks.prepareBuildSources(sourceUris))) return;
      // Saving or a confirmation may yield to external config/file changes. Do not silently retarget.
      const current = await discoverFormBuildInputs(folder);
      if (JSON.stringify(current) !== JSON.stringify(inputs)) {
        throw new Error(
          vscode.l10n.t("Form build inputs changed. Review the project and start the build again."),
        );
      }
      this.channels.output.show(true);
      const outcome = await vscode.window.withProgress(
        {
          location: vscode.ProgressLocation.Notification,
          title: vscode.l10n.t("Building UserForms"),
          cancellable: false,
        },
        async () =>
          executeFormBuild(
            selected.map((source) => ({
              ...source,
              specPath: vscode.Uri.joinPath(inputs.formsRoot, "specs", source.specPath).fsPath,
            })),
            inputs.workbook,
            folder.uri.fsPath,
            (args) =>
              runXlflowJsonCommand(args, "xlflow form build workflow", this.channels.output, {
                requireWorkspace: true,
                workspaceFolder: folder,
              }),
          ),
      );
      this.report(outcome);
    } catch (error) {
      this.channels.output.appendLine(`[error] UserForm build: ${String(error)}`);
      void vscode.window.showErrorMessage(
        vscode.l10n.t("UserForm build failed: {error}", { error: String(error) }),
      );
    } finally {
      try {
        const refreshes = await Promise.allSettled([
          this.hooks.refreshUserForms(),
          this.sessions.refreshStatus(),
        ]);
        for (const refresh of refreshes) {
          if (refresh.status === "rejected")
            this.channels.output.appendLine(
              `[error] UserForm build refresh: ${String(refresh.reason)}`,
            );
        }
      } finally {
        this.building = false;
        await vscode.commands.executeCommand("setContext", "xlflow.formBuildRunning", false);
      }
    }
  }

  private report(outcome: FormBuildOutcome): void {
    if (outcome.phase === "complete") {
      void vscode.window.showInformationMessage(
        vscode.l10n.t("Built and saved {count} UserForm(s).", { count: outcome.succeeded.length }),
      );
      return;
    }
    const message = vscode.l10n.t(
      "UserForm build stopped during {phase}.\nSucceeded: {succeeded}\nFailed: {failed}\nPending: {pending}\n{error}\nSuccessful builds may remain unsaved in Excel. Review the session state before saving or stopping it.",
      {
        phase: outcome.phase,
        succeeded: outcome.succeeded.join(", ") || "-",
        failed: outcome.failed ?? "-",
        pending: outcome.pending.join(", ") || "-",
        error: outcome.error ?? "",
      },
    );
    this.channels.output.appendLine(message);
    void vscode.window.showErrorMessage(message);
  }
}
