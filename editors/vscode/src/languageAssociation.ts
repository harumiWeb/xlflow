import * as path from "path";
import * as vscode from "vscode";
import { selectedWorkspaceFolder } from "./projectState";

const vbaLanguageAssociationSuppressionKey = "xlflow.vbaLanguageAssociationNotice.dismissed";
const notifiedWorkspaceKeys = new Set<string>();

export interface VbaLanguageDocument {
  uri: vscode.Uri;
  languageId: string;
}

export function isVbaSourcePath(filePath: string): boolean {
  return /\.(bas|cls|frm)$/i.test(filePath);
}

export function mismatchedVbaLanguageDocuments<T extends VbaLanguageDocument>(
  documents: readonly T[],
): T[] {
  return documents.filter(
    (document) => isVbaSourcePath(document.uri.fsPath) && document.languageId !== "vba",
  );
}

export function vbaFilesAssociationUpdate(
  existing: Record<string, string> | undefined,
): Record<string, string> {
  return {
    ...existing,
    "*.bas": "vba",
    "*.cls": "vba",
    "*.frm": "vba",
  };
}

export async function checkVbaLanguageAssociation(
  context: vscode.ExtensionContext,
  options: { force?: boolean } = {},
): Promise<void> {
  const mismatches = mismatchedVbaLanguageDocuments(vscode.workspace.textDocuments);
  if (mismatches.length === 0) {
    return;
  }

  const workspaceKey = languageAssociationWorkspaceKey(mismatches[0]);
  if (context.globalState.get<boolean>(suppressionKey(workspaceKey)) === true) {
    return;
  }
  if (options.force !== true && notifiedWorkspaceKeys.has(workspaceKey)) {
    return;
  }
  notifiedWorkspaceKeys.add(workspaceKey);

  const configure = vscode.l10n.t("Configure Workspace Association");
  const openSettings = vscode.l10n.t("Open Settings");
  const dismiss = vscode.l10n.t("Don't Show Again");
  const sample = path.basename(mismatches[0].uri.fsPath);
  const selected = await vscode.window.showWarningMessage(
    vscode.l10n.t(
      "{file} is not opened as VBA, so xlflow LSP features such as completion and diagnostics will not attach.",
      { file: sample },
    ),
    configure,
    openSettings,
    dismiss,
  );

  if (selected === configure) {
    await configureWorkspaceVbaAssociations();
    vscode.window.showInformationMessage(
      vscode.l10n.t("Configured workspace file associations for VBA source files."),
    );
  } else if (selected === openSettings) {
    await vscode.commands.executeCommand("workbench.action.openSettings", "files.associations");
  } else if (selected === dismiss) {
    await context.globalState.update(suppressionKey(workspaceKey), true);
  }
}

async function configureWorkspaceVbaAssociations(): Promise<void> {
  const config = vscode.workspace.getConfiguration("files");
  const existing = config.get<Record<string, string>>("associations", {});
  await config.update(
    "associations",
    vbaFilesAssociationUpdate(existing),
    vscode.ConfigurationTarget.Workspace,
  );
}

function languageAssociationWorkspaceKey(document: VbaLanguageDocument): string {
  const folder = vscode.workspace.getWorkspaceFolder(document.uri) ?? selectedWorkspaceFolder();
  return folder?.uri.toString() ?? "global";
}

function suppressionKey(workspaceKey: string): string {
  return `${vbaLanguageAssociationSuppressionKey}.${workspaceKey}`;
}
