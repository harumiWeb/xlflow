import * as vscode from "vscode";
import { readFormsRootFromToml } from "../sidebar";
import { isFormSpecPath } from "./document";
import { designerResourceUri } from "./resource";

export const openDesignerCommand = "xlflow.openUserFormDesigner";

export async function isDesignerSource(source: vscode.Uri): Promise<boolean> {
  const folder = vscode.workspace.getWorkspaceFolder(source);
  if (source.scheme !== "file" || !folder) return false;
  const config = await vscode.workspace.fs.readFile(vscode.Uri.joinPath(folder.uri, "xlflow.toml"));
  return isFormSpecPath(
    source.fsPath,
    folder.uri.fsPath,
    readFormsRootFromToml(Buffer.from(config).toString("utf8")),
  );
}

/** Use the same path boundary and resource identity as the custom editor. */
export async function openDesigner(uri?: vscode.Uri): Promise<void> {
  const source = uri ?? vscode.window.activeTextEditor?.document.uri;
  const folder = source && vscode.workspace.getWorkspaceFolder(source);
  if (!source || source.scheme !== "file" || !folder) {
    await vscode.window.showWarningMessage(vscode.l10n.t("Open a FormSpec in an xlflow project."));
    return;
  }
  try {
    if (!(await isDesignerSource(source))) {
      await vscode.window.showWarningMessage(
        vscode.l10n.t(
          "Open a FormSpec directly under the configured forms root's specs directory.",
        ),
      );
      return;
    }
    // Loading retains any unsaved TextDocument buffer; opening never publishes it.
    const document = await vscode.workspace.openTextDocument(source);
    await vscode.commands.executeCommand(
      "vscode.openWith",
      designerResourceUri(document.uri),
      "xlflow.userFormDesigner",
    );
  } catch (error) {
    await vscode.window.showErrorMessage(
      vscode.l10n.t("Could not open the UserForm Designer: {error}", { error: String(error) }),
    );
  }
}

export function registerOpenDesignerCommand(): vscode.Disposable {
  return vscode.commands.registerCommand(openDesignerCommand, openDesigner);
}
