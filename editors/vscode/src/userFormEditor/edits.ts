import * as vscode from "vscode";
import type { DesignerError, EditResult, GeometryOperation, SourceTextEdit } from "./protocol";

export interface EditClient {
  requestUserFormEdit(
    document: vscode.TextDocument,
    version: number,
    text: string,
    operations: GeometryOperation[],
  ): Promise<EditResult>;
}
// Shared across panels: no concurrent edit generation/application for one URI.
export class DocumentEditQueue {
  private pending = new Map<string, Promise<unknown>>();
  async run(
    document: vscode.TextDocument,
    version: number,
    operations: GeometryOperation[],
    client: EditClient,
    current: () => boolean,
  ): Promise<DesignerError | undefined> {
    const key = document.uri.toString();
    const previous = this.pending.get(key) ?? Promise.resolve();
    const job = previous
      .catch(() => {})
      .then(async () => {
        const text = document.getText();
        const stale = () =>
          !current() ||
          document.isClosed ||
          document.version !== version ||
          document.getText() !== text;
        if (stale())
          return {
            code: "staleDocument",
            message: vscode.l10n.t(
              "The document changed. Retry the operation on the updated form.",
            ),
          };
        try {
          const result = await client.requestUserFormEdit(document, version, text, operations);
          if (stale() || result.version !== version)
            return {
              code: "staleDocument",
              message: vscode.l10n.t(
                "The document changed. Retry the operation on the updated form.",
              ),
            };
          if (result.error) return result.error;
          const edits = result.edits;
          if (!Array.isArray(edits) || !validTextEdits(document, edits))
            return {
              code: "invalidResponse",
              message: vscode.l10n.t("The language server returned invalid edit ranges."),
            };
          if (!edits.length) return undefined;
          const edit = new vscode.WorkspaceEdit();
          const label = operations.some((op) => op.type === "resizeControl")
            ? vscode.l10n.t("Resize control")
            : vscode.l10n.t("Move control");
          edit.set(
            document.uri,
            edits.map((e): [vscode.TextEdit, vscode.WorkspaceEditEntryMetadata] => [
              new vscode.TextEdit(
                new vscode.Range(
                  e.range.start.line,
                  e.range.start.character,
                  e.range.end.line,
                  e.range.end.character,
                ),
                e.newText,
              ),
              { label, needsConfirmation: false },
            ]),
          );
          if (!(await vscode.workspace.applyEdit(edit)))
            return {
              code: "editRejected",
              message: vscode.l10n.t("VS Code could not apply the UserForm edit."),
            };
          return undefined;
        } catch (error) {
          return { code: "editUnavailable", message: String(error) };
        }
      });
    this.pending.set(key, job);
    try {
      return await job;
    } finally {
      if (this.pending.get(key) === job) this.pending.delete(key);
    }
  }
}
function validTextEdits(document: vscode.TextDocument, edits: SourceTextEdit[]): boolean {
  let previousEnd = -1;
  for (const edit of edits) {
    if (!edit || typeof edit.newText !== "string" || !edit.range) return false;
    const validPosition = (position: { line: number; character: number }) =>
      position &&
      Number.isSafeInteger(position.line) &&
      Number.isSafeInteger(position.character) &&
      position.line >= 0 &&
      position.line < document.lineCount &&
      position.character >= 0 &&
      position.character <= document.lineAt(position.line).text.length;
    if (!validPosition(edit.range.start) || !validPosition(edit.range.end)) return false;
    const start = document.offsetAt(
      new vscode.Position(edit.range.start.line, edit.range.start.character),
    );
    const end = document.offsetAt(
      new vscode.Position(edit.range.end.line, edit.range.end.character),
    );
    const boundary = (offset: number) =>
      !(
        offset > 0 &&
        /[\uD800-\uDBFF]/.test(document.getText()[offset - 1]) &&
        /[\uDC00-\uDFFF]/.test(document.getText()[offset] ?? "")
      );
    if (start < previousEnd || end < start || !boundary(start) || !boundary(end)) return false;
    previousEnd = end;
  }
  return true;
}
