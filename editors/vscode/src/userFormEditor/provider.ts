import * as vscode from "vscode";
import { randomBytes } from "crypto";
import type { XlflowLanguageClientManager } from "../client";
import { readFormsRootFromToml } from "../sidebar";
import { isFormSpecPath, PreviewSynchronizer } from "./document";
import { designerStrings, isProductivityOperation, isWebviewMessage } from "./protocol";
import type {
  HostMessage,
  DesignerDocument,
  StructuralOperation,
  ParentOperation,
  DesignerError,
  PageOperation,
  TabOperation,
  DesignerStrings,
  ProductivityOperation,
  SemanticOperation,
  WebviewMessage,
  SpecControl,
} from "./protocol";
import { DocumentEditQueue } from "./edits";
import { controlMetadata, subtreeIds } from "./toolbox";
import { isDesignerFrame, isDesignerPage } from "./geometry";
import { isPageOperation, isTabOperation } from "./protocol";
import {
  DesignerClipboard,
  pasteOperation,
  prepareProductivityDeletion,
  projectKey,
  siblingSelection,
  validCopyResult,
} from "./clipboard";
import type { ClipboardContent } from "./clipboard";
import { designerDocument } from "./model";
import { tmpdir } from "node:os";
import { UserFormCustomDocument, virtualTextDocument } from "./customDocument";
import { DraftCheckpointStore, decodeDraftCheckpoint } from "./checkpoint";
import { pictureAssetReferenced } from "./pictureReferences";
import { DraftState, computeSourceChanges } from "./draft";
import type { DraftFrame } from "./draft";
import { designerResourceUri, designerSourceUri } from "./resource";
import { PictureImportSession } from "./pictures";
import type { PictureCommitReceipt } from "./pictures";

export type DeleteConfirmation = (name: string, descendants: number) => PromiseLike<boolean>;
const confirmSubtreeDeletion: DeleteConfirmation = async (name, descendants) => {
  const action = vscode.l10n.t("Delete controls");
  return (
    (await vscode.window.showWarningMessage(
      vscode.l10n.t('Delete "{0}" and its {1} descendant control(s)?', name, descendants),
      { modal: true },
      action,
    )) === action
  );
};

const confirmConflictResolution = async (choice: "source" | "designer"): Promise<boolean> => {
  const action =
    choice === "source"
      ? vscode.l10n.t("Use YAML/JSON changes")
      : vscode.l10n.t("Use Designer draft");
  const warning =
    choice === "source"
      ? vscode.l10n.t("Discard the conflicting Designer draft and use the current YAML/JSON text?")
      : vscode.l10n.t(
          "Use the Designer draft instead of the current YAML/JSON text? The source will change only when you save.",
        );
  return (await vscode.window.showWarningMessage(warning, { modal: true }, action)) === action;
};

export async function prepareStructuralEdit(
  operation: StructuralOperation,
  snapshot: { version: number; document: DesignerDocument } | undefined,
  version: number,
  confirm: DeleteConfirmation,
  pageEditable = false,
  initialPages: SpecControl[] = [],
): Promise<DesignerError | undefined> {
  if (!snapshot || snapshot.version !== version)
    return {
      code: "staleDocument",
      message: vscode.l10n.t("The document changed. Retry the operation on the updated form."),
    };
  const controls = snapshot.document.controls;
  if (operation.type === "addControl") {
    const name = operation.control.name.toLowerCase();
    if (
      controls.some((c) => c.id === operation.control.id || c.name.toLowerCase() === name) ||
      snapshot.document.name.toLowerCase() === name ||
      initialPages.some(
        (page) =>
          controls.some(
            (c) => c.id === page.id || c.name.toLowerCase() === page.name.toLowerCase(),
          ) || snapshot.document.name.toLowerCase() === page.name.toLowerCase(),
      )
    )
      return {
        code: "controlConflict",
        message: vscode.l10n.t("The new control ID or name is already in use."),
      };
    if (operation.control.parentId !== undefined) {
      const parent = controls.find((control) => control.id === operation.control.parentId);
      if (
        !parent ||
        (!isDesignerFrame(parent) && !(pageEditable && isDesignerPage(snapshot.document, parent)))
      )
        return {
          code: "invalidParent",
          message: vscode.l10n.t("This control cannot be placed in this container."),
        };
    }
    return;
  }
  const control = controls.find((c) => c.id === operation.controlId);
  if (!control || control.type.toLowerCase() === "page")
    return {
      code: "unsupportedDeletion",
      message: vscode.l10n.t("Select a control other than Page to delete."),
    };
  const descendants = subtreeIds(controls, control.id).size - 1;
  if (descendants && !(await confirm(control.name, descendants)))
    return { code: "editCancelled", message: vscode.l10n.t("Control deletion cancelled.") };
}

export function prepareParentEdit(
  operation: ParentOperation,
  snapshot: { version: number; document: DesignerDocument } | undefined,
  version: number,
  pageEditable = false,
): DesignerError | undefined {
  if (!snapshot || snapshot.version !== version)
    return {
      code: "staleDocument",
      message: vscode.l10n.t("The document changed. Retry the operation on the updated form."),
    };

  const { controls } = snapshot.document;
  const control = controls.find((candidate) => candidate.id === operation.controlId);
  const invalid = (): DesignerError => ({
    code: "invalidParent",
    message: vscode.l10n.t("This control cannot be placed in this container."),
  });
  if (
    !control ||
    !controlMetadata.some((metadata) => metadata.type.toLowerCase() === control.type.toLowerCase())
  )
    return invalid();

  if (control.parentId) {
    const currentParent = controls.find((candidate) => candidate.id === control.parentId);
    if (
      !currentParent ||
      (!isDesignerFrame(currentParent) &&
        !(pageEditable && isDesignerPage(snapshot.document, currentParent)))
    )
      return invalid();
  }

  if (operation.parentId === "") return;
  const parent = controls.find((candidate) => candidate.id === operation.parentId);
  if (
    !parent ||
    (!isDesignerFrame(parent) && !(pageEditable && isDesignerPage(snapshot.document, parent)))
  )
    return invalid();
  if (subtreeIds(controls, control.id).has(parent.id)) return invalid();
}

export async function prepareCollectionEdit(
  operation: PageOperation | TabOperation,
  snapshot: { version: number; document: DesignerDocument } | undefined,
  version: number,
  confirm: DeleteConfirmation,
): Promise<DesignerError | undefined> {
  if (!snapshot || snapshot.version !== version)
    return {
      code: "staleDocument",
      message: vscode.l10n.t("The document changed. Retry the operation on the updated form."),
    };
  const { document } = snapshot;
  const invalid = (): DesignerError => ({
    code: "invalidCollection",
    message: vscode.l10n.t("This control cannot be placed in this container."),
  });
  if (operation.type === "addPage") {
    const parent = document.controls.find((c) => c.id === operation.control.parentId);
    if (
      !parent ||
      parent.type.toLowerCase() !== "multipage" ||
      (parent.progId && parent.progId.toLowerCase() !== "forms.multipage.1")
    )
      return invalid();
    const name = operation.control.name.toLowerCase();
    if (
      document.name.toLowerCase() === name ||
      document.controls.some((c) => c.id === operation.control.id || c.name.toLowerCase() === name)
    )
      return {
        code: "controlConflict",
        message: vscode.l10n.t("The new control ID or name is already in use."),
      };
    return;
  }
  const control = document.controls.find((c) => c.id === operation.controlId);
  if (!control) return invalid();
  if (operation.type === "removePage" || operation.type === "reorderPage") {
    if (!isDesignerPage(document, control)) return invalid();
    if (operation.type === "removePage") {
      const descendants = subtreeIds(document.controls, control.id).size - 1;
      if (descendants && !operation.cascade) return invalid();
      if (descendants && !(await confirm(control.name, descendants)))
        return { code: "editCancelled", message: vscode.l10n.t("Control deletion cancelled.") };
    }
    return;
  }
  if (
    control.type.toLowerCase() !== "tabstrip" ||
    (control.progId && control.progId.toLowerCase() !== "forms.tabstrip.1")
  )
    return invalid();
}

export const designerViewType = "xlflow.userFormDesigner";
export const designerSaveCommand = "xlflow.userFormDesigner.save";

export class UserFormEditorProvider
  implements vscode.CustomEditorProvider<UserFormCustomDocument>, vscode.Disposable
{
  private readonly edits = new DocumentEditQueue();
  private readonly clipboard = new DesignerClipboard();
  private readonly documents = new Map<string, UserFormCustomDocument>();
  private readonly customEdits = new vscode.EventEmitter<
    vscode.CustomDocumentEditEvent<UserFormCustomDocument>
  >();
  public readonly onDidChangeCustomDocument = this.customEdits.event;
  private readonly restored = new Set<UserFormCustomDocument>();
  private readonly redirectDrafts = new Map<string, DraftState>();
  private readonly checkpoints: DraftCheckpointStore;
  private readonly panels = new Map<vscode.WebviewPanel, UserFormCustomDocument>();
  private readonly checkpointCancellation = new vscode.CancellationTokenSource();
  public constructor(
    private readonly extensionUri: vscode.Uri,
    private readonly clients: XlflowLanguageClientManager,
    private readonly confirmDeletion: DeleteConfirmation = confirmSubtreeDeletion,
    private readonly confirmResolution: (
      choice: "source" | "designer",
    ) => PromiseLike<boolean> = confirmConflictResolution,
    checkpointStorageRoot = vscode.Uri.joinPath(
      vscode.Uri.file(tmpdir()),
      "xlflow-userform-designer-drafts",
    ),
  ) {
    this.checkpoints = new DraftCheckpointStore(checkpointStorageRoot);
  }

  dispose(): void {
    this.documents.forEach((document) => document.dispose());
    this.documents.clear();
    this.redirectDrafts.clear();
    this.panels.clear();
    this.checkpointCancellation.dispose();
    this.customEdits.dispose();
    this.clipboard.dispose();
  }

  public async openCustomDocument(
    uri: vscode.Uri,
    context: vscode.CustomDocumentOpenContext,
    token: vscode.CancellationToken,
  ): Promise<UserFormCustomDocument> {
    const existing = this.documents.get(uri.toString());
    if (existing && !existing.closed) return existing;
    const source = await vscode.workspace.openTextDocument(designerSourceUri(uri));
    let state: DraftState | undefined = this.redirectDrafts.get(uri.toString());
    // A freshly closed panel can still have a queued checkpoint write. Read
    // behind that same source queue so reopening cannot miss its latest draft.
    let recovery = await this.edits.exclusive(source.uri, () => this.checkpoints.read(source.uri));
    if (context.backupId) {
      const backup = decodeDraftCheckpoint(
        JSON.parse(
          Buffer.from(
            await vscode.workspace.fs.readFile(vscode.Uri.parse(context.backupId)),
          ).toString("utf8"),
        ),
        source.uri,
      );
      this.checkpoints.observe(backup);
      if (!recovery || backup.updatedAt > recovery.updatedAt) recovery = backup;
    }
    if (!state && recovery) {
      const restored = DraftState.restore(recovery.state);
      // VS Code does not persist custom undo callbacks across extension reloads.
      // Recover the pending draft as one native edit against its source baseline.
      state = new DraftState(restored.baseText);
      state.commit(restored.text, vscode.l10n.t("Restore Designer draft"));
      try {
        if (state.dirty && !(await this.sourceMergeSafe(source, state.baseText, source.getText())))
          state.retainConflict(source.getText());
        else state.synchronize(source.getText());
      } catch {
        state.retainConflict(source.getText());
      }
    }
    if (token.isCancellationRequested) throw new vscode.CancellationError();
    const document: UserFormCustomDocument = new UserFormCustomDocument(
      source,
      state,
      uri,
      (base, external) => this.sourceMergeSafe(document.sourceDocument, base, external),
      (action) => this.edits.exclusive(source.uri, action),
    );
    if (state) this.restored.add(document);
    this.documents.set(uri.toString(), document);
    this.redirectDrafts.delete(uri.toString());
    let checkpointRevision = document.state.revision;
    const checkpointChanges = document.onDidChange(() => {
      if (document.uri.scheme === "file") return; // Temporary Open With redirect, not the draft owner.
      if (checkpointRevision === document.state.revision) return;
      checkpointRevision = document.state.revision;
      // Persist source synchronization and Undo/Redo as well as native saves:
      // a clean native tab can still contain changes not yet applied to YAML.
      void this.edits
        .exclusive(document.sourceUri, () =>
          this.checkpointDocument(document, this.checkpointCancellation.token),
        )
        .catch((error) => {
          document.saveError = String(error);
          if (!document.closed) document.notify();
        });
    });
    document.onDidDispose(() => {
      checkpointChanges.dispose();
      this.restored.delete(document);
      if (this.documents.get(uri.toString()) === document) this.documents.delete(uri.toString());
    });
    return document;
  }

  private async sourceMergeSafe(
    source: vscode.TextDocument,
    baseText: string,
    externalText: string,
  ): Promise<boolean> {
    if (baseText === externalText) return true;
    const snapshot = (text: string) =>
      virtualTextDocument(
        () => source,
        () => text,
        () => source.version,
        () => source.isClosed,
      );
    const [base, external] = await Promise.all([
      this.clients.requestUserFormPreview(snapshot(baseText)),
      this.clients.requestUserFormPreview(snapshot(externalText)),
    ]);
    if (!base.document || !external.document || base.error || external.error) return false;
    // A textual merge must not bind a pending edit to a different control after
    // a source reorder, deletion, reparenting or type change. Canonical LSP
    // projections identify these boundaries; the host does not parse FormSpec.
    const identity = (document: import("./protocol").PreviewDocument) =>
      JSON.stringify(
        document.controls.map((control) => [
          control.id,
          control.type.toLowerCase(),
          control.progId?.toLowerCase() ?? "",
          control.parentId ?? "",
        ]),
      );
    return identity(base.document) === identity(external.document);
  }

  private recordEdit(document: UserFormCustomDocument, frame: DraftFrame, label: string): void {
    this.customEdits.fire({
      document,
      label,
      undo: () =>
        this.edits.exclusive(document.sourceUri, async () => {
          document.undo(frame);
        }),
      redo: () =>
        this.edits.exclusive(document.sourceUri, async () => {
          document.redo(frame);
        }),
    });
  }

  public async saveCustomDocument(
    document: UserFormCustomDocument,
    token: vscode.CancellationToken,
  ): Promise<void> {
    try {
      await this.edits.exclusive(document.sourceUri, () =>
        this.checkpointDocument(document, token),
      );
    } catch (error) {
      if (!(error instanceof vscode.CancellationError)) {
        document.saveError = String(error);
        document.notify();
      }
      throw error;
    }
  }

  private async checkpointDocument(
    document: UserFormCustomDocument,
    token: vscode.CancellationToken,
  ): Promise<void> {
    await this.checkpoints.write(document.sourceUri, document.state.serialize(), token);
  }

  /** Publish only through an explicit Designer action, never native Auto Save. */
  public async publishCustomDocument(
    document: UserFormCustomDocument,
    token: vscode.CancellationToken,
  ): Promise<void> {
    await this.edits.exclusive(document.sourceUri, () =>
      (async () => {
        await document.save(async () => {
          const result = await this.clients.requestUserFormPreview(document.document);
          if (result.error || !result.document)
            throw new Error(result.error?.message ?? "The UserForm draft is invalid.");
        }, token);
        await this.checkpointDocument(document, token);
      })(),
    );
  }

  public registerSaveCommand(command = designerSaveCommand): vscode.Disposable {
    return vscode.commands.registerCommand(command, () => this.publishActiveDocument());
  }

  public async publishActiveDocument(): Promise<void> {
    const active = [...this.panels.entries()].find(([panel]) => panel.active);
    if (!active) return;
    await this.publishPanel(active[0], active[1]);
  }

  private async publishPanel(
    panel: vscode.WebviewPanel,
    document: UserFormCustomDocument,
  ): Promise<void> {
    const cancellation = new vscode.CancellationTokenSource();
    try {
      await this.publishCustomDocument(document, cancellation.token);
      // Refresh the native draft checkpoint even when Auto Save already made
      // the tab clean. This command never publishes source by itself.
      if (!document.closed && panel.visible) {
        panel.reveal(panel.viewColumn);
        await vscode.commands.executeCommand("workbench.action.files.save");
      }
    } finally {
      cancellation.dispose();
    }
  }

  public async saveCustomDocumentAs(
    document: UserFormCustomDocument,
    destination: vscode.Uri,
    token: vscode.CancellationToken,
  ): Promise<void> {
    if (
      destination.toString() === document.uri.toString() ||
      destination.toString() === document.sourceUri.toString()
    )
      return this.publishCustomDocument(document, token);
    await this.edits.exclusive(document.sourceUri, async () => {
      document.saving = true;
      document.saveError = undefined;
      document.notify();
      try {
        if (!(await document.reconcileSource(document.sourceDocument.getText())))
          throw new Error("The source changed during Save As.");
        if (document.state.conflicted)
          throw new Error(
            vscode.l10n.t("Resolve the YAML/JSON and Designer conflict before saving."),
          );
        const format = (uri: vscode.Uri) =>
          /\.json$/i.test(uri.path) ? "json" : /\.ya?ml$/i.test(uri.path) ? "yaml" : undefined;
        if (
          destination.scheme !== "file" ||
          !format(destination) ||
          format(destination) !== format(document.sourceUri)
        )
          throw new Error(
            vscode.l10n.t("Save the Designer draft using the same YAML/JSON format."),
          );
        const revision = document.state.revision;
        const text = document.state.text;
        const result = await this.clients.requestUserFormPreview(document.document);
        if (result.error || !result.document)
          throw new Error(result.error?.message ?? "The UserForm draft is invalid.");
        if (token.isCancellationRequested) throw new vscode.CancellationError();
        if (document.state.revision !== revision)
          throw new Error("The Designer draft changed during Save As.");
        let target: vscode.TextDocument;
        try {
          await vscode.workspace.fs.stat(destination);
        } catch (error) {
          if (!(error instanceof vscode.FileSystemError) || error.code !== "FileNotFound")
            throw error;
          const create = new vscode.WorkspaceEdit();
          create.createFile(destination, { overwrite: false, ignoreIfExists: false });
          if (!(await vscode.workspace.applyEdit(create)))
            throw new Error("Could not create the Save As destination.");
        }
        target = await vscode.workspace.openTextDocument(destination);
        if (target.isDirty)
          throw new Error(
            vscode.l10n.t("Save or revert the destination's unsaved text changes before Save As."),
          );
        if (token.isCancellationRequested) throw new vscode.CancellationError();
        if (document.closed || document.state.revision !== revision)
          throw new Error("The Designer draft changed during Save As.");
        const edit = new vscode.WorkspaceEdit();
        edit.set(destination, [
          vscode.TextEdit.setEndOfLine(document.sourceDocument.eol),
          ...computeSourceChanges(target.getText(), text).map(
            (change) =>
              new vscode.TextEdit(
                new vscode.Range(target.positionAt(change.start), target.positionAt(change.end)),
                change.newText,
              ),
          ),
        ]);
        if (!(await vscode.workspace.applyEdit(edit)) || !(await target.save()))
          throw new Error("Could not save the Designer draft.");
        if (target.getText() !== text || document.state.revision !== revision)
          throw new Error(
            "The document changed during Save As. The Designer draft has been retained.",
          );
        // Save As replaces the editor with the destination. Retain the latest
        // source for the original URI so reopening it does not resurrect the
        // draft that was explicitly published to a different file.
        document.state.revert(document.sourceDocument.getText());
        await this.checkpointDocument(document, token);
      } catch (error) {
        document.saveError = String(error);
        throw error;
      } finally {
        document.saving = false;
        document.notify();
      }
    });
  }

  public async revertCustomDocument(
    document: UserFormCustomDocument,
    token: vscode.CancellationToken,
  ): Promise<void> {
    await this.edits.exclusive(document.sourceUri, async () => {
      document.sourceDocument = await vscode.workspace.openTextDocument(document.sourceUri);
      const text = document.sourceDocument.getText();
      if (token.isCancellationRequested) throw new vscode.CancellationError();
      document.state.revert(text);
      await this.checkpointDocument(document, token);
      document.notify();
    });
  }

  public async backupCustomDocument(
    document: UserFormCustomDocument,
    context: vscode.CustomDocumentBackupContext,
    token: vscode.CancellationToken,
  ): Promise<vscode.CustomDocumentBackup> {
    return this.edits.exclusive(document.sourceUri, async () => {
      if (token.isCancellationRequested) throw new vscode.CancellationError();
      const checkpoint = await this.checkpoints.write(
        document.sourceUri,
        document.state.serialize(),
        token,
      );
      await vscode.workspace.fs.writeFile(
        context.destination,
        Buffer.from(JSON.stringify(checkpoint), "utf8"),
      );
      if (token.isCancellationRequested) throw new vscode.CancellationError();
      return {
        id: context.destination.toString(),
        delete: () => {
          void vscode.workspace.fs.delete(context.destination).then(undefined, () => {});
        },
      };
    });
  }

  public async resolveCustomEditor(
    customDocument: UserFormCustomDocument,
    panel: vscode.WebviewPanel,
    token: vscode.CancellationToken,
  ): Promise<void> {
    if (customDocument.uri.scheme === "file") {
      const resource = designerResourceUri(customDocument.sourceUri);
      if (customDocument.state.dirty || customDocument.state.conflicted) {
        const existing = this.documents.get(resource.toString());
        if (existing && existing.state.text !== customDocument.state.text)
          throw new Error(
            "Another Designer draft is already open for this form. Close it before restoring this backup.",
          );
        if (!existing)
          this.redirectDrafts.set(
            resource.toString(),
            DraftState.restore(customDocument.state.serialize()),
          );
      }
      await vscode.commands.executeCommand(
        "vscode.openWith",
        resource,
        panel.viewType || designerViewType,
        { viewColumn: panel.viewColumn },
      );
      panel.dispose();
      return;
    }
    this.panels.set(panel, customDocument);
    panel.onDidDispose(() => this.panels.delete(panel));
    const document = customDocument.document;
    const resources = vscode.Uri.joinPath(this.extensionUri, "dist", "userFormDesigner");
    panel.webview.options = { enableScripts: true, localResourceRoots: [resources] };
    panel.webview.html = designerHTML(panel.webview, resources);
    const sync = new PreviewSynchronizer();
    let ready = false;
    let disposed = false;
    let timer: NodeJS.Timeout | undefined;
    let editable = false;
    let propertyEditable = false;
    let propertyEditorVersion = 0;
    let structuralEditable = false;
    let parentEditable = false;
    let pageEditable = false;
    let tabEditable = false;
    let productivityEditable = false;
    let latestDocument: { version: number; document: DesignerDocument } | undefined;
    let generation = 0;
    let contextGeneration = 0;
    const replies = new Map<
      number,
      {
        message: Extract<HostMessage, { type: "editResult" | "pictureSelection" }>;
        sending: boolean;
      }
    >();
    const flushReplies = () => {
      if (!ready || disposed || !panel.visible) return;
      for (const [id, reply] of replies) {
        if (reply.sending) continue;
        reply.sending = true;
        void panel.webview.postMessage(reply.message).then(
          (delivered) => {
            reply.sending = false;
            if (delivered && replies.get(id) === reply) replies.delete(id);
          },
          () => {
            reply.sending = false;
          },
        );
      }
    };
    const send = (message: HostMessage) => {
      if (
        message.type === "document" &&
        (customDocument.state.conflicted || customDocument.saving || customDocument.reconciling)
      ) {
        message = {
          ...message,
          editable: false,
          propertyEditable: false,
          structuralEditable: false,
          parentEditable: false,
          pageEditable: false,
          tabEditable: false,
          productivityEditable: false,
        };
      }
      if (message.type === "document") editable = message.editable === true;
      if (message.type === "document") propertyEditable = message.propertyEditable === true;
      if (message.type === "document") {
        propertyEditorVersion = message.propertyEditorVersion ?? 0;
        structuralEditable = message.structuralEditable === true;
        parentEditable = message.parentEditable === true;
        pageEditable = message.pageEditable === true;
        tabEditable = message.tabEditable === true;
        productivityEditable = message.productivityEditable === true;
        latestDocument = { version: message.version, document: message.document };
      }
      if (message.type === "invalidDocument" || message.type === "editingUnavailable") {
        editable = false;
        propertyEditorVersion = 0;
        propertyEditable = false;
        structuralEditable = false;
        parentEditable = false;
        pageEditable = false;
        tabEditable = false;
        productivityEditable = false;
      }
      if (ready && !disposed && panel.visible) void panel.webview.postMessage(message);
    };
    let pictures: PictureImportSession | undefined;
    let pictureBinding:
      | { version: number; controlId: string; generation: number; context: number }
      | undefined;
    const pictureCurrent = (version: number, controlId: string) =>
      !disposed &&
      ready &&
      panel.visible &&
      propertyEditable &&
      propertyEditorVersion === 1 &&
      !customDocument.state.conflicted &&
      !customDocument.reconciling &&
      !customDocument.saving &&
      document.version === version &&
      latestDocument?.version === version &&
      latestDocument.document.controls.some(
        (control) =>
          control.id === controlId &&
          control.type === "Image" &&
          (!control.progId || control.progId.toLowerCase() === "forms.image.1"),
      );
    const pictureReferenced = async (relative: string) => {
      const folder = vscode.workspace.getWorkspaceFolder(document.uri);
      if (!folder) return true;
      const documents = this.documents;
      const redirectDrafts = this.redirectDrafts;
      const checkpoints = this.checkpoints;
      const checkpointRevision = checkpoints.revision;
      if (checkpoints.writing) return true;
      return pictureAssetReferenced(relative, {
        checkpointStable: () => !checkpoints.writing && checkpoints.revision === checkpointRevision,
        *openTexts() {
          // Text documents can be dirty without any Designer being open, and
          // untitled buffers or other workspace folders can retain references.
          for (const text of vscode.workspace.textDocuments) yield text.getText();
          for (const open of documents.values()) {
            yield JSON.stringify(open.state.serialize());
            yield open.sourceDocument.getText();
          }
          for (const state of redirectDrafts.values()) yield JSON.stringify(state.serialize());
        },
        async *sourceTexts() {
          const files = await vscode.workspace.findFiles(
            new vscode.RelativePattern(folder, "**/*.{yaml,yml,json}"),
            "**/{node_modules,.git,tmp_workspaces}/**",
          );
          for (const file of files)
            yield Buffer.from(await vscode.workspace.fs.readFile(file)).toString("utf8");
        },
        async *checkpointTexts() {
          // Checkpoints live outside the project and remain after their native
          // editor closes. Scan the entire store, including in-flight .tmp files.
          const root = vscode.Uri.joinPath(checkpoints.location(document.uri), "..");
          let entries: [string, vscode.FileType][];
          try {
            entries = await vscode.workspace.fs.readDirectory(root);
          } catch (error) {
            if (error instanceof vscode.FileSystemError && error.code === "FileNotFound") return;
            throw error;
          }
          for (const [name, type] of entries) {
            if (type !== vscode.FileType.File)
              throw new Error("Unknown Designer checkpoint entry.");
            yield Buffer.from(
              await vscode.workspace.fs.readFile(vscode.Uri.joinPath(root, name)),
            ).toString("utf8");
          }
        },
      });
    };
    const handlePicture = async (
      message: Extract<WebviewMessage, { type: "picturePick" | "pictureCommit" }>,
    ) => {
      const requestContext = contextGeneration;
      const reply = (result: Extract<HostMessage, { type: "editResult" | "pictureSelection" }>) => {
        if (disposed || requestContext !== contextGeneration) return;
        replies.set(message.requestId, { message: result, sending: false });
        flushReplies();
      };
      const failed = (error: unknown) =>
        reply({
          type: message.type === "picturePick" ? "pictureSelection" : "editResult",
          requestId: message.requestId,
          error: { code: "pictureUnavailable", message: String(error) },
        });
      if (!pictureCurrent(message.version, message.controlId)) {
        failed(vscode.l10n.t("The document changed. Cancel and reopen the editor."));
        return;
      }
      try {
        if (message.type === "picturePick") {
          const folder = vscode.workspace.getWorkspaceFolder(document.uri)!;
          const config = Buffer.from(
            await vscode.workspace.fs.readFile(vscode.Uri.joinPath(folder.uri, "xlflow.toml")),
          ).toString("utf8");
          pictures?.cancel();
          const session = new PictureImportSession({
            projectRoot: folder.uri.fsPath,
            formsRoot: readFormsRootFromToml(config),
            sourceURI: document.uri.toString(),
            validate: (request) => this.clients.requestUserFormPicture(request),
          });
          pictures = session;
          const binding = {
            version: message.version,
            controlId: message.controlId,
            generation,
            context: requestContext,
          };
          pictureBinding = binding;
          const selection = await session.open();
          if (
            !pictureCurrent(message.version, message.controlId) ||
            binding.generation !== generation ||
            requestContext !== contextGeneration ||
            pictures !== session
          ) {
            session.cancel();
            throw new Error(vscode.l10n.t("The document changed. Cancel and reopen the editor."));
          }
          reply({ type: "pictureSelection", requestId: message.requestId, selection });
          return;
        }
        const session = pictures;
        const binding = pictureBinding;
        if (
          !session ||
          !binding ||
          binding.version !== message.version ||
          binding.controlId !== message.controlId ||
          binding.generation !== generation ||
          binding.context !== requestContext
        )
          throw new Error(vscode.l10n.t("The picture selection expired. Select it again."));
        let receipt: PictureCommitReceipt | undefined;
        const operations: SemanticOperation[] = [];
        const current = () =>
          pictureCurrent(message.version, message.controlId) &&
          binding.generation === generation &&
          requestContext === contextGeneration;
        let error: DesignerError | undefined;
        try {
          error = await this.edits.run(
            document,
            message.version,
            operations,
            this.clients,
            current,
            async () => {
              receipt = await session.commit(message.token);
              operations.push({
                type: "setControlProperty",
                controlId: message.controlId,
                field: "picture",
                value: receipt.picture,
              });
              return undefined;
            },
            applyDraft,
          );
          if (!error && receipt) session.accept(receipt);
        } catch (caught) {
          error = { code: "pictureUnavailable", message: String(caught) };
        }
        if (error) await session.rollback({ token: message.token }, pictureReferenced);
        reply({ type: "editResult", requestId: message.requestId, error });
        await update();
      } catch (error) {
        failed(error);
      }
    };
    const sendDraftState = () =>
      send({
        type: "draftState",
        dirty: customDocument.state.dirty,
        conflicted: customDocument.state.conflicted,
        saving: customDocument.saving || customDocument.reconciling,
        error: customDocument.saveError,
      });
    const applyDraft = async (edits: import("./protocol").SourceTextEdit[], label: string) => {
      const frame = customDocument.apply(edits, label);
      if (!frame) return;
      this.recordEdit(customDocument, frame, label);
    };
    const sendLocalization = () =>
      send({
        type: "localization",
        strings: Object.fromEntries(
          Object.entries(designerStrings).map(([key, value]) => [key, vscode.l10n.t(value)]),
        ) as DesignerStrings,
      });
    const update = async () => {
      if (disposed || token.isCancellationRequested) return;
      if (timer) {
        clearTimeout(timer);
        timer = undefined;
      }
      const version = document.version;
      await sync.update(
        version,
        async () => {
          const folder = vscode.workspace.getWorkspaceFolder(document.uri);
          if (!folder || document.uri.scheme !== "file") {
            return {
              version,
              error: {
                code: "notFormSpec",
                message: vscode.l10n.t("Open a FormSpec in an xlflow project."),
              },
            };
          }
          let config: string;
          try {
            const configUri = vscode.Uri.joinPath(folder.uri, "xlflow.toml");
            // Match the LSP's saved configuration, not an unsaved config buffer.
            config = Buffer.from(await vscode.workspace.fs.readFile(configUri)).toString("utf8");
          } catch {
            return {
              version,
              error: {
                code: "notFormSpec",
                message: vscode.l10n.t("This workspace has no xlflow.toml."),
              },
            };
          }
          if (
            !isFormSpecPath(document.uri.fsPath, folder.uri.fsPath, readFormsRootFromToml(config))
          ) {
            return {
              version,
              error: {
                code: "notFormSpec",
                message: vscode.l10n.t(
                  "Open a FormSpec directly under the configured forms root's specs directory.",
                ),
              },
            };
          }
          const preview = await this.clients.requestUserFormPreview(document);
          if (preview.propertyEditorVersion === 1 && preview.document) {
            await Promise.all(
              preview.document.controls.map(async (control) => {
                if (control.type !== "Image" || !control.picture?.path) return;
                try {
                  const image = await this.clients.requestUserFormPicture({
                    uri: document.uri.toString(),
                    path: control.picture.path,
                  });
                  if (!image.error) control.picturePreview = image.dataURI;
                } catch {
                  /* The typed source remains available when an asset is unavailable. */
                }
              }),
            );
          }
          return preview;
        },
        () => document.version,
        send,
      );
    };
    const schedule = () => {
      generation++;
      sync.invalidate();
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => {
        timer = undefined;
        void update();
      }, 100);
    };
    const sendClipboardState = () =>
      send({ type: "clipboardState", available: this.clipboard.available });
    const handleClipboard = async (
      message: Extract<WebviewMessage, { type: "clipboard" }>,
    ): Promise<{ error?: DesignerError; selectionIds?: string[] }> => {
      const editContext = contextGeneration;
      const editGeneration = generation;
      const key = projectKey(document);
      let copyRequest: number | undefined;
      const current = () =>
        !customDocument.state.conflicted &&
        !customDocument.reconciling &&
        !customDocument.saving &&
        !disposed &&
        panel.visible &&
        contextGeneration === editContext &&
        generation === editGeneration &&
        productivityEditable &&
        (message.action === "copy" || structuralEditable) &&
        (message.action !== "cut" ||
          copyRequest === undefined ||
          this.clipboard.isCurrentCopy(copyRequest));
      const stale: DesignerError = {
        code: "staleDocument",
        message: vscode.l10n.t("The document changed. Retry the operation on the updated form."),
      };
      if (!key || !current() || latestDocument?.version !== message.version)
        return { error: stale };
      if (message.action !== "paste") {
        const invalid = siblingSelection(latestDocument.document, message.controlIds);
        if (invalid) return { error: invalid };
      }
      const destinationParent =
        message.action === "duplicate"
          ? (latestDocument.document.controls.find((c) => c.id === message.controlIds[0])
              ?.parentId ?? "")
          : message.parentId;
      copyRequest =
        message.action === "copy" || message.action === "cut"
          ? this.clipboard.beginCopy()
          : undefined;
      const capture = async (text: string): Promise<ClipboardContent | DesignerError> => {
        const selectionError = siblingSelection(latestDocument!.document, message.controlIds);
        if (selectionError) return selectionError;
        const result = await this.clients.requestUserFormCopy(
          document,
          message.version,
          text,
          message.controlIds,
        );
        if (result.error) return result.error;
        if (!validCopyResult(result, message.version))
          return {
            code: "invalidResponse",
            message: vscode.l10n.t("The language server returned an invalid control snapshot."),
          };
        return {
          snapshot: result.snapshot!,
          document: designerDocument(result.document!),
          projectKey: key,
          hasPicturePaths: result.hasPicturePaths!,
        };
      };
      if (message.action === "copy") {
        const result = await this.edits.capture(document, message.version, current, capture);
        if (result.error) return { error: result.error };
        if (!result.value) return { error: stale };
        if ("code" in result.value) return { error: result.value };
        if (!this.clipboard.finishCopy(copyRequest!, result.value))
          return {
            error: {
              code: "copySuperseded",
              message: vscode.l10n.t("A newer copy request replaced this selection."),
            },
          };
        return {};
      }
      const operations: SemanticOperation[] = [];
      let copied: ClipboardContent | undefined;
      let selectionIds: string[] | undefined;
      let clipboardGeneration: number | undefined;
      const prepare = async (): Promise<DesignerError | undefined> => {
        if (latestDocument?.version !== message.version) return stale;
        if (message.action !== "paste") {
          const result = await capture(document.getText());
          if ("code" in result) return result;
          copied = result;
          if (!current() || document.version !== message.version) return stale;
        }
        if (message.action === "cut") {
          const operation: Extract<ProductivityOperation, { type: "removeControls" }> = {
            type: "removeControls",
            controlIds: message.controlIds,
            cascade: true,
          };
          const error = await prepareProductivityDeletion(
            operation,
            latestDocument.document,
            this.confirmDeletion,
          );
          if (error) return error;
          operations.push(operation);
          return;
        }
        const content = copied ?? this.clipboard.read();
        if (!content)
          return {
            code: "emptyClipboard",
            message: vscode.l10n.t("Copy controls before pasting."),
          };
        if (
          (destinationParent && !parentEditable) ||
          (content.document.controls.some((c) => !!c.parentId) && !parentEditable) ||
          (content.document.controls.some((c) => c.type.toLowerCase() === "page") &&
            !pageEditable) ||
          (destinationParent &&
            latestDocument.document.controls.some(
              (c) => c.id === destinationParent && c.type.toLowerCase() === "page",
            ) &&
            !pageEditable)
        )
          return {
            code: "unsupportedPaste",
            message: vscode.l10n.t("Update xlflow to support this container hierarchy."),
          };
        clipboardGeneration = this.clipboard.generation;
        const operation = pasteOperation(
          content,
          latestDocument.document,
          key,
          destinationParent,
          message.action === "duplicate"
            ? 8
            : this.clipboard.nextOffset(document.uri.toString(), destinationParent),
          pageEditable,
        );
        if ("code" in operation) return operation;
        const roots = new Set(content.snapshot.controlIds);
        selectionIds = operation.allocations.filter((a) => roots.has(a.sourceId)).map((a) => a.id);
        operations.push(operation);
      };
      const error = await this.edits.run(
        document,
        message.version,
        operations,
        this.clients,
        current,
        prepare,
        applyDraft,
      );
      if (error) return { error };
      if (message.action === "cut" && copied) this.clipboard.finishCopy(copyRequest!, copied);
      if (message.action === "paste")
        this.clipboard.pasted(document.uri.toString(), destinationParent, clipboardGeneration);
      return { selectionIds };
    };
    const watcher = vscode.workspace.createFileSystemWatcher("**/xlflow.toml");
    const subscriptions: vscode.Disposable[] = [
      watcher,
      watcher.onDidChange(schedule),
      watcher.onDidCreate(schedule),
      watcher.onDidDelete(schedule),
      customDocument.onDidChange(() => {
        sendDraftState();
        schedule();
      }),
      this.clients.onDidChangePreviewConnection(() => {
        send({ type: "editingUnavailable" });
        schedule();
      }),
      this.clipboard.onDidChange(sendClipboardState),
      vscode.window.onDidChangeActiveColorTheme(() => send({ type: "themeChanged" })),
      panel.onDidChangeViewState(() => {
        if (panel.visible && ready) {
          sendLocalization();
          sendDraftState();
          sendClipboardState();
          sync.replay(send);
          flushReplies();
          schedule();
        }
      }),
      panel.webview.onDidReceiveMessage((message: unknown) => {
        if (!isWebviewMessage(message)) return;
        if (message.type === "ready") {
          if (this.restored.delete(customDocument)) {
            for (const frame of customDocument.state.appliedFrames)
              this.recordEdit(
                customDocument,
                frame,
                frame.label ?? vscode.l10n.t("Restore Designer draft"),
              );
          }
          // A new App has no pending requests and starts request IDs over.
          // Never deliver an old context's reply to its replacement.
          contextGeneration++;
          replies.clear();
          ready = true;
          sendLocalization();
          sendDraftState();
          sendClipboardState();
          sync.replay(send);
          void update();
        } else if (message.type === "openText") {
          void vscode.window.showTextDocument(customDocument.sourceDocument, {
            viewColumn: vscode.ViewColumn.Beside,
          });
        } else if (message.type === "save") {
          void this.publishPanel(panel, customDocument).then(undefined, (error) =>
            vscode.window.showErrorMessage(String(error)),
          );
        } else if (message.type === "pictureCancel") {
          pictures?.cancel();
          pictureBinding = undefined;
        } else if (message.type === "picturePick" || message.type === "pictureCommit") {
          void handlePicture(message);
        } else if (message.type === "resolveConflict") {
          void this.edits
            .exclusive(document.uri, async () => {
              if (!customDocument.state.conflicted) return;
              const choice = message.choice;
              const revision = document.version;
              if (!(await this.confirmResolution(choice)) || document.version !== revision) return;
              const frame = customDocument.state.resolveConflict(choice);
              this.recordEdit(customDocument, frame, vscode.l10n.t("Resolve Designer conflict"));
              customDocument.notify();
            })
            .catch((error) => vscode.window.showErrorMessage(String(error)));
        } else if (message.type === "clipboard") {
          const editContext = contextGeneration;
          void handleClipboard(message)
            .catch((error) => ({
              error: { code: "copyUnavailable", message: String(error) },
            }))
            .then(async (result) => {
              if (disposed || contextGeneration !== editContext) return;
              replies.set(message.requestId, {
                message: { type: "editResult", requestId: message.requestId, ...result },
                sending: false,
              });
              flushReplies();
              await update();
            });
        } else {
          const editGeneration = generation;
          const editContext = contextGeneration;
          const property =
            message.operations[0].type === "setFormProperty" ||
            message.operations[0].type === "setControlProperty";
          const operation = message.operations[0];
          const initialPages = message.operations.flatMap((op) =>
            op.type === "addPage" && operation.type === "addControl" ? [op.control] : [],
          );
          const initialTabs =
            operation.type === "addControl" && operation.control.type === "TabStrip";
          const initialMultiPage =
            operation.type === "addControl" && operation.control.type === "MultiPage";
          const structural = operation.type === "addControl" || operation.type === "removeControl";
          const parentOperation = operation.type === "setParent";
          const parentedAdd =
            operation.type === "addControl" && operation.control.parentId !== undefined;
          const page = isPageOperation(operation);
          const tab = isTabOperation(operation);
          const productivity = isProductivityOperation(operation);
          const productivityAllowed = () => {
            if (!productivity || !productivityEditable) return false;
            if (operation.type === "removeControls") return structuralEditable;
            if (operation.type === "setTabOrder") return propertyEditable;
            // Paste snapshots belong to the host clipboard, never an untrusted Webview.
            if (operation.type === "pasteControls") return false;
            return editable;
          };
          const pageParent =
            parentOperation || parentedAdd
              ? latestDocument?.document.controls.some(
                  (c) =>
                    c.type.toLowerCase() === "page" &&
                    (c.id ===
                      (operation.type === "setParent"
                        ? operation.parentId
                        : operation.type === "addControl"
                          ? operation.control.parentId
                          : undefined) ||
                      (operation.type === "setParent" &&
                        latestDocument?.document.controls.find((s) => s.id === operation.controlId)
                          ?.parentId === c.id)),
                )
              : false;
          const typedEditor =
            operation.type === "setGeometry" ||
            (property &&
              "value" in operation &&
              typeof operation.value === "object" &&
              operation.value !== null);
          const permitted = () =>
            (!typedEditor || propertyEditorVersion === 1) &&
            !customDocument.state.conflicted &&
            !customDocument.reconciling &&
            !customDocument.saving &&
            !disposed &&
            contextGeneration === editContext &&
            panel.visible &&
            (!pageParent || pageEditable) &&
            (!initialPages.length || pageEditable) &&
            (!initialMultiPage || pageEditable) &&
            (!initialTabs || tabEditable) &&
            (productivity
              ? productivityAllowed()
              : page
                ? pageEditable
                : tab
                  ? tabEditable
                  : parentOperation
                    ? editable && parentEditable
                    : parentedAdd
                      ? structuralEditable && parentEditable
                      : structural
                        ? structuralEditable
                        : property || operation.type === "setGeometry"
                          ? propertyEditable
                          : editable) &&
            generation === editGeneration;
          const prepare = productivity
            ? async () => {
                if (latestDocument?.version !== message.version)
                  return {
                    code: "staleDocument",
                    message: vscode.l10n.t(
                      "The document changed. Retry the operation on the updated form.",
                    ),
                  };
                if (operation.type === "removeControls")
                  return prepareProductivityDeletion(
                    operation,
                    latestDocument.document,
                    this.confirmDeletion,
                  );
                return undefined;
              }
            : structural
              ? () =>
                  prepareStructuralEdit(
                    operation,
                    latestDocument,
                    message.version,
                    this.confirmDeletion,
                    pageEditable,
                    initialPages,
                  )
              : parentOperation
                ? async () =>
                    prepareParentEdit(operation, latestDocument, message.version, pageEditable)
                : page || tab
                  ? () =>
                      prepareCollectionEdit(
                        operation as PageOperation | TabOperation,
                        latestDocument,
                        message.version,
                        this.confirmDeletion,
                      )
                  : undefined;
          void this.edits
            .run(
              document,
              message.version,
              message.operations,
              this.clients,
              permitted,
              prepare,
              applyDraft,
            )
            .then(async (error) => {
              if (disposed || contextGeneration !== editContext) return;
              replies.set(message.requestId, {
                message: { type: "editResult", requestId: message.requestId, error },
                sending: false,
              });
              flushReplies();
              await update();
            });
        }
      }),
    ];
    const dispose = () => {
      disposed = true;
      replies.clear();
      if (timer) clearTimeout(timer);
      pictures?.cancel();
      sync.dispose();
      subscriptions.forEach((subscription) => subscription.dispose());
    };
    subscriptions.push(token.onCancellationRequested(dispose), panel.onDidDispose(dispose));
  }
}

export function designerHTML(webview: vscode.Webview, resources: vscode.Uri): string {
  const nonce = randomBytes(24).toString("hex");
  const script = webview.asWebviewUri(vscode.Uri.joinPath(resources, "app.js"));
  const style = webview.asWebviewUri(vscode.Uri.joinPath(resources, "app.css"));
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src ${webview.cspSource}; font-src ${webview.cspSource}; img-src data:;">
<link rel="stylesheet" href="${style}"><title>xlflow UserForm Designer</title></head>
<body><div id="app"></div><script nonce="${nonce}" src="${script}"></script></body></html>`;
}
