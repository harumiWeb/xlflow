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

export async function prepareStructuralEdit(
  operation: StructuralOperation,
  snapshot: { version: number; document: DesignerDocument } | undefined,
  version: number,
  confirm: DeleteConfirmation,
  pageEditable = false,
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
      snapshot.document.name.toLowerCase() === name
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

export class UserFormEditorProvider implements vscode.CustomTextEditorProvider, vscode.Disposable {
  private readonly edits = new DocumentEditQueue();
  private readonly clipboard = new DesignerClipboard();
  public constructor(
    private readonly extensionUri: vscode.Uri,
    private readonly clients: XlflowLanguageClientManager,
    private readonly confirmDeletion: DeleteConfirmation = confirmSubtreeDeletion,
  ) {}

  dispose(): void {
    this.clipboard.dispose();
  }

  public async resolveCustomTextEditor(
    document: vscode.TextDocument,
    panel: vscode.WebviewPanel,
    token: vscode.CancellationToken,
  ): Promise<void> {
    const resources = vscode.Uri.joinPath(this.extensionUri, "dist", "userFormDesigner");
    panel.webview.options = { enableScripts: true, localResourceRoots: [resources] };
    panel.webview.html = designerHTML(panel.webview, resources);
    const sync = new PreviewSynchronizer();
    let ready = false;
    let disposed = false;
    let timer: NodeJS.Timeout | undefined;
    let editable = false;
    let propertyEditable = false;
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
      { message: Extract<HostMessage, { type: "editResult" }>; sending: boolean }
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
      if (message.type === "document") editable = message.editable === true;
      if (message.type === "document") propertyEditable = message.propertyEditable === true;
      if (message.type === "document") {
        structuralEditable = message.structuralEditable === true;
        parentEditable = message.parentEditable === true;
        pageEditable = message.pageEditable === true;
        tabEditable = message.tabEditable === true;
        productivityEditable = message.productivityEditable === true;
        latestDocument = { version: message.version, document: message.document };
      }
      if (message.type === "invalidDocument" || message.type === "editingUnavailable") {
        editable = false;
        propertyEditable = false;
        structuralEditable = false;
        parentEditable = false;
        pageEditable = false;
        tabEditable = false;
        productivityEditable = false;
      }
      if (ready && !disposed && panel.visible) void panel.webview.postMessage(message);
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
          return this.clients.requestUserFormPreview(document);
        },
        () => document.version,
        send,
      );
    };
    const schedule = () => {
      generation++;
      send({ type: "editingUnavailable" });
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
      vscode.workspace.onDidChangeTextDocument((event) => {
        if (event.document.uri.toString() === document.uri.toString()) schedule();
      }),
      this.clients.onDidChangePreviewConnection(schedule),
      this.clipboard.onDidChange(sendClipboardState),
      vscode.window.onDidChangeActiveColorTheme(() => send({ type: "themeChanged" })),
      panel.onDidChangeViewState(() => {
        if (panel.visible && ready) {
          sendLocalization();
          sendClipboardState();
          sync.replay(send);
          flushReplies();
          schedule();
        }
      }),
      panel.webview.onDidReceiveMessage((message: unknown) => {
        if (!isWebviewMessage(message)) return;
        if (message.type === "ready") {
          // A new App has no pending requests and starts request IDs over.
          // Never deliver an old context's reply to its replacement.
          contextGeneration++;
          replies.clear();
          ready = true;
          sendLocalization();
          sendClipboardState();
          sync.replay(send);
          void update();
        } else if (message.type === "openText") {
          void vscode.commands.executeCommand("vscode.openWith", document.uri, "default");
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
          const permitted = () =>
            !disposed &&
            contextGeneration === editContext &&
            panel.visible &&
            (!pageParent || pageEditable) &&
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
                        : property
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
            .run(document, message.version, message.operations, this.clients, permitted, prepare)
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
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src ${webview.cspSource};">
<link rel="stylesheet" href="${style}"><title>xlflow UserForm Designer</title></head>
<body><div id="app"></div><script nonce="${nonce}" src="${script}"></script></body></html>`;
}
