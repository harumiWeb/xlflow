import * as vscode from "vscode";
import { randomBytes } from "crypto";
import type { XlflowLanguageClientManager } from "../client";
import { readFormsRootFromToml } from "../sidebar";
import { isFormSpecPath, PreviewSynchronizer } from "./document";
import { isWebviewMessage } from "./protocol";
import type {
  HostMessage,
  DesignerDocument,
  StructuralOperation,
  ParentOperation,
  DesignerError,
  PageOperation,
  TabOperation,
} from "./protocol";
import { DocumentEditQueue } from "./edits";
import { controlMetadata, subtreeIds } from "./toolbox";
import { isDesignerFrame, isDesignerPage } from "./geometry";
import { isPageOperation, isTabOperation } from "./protocol";

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

export class UserFormEditorProvider implements vscode.CustomTextEditorProvider {
  private readonly edits = new DocumentEditQueue();
  public constructor(
    private readonly extensionUri: vscode.Uri,
    private readonly clients: XlflowLanguageClientManager,
    private readonly confirmDeletion: DeleteConfirmation = confirmSubtreeDeletion,
  ) {}

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
        latestDocument = { version: message.version, document: message.document };
      }
      if (message.type === "invalidDocument" || message.type === "editingUnavailable") {
        editable = false;
        propertyEditable = false;
        structuralEditable = false;
        parentEditable = false;
        pageEditable = false;
        tabEditable = false;
      }
      if (ready && !disposed && panel.visible) void panel.webview.postMessage(message);
    };
    const sendLocalization = () =>
      send({
        type: "localization",
        strings: {
          pages: vscode.l10n.t("Pages"),
          tabs: vscode.l10n.t("Tabs"),
          addPage: vscode.l10n.t("Add Page"),
          addTab: vscode.l10n.t("Add tab"),
          deleteItem: vscode.l10n.t("Delete"),
          moveUp: vscode.l10n.t("Move up"),
          moveDown: vscode.l10n.t("Move down"),
          moveToPage: vscode.l10n.t("Move to Page"),
          choosePage: vscode.l10n.t("Choose destination Page"),
          collectionReadOnly: vscode.l10n.t(
            "Collection editing requires an xlflow version supporting Page or tab edits.",
          ),
          pagePlacement: vscode.l10n.t(
            "Click the form, a Frame or the active Page to place the control. Escape cancels.",
          ),
          header: vscode.l10n.t("xlflow UserForm Designer"),
          toolbox: vscode.l10n.t("Toolbox"),
          pointer: vscode.l10n.t("Pointer"),
          deleteControl: vscode.l10n.t("Delete control"),
          structuralReadOnly: vscode.l10n.t(
            "Adding and deleting controls requires an xlflow version supporting structural edits.",
          ),
          cannotInsert: vscode.l10n.t("This control's default size is larger than the form."),
          placement: vscode.l10n.t(
            "Click the form to place the control. Escape cancels. Controls are added at the form root.",
          ),
          hierarchy: vscode.l10n.t("Selection path"),
          invalidDrop: vscode.l10n.t("This control cannot be placed in this container."),
          dropTarget: vscode.l10n.t("Drop target"),
          parentReadOnly: vscode.l10n.t(
            "Changing parents requires an xlflow version supporting parent edits.",
          ),
          parentPlacement: vscode.l10n.t(
            "Click the form or a Frame to place the control. Escape cancels.",
          ),
          cannotInsertContainer: vscode.l10n.t(
            "This control's default size is larger than the container.",
          ),
          openText: vscode.l10n.t("Open text editor"),
          lastValid: vscode.l10n.t("Showing the last valid document."),
          loading: vscode.l10n.t("Loading FormSpec…"),
          approximate: vscode.l10n.t(
            "Approximate preview. Missing dimensions use display defaults; Page bounds derive from MultiPage.",
          ),
          approximateBounds: vscode.l10n.t("approximate bounds"),
          grid: vscode.l10n.t("Show grid"),
          snap: vscode.l10n.t("Snap to grid"),
          zoom: vscode.l10n.t("Zoom"),
          readOnly: vscode.l10n.t("Editing requires an xlflow version supporting UserForm edits."),
          cannotMove: vscode.l10n.t(
            "This control is larger than its parent. Resize it to fit before moving.",
          ),
          properties: vscode.l10n.t("Properties"),
          identity: vscode.l10n.t("Identity"),
          appearance: vscode.l10n.t("Appearance"),
          layout: vscode.l10n.t("Layout"),
          behavior: vscode.l10n.t("Behavior"),
          navigation: vscode.l10n.t("Navigation"),
          advanced: vscode.l10n.t("Advanced"),
          propertyName: vscode.l10n.t("Name"),
          propertyCaption: vscode.l10n.t("Caption"),
          propertyText: vscode.l10n.t("Text"),
          propertyValue: vscode.l10n.t("Value"),
          propertyLeft: vscode.l10n.t("Left"),
          propertyTop: vscode.l10n.t("Top"),
          propertyWidth: vscode.l10n.t("Width"),
          propertyHeight: vscode.l10n.t("Height"),
          propertyTabIndex: vscode.l10n.t("TabIndex"),
          propertySelectedIndex: vscode.l10n.t("SelectedIndex"),
          propertyEnabled: vscode.l10n.t("Enabled"),
          propertyVisible: vscode.l10n.t("Visible"),
          propertyTag: vscode.l10n.t("Tag"),
          propertyControlTipText: vscode.l10n.t("ControlTipText"),
          propertyAccelerator: vscode.l10n.t("Accelerator"),
          propertyBuildCaption: vscode.l10n.t("Build caption"),
          propertyBuildWidth: vscode.l10n.t("Build width"),
          propertyBuildHeight: vscode.l10n.t("Build height"),
          propertyClientWidth: vscode.l10n.t("Client width"),
          propertyClientHeight: vscode.l10n.t("Client height"),
          unset: vscode.l10n.t("Not set"),
          nullValue: vscode.l10n.t("null"),
          valueType: vscode.l10n.t("Value type"),
          typeString: vscode.l10n.t("String"),
          typeNumber: vscode.l10n.t("Number"),
          typeBoolean: vscode.l10n.t("Boolean"),
          invalidNumber: vscode.l10n.t("Enter a finite number."),
          invalidNumberPrecision: vscode.l10n.t(
            "This number would lose precision. Use the source editor.",
          ),
          invalidInteger: vscode.l10n.t("Enter a safe integer."),
          invalidValue: vscode.l10n.t("Enter a valid value."),
          staleProperty: vscode.l10n.t(
            "The document changed. Re-enter this value on the updated form.",
          ),
          propertyReadOnly: vscode.l10n.t(
            "Property editing requires an xlflow version supporting UserForm property edits.",
          ),
        },
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
      vscode.window.onDidChangeActiveColorTheme(() => send({ type: "themeChanged" })),
      panel.onDidChangeViewState(() => {
        if (panel.visible && ready) {
          sendLocalization();
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
          sync.replay(send);
          void update();
        } else if (message.type === "openText") {
          void vscode.commands.executeCommand("vscode.openWith", document.uri, "default");
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
            (page
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
          const prepare = structural
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
