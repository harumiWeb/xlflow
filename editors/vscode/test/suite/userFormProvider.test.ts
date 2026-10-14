import * as assert from "assert";
import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "../../src/client";
import {
  UserFormEditorProvider,
  prepareParentEdit,
  prepareStructuralEdit,
  prepareCollectionEdit,
} from "../../src/userFormEditor/provider";
import { designerDocument } from "../../src/userFormEditor/model";
import { createControl } from "../../src/userFormEditor/toolbox";
import { initialControlOperations } from "../../src/userFormEditor/collections";
import type { DocumentEditQueue } from "../../src/userFormEditor/edits";
import type { DesignerError, HostMessage, WebviewMessage } from "../../src/userFormEditor/protocol";
import {
  DesignerClipboard,
  pasteOperation,
  prepareProductivityDeletion,
} from "../../src/userFormEditor/clipboard";
import type { ClipboardContent } from "../../src/userFormEditor/clipboard";

// Use the real provider and VS Code events with a controllable panel transport
// and deferred queue result, so visibility cannot race the test driver.
export async function runUserFormProviderAssertions(extensionUri: vscode.Uri): Promise<void> {
  await runClientCapabilityAssertions();
  await runInitialCollectionCapabilityAssertions();
  await runClipboardAssertions();
  await runProductivityCapabilityAssertions();
  await runStructuralPreparationAssertions();
  await runCollectionPreparationAssertions();
  const incoming = new vscode.EventEmitter<WebviewMessage>();
  const view = new vscode.EventEmitter<vscode.WebviewPanelOnDidChangeViewStateEvent>();
  const disposed = new vscode.EventEmitter<void>();
  const connection = new vscode.EventEmitter<void>();
  const cancel = new vscode.CancellationTokenSource();
  const messages: HostMessage[] = [];
  let visible = true;
  let delivered = true;
  const panel = {
    get visible() {
      return visible;
    },
    onDidChangeViewState: view.event,
    onDidDispose: disposed.event,
    webview: {
      options: {},
      html: "",
      cspSource: "https://test.invalid",
      asWebviewUri: (uri: vscode.Uri) => uri,
      onDidReceiveMessage: incoming.event,
      postMessage: async (message: HostMessage) => {
        if (delivered) messages.push(message);
        return delivered;
      },
    },
  } as unknown as vscode.WebviewPanel;
  const client = {
    onDidChangePreviewConnection: connection.event,
  } as unknown as XlflowLanguageClientManager;
  const provider = new UserFormEditorProvider(extensionUri, client);
  const document = await vscode.workspace.openTextDocument({
    language: "yaml",
    content: "form: {name: Main, caption: Original}\ncontrols: []\n",
  });
  const customDocument = await provider.openCustomDocument(
    document.uri,
    { backupId: undefined, untitledDocumentData: undefined },
    cancel.token,
  );
  let finish!: (error: DesignerError | undefined) => void;
  let current!: () => boolean;
  (provider as unknown as { edits: Pick<DocumentEditQueue, "run"> }).edits = {
    run: async (_document, _version, _operations, _client, permitted) => {
      current = permitted;
      return new Promise<DesignerError | undefined>((resolve) => {
        finish = resolve;
      });
    },
  };
  const source = document.getText();
  const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0));
  const show = async () => {
    visible = true;
    view.fire({ webviewPanel: panel });
    await tick();
  };
  const edit = (requestId: number) =>
    incoming.fire({
      type: "edit",
      requestId,
      version: customDocument.document.version,
      operations: [{ type: "setFormProperty", field: "caption", value: "Draft" }],
    });
  const replies = () => messages.filter((m) => m.type === "editResult");
  try {
    await provider.resolveCustomEditor(customDocument, panel, cancel.token);
    incoming.fire({ type: "ready" });
    await tick();
    const errors: (DesignerError | undefined)[] = [
      { code: "staleDocument", message: "Hidden edit is stale" },
      {
        code: "invalid",
        message: "Invalid caption",
        diagnostics: [
          { operationIndex: 0, field: "caption", code: "invalid_caption", message: "Rejected" },
        ],
      },
      undefined,
    ];
    for (const [index, error] of errors.entries()) {
      const requestId = index + 1;
      const before = replies().length;
      edit(requestId);
      visible = false;
      assert.strictEqual(current(), false, "hidden panels cannot apply a late source edit");
      finish(error);
      await tick();
      assert.strictEqual(replies().length, before, "hidden replies are retained, not posted");
      await show();
      assert.deepStrictEqual(replies().at(-1), { type: "editResult", requestId, error });
      assert.strictEqual(replies().length, before + 1);
      await show();
      assert.strictEqual(replies().length, before + 1, "delivered replies are not replayed twice");
      assert.strictEqual(document.getText(), source);
    }
    // A transport rejection must remain queued until a subsequent visible flush.
    edit(4);
    delivered = false;
    finish(undefined);
    await tick();
    const beforeRetry = replies().length;
    delivered = true;
    await show();
    assert.strictEqual(replies().length, beforeRetry + 1);
    assert.deepStrictEqual(replies().at(-1), {
      type: "editResult",
      requestId: 4,
      error: undefined,
    });

    // Recreated Webviews restart IDs: neither queued nor still-running old
    // replies may acknowledge a new context's request with the same ID.
    edit(5);
    visible = false;
    finish(undefined);
    await tick();
    edit(6);
    const beforeReset = replies().length;
    visible = true;
    incoming.fire({ type: "ready" });
    assert.strictEqual(current(), false, "a new context invalidates old in-flight edits");
    finish(undefined);
    await tick();
    await show();
    assert.strictEqual(replies().length, beforeReset);
    edit(1);
    finish(undefined);
    await tick();
    assert.strictEqual(
      replies().length,
      beforeReset + 1,
      "new context edits still receive replies",
    );
  } finally {
    disposed.fire();
    incoming.dispose();
    view.dispose();
    disposed.dispose();
    connection.dispose();
    cancel.dispose();
  }
}

async function runClipboardAssertions(): Promise<void> {
  const clipboard = new DesignerClipboard();
  const content: ClipboardContent = {
    projectKey: "project-a",
    hasPicturePaths: true,
    snapshot: { format: "yaml", text: "opaque source with explicit null", controlIds: ["frame"] },
    document: designerDocument({
      form: { name: "Source" },
      controls: [
        { id: "frame", name: "Frame1", type: "Frame" },
        { id: "child", name: "Label1", type: "Label", parentId: "frame" },
      ],
    }),
  };
  try {
    clipboard.write(content);
    const earlierCopy = clipboard.beginCopy();
    const laterCopy = clipboard.beginCopy();
    assert.strictEqual(clipboard.finishCopy(laterCopy, content), true);
    assert.strictEqual(
      clipboard.finishCopy(earlierCopy, { ...content, projectKey: "obsolete" }),
      false,
    );
    assert.strictEqual(
      clipboard.read()!.projectKey,
      "project-a",
      "late old copies cannot overwrite a newer selection",
    );
    content.snapshot.text = "mutated outside clipboard";
    assert.strictEqual(clipboard.read()!.snapshot.text, "opaque source with explicit null");
    const read = clipboard.read()!;
    read.document.controls.length = 0;
    assert.strictEqual(clipboard.read()!.document.controls.length, 2);
    assert.strictEqual(clipboard.nextOffset("form-b", ""), 8);
    const generation = clipboard.generation;
    clipboard.pasted("form-b", "", generation);
    assert.strictEqual(clipboard.nextOffset("form-b", ""), 16);
    assert.strictEqual(clipboard.nextOffset("form-c", ""), 8);
    clipboard.write(clipboard.read()!);
    clipboard.pasted("form-b", "", generation);
    assert.strictEqual(
      clipboard.nextOffset("form-b", ""),
      8,
      "an older paste cannot advance a new clipboard",
    );
    const target = designerDocument({
      form: { name: "LABEL1" },
      controls: [
        { id: "keep", name: "FRAME1", type: "Frame" },
        { id: "label", name: "Label2", type: "Label" },
      ],
    });
    let sequence = 0;
    const operation = pasteOperation(clipboard.read()!, target, "project-a", "", 8, false, () =>
      String(++sequence),
    );
    assert.ok("type" in operation);
    if ("type" in operation) {
      assert.deepStrictEqual(
        operation.allocations.map((a) => a.name),
        ["Frame2", "Label3"],
      );
      assert.deepStrictEqual(
        operation.allocations.map((a) => a.id),
        ["control-1", "control-2"],
      );
      assert.strictEqual(operation.snapshot.text, "opaque source with explicit null");
    }
    const crossProject = pasteOperation(clipboard.read()!, target, "project-b", "", 8, false);
    assert.ok("code" in crossProject);
    if ("code" in crossProject) assert.strictEqual(crossProject.code, "crossProjectPicture");
    const noPicture = { ...clipboard.read()!, hasPicturePaths: false };
    assert.ok("type" in pasteOperation(noPicture, target, "project-b", "", 8, false));
    assert.ok("code" in pasteOperation(noPicture, target, "project-a", "label", 8, false));
    let confirmations = 0;
    const selection = designerDocument({
      form: { name: "Main" },
      controls: [
        { id: "a", name: "Frame1", type: "Frame" },
        { id: "b", name: "Frame2", type: "Frame" },
        { id: "c", name: "Label1", type: "Label", parentId: "a" },
        { id: "d", name: "Label2", type: "Label", parentId: "b" },
      ],
    });
    assert.strictEqual(
      await prepareProductivityDeletion(
        {
          type: "removeControls",
          controlIds: ["a", "b"],
          cascade: true,
        },
        selection,
        async (names, descendants) => {
          confirmations++;
          assert.strictEqual(names, "Frame1, Frame2");
          assert.strictEqual(descendants, 2);
          return true;
        },
      ),
      undefined,
    );
    assert.strictEqual(confirmations, 1, "one confirmation covers all selected subtrees");
    assert.strictEqual(
      (
        await prepareProductivityDeletion(
          {
            type: "removeControls",
            controlIds: ["a", "c"],
            cascade: true,
          },
          selection,
          async () => true,
        )
      )?.code,
      "invalidSelection",
    );
  } finally {
    clipboard.dispose();
  }
}

async function runProductivityCapabilityAssertions(): Promise<void> {
  const document = await vscode.workspace.openTextDocument({ content: "form: {name: Main}" });
  for (const enabled of [undefined, false, true]) {
    let requests = 0;
    const manager = Object.assign(Object.create(XlflowLanguageClientManager.prototype), {
      previewGeneration: 0,
      workspaceFolderKey: vscode.workspace.getWorkspaceFolder(document.uri)?.uri.toString(),
      client: {
        isRunning: () => true,
        initializeResult: {
          capabilities: {
            experimental: {
              userFormPreview: true,
              userFormEdit: true,
              userFormStructuralEdit: true,
              userFormPropertyEdit: true,
              userFormProductivityEdit: enabled,
            },
          },
        },
        sendRequest: async () => {
          requests++;
          return {
            version: document.version,
            edits: [],
            document: { form: { name: "Main" }, controls: [] },
          };
        },
      },
    }) as XlflowLanguageClientManager;
    assert.strictEqual(
      (await manager.requestUserFormPreview(document)).productivityEditable,
      enabled === true,
    );
    const before = requests;
    const request = manager.requestUserFormEdit(document, document.version, document.getText(), [
      {
        type: "moveControls",
        moves: [
          { controlId: "a", left: 1, top: 2 },
          { controlId: "b", left: 3, top: 4 },
        ],
      },
    ]);
    if (enabled === true) await request;
    else await assert.rejects(request);
    const copy = manager.requestUserFormCopy(document, document.version, document.getText(), ["a"]);
    if (enabled === true) await copy;
    else await assert.rejects(copy);
    assert.strictEqual(requests - before, enabled === true ? 2 : 0);
  }
}

async function runCollectionPreparationAssertions(): Promise<void> {
  const document = designerDocument({
    form: { name: "Main" },
    controls: [
      { id: "multi", name: "Multi", type: "MultiPage" },
      { id: "page", name: "Page1", type: "Page", parentId: "multi" },
      { id: "child", name: "Label1", type: "Label", parentId: "page" },
      { id: "tabs", name: "Tabs", type: "TabStrip", tabs: [] },
    ],
  });
  const snapshot = { version: 7, document };
  let confirmations = 0;
  const confirm = async (name: string, count: number) => {
    confirmations++;
    assert.strictEqual(name, "Page1");
    assert.strictEqual(count, 1);
    return true;
  };
  const remove = { type: "removePage", controlId: "page", cascade: true } as const;
  assert.strictEqual(await prepareCollectionEdit(remove, snapshot, 7, confirm), undefined);
  assert.strictEqual(confirmations, 1);
  assert.strictEqual(
    (await prepareCollectionEdit(remove, snapshot, 7, async () => false))?.code,
    "editCancelled",
  );
  assert.strictEqual(
    (await prepareCollectionEdit(remove, snapshot, 8, confirm))?.code,
    "staleDocument",
  );
  assert.strictEqual(
    (await prepareCollectionEdit({ ...remove, cascade: false }, snapshot, 7, confirm))?.code,
    "invalidCollection",
  );
  assert.strictEqual(
    (await prepareCollectionEdit({ ...remove, controlId: "multi" }, snapshot, 7, confirm))?.code,
    "invalidCollection",
  );
  const add = {
    type: "addPage",
    control: { id: "fresh", name: "Page2", type: "Page", parentId: "multi" },
  } as const;
  assert.strictEqual(await prepareCollectionEdit(add, snapshot, 7, confirm), undefined);
  assert.strictEqual(
    (
      await prepareCollectionEdit(
        { ...add, control: { ...add.control, name: "PAGE1" } },
        snapshot,
        7,
        confirm,
      )
    )?.code,
    "controlConflict",
  );
  const parent = { type: "setParent", controlId: "child", parentId: "", left: 1, top: 2 } as const;
  assert.strictEqual(prepareParentEdit(parent, snapshot, 7, true), undefined);
  assert.strictEqual(prepareParentEdit(parent, snapshot, 7)?.code, "invalidParent");
  assert.strictEqual(
    await prepareCollectionEdit(
      { type: "addTab", controlId: "tabs", tab: { name: "Tab1" } },
      snapshot,
      7,
      confirm,
    ),
    undefined,
  );
  assert.strictEqual(
    (
      await prepareCollectionEdit(
        { type: "addTab", controlId: "multi", tab: { name: "Tab1" } },
        snapshot,
        7,
        confirm,
      )
    )?.code,
    "invalidCollection",
  );
  for (const userFormPageEdit of [undefined, false, true])
    for (const userFormTabEdit of [undefined, false, true]) {
      let requests = 0;
      const doc = await vscode.workspace.openTextDocument({
        language: "yaml",
        content: "form: {name: Main}\ncontrols: []\n",
      });
      const manager = Object.assign(Object.create(XlflowLanguageClientManager.prototype), {
        previewGeneration: 0,
        workspaceFolderKey: undefined,
        client: {
          isRunning: () => true,
          initializeResult: {
            capabilities: {
              experimental: { userFormPreview: true, userFormPageEdit, userFormTabEdit },
            },
          },
          sendRequest: async () => {
            requests++;
            return {
              version: doc.version,
              document: { form: { name: "Main" }, controls: [] },
              edits: [],
            };
          },
        },
      }) as XlflowLanguageClientManager;
      const preview = await manager.requestUserFormPreview(doc);
      assert.strictEqual(preview.pageEditable, userFormPageEdit === true);
      assert.strictEqual(preview.tabEditable, userFormTabEdit === true);
      const before = requests;
      if (userFormPageEdit === true)
        await manager.requestUserFormEdit(doc, doc.version, doc.getText(), [add]);
      else
        await assert.rejects(manager.requestUserFormEdit(doc, doc.version, doc.getText(), [add]));
      if (userFormTabEdit === true)
        await manager.requestUserFormEdit(doc, doc.version, doc.getText(), [
          { type: "addTab", controlId: "tabs", tab: { name: "Tab1" } },
        ]);
      else
        await assert.rejects(
          manager.requestUserFormEdit(doc, doc.version, doc.getText(), [
            { type: "addTab", controlId: "tabs", tab: { name: "Tab1" } },
          ]),
        );
      assert.strictEqual(
        requests - before,
        Number(userFormPageEdit === true) + Number(userFormTabEdit === true),
      );
    }
}

async function runClientCapabilityAssertions(): Promise<void> {
  const document = await vscode.workspace.openTextDocument({ content: "form: {name: Main}" });
  const property = { type: "setFormProperty", field: "caption", value: "New" } as const;
  const geometry = { type: "moveControl", controlId: "button", left: 1, top: 2 } as const;
  const structural = { type: "removeControl", controlId: "button", cascade: true } as const;
  const rootAdd = {
    type: "addControl",
    control: { id: "root", name: "Label1", type: "Label", left: 1, top: 2, width: 72, height: 18 },
  } as const;
  const parentedAdd = {
    type: "addControl",
    control: {
      id: "nested",
      name: "Label2",
      type: "Label",
      left: 1,
      top: 2,
      width: 72,
      height: 18,
      parentId: "frame",
    },
  } as const;
  const parent = {
    type: "setParent",
    controlId: "button",
    parentId: "",
    left: 0,
    top: 0,
  } as const;
  for (const userFormParentEdit of [undefined, false, true]) {
    const manager = Object.assign(Object.create(XlflowLanguageClientManager.prototype), {
      previewGeneration: 0,
      workspaceFolderKey: undefined,
      client: {
        isRunning: () => true,
        initializeResult: {
          capabilities: {
            experimental: { userFormPreview: true, userFormParentEdit },
          },
        },
        sendRequest: async () => ({
          version: document.version,
          document: { form: { name: "Main" }, controls: [] },
        }),
      },
    }) as XlflowLanguageClientManager;
    const preview = await manager.requestUserFormPreview(document);
    assert.strictEqual(preview.parentEditable, userFormParentEdit === true);
  }
  for (const userFormEdit of [undefined, false, true]) {
    for (const userFormPropertyEdit of [undefined, false, true]) {
      for (const userFormStructuralEdit of [undefined, false, true]) {
        for (const userFormParentEdit of [undefined, false, true]) {
          let requests = 0;
          const response = { edits: [] };
          const manager = Object.assign(Object.create(XlflowLanguageClientManager.prototype), {
            previewGeneration: 0,
            workspaceFolderKey: vscode.workspace.getWorkspaceFolder(document.uri)?.uri.toString(),
            client: {
              isRunning: () => true,
              initializeResult: {
                capabilities: {
                  experimental: {
                    userFormEdit,
                    userFormPropertyEdit,
                    userFormStructuralEdit,
                    userFormParentEdit,
                  },
                },
              },
              sendRequest: async () => {
                requests++;
                return response;
              },
            },
          }) as XlflowLanguageClientManager;
          for (const operations of [
            [property],
            [geometry],
            [structural],
            [rootAdd],
            [parentedAdd],
            [parent],
            [property, geometry],
          ]) {
            const before = requests;
            const allowed = operations.every((op) => {
              if (op.type === "setFormProperty") return userFormPropertyEdit === true;
              if (op.type === "setParent")
                return userFormEdit === true && userFormParentEdit === true;
              if (op.type === "addControl")
                return (
                  userFormStructuralEdit === true &&
                  (!("parentId" in op.control) ||
                    op.control.parentId === undefined ||
                    userFormParentEdit === true)
                );
              if (op.type === "removeControl") return userFormStructuralEdit === true;
              return userFormEdit === true;
            });
            const request = manager.requestUserFormEdit(
              document,
              document.version,
              document.getText(),
              operations,
            );
            if (allowed) assert.strictEqual(await request, response);
            else await assert.rejects(request);
            assert.strictEqual(
              requests,
              before + Number(allowed),
              "only required capabilities gate transport",
            );
          }
        }
      }
    }
  }
}

async function runInitialCollectionCapabilityAssertions(): Promise<void> {
  const document = await vscode.workspace.openTextDocument({
    language: "json",
    content: '{"form":{"name":"Main"},"controls":[]}',
  });
  const model = designerDocument({ form: { name: "Main" }, controls: [] });
  for (const structural of [false, true])
    for (const pages of [false, true])
      for (const tabs of [false, true]) {
        let requests = 0;
        const manager = Object.assign(Object.create(XlflowLanguageClientManager.prototype), {
          previewGeneration: 0,
          workspaceFolderKey: vscode.workspace.getWorkspaceFolder(document.uri)?.uri.toString(),
          client: {
            isRunning: () => true,
            initializeResult: {
              capabilities: {
                experimental: {
                  userFormStructuralEdit: structural,
                  userFormPageEdit: pages,
                  userFormTabEdit: tabs,
                },
              },
            },
            sendRequest: async () => {
              requests++;
              return { version: document.version, edits: [] };
            },
          },
        }) as XlflowLanguageClientManager;
        for (const type of ["MultiPage", "TabStrip"] as const) {
          const control = createControl(model, type, 0, 0, false)!;
          const operations = initialControlOperations(model, control)!;
          const before = requests;
          const allowed = structural && (type === "MultiPage" ? pages : tabs);
          const request = manager.requestUserFormEdit(
            document,
            document.version,
            document.getText(),
            operations,
          );
          if (allowed) await request;
          else await assert.rejects(request);
          assert.strictEqual(
            requests,
            before + Number(allowed),
            "initial members require only structural and the matching collection capability",
          );
          if (!(type === "MultiPage" ? pages : tabs)) {
            const beforeEmpty = requests;
            await assert.rejects(
              manager.requestUserFormEdit(document, document.version, document.getText(), [
                {
                  type: "addControl",
                  control: {
                    ...control,
                    selectedIndex: -1,
                    ...(type === "TabStrip" ? { tabs: [] } : {}),
                  },
                },
              ]),
            );
            assert.strictEqual(
              requests,
              beforeEmpty,
              "empty collection payloads cannot bypass the owner capability gate",
            );
          }
        }
      }
}

async function runStructuralPreparationAssertions(): Promise<void> {
  const document = designerDocument({
    form: { name: "Main", width: 480, height: 360 },
    controls: [
      { id: "frame", type: "Frame", name: "Frame1" },
      { id: "nested", type: "Frame", name: "Frame2", parentId: "frame" },
      { id: "child", type: "TextBox", name: "TextBox1", parentId: "nested" },
      { id: "other", type: "Label", name: "Label1" },
      { id: "page", type: "Page", name: "Page1", parentId: "multi" },
      { id: "multi", type: "MultiPage", name: "MultiPage1" },
    ],
  });
  const snapshot = { version: 7, document };
  const operation = { type: "removeControl", controlId: "frame", cascade: true } as const;
  let confirmations = 0;
  const confirm = async (name: string, count: number) => {
    confirmations++;
    assert.strictEqual(name, "Frame1");
    assert.strictEqual(count, 2, "all descendants, excluding siblings, are shown");
    return true;
  };
  assert.strictEqual(await prepareStructuralEdit(operation, snapshot, 7, confirm), undefined);
  assert.strictEqual(confirmations, 1);
  assert.strictEqual(
    (await prepareStructuralEdit(operation, snapshot, 7, async () => false))?.code,
    "editCancelled",
  );
  assert.strictEqual(
    (await prepareStructuralEdit(operation, snapshot, 8, confirm))?.code,
    "staleDocument",
  );
  assert.strictEqual(
    (await prepareStructuralEdit(operation, undefined, 7, confirm))?.code,
    "staleDocument",
  );
  assert.strictEqual(
    await prepareStructuralEdit({ ...operation, controlId: "other" }, snapshot, 7, confirm),
    undefined,
  );
  assert.strictEqual(confirmations, 1, "leaf deletion does not open a confirmation");
  for (const controlId of ["missing", "page"])
    assert.strictEqual(
      (await prepareStructuralEdit({ ...operation, controlId }, snapshot, 7, confirm))?.code,
      "unsupportedDeletion",
    );
  const control = createControl(document, "Label", 0, 0, false, () => "fresh")!;
  assert.strictEqual(
    await prepareStructuralEdit({ type: "addControl", control }, snapshot, 7, confirm),
    undefined,
  );
  for (const conflict of [
    { ...control, id: "child" },
    { ...control, name: "LABEL1" },
    { ...control, name: "MAIN" },
  ])
    assert.strictEqual(
      (await prepareStructuralEdit({ type: "addControl", control: conflict }, snapshot, 7, confirm))
        ?.code,
      "controlConflict",
    );

  const parentOperation = {
    type: "setParent",
    controlId: "child",
    parentId: "frame",
    left: 12,
    top: 16,
  } as const;
  assert.strictEqual(prepareParentEdit(parentOperation, snapshot, 7), undefined);
  for (const progId of ["Forms.Frame.1", "forms.FRAME.1", "Vendor.Frame"]) {
    const frameSnapshot = {
      version: 7,
      document: {
        ...document,
        controls: document.controls.map((c) => (c.id === "frame" ? { ...c, progId } : c)),
      },
    };
    const expected = progId === "Vendor.Frame" ? "invalidParent" : undefined;
    assert.strictEqual(prepareParentEdit(parentOperation, frameSnapshot, 7)?.code, expected);
    assert.strictEqual(
      (
        await prepareStructuralEdit(
          { type: "addControl", control: { ...control, parentId: "frame" } },
          frameSnapshot,
          7,
          confirm,
        )
      )?.code,
      expected,
    );
    assert.strictEqual(
      prepareParentEdit({ ...parentOperation, controlId: "nested", parentId: "" }, frameSnapshot, 7)
        ?.code,
      expected,
      "current custom Frame parent is also rejected",
    );
  }
  assert.strictEqual(
    prepareParentEdit({ ...parentOperation, parentId: "" }, snapshot, 7),
    undefined,
    "empty parentId explicitly reparents to the form root",
  );
  for (const parentId of ["missing", "other", "child"])
    assert.strictEqual(
      prepareParentEdit({ ...parentOperation, parentId }, snapshot, 7)?.code,
      "invalidParent",
    );
  assert.strictEqual(
    prepareParentEdit({ ...parentOperation, controlId: "frame", parentId: "nested" }, snapshot, 7)
      ?.code,
    "invalidParent",
    "a descendant cannot become its ancestor's parent",
  );
  for (const controlId of ["missing", "page"])
    assert.strictEqual(
      prepareParentEdit({ ...parentOperation, controlId }, snapshot, 7)?.code,
      "invalidParent",
    );
  assert.strictEqual(prepareParentEdit(parentOperation, snapshot, 8)?.code, "staleDocument");
  const orphan = designerDocument({
    form: { name: "Main" },
    controls: [
      { id: "frame", type: "Frame", name: "Frame1" },
      { id: "child", type: "Label", name: "Label1", parentId: "missing" },
    ],
  });
  assert.strictEqual(
    prepareParentEdit(parentOperation, { version: 7, document: orphan }, 7)?.code,
    "invalidParent",
    "the current direct parent must resolve to root or a Frame",
  );
  assert.strictEqual(
    await prepareStructuralEdit(
      { type: "addControl", control: { ...control, parentId: "frame" } },
      snapshot,
      7,
      confirm,
    ),
    undefined,
  );
  for (const parentId of ["missing", "other", "page"])
    assert.strictEqual(
      (
        await prepareStructuralEdit(
          { type: "addControl", control: { ...control, parentId } },
          snapshot,
          7,
          confirm,
        )
      )?.code,
      "invalidParent",
    );
}
