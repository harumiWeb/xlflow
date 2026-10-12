import * as assert from "assert";
import { promises as fs } from "fs";
import * as path from "path";
import * as vscode from "vscode";
import { XlflowLanguageClientManager } from "../../src/client";
import { designerViewType, UserFormEditorProvider } from "../../src/userFormEditor/provider";
import type {
  ClipboardSnapshot,
  CopyResult,
  EditResult,
  HostMessage,
  PreviewDocument,
  PreviewResult,
  SemanticOperation,
  SpecControl,
  WebviewMessage,
} from "../../src/userFormEditor/protocol";

interface CopyFixture {
  snapshot: ClipboardSnapshot;
  document: PreviewDocument;
}

interface PendingCopy {
  document: vscode.TextDocument;
  version: number;
  text: string;
  controlIds: string[];
  resolve: (result: CopyResult) => void;
}

interface RecordedEdit {
  document: vscode.TextDocument;
  version: number;
  text: string;
  operations: SemanticOperation[];
}

interface FakePanel {
  panel: vscode.WebviewPanel;
  messages: HostMessage[];
  receive: (message: unknown) => void;
  setVisible: (visible: boolean) => void;
  dispose: () => void;
}

export async function runUserFormClipboardAssertions(extensionUri: vscode.Uri): Promise<void> {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, "the VS Code test runner should open the temporary xlflow workspace");
  const projectRoot = folder.uri.fsPath;
  const specsRoot = path.join(projectRoot, "src", "forms", "specs");
  const sourcePath = path.join(specsRoot, "ClipboardSource.yaml");
  const targetPath = path.join(specsRoot, "ClipboardTarget.yaml");
  const sourceInitialText = initialFormSpec("ClipboardSource", sourceControls);
  const targetInitialText = initialFormSpec("ClipboardTarget", targetControls);
  await fs.mkdir(specsRoot, { recursive: true });
  await fs.writeFile(sourcePath, sourceInitialText, "utf8");
  await fs.writeFile(targetPath, targetInitialText, "utf8");

  const sourceDocument = await vscode.workspace.openTextDocument(vscode.Uri.file(sourcePath));
  const targetDocument = await vscode.workspace.openTextDocument(vscode.Uri.file(targetPath));
  assert.strictEqual(
    vscode.workspace.getWorkspaceFolder(sourceDocument.uri)?.uri.toString(),
    folder.uri.toString(),
    "the source FormSpec should resolve to the launched xlflow project key",
  );
  assert.strictEqual(
    vscode.workspace.getWorkspaceFolder(targetDocument.uri)?.uri.toString(),
    folder.uri.toString(),
    "the target FormSpec should resolve to the same launched xlflow project key",
  );

  const sourceFixture = copyFixture("source", "ClipboardSource", "source-root", "source-child");
  const targetFixture = copyFixture("target", "ClipboardTarget", "target-root", "target-child");
  const fixtureBySnapshot = new Map([
    [sourceFixture.snapshot.text, sourceFixture],
    [targetFixture.snapshot.text, targetFixture],
  ]);
  const pendingCopies: PendingCopy[] = [];
  const recordedEdits: RecordedEdit[] = [];
  const previewChanged = new vscode.EventEmitter<void>();
  const previewDocuments = new Map<string, SpecControl[]>([
    [path.resolve(sourcePath), cloneControls(sourceControls)],
    [path.resolve(targetPath), cloneControls(targetControls)],
  ]);
  let targetStructuralEditable = true;
  let confirmCut = true;
  let failNextEdit = false;

  const clients = {
    onDidChangePreviewConnection: previewChanged.event,
    requestUserFormPreview: async (document: vscode.TextDocument): Promise<PreviewResult> => {
      const controls = previewDocuments.get(path.resolve(document.uri.fsPath));
      assert.ok(controls, `unexpected FormSpec preview request: ${document.uri.fsPath}`);
      const target = path.resolve(document.uri.fsPath) === path.resolve(targetPath);
      return {
        version: document.version,
        editable: true,
        propertyEditable: true,
        structuralEditable: !target || targetStructuralEditable,
        parentEditable: true,
        pageEditable: true,
        tabEditable: true,
        productivityEditable: true,
        document: {
          form: { name: target ? "ClipboardTarget" : "ClipboardSource", caption: "Clipboard" },
          controls: cloneControls(controls!),
        },
      };
    },
    requestUserFormCopy: (
      document: vscode.TextDocument,
      version: number,
      text: string,
      controlIds: string[],
    ): Promise<CopyResult> =>
      new Promise((resolve) => {
        pendingCopies.push({ document, version, text, controlIds: [...controlIds], resolve });
      }),
    requestUserFormEdit: async (
      document: vscode.TextDocument,
      version: number,
      text: string,
      operations: SemanticOperation[],
    ): Promise<EditResult> => {
      recordedEdits.push({ document, version, text, operations: structuredClone(operations) });
      if (failNextEdit) {
        failNextEdit = false;
        return { version, error: { code: "testBackendFailure", message: "Rejected cut." } };
      }
      assert.strictEqual(
        operations.length,
        1,
        "clipboard actions should send one semantic operation",
      );
      const operation = operations[0];
      if (operation.type === "removeControls") {
        const existing = previewDocuments.get(path.resolve(document.uri.fsPath)) ?? [];
        const removed = new Set(operation.controlIds);
        for (let changed = true; changed; ) {
          changed = false;
          for (const control of existing) {
            if (control.parentId && removed.has(control.parentId) && !removed.has(control.id)) {
              removed.add(control.id);
              changed = true;
            }
          }
        }
        const remaining = existing.filter((control) => !removed.has(control.id));
        previewDocuments.set(path.resolve(document.uri.fsPath), remaining);
        const end = document.positionAt(text.length);
        return {
          version,
          edits: [
            {
              range: {
                start: { line: 0, character: 0 },
                end: { line: end.line, character: end.character },
              },
              newText: initialFormSpec("ClipboardSource", remaining),
            },
          ],
        };
      }
      if (operation.type !== "pasteControls") {
        return {
          version,
          error: { code: "unexpectedOperation", message: `Unexpected ${operation.type}.` },
        };
      }
      const fixture = fixtureBySnapshot.get(operation.snapshot.text);
      assert.ok(fixture, "paste should carry the opaque snapshot returned by the copy request");
      const allocations = new Map(
        operation.allocations.map((allocation) => [allocation.sourceId, allocation]),
      );
      const addedControls = fixture.document.controls.map((control): SpecControl => {
        const allocation = allocations.get(control.id);
        assert.ok(allocation, `paste should allocate an ID and name for ${control.id}`);
        const parentId = control.parentId
          ? (allocations.get(control.parentId)?.id ?? control.parentId)
          : operation.parentId || undefined;
        return {
          ...control,
          id: allocation.id,
          name: allocation.name,
          ...(parentId ? { parentId } : {}),
          left: (control.left ?? 0) + operation.offset.left,
          top: (control.top ?? 0) + operation.offset.top,
        };
      });
      previewDocuments.set(path.resolve(targetPath), [
        ...(previewDocuments.get(path.resolve(targetPath)) ?? []),
        ...addedControls,
      ]);
      const insertion = renderYamlControls(addedControls);
      const position = document.positionAt(text.length);
      return {
        version,
        edits: [
          {
            range: {
              start: { line: position.line, character: position.character },
              end: { line: position.line, character: position.character },
            },
            newText: insertion,
          },
        ],
      };
    },
  } as unknown as XlflowLanguageClientManager;

  const provider = new UserFormEditorProvider(extensionUri, clients, async () => confirmCut);
  const sourcePanel = createFakePanel(true);
  const targetPanel = createFakePanel(true);
  const sourceCancellation = new vscode.CancellationTokenSource();
  const targetCancellation = new vscode.CancellationTokenSource();
  const targetChanges: vscode.TextDocumentChangeEvent[] = [];
  const targetChangeSubscription = vscode.workspace.onDidChangeTextDocument((event) => {
    if (
      event.document.uri.toString() === targetDocument.uri.toString() &&
      event.contentChanges.length > 0
    )
      targetChanges.push(event);
  });

  try {
    console.log("[userFormClipboard] opening source and target preview panels");
    await provider.resolveCustomTextEditor(
      sourceDocument,
      sourcePanel.panel,
      sourceCancellation.token,
    );
    await provider.resolveCustomTextEditor(
      targetDocument,
      targetPanel.panel,
      targetCancellation.token,
    );
    sourcePanel.receive({ type: "ready" });
    targetPanel.receive({ type: "ready" });
    await Promise.all([
      waitForDocument(sourcePanel, sourceDocument, "the source workspace FormSpec preview"),
      waitForDocument(targetPanel, targetDocument, "the target workspace FormSpec preview"),
    ]);
    console.log("[userFormClipboard] both workspace previews are ready");
    targetPanel.setVisible(false);

    const initialCopyIndex = pendingCopies.length;
    startClipboardAction(sourcePanel, sourceDocument, 1, "copy", ["source-root"]);
    await waitFor(() => pendingCopies.length > initialCopyIndex, "the source copy request");
    assert.strictEqual(
      pendingCopies[initialCopyIndex].document.uri.toString(),
      sourceDocument.uri.toString(),
    );
    assert.strictEqual(pendingCopies[initialCopyIndex].text, sourceDocument.getText());
    assert.deepStrictEqual(pendingCopies[initialCopyIndex].controlIds, ["source-root"]);
    resolveCopy(pendingCopies[initialCopyIndex], sourceFixture);
    const initialCopyReply = await waitForReply(sourcePanel, 1);
    assert.strictEqual(initialCopyReply.error, undefined, "Copy A should publish a valid snapshot");
    assert.ok(
      !targetPanel.messages.some(
        (message) => message.type === "clipboardState" && message.available,
      ),
      "a hidden second panel should not receive the clipboard update immediately",
    );

    targetPanel.setVisible(true);
    await waitFor(
      () =>
        targetPanel.messages.some(
          (message) => message.type === "clipboardState" && message.available,
        ),
      "the hidden target panel to replay shared clipboard availability when shown",
    );
    await waitForDocument(targetPanel, targetDocument, "the visible target panel preview");

    const originalTarget = targetDocument.getText();
    const targetChangeBaseline = targetChanges.length;
    startClipboardAction(targetPanel, targetDocument, 2, "paste", ["target-root"]);
    const pasteReply = await waitForReply(targetPanel, 2);
    assert.strictEqual(pasteReply.error, undefined, "Paste B should apply successfully");
    const firstPasteEdit = recordedEdits.at(-1);
    assert.ok(firstPasteEdit);
    assert.strictEqual(firstPasteEdit.document.uri.toString(), targetDocument.uri.toString());
    assert.strictEqual(firstPasteEdit.operations.length, 1);
    const firstPaste = firstPasteEdit.operations[0];
    assert.strictEqual(firstPaste.type, "pasteControls");
    if (firstPaste.type !== "pasteControls") throw new Error("expected pasteControls");
    assert.deepStrictEqual(
      firstPaste.snapshot,
      sourceFixture.snapshot,
      "paste should retain the opaque copy snapshot",
    );
    assert.deepStrictEqual(
      firstPaste.allocations.map(({ sourceId }) => sourceId),
      ["source-root", "source-child"],
    );
    assert.deepStrictEqual(
      firstPaste.allocations.map(({ name }) => name),
      ["Frame2", "Label2"],
    );
    assert.strictEqual(new Set(firstPaste.allocations.map(({ id }) => id)).size, 2);
    assert.ok(firstPaste.allocations.every(({ id }) => id.startsWith("control-")));
    assert.deepStrictEqual(firstPaste.offset, { left: 8, top: 8 });
    assert.deepStrictEqual(pasteReply.selectionIds, [firstPaste.allocations[0].id]);
    await waitFor(
      () => targetChanges.length === targetChangeBaseline + 1,
      "one real WorkspaceEdit change for the pasted subtree",
    );
    const afterFirstPaste = targetDocument.getText();
    assert.ok(
      afterFirstPaste.startsWith(originalTarget),
      "the WorkspaceEdit should preserve the original source bytes",
    );
    assert.ok(afterFirstPaste.includes(`id: ${JSON.stringify(firstPaste.allocations[0].id)}`));
    assert.ok(afterFirstPaste.includes(`name: ${JSON.stringify("Frame2")}`));
    assert.ok(
      afterFirstPaste.includes(`parentId: ${JSON.stringify(firstPaste.allocations[0].id)}`),
      "the copied child should reference its newly allocated root ID",
    );

    await vscode.commands.executeCommand("vscode.openWith", targetDocument.uri, designerViewType);
    await waitFor(
      () => activeCustomEditorIs(designerViewType, targetDocument.uri),
      "the real custom editor for target FormSpec B to become active",
    );
    await vscode.commands.executeCommand("undo");
    await waitFor(
      () => targetDocument.getText() === originalTarget,
      "Undo in custom editor B to remove the pasted subtree",
    );
    await vscode.commands.executeCommand("redo");
    await waitFor(
      () => targetDocument.getText() === afterFirstPaste,
      "Redo in custom editor B to restore the pasted subtree",
    );
    await waitForDocument(targetPanel, targetDocument, "the target preview after Undo and Redo");
    console.log("[userFormClipboard] Paste, custom-editor Undo, and Redo completed");

    const olderCopyIndex = pendingCopies.length;
    startClipboardAction(sourcePanel, sourceDocument, 3, "copy", ["source-root"]);
    await waitFor(() => pendingCopies.length > olderCopyIndex, "the earlier panel copy request");
    const newerCopyIndex = pendingCopies.length;
    startClipboardAction(targetPanel, targetDocument, 4, "copy", ["target-root"]);
    await waitFor(
      () => pendingCopies.length > newerCopyIndex,
      "two overlapping copy requests from different panels",
    );
    resolveCopy(pendingCopies[newerCopyIndex], targetFixture);
    const newerCopyReply = await waitForReply(targetPanel, 4);
    assert.strictEqual(newerCopyReply.error, undefined);
    resolveCopy(pendingCopies[olderCopyIndex], sourceFixture);
    const olderCopyReply = await waitForReply(sourcePanel, 3);
    assert.strictEqual(
      olderCopyReply.error?.code,
      "copySuperseded",
      "a slower earlier copy must not replace the newer selection",
    );
    console.log("[userFormClipboard] overlapping Copy requests kept the newest clipboard snapshot");

    const editsBeforeSupersededCut = recordedEdits.length;
    const sourceBeforeSupersededCut = sourceDocument.getText();
    const supersededCutIndex = pendingCopies.length;
    startClipboardAction(sourcePanel, sourceDocument, 5, "cut", ["source-root"]);
    await waitFor(() => pendingCopies.length > supersededCutIndex, "the source Cut copy request");
    const supersedingCopyIndex = pendingCopies.length;
    startClipboardAction(targetPanel, targetDocument, 6, "copy", ["target-root"]);
    await waitFor(
      () => pendingCopies.length > supersedingCopyIndex,
      "the newer target Copy request",
    );
    resolveCopy(pendingCopies[supersedingCopyIndex], targetFixture);
    const supersedingCopyReply = await waitForReply(targetPanel, 6);
    assert.strictEqual(supersedingCopyReply.error, undefined);
    resolveCopy(pendingCopies[supersededCutIndex], sourceFixture);
    const supersededCutReply = await waitForReply(sourcePanel, 5);
    assert.strictEqual(supersededCutReply.error?.code, "staleDocument");
    assert.strictEqual(sourceDocument.getText(), sourceBeforeSupersededCut);
    assert.strictEqual(
      recordedEdits.length,
      editsBeforeSupersededCut,
      "a Cut superseded by a newer Copy must stop before requesting deletion",
    );

    const changedCopyIndex = pendingCopies.length;
    startClipboardAction(sourcePanel, sourceDocument, 7, "copy", ["source-root"]);
    await waitFor(
      () => pendingCopies.length > changedCopyIndex,
      "a delayed copy before source mutation",
    );
    const originalSource = sourceDocument.getText();
    const sourcePosition = sourceDocument.positionAt(originalSource.length);
    const sourceMutation = new vscode.WorkspaceEdit();
    sourceMutation.insert(sourceDocument.uri, sourcePosition, "# changed while copy was pending\n");
    assert.ok(
      await vscode.workspace.applyEdit(sourceMutation),
      "the source-change setup edit should apply",
    );
    resolveCopy(pendingCopies[changedCopyIndex], sourceFixture);
    const changedCopyReply = await waitForReply(sourcePanel, 7);
    assert.strictEqual(changedCopyReply.error?.code, "staleDocument");
    assert.strictEqual(
      recordedEdits.length,
      1,
      "copy requests and stale copies should never issue a source edit",
    );
    await waitForDocument(
      sourcePanel,
      sourceDocument,
      "the source panel to refresh after the delayed copy becomes stale",
      (message) => message.productivityEditable === true,
    );

    targetStructuralEditable = false;
    previewChanged.fire();
    await waitForDocument(
      targetPanel,
      targetDocument,
      "the target panel preview with structural editing disabled",
      (message) => message.structuralEditable === false && message.productivityEditable === true,
    );
    const editCountBeforeDeniedPaste = recordedEdits.length;
    const copyCountBeforeDeniedPaste = pendingCopies.length;
    startClipboardAction(targetPanel, targetDocument, 8, "paste", ["target-root"]);
    const deniedPasteReply = await waitForReply(targetPanel, 8);
    assert.ok(
      deniedPasteReply.error,
      "paste should be rejected when its structural capability is absent",
    );
    assert.strictEqual(recordedEdits.length, editCountBeforeDeniedPaste);
    assert.strictEqual(pendingCopies.length, copyCountBeforeDeniedPaste);
    targetStructuralEditable = true;
    previewChanged.fire();
    await waitForDocument(
      targetPanel,
      targetDocument,
      "the target panel to restore its structural-edit capability",
      (message) => message.structuralEditable === true,
    );
    console.log(
      "[userFormClipboard] capability denial produced no edit; capability recovery succeeded",
    );

    startClipboardAction(targetPanel, targetDocument, 9, "paste", ["target-root"]);
    const newerClipboardPasteReply = await waitForReply(targetPanel, 9);
    assert.strictEqual(newerClipboardPasteReply.error, undefined);
    const newerClipboardPaste = recordedEdits.at(-1)?.operations[0];
    assert.ok(newerClipboardPaste?.type === "pasteControls");
    if (!newerClipboardPaste || newerClipboardPaste.type !== "pasteControls")
      throw new Error("expected pasteControls from the newest clipboard snapshot");
    assert.deepStrictEqual(newerClipboardPaste.snapshot, targetFixture.snapshot);
    assert.deepStrictEqual(
      newerClipboardPaste.allocations.map(({ name }) => name),
      ["Frame3", "Label3"],
    );

    await waitForDocument(
      sourcePanel,
      sourceDocument,
      "the source panel before the cut checks",
      (message) => message.productivityEditable === true,
    );
    const sourceBeforeFailedCut = sourceDocument.getText();
    const failedCutIndex = pendingCopies.length;
    startClipboardAction(sourcePanel, sourceDocument, 10, "cut", ["source-root"]);
    await waitFor(
      () => pendingCopies.length > failedCutIndex,
      "the cut copy snapshot before a failed edit",
    );
    failNextEdit = true;
    resolveCopy(pendingCopies[failedCutIndex], sourceFixture);
    const failedCutReply = await waitForReply(sourcePanel, 10);
    assert.strictEqual(failedCutReply.error?.code, "testBackendFailure");
    assert.strictEqual(
      sourceDocument.getText(),
      sourceBeforeFailedCut,
      "a failed Cut must leave its source intact",
    );

    confirmCut = false;
    const editCountBeforeCancelledCut = recordedEdits.length;
    const sourceBeforeCancelledCut = sourceDocument.getText();
    const cancelledCutIndex = pendingCopies.length;
    startClipboardAction(sourcePanel, sourceDocument, 11, "cut", ["source-root"]);
    await waitFor(
      () => pendingCopies.length > cancelledCutIndex,
      "the cut copy snapshot before confirmation cancellation",
    );
    resolveCopy(pendingCopies[cancelledCutIndex], sourceFixture);
    const cancelledCutReply = await waitForReply(sourcePanel, 11);
    assert.strictEqual(cancelledCutReply.error?.code, "editCancelled");
    assert.strictEqual(
      sourceDocument.getText(),
      sourceBeforeCancelledCut,
      "a cancelled Cut must leave its source intact",
    );
    assert.strictEqual(
      recordedEdits.length,
      editCountBeforeCancelledCut,
      "a cancelled Cut must not request an edit",
    );
    assert.deepStrictEqual(
      recordedEdits.filter((edit) => edit.operations[0]?.type === "removeControls").length,
      1,
      "only the failed Cut should reach the edit client",
    );
    await waitForDocument(
      targetPanel,
      targetDocument,
      "the target before checking retained clipboard",
    );
    startClipboardAction(targetPanel, targetDocument, 12, "paste", ["target-root"]);
    assert.strictEqual((await waitForReply(targetPanel, 12)).error, undefined);
    const retainedPaste = recordedEdits.at(-1)!.operations[0];
    assert.ok(retainedPaste.type === "pasteControls");
    if (retainedPaste.type !== "pasteControls")
      throw new Error("expected retained clipboard paste");
    assert.deepStrictEqual(
      retainedPaste.snapshot,
      targetFixture.snapshot,
      "failed and cancelled Cut retain the previous clipboard",
    );
    assert.deepStrictEqual(
      retainedPaste.offset,
      { left: 16, top: 16 },
      "failed and cancelled Cut retain the paste repetition counter",
    );
    confirmCut = true;
    await waitForDocument(sourcePanel, sourceDocument, "the source before successful Cut");
    const successfulCutIndex = pendingCopies.length;
    startClipboardAction(sourcePanel, sourceDocument, 13, "cut", ["source-root"]);
    await waitFor(() => pendingCopies.length > successfulCutIndex, "the successful Cut snapshot");
    resolveCopy(pendingCopies[successfulCutIndex], sourceFixture);
    assert.strictEqual((await waitForReply(sourcePanel, 13)).error, undefined);
    assert.ok(
      !sourceDocument.getText().includes("source-root") &&
        !sourceDocument.getText().includes("source-child"),
      "successful Cut removes the complete subtree",
    );
    await waitForDocument(targetPanel, targetDocument, "the target after successful Cut");
    startClipboardAction(targetPanel, targetDocument, 14, "paste", ["target-root"]);
    assert.strictEqual((await waitForReply(targetPanel, 14)).error, undefined);
    const cutPaste = recordedEdits.at(-1)!.operations[0];
    assert.ok(cutPaste.type === "pasteControls");
    if (cutPaste.type !== "pasteControls") throw new Error("expected cut clipboard paste");
    assert.deepStrictEqual(
      cutPaste.snapshot,
      sourceFixture.snapshot,
      "successful Cut publishes the captured source snapshot",
    );
    assert.deepStrictEqual(
      cutPaste.offset,
      { left: 8, top: 8 },
      "successful Cut starts a new paste generation",
    );
    console.log(
      "[userFormClipboard] failed/cancelled Cut retained clipboard; successful Cut published its subtree",
    );
  } finally {
    targetChangeSubscription.dispose();
    sourcePanel.dispose();
    targetPanel.dispose();
    sourceCancellation.dispose();
    targetCancellation.dispose();
    previewChanged.dispose();
    provider.dispose();
    console.log("[userFormClipboard] restoring only the two clipboard fixture documents");
    await restoreFixtureDocument(sourceDocument, sourceInitialText);
    await restoreFixtureDocument(targetDocument, targetInitialText);
    console.log("[userFormClipboard] closing only the restored clipboard fixture editor tabs");
    await closeFixtureEditors([sourceDocument.uri, targetDocument.uri]);
    console.log("[userFormClipboard] fixture editor tabs closed");
  }
}

const sourceControls: SpecControl[] = [
  {
    id: "source-root",
    name: "Frame1",
    type: "Frame",
    caption: "Copied frame",
    left: 10,
    top: 12,
    width: 140,
    height: 90,
  },
  {
    id: "source-child",
    parentId: "source-root",
    name: "Label1",
    type: "Label",
    caption: "Nested label",
    left: 8,
    top: 9,
    width: 60,
    height: 18,
  },
];

const targetControls: SpecControl[] = [
  {
    id: "target-root",
    name: "Frame1",
    type: "Frame",
    caption: "Existing frame",
    left: 0,
    top: 0,
    width: 160,
    height: 110,
  },
  {
    id: "target-child",
    parentId: "target-root",
    name: "Label1",
    type: "Label",
    caption: "Existing label",
    left: 5,
    top: 5,
    width: 45,
    height: 16,
  },
];

function copyFixture(
  prefix: string,
  formName: string,
  rootId: string,
  childId: string,
): CopyFixture {
  const document: PreviewDocument = {
    form: { name: formName, caption: "Copied form" },
    controls: [
      {
        ...sourceControls[0],
        id: rootId,
      },
      {
        ...sourceControls[1],
        id: childId,
        parentId: rootId,
      },
    ],
  };
  return {
    snapshot: {
      format: "yaml",
      text: `# opaque ${prefix} snapshot\ncontrols: [${rootId}, ${childId}]\n`,
      controlIds: [rootId],
    },
    document,
  };
}

function initialFormSpec(name: string, controls: SpecControl[]): string {
  const lines = [
    "schemaVersion: 1",
    "kind: xlflow.userform",
    "basis: designer",
    "form:",
    `  name: ${JSON.stringify(name)}`,
    `  caption: ${JSON.stringify(name)}`,
    "controls:",
  ];
  for (const control of controls) {
    lines.push(`  - id: ${JSON.stringify(control.id)}`);
    if (control.parentId) lines.push(`    parentId: ${JSON.stringify(control.parentId)}`);
    lines.push(
      `    name: ${JSON.stringify(control.name)}`,
      `    type: ${JSON.stringify(control.type)}`,
      `    caption: ${JSON.stringify(control.caption ?? "")}`,
      `    left: ${control.left ?? 0}`,
      `    top: ${control.top ?? 0}`,
      `    width: ${control.width ?? 80}`,
      `    height: ${control.height ?? 20}`,
    );
  }
  return `${lines.join("\n")}\n`;
}

function cloneControls(controls: SpecControl[]): SpecControl[] {
  return controls.map((control) => ({ ...control }));
}

function renderYamlControls(controls: SpecControl[]): string {
  const lines: string[] = [];
  for (const control of controls) {
    lines.push(`  - id: ${JSON.stringify(control.id)}`);
    if (control.parentId) lines.push(`    parentId: ${JSON.stringify(control.parentId)}`);
    lines.push(
      `    name: ${JSON.stringify(control.name)}`,
      `    type: ${JSON.stringify(control.type)}`,
      `    caption: ${JSON.stringify(control.caption ?? "")}`,
      `    left: ${control.left ?? 0}`,
      `    top: ${control.top ?? 0}`,
      `    width: ${control.width ?? 80}`,
      `    height: ${control.height ?? 20}`,
    );
  }
  return `${lines.join("\n")}\n`;
}

function createFakePanel(initiallyVisible: boolean): FakePanel {
  const incoming = new vscode.EventEmitter<unknown>();
  const viewState = new vscode.EventEmitter<vscode.WebviewPanelOnDidChangeViewStateEvent>();
  const disposed = new vscode.EventEmitter<void>();
  const messages: HostMessage[] = [];
  let visible = initiallyVisible;
  const panel = {
    get visible() {
      return visible;
    },
    onDidChangeViewState: viewState.event,
    onDidDispose: disposed.event,
    webview: {
      options: {},
      html: "",
      cspSource: "https://test.invalid",
      asWebviewUri: (uri: vscode.Uri) => uri,
      onDidReceiveMessage: incoming.event,
      postMessage: async (message: HostMessage) => {
        if (visible) messages.push(message);
        return visible;
      },
    },
  } as unknown as vscode.WebviewPanel;
  return {
    panel,
    messages,
    receive: (message) => incoming.fire(message),
    setVisible: (next) => {
      visible = next;
      viewState.fire({ webviewPanel: panel });
    },
    dispose: () => {
      disposed.fire();
      incoming.dispose();
      viewState.dispose();
      disposed.dispose();
    },
  };
}

function startClipboardAction(
  panel: FakePanel,
  document: vscode.TextDocument,
  requestId: number,
  action: Extract<WebviewMessage, { type: "clipboard" }>["action"],
  controlIds: string[],
): void {
  panel.receive({
    type: "clipboard",
    requestId,
    version: document.version,
    action,
    controlIds,
    parentId: "",
  });
}

function resolveCopy(request: PendingCopy, fixture: CopyFixture): void {
  request.resolve({
    version: request.version,
    snapshot: fixture.snapshot,
    document: fixture.document,
    hasPicturePaths: false,
  });
}

type DocumentMessage = Extract<HostMessage, { type: "document" }>;

type PreviewStateMessage = Extract<
  HostMessage,
  { type: "document" | "invalidDocument" | "editingUnavailable" }
>;

function latestPreviewState(panel: FakePanel): PreviewStateMessage | undefined {
  return [...panel.messages]
    .reverse()
    .find(
      (message): message is PreviewStateMessage =>
        message.type === "document" ||
        message.type === "invalidDocument" ||
        message.type === "editingUnavailable",
    );
}

function waitForDocument(
  panel: FakePanel,
  document: vscode.TextDocument,
  description: string,
  predicate: (message: DocumentMessage) => boolean = () => true,
): Promise<DocumentMessage> {
  return waitFor(() => {
    const state = latestPreviewState(panel);
    return state?.type === "document" && state.version === document.version && predicate(state)
      ? state
      : undefined;
  }, description);
}

async function waitForReply(panel: FakePanel, requestId: number) {
  return waitFor(
    () =>
      panel.messages.find(
        (message): message is Extract<HostMessage, { type: "editResult" }> =>
          message.type === "editResult" && message.requestId === requestId,
      ),
    `clipboard reply ${requestId}`,
  );
}

function activeCustomEditorIs(viewType: string, uri: vscode.Uri): boolean {
  return vscode.window.tabGroups.all.some((group) =>
    group.tabs.some((tab) => {
      const input = tab.input as unknown as { viewType?: string; uri?: vscode.Uri };
      return (
        tab.isActive && input.viewType === viewType && input.uri?.toString() === uri.toString()
      );
    }),
  );
}

async function restoreFixtureDocument(
  document: vscode.TextDocument,
  originalText: string,
): Promise<void> {
  if (document.getText() !== originalText) {
    const edit = new vscode.WorkspaceEdit();
    edit.replace(
      document.uri,
      new vscode.Range(document.positionAt(0), document.positionAt(document.getText().length)),
      originalText,
    );
    assert.ok(
      await vscode.workspace.applyEdit(edit),
      `could not restore fixture ${document.uri.fsPath}`,
    );
  }
  if (document.isDirty) {
    assert.ok(await document.save(), `could not save restored fixture ${document.uri.fsPath}`);
  }
}

async function closeFixtureEditors(uris: readonly vscode.Uri[]): Promise<void> {
  const fixtureUris = new Set(uris.map((uri) => uri.toString()));
  for (const group of vscode.window.tabGroups.all) {
    for (const tab of group.tabs.slice()) {
      const input = tab.input;
      if (
        (input instanceof vscode.TabInputText || input instanceof vscode.TabInputCustom) &&
        fixtureUris.has(input.uri.toString())
      ) {
        await vscode.window.tabGroups.close(tab, true);
      }
    }
  }
}

async function waitFor<T>(read: () => T | undefined | false, description: string): Promise<T> {
  const deadline = Date.now() + 8_000;
  while (Date.now() < deadline) {
    const value = read();
    if (value !== undefined && value !== false) return value as T;
    await new Promise<void>((resolve) => setTimeout(resolve, 20));
  }
  assert.fail(`Timed out waiting for ${description}`);
}
