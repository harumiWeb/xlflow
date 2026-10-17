import * as assert from "assert";
import { NativeDesignerMouse } from "./designerNativeMouse";
import { mkdtemp, mkdir, readFile, rm, writeFile } from "fs/promises";
import * as path from "path";
import * as vscode from "vscode";
import type { XlflowLanguageClientManager } from "../../src/client";
import { DraftCheckpointStore } from "../../src/userFormEditor/checkpoint";
import { UserFormCustomDocument } from "../../src/userFormEditor/customDocument";
import { UserFormEditorProvider } from "../../src/userFormEditor/provider";
import { designerResourceUri } from "../../src/userFormEditor/resource";
import { createControl } from "../../src/userFormEditor/toolbox";
import { initialControlOperations } from "../../src/userFormEditor/collections";
import { designerDocument } from "../../src/userFormEditor/model";
import { PictureImportSession } from "../../src/userFormEditor/pictures";
import type { PictureSelection } from "../../src/userFormEditor/pictures";
import type {
  HostMessage,
  SemanticOperation,
  SpecControl,
  WebviewMessage,
} from "../../src/userFormEditor/protocol";

const yamlSource = [
  "# keep the file header comment",
  "schemaVersion: 1",
  "kind: xlflow.userform",
  "form:",
  "  name: Main",
  '  caption: "Original" # keep the inline comment',
  "controls: []",
  "",
].join("\n");
const crlfYamlSource = yamlSource.replace(/\n/g, "\r\n");

const jsonSource = [
  "{",
  '  "schemaVersion": 1,',
  '  "kind": "xlflow.userform",',
  '  "form": {',
  '    "name": "Main",',
  '    "caption": "Original"',
  "  },",
  '  "controls": []',
  "}",
  "",
].join("\n");

const jsonControlsSource =
  JSON.stringify(
    {
      schemaVersion: 1,
      kind: "xlflow.userform",
      form: { name: "Main", caption: "Original" },
      controls: [
        {
          id: "control-a",
          type: "CommandButton",
          progId: "Forms.CommandButton.1",
          name: "ButtonA",
          left: 16,
          top: 16,
          width: 80,
          height: 24,
        },
        {
          id: "control-b",
          type: "Label",
          progId: "Forms.Label.1",
          name: "LabelB",
          left: 120,
          top: 16,
          width: 80,
          height: 24,
        },
      ],
    },
    null,
    2,
  ) + "\n";

interface TestClient {
  client: XlflowLanguageClientManager;
  dispose: () => void;
}

interface FakePanel {
  panel: vscode.WebviewPanel;
  messages: HostMessage[];
  receive: (message: WebviewMessage) => void;
  dispose: () => void;
}

interface TestEnvironment {
  directory: string;
  checkpointRoot: vscode.Uri;
  uri: vscode.Uri;
  sourceDocument: vscode.TextDocument;
  customDocument: UserFormCustomDocument;
  provider: UserFormEditorProvider;
  client: TestClient;
  cancellation: vscode.CancellationTokenSource;
  panel: FakePanel;
}

interface DraftProviderTestExtension {
  register(provider: vscode.CustomEditorProvider<UserFormCustomDocument>): vscode.Disposable;
  registerSaveCommand(provider: UserFormEditorProvider, commandId: string): vscode.Disposable;
}

export async function runUserFormDraftProviderAssertions(extensionUri: vscode.Uri): Promise<void> {
  await assertDraftEditUndoRedoAndSave(extensionUri, "yaml", yamlSource);
  await assertPrepareBuildSources(extensionUri);
  await assertReplacementPictureSelection(extensionUri);
  await assertDraftEditUndoRedoAndSave(extensionUri, "yaml", crlfYamlSource);
  await assertDraftEditUndoRedoAndSave(extensionUri, "json", jsonSource);
  await assertMultilineCRLFPublication(extensionUri);
  await assertSaveParticipantPublication(extensionUri);
  await assertInitialCollectionUndoRedo(extensionUri);
  await assertRepeatedOpenSharesDraft(extensionUri);
  await assertExternalChangesMergeIntoDraft(extensionUri);
  await assertExternalControlReorderConflictsWithDraft(extensionUri);
  await assertConflictingSourceChangeBlocksPublish(extensionUri);
  await assertRevertPreservesUnsavedSourceBuffer(extensionUri);
  await assertBackupRestoreRevertAndSaveAs(extensionUri);
  await assertRegisteredProviderNativeSave(extensionUri);
  await assertCheckpointRecovery(extensionUri, "yaml", yamlSource);
  await assertCheckpointRecovery(extensionUri, "json", jsonControlsSource);
  await assertMalformedCheckpointIsRetained(extensionUri);
}

async function assertInitialCollectionUndoRedo(extensionUri: vscode.Uri): Promise<void> {
  for (const type of ["MultiPage", "TabStrip"] as const) {
    await withTestEnvironment(extensionUri, "json", jsonSource, async (environment) => {
      const { customDocument, provider, cancellation, sourceDocument, uri, panel } = environment;
      const model = designerDocument(JSON.parse(jsonSource));
      const control = createControl(model, type, 0, 0, false)!;
      const operations = initialControlOperations(model, control)!;
      const events: vscode.CustomDocumentEditEvent<UserFormCustomDocument>[] = [];
      const subscription = provider.onDidChangeCustomDocument((event) => events.push(event));
      try {
        panel.receive({
          type: "edit",
          requestId: 1,
          version: customDocument.document.version,
          operations,
        });
        await waitFor(
          () =>
            panel.messages.some(
              (message) => message.type === "editResult" && message.requestId === 1,
            ),
          "initial collection creation",
        );
        const reply = panel.messages.find(
          (message) => message.type === "editResult" && message.requestId === 1,
        );
        assert.ok(reply?.type === "editResult" && !reply.error);
        assert.strictEqual(
          events.length,
          1,
          "owner and initial members share one native Undo event",
        );
        const draft = customDocument.state.text;
        const controls = (JSON.parse(draft) as { controls: SpecControl[] }).controls;
        assert.strictEqual(controls[0].selectedIndex, 0);
        assert.strictEqual(
          type === "MultiPage" ? controls.length - 1 : controls[0].tabs?.length,
          2,
        );
        assert.strictEqual(sourceDocument.getText(), jsonSource);
        assert.strictEqual(await readFile(uri.fsPath, "utf8"), jsonSource);
        await events[0].undo();
        assert.strictEqual(customDocument.state.text, jsonSource);
        await events[0].redo();
        assert.strictEqual(customDocument.state.text, draft);
        await provider.publishCustomDocument(customDocument, cancellation.token);
        assert.strictEqual(await readFile(uri.fsPath, "utf8"), draft);
      } finally {
        subscription.dispose();
      }
    });
  }
}

async function assertMultilineCRLFPublication(extensionUri: vscode.Uri): Promise<void> {
  await withTestEnvironment(extensionUri, "yaml", crlfYamlSource, async (environment) => {
    const { customDocument, provider, cancellation, sourceDocument, uri } = environment;
    // Exercise LF edit text against a native CRLF document.
    const frame = customDocument.apply(
      [
        {
          range: { start: { line: 6, character: 0 }, end: { line: 6, character: 0 } },
          newText: "  width: 500\n  height: 450\n",
        },
      ],
      "Resize form",
    );
    assert.ok(frame);
    const expected = crlfYamlSource.replace(
      "controls: []",
      "  width: 500\r\n  height: 450\r\ncontrols: []",
    );
    await provider.publishCustomDocument(customDocument, cancellation.token);
    assert.strictEqual(customDocument.state.conflicted, false);
    assert.strictEqual(customDocument.state.dirty, false);
    assert.strictEqual(customDocument.state.text, expected);
    assert.strictEqual(sourceDocument.getText(), expected);
    assert.strictEqual(await readFile(uri.fsPath, "utf8"), expected);
    customDocument.undo(frame);
    assert.strictEqual(customDocument.state.text, crlfYamlSource);
    customDocument.redo(frame);
    assert.strictEqual(customDocument.state.text, expected);
    await provider.publishCustomDocument(customDocument, cancellation.token);
  });
}

async function assertSaveParticipantPublication(extensionUri: vscode.Uri): Promise<void> {
  await withTestEnvironment(extensionUri, "yaml", yamlSource, async (environment) => {
    const { customDocument, provider, cancellation, sourceDocument, uri } = environment;
    const subscription = vscode.workspace.onWillSaveTextDocument((event) => {
      if (event.document.uri.toString() !== uri.toString()) return;
      // A real VS Code save participant, like insertFinalNewline or a formatter.
      if (!event.document.getText().endsWith("\n\n"))
        event.waitUntil(
          Promise.resolve([
            vscode.TextEdit.insert(
              event.document.positionAt(event.document.getText().length),
              "\n",
            ),
          ]),
        );
    });
    try {
      await sendCaptionEdit(environment, 1, "Saved with formatter");
      await provider.publishCustomDocument(customDocument, cancellation.token);
      const expected = replaceCaption(yamlSource, "Saved with formatter") + "\n";
      await waitFor(() => !customDocument.reconciling, "save participant reconciliation");
      assert.strictEqual(sourceDocument.getText(), expected);
      assert.strictEqual(await readFile(uri.fsPath, "utf8"), expected);
      assert.strictEqual(customDocument.state.text, expected);
      assert.strictEqual(customDocument.state.dirty, false);
      assert.strictEqual(customDocument.state.conflicted, false);
      const frame = customDocument.state.appliedFrames.at(-1)!;
      customDocument.undo(frame);
      assert.strictEqual(customDocument.state.text, yamlSource + "\n");
      customDocument.redo(frame);
      assert.strictEqual(customDocument.state.text, expected);
      await sendCaptionEdit(environment, 2, "Second save");
      await provider.publishCustomDocument(customDocument, cancellation.token);
      assert.strictEqual(customDocument.state.conflicted, false);
    } finally {
      subscription.dispose();
    }
  });
}

async function assertDraftEditUndoRedoAndSave(
  extensionUri: vscode.Uri,
  format: "yaml" | "json",
  initial: string,
): Promise<void> {
  await withTestEnvironment(extensionUri, format, initial, async (environment) => {
    const { cancellation, customDocument, provider, sourceDocument, uri } = environment;
    const events: vscode.CustomDocumentEditEvent<UserFormCustomDocument>[] = [];
    const subscription = provider.onDidChangeCustomDocument((event) => events.push(event));
    const diskBefore = await readFile(uri.fsPath, "utf8");
    const sourceBefore = sourceDocument.getText();

    await sendCaptionEdit(environment, 1, "Draft");
    const draft = replaceCaption(initial, "Draft");
    assert.strictEqual(customDocument.document.getText(), draft);
    assert.strictEqual(
      events.length,
      1,
      "one semantic edit emits one native custom-document event",
    );
    const editEvent = events[0];
    assert.strictEqual(editEvent.document, customDocument);
    assert.strictEqual(
      sourceDocument.getText(),
      sourceBefore,
      "draft edits leave the source buffer untouched",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      diskBefore,
      "draft edits leave disk untouched",
    );

    await editEvent.undo();
    await waitFor(
      () => customDocument.document.getText() === initial,
      "the native custom-document Undo callback",
    );
    assert.strictEqual(sourceDocument.getText(), sourceBefore);
    assert.strictEqual(await readFile(uri.fsPath, "utf8"), diskBefore);

    await editEvent.redo();
    await waitFor(
      () => customDocument.document.getText() === draft,
      "the native custom-document Redo callback",
    );
    assert.strictEqual(sourceDocument.getText(), sourceBefore);
    assert.strictEqual(await readFile(uri.fsPath, "utf8"), diskBefore);
    if (format === "yaml") {
      assert.ok(draft.includes("# keep the file header comment"));
      assert.ok(draft.includes("# keep the inline comment"));
    }

    await provider.publishCustomDocument(customDocument, cancellation.token);
    assert.strictEqual(sourceDocument.getText(), draft);
    assert.strictEqual(await readFile(uri.fsPath, "utf8"), draft);
    assert.strictEqual(customDocument.state.dirty, false);
    assert.strictEqual(sourceDocument.isDirty, false);
    subscription.dispose();
  });
}

async function assertRepeatedOpenSharesDraft(extensionUri: vscode.Uri): Promise<void> {
  await withTestEnvironment(extensionUri, "yaml", yamlSource, async (environment) => {
    const { cancellation, customDocument, panel, provider, sourceDocument, uri } = environment;
    const editorUri = designerResourceUri(uri);
    const reopened = await provider.openCustomDocument(
      editorUri,
      { backupId: undefined, untitledDocumentData: undefined },
      cancellation.token,
    );
    assert.strictEqual(
      reopened,
      customDocument,
      "reopening one virtual URI reuses its draft document",
    );

    const secondPanel = createFakePanel();
    try {
      await provider.resolveCustomEditor(reopened, secondPanel.panel, cancellation.token);
      secondPanel.receive({ type: "ready" });
      await waitForDocument(
        secondPanel,
        reopened.document,
        "the second panel's initial shared draft",
      );

      await sendCaptionEdit(environment, 1, "Shared across panels");
      const sharedDraft = replaceCaption(yamlSource, "Shared across panels");
      await waitForDocument(
        secondPanel,
        reopened.document,
        "the first panel's edit in the second panel",
      );
      assert.strictEqual(reopened.document.getText(), sharedDraft);
      assert.strictEqual(
        sourceDocument.getText(),
        yamlSource,
        "shared draft edits leave the source buffer untouched",
      );
      assert.strictEqual(
        await readFile(uri.fsPath, "utf8"),
        yamlSource,
        "shared draft edits leave disk untouched",
      );
      assert.ok(
        panel.messages.some((message) => message.type === "editResult" && message.requestId === 1),
      );
    } finally {
      secondPanel.dispose();
    }
  });
}

async function assertExternalChangesMergeIntoDraft(extensionUri: vscode.Uri): Promise<void> {
  await withTestEnvironment(extensionUri, "yaml", yamlSource, async (environment) => {
    const { customDocument, provider, sourceDocument, uri } = environment;
    await sendCaptionEdit(environment, 1, "Designer");
    const localDraft = replaceCaption(yamlSource, "Designer");
    const externalPrefix = "# edited in the text editor\n";
    const externalEdit = new vscode.WorkspaceEdit();
    externalEdit.insert(sourceDocument.uri, new vscode.Position(0, 0), externalPrefix);
    assert.ok(
      await vscode.workspace.applyEdit(externalEdit),
      "the external source edit should apply",
    );

    const mergedDraft = externalPrefix + localDraft;
    await waitFor(
      () => !customDocument.reconciling && customDocument.document.getText() === mergedDraft,
      "the external TextDocument change to merge into the current draft",
    );
    assert.strictEqual(customDocument.state.conflicted, false);
    assert.strictEqual(sourceDocument.getText(), externalPrefix + yamlSource);
    assert.strictEqual(await readFile(uri.fsPath, "utf8"), yamlSource);

    await provider.publishCustomDocument(customDocument, environment.cancellation.token);
    assert.strictEqual(sourceDocument.getText(), mergedDraft);
    assert.strictEqual(await readFile(uri.fsPath, "utf8"), mergedDraft);
  });
}

async function assertConflictingSourceChangeBlocksPublish(extensionUri: vscode.Uri): Promise<void> {
  await withTestEnvironment(extensionUri, "yaml", yamlSource, async (environment) => {
    const { customDocument, provider, sourceDocument, uri } = environment;
    const events: vscode.CustomDocumentEditEvent<UserFormCustomDocument>[] = [];
    const subscription = provider.onDidChangeCustomDocument((event) => events.push(event));
    await sendCaptionEdit(environment, 1, "Designer");
    const localDraft = replaceCaption(yamlSource, "Designer");
    const externalSource = replaceCaption(yamlSource, "External");
    const externalEdit = replaceWholeDocument(sourceDocument, externalSource);
    assert.ok(
      await vscode.workspace.applyEdit(externalEdit),
      "the conflicting source edit should apply",
    );

    await waitFor(
      () => customDocument.state.conflicted && !customDocument.reconciling,
      "the overlapping TextDocument and Designer edits to enter conflict",
    );
    assert.strictEqual(customDocument.document.getText(), externalSource);
    assert.strictEqual(
      customDocument.state.text,
      localDraft,
      "the conflicting Designer draft is retained",
    );
    await provider.saveCustomDocument(customDocument, environment.cancellation.token);
    const checkpoint = await new DraftCheckpointStore(environment.checkpointRoot).read(uri);
    assert.strictEqual(
      checkpoint?.state.text,
      localDraft,
      "checkpoint save retains the conflict draft",
    );
    await assert.rejects(
      provider.publishCustomDocument(customDocument, environment.cancellation.token),
      /conflict/i,
    );
    assert.strictEqual(sourceDocument.getText(), externalSource, "the external buffer is retained");
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      yamlSource,
      "blocked save leaves disk untouched",
    );
    assert.strictEqual(
      customDocument.state.text,
      localDraft,
      "blocked save retains the draft text",
    );

    const oldEdit = events[0];
    assert.ok(oldEdit, "the pre-conflict Designer edit should have a native callback");
    customDocument.state.resolveConflict("source");
    customDocument.notify();
    await oldEdit.undo();
    await oldEdit.redo();
    assert.strictEqual(customDocument.document.getText(), externalSource);
    assert.strictEqual(customDocument.state.conflicted, false);
    assert.strictEqual(customDocument.state.dirty, false);
    assert.strictEqual(
      sourceDocument.getText(),
      externalSource,
      "stale callbacks preserve the source buffer",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      yamlSource,
      "stale callbacks preserve disk",
    );
    subscription.dispose();
  });
}

async function assertRevertPreservesUnsavedSourceBuffer(extensionUri: vscode.Uri): Promise<void> {
  await withTestEnvironment(extensionUri, "yaml", yamlSource, async (environment) => {
    const { customDocument, provider, sourceDocument, uri, cancellation } = environment;
    await sendCaptionEdit(environment, 1, "Designer");
    const localDraft = replaceCaption(yamlSource, "Designer");
    const externalSource = replaceCaption(yamlSource, "Unsaved source");
    const externalEdit = replaceWholeDocument(sourceDocument, externalSource);
    assert.ok(
      await vscode.workspace.applyEdit(externalEdit),
      "the unsaved source edit should apply",
    );
    await waitFor(
      () => customDocument.state.conflicted && !customDocument.reconciling,
      "source reconciliation to finish before reverting the Designer draft",
    );
    assert.strictEqual(
      customDocument.state.text,
      localDraft,
      "the Designer draft remains available before revert",
    );
    assert.strictEqual(sourceDocument.getText(), externalSource);
    assert.strictEqual(sourceDocument.isDirty, true, "the direct YAML edit remains unsaved");
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      yamlSource,
      "disk still contains the saved baseline",
    );

    await provider.revertCustomDocument(customDocument, cancellation.token);

    assert.strictEqual(
      customDocument.state.text,
      externalSource,
      "Designer revert adopts the latest source buffer",
    );
    assert.strictEqual(customDocument.document.getText(), externalSource);
    assert.strictEqual(customDocument.state.conflicted, false);
    assert.strictEqual(customDocument.state.dirty, false);
    assert.strictEqual(
      sourceDocument.getText(),
      externalSource,
      "revert does not replace the source buffer",
    );
    assert.strictEqual(
      sourceDocument.isDirty,
      true,
      "revert preserves the unsaved source dirty state",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      yamlSource,
      "revert does not write the source buffer to disk",
    );
  });
}

async function assertExternalControlReorderConflictsWithDraft(
  extensionUri: vscode.Uri,
): Promise<void> {
  await withTestEnvironment(extensionUri, "json", jsonControlsSource, async (environment) => {
    const { customDocument, provider, sourceDocument, uri } = environment;
    const events: vscode.CustomDocumentEditEvent<UserFormCustomDocument>[] = [];
    const subscription = provider.onDidChangeCustomDocument((event) => events.push(event));
    await sendCaptionEdit(environment, 1, "Designer");
    const localDraft = replaceCaption(jsonControlsSource, "Designer");
    const parsed = JSON.parse(jsonControlsSource) as { controls: SpecControl[] };
    const reorderedSource =
      JSON.stringify({ ...parsed, controls: [...parsed.controls].reverse() }, null, 2) + "\n";
    const externalEdit = replaceWholeDocument(sourceDocument, reorderedSource);
    assert.ok(
      await vscode.workspace.applyEdit(externalEdit),
      "the external control reorder should apply",
    );

    await waitFor(
      () => customDocument.state.conflicted && !customDocument.reconciling,
      "the semantic source guard to reject a control-ID reorder during a pending root caption edit",
    );
    assert.strictEqual(
      customDocument.state.text,
      localDraft,
      "the original Designer draft is retained",
    );
    assert.strictEqual(
      customDocument.document.getText(),
      reorderedSource,
      "the conflicted preview follows the latest external source buffer",
    );
    assert.strictEqual(sourceDocument.getText(), reorderedSource);
    assert.strictEqual(await readFile(uri.fsPath, "utf8"), jsonControlsSource);
    await provider.saveCustomDocument(customDocument, environment.cancellation.token);
    const checkpoint = await new DraftCheckpointStore(environment.checkpointRoot).read(uri);
    assert.strictEqual(checkpoint?.state.text, localDraft);
    await assert.rejects(
      provider.publishCustomDocument(customDocument, environment.cancellation.token),
      /conflict/i,
      "a semantic identity conflict must block publishing",
    );
    assert.strictEqual(customDocument.state.text, localDraft);
    assert.strictEqual(customDocument.document.getText(), reorderedSource);
    assert.strictEqual(
      sourceDocument.getText(),
      reorderedSource,
      "blocked save preserves the latest source buffer",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      jsonControlsSource,
      "blocked save preserves disk",
    );
    subscription.dispose();
  });
}

async function assertBackupRestoreRevertAndSaveAs(extensionUri: vscode.Uri): Promise<void> {
  await withTestEnvironment(extensionUri, "yaml", yamlSource, async (environment) => {
    const { customDocument, directory, provider, sourceDocument, uri, cancellation } = environment;
    await sendCaptionEdit(environment, 1, "First backup");
    await sendCaptionEdit(environment, 2, "Backed up");
    const backedUpText = replaceCaption(yamlSource, "Backed up");
    const backupUri = vscode.Uri.file(path.join(directory, "Main.backup"));
    const backup = await provider.backupCustomDocument(
      customDocument,
      { destination: backupUri },
      cancellation.token,
    );
    assert.strictEqual(
      await readFile(backupUri.fsPath, "utf8").then((text) => text.length > 0),
      true,
    );

    provider.dispose();
    const restoredClient = createTestClient();
    const restoredProvider = new UserFormEditorProvider(
      extensionUri,
      restoredClient.client,
      undefined,
      undefined,
      environment.checkpointRoot,
    );
    const restoredPanel = createFakePanel();
    const restoredEvents: vscode.CustomDocumentEditEvent<UserFormCustomDocument>[] = [];
    const restoredSubscription = restoredProvider.onDidChangeCustomDocument((event) =>
      restoredEvents.push(event),
    );
    try {
      const restored = await restoredProvider.openCustomDocument(
        designerResourceUri(uri),
        { backupId: backup.id, untitledDocumentData: undefined },
        cancellation.token,
      );
      assert.strictEqual(
        restored.document.getText(),
        backedUpText,
        "backup restore recovers the unsaved draft",
      );
      await restoredProvider.resolveCustomEditor(restored, restoredPanel.panel, cancellation.token);
      restoredPanel.receive({ type: "ready" });
      await waitForDocument(restoredPanel, restored.document, "the restored draft preview");
      await waitFor(
        () => restoredEvents.length === 1,
        "the restored draft history to publish one aggregate native edit callback",
      );
      assert.ok(restoredEvents.every((event) => event.document === restored));
      await restoredEvents[0].undo();
      await waitFor(
        () => restored.document.getText() === yamlSource,
        "Undo of the restored aggregate custom-document edit",
      );
      assert.strictEqual(
        sourceDocument.getText(),
        yamlSource,
        "restored Undo leaves the source buffer untouched",
      );
      assert.strictEqual(
        await readFile(uri.fsPath, "utf8"),
        yamlSource,
        "restored Undo leaves disk untouched",
      );
      await restoredEvents[0].redo();
      await waitFor(
        () => restored.document.getText() === backedUpText,
        "Redo of the restored aggregate custom-document edit",
      );
      assert.strictEqual(
        sourceDocument.getText(),
        yamlSource,
        "restored Redo leaves the source buffer untouched",
      );
      assert.strictEqual(
        await readFile(uri.fsPath, "utf8"),
        yamlSource,
        "restored Redo leaves disk untouched",
      );

      await restoredProvider.revertCustomDocument(restored, cancellation.token);
      assert.strictEqual(
        restored.document.getText(),
        yamlSource,
        "revert returns to the file on disk",
      );

      const copiedText = replaceCaption(yamlSource, "Saved As");
      const range = captionRange(restored.document, "Original");
      restored.apply([{ range, newText: JSON.stringify("Saved As") }]);
      const destination = vscode.Uri.file(path.join(directory, "Copy.yaml"));
      await restoredProvider.saveCustomDocumentAs(restored, destination, cancellation.token);
      assert.strictEqual(await readFile(destination.fsPath, "utf8"), copiedText);
      assert.strictEqual(await readFile(uri.fsPath, "utf8"), yamlSource);
      assert.strictEqual(sourceDocument.getText(), yamlSource);
    } finally {
      restoredSubscription.dispose();
      restoredPanel.dispose();
      restoredProvider.dispose();
      restoredClient.dispose();
      backup.delete();
    }
  });
}

async function assertRegisteredProviderNativeSave(extensionUri: vscode.Uri): Promise<void> {
  const fixture = await createWorkspaceFixture("yaml");
  const { directory, uri } = fixture;
  const designerUri = designerResourceUri(uri);
  const viewType = "xlflow.userFormDraftProviderTest";
  const cancellation = new vscode.CancellationTokenSource();
  const client = createTestClient();
  const checkpointRoot = vscode.Uri.file(path.join(directory, "checkpoints"));
  const confirmedResolutions: ("source" | "designer")[] = [];
  const provider = new UserFormEditorProvider(
    extensionUri,
    client.client,
    undefined,
    async (choice) => {
      confirmedResolutions.push(choice);
      return true;
    },
    checkpointRoot,
  );
  const hostMessages: HostMessage[] = [];
  const events: vscode.CustomDocumentEditEvent<UserFormCustomDocument>[] = [];
  let nativePanel: vscode.WebviewPanel | undefined;
  let transportPanel: vscode.WebviewPanel | undefined;
  let editSubscription: vscode.Disposable | undefined;
  let registration: vscode.Disposable | undefined;
  let saveCommandRegistration: vscode.Disposable | undefined;
  let customDocument: UserFormCustomDocument | undefined;
  let placeholderDisposed = false;
  const resolvedEditorUris: string[] = [];
  const placeholderSubscriptions: vscode.Disposable[] = [];

  try {
    await writeFile(uri.fsPath, yamlSource, "utf8");
    const sourceDocument = await vscode.workspace.openTextDocument(uri);
    customDocument = await provider.openCustomDocument(
      designerUri,
      { backupId: undefined, untitledDocumentData: undefined },
      cancellation.token,
    );
    assert.strictEqual(customDocument.uri.toString(), designerUri.toString());
    assert.strictEqual(customDocument.sourceUri.toString(), uri.toString());
    assert.strictEqual(customDocument.document.uri.toString(), uri.toString());
    editSubscription = provider.onDidChangeCustomDocument((event) => events.push(event));

    const resolve = provider.resolveCustomEditor.bind(provider);
    provider.resolveCustomEditor = async (document, panel, token) => {
      resolvedEditorUris.push(document.uri.toString());
      if (document.uri.scheme === "file") {
        placeholderSubscriptions.push(panel.onDidDispose(() => (placeholderDisposed = true)));
        await resolve(document, panel, token);
        return;
      }
      assert.strictEqual(document.uri.toString(), designerUri.toString());
      if (customDocument && !customDocument.closed)
        assert.strictEqual(
          document,
          customDocument,
          "the initial redirect reuses the preopened draft",
        );
      else
        assert.notStrictEqual(document, customDocument, "reopen creates a fresh custom document");
      customDocument = document;
      nativePanel = panel;
      transportPanel = instrumentPanel(panel, hostMessages);
      await resolve(document, transportPanel, token);
    };
    const fixtureExtension = vscode.extensions.getExtension<DraftProviderTestExtension>(
      "xlflowtest.userform-draft-fixture",
    );
    assert.ok(
      fixtureExtension,
      "the test-only custom editor fixture extension should be installed",
    );
    const fixtureApi = await fixtureExtension.activate();
    registration = fixtureApi.register(provider);
    const testSaveCommand = "xlflow.userFormDraftProviderTest.save";
    saveCommandRegistration = fixtureApi.registerSaveCommand(provider, testSaveCommand);
    await vscode.commands.executeCommand("vscode.openWith", uri, viewType);

    try {
      await waitFor(
        () =>
          findCustomTab(viewType, designerUri) !== undefined &&
          transportPanel !== undefined &&
          placeholderDisposed,
        "opening the source file to redirect to the registered virtual Custom Editor",
      );
      assert.deepStrictEqual(resolvedEditorUris, [uri.toString(), designerUri.toString()]);
    } catch (error) {
      const tabs = vscode.window.tabGroups.all.map((group, index) => ({
        group: index,
        active: group.isActive,
        tabs: group.tabs.map((tab) => {
          const input = tab.input;
          return {
            label: tab.label,
            active: tab.isActive,
            dirty: tab.isDirty,
            input:
              input instanceof vscode.TabInputCustom
                ? "custom:" + input.viewType + ":" + input.uri.toString()
                : input instanceof vscode.TabInputText
                  ? "text:" + input.uri.toString()
                  : "other",
          };
        }),
      }));
      throw new Error(
        String(error) +
          "; extension order=" +
          vscode.extensions.all.map((extension) => extension.id).join(",") +
          "; tabs=" +
          JSON.stringify(tabs) +
          "; nativePanel=" +
          (nativePanel !== undefined) +
          "; transportPanel=" +
          (transportPanel !== undefined) +
          "; placeholderDisposed=" +
          placeholderDisposed +
          "; resolvedEditorUris=" +
          JSON.stringify(resolvedEditorUris),
      );
    }
    await waitFor(
      () => hostMessages.some((message) => message.type === "document"),
      "the registered Webview transport to become ready",
    );

    const firstExpected = replaceCaption(yamlSource, "Native Save");
    const propertyMouse = await NativeDesignerMouse.connect();
    try {
      if (!(await propertyMouse.isVisible(".properties-pane")))
        await propertyMouse.clickSelector('.pane-toggle[data-pane="properties"]');
      await propertyMouse.clickSelector("#designer-tab-properties");
      await propertyMouse.replaceText('[data-property-field="caption"] input', "Native Save");
      await propertyMouse.clickSelector(
        (await propertyMouse.isVisible(".properties-pane .pane-close"))
          ? ".properties-pane .pane-close"
          : ".form-title",
      );
    } finally {
      propertyMouse.close();
    }
    await waitFor(
      () => events.length === 1,
      "the registered provider to emit its native edit event",
    );
    await waitFor(
      () => findCustomTab(viewType, designerUri)?.isDirty === true,
      "the native custom-tab dirty state after the property edit",
    );
    const originalUndo = customDocument.undo.bind(customDocument);
    const originalRedo = customDocument.redo.bind(customDocument);
    let undoCallbacks = 0;
    let redoCallbacks = 0;
    customDocument.undo = (frame) => {
      undoCallbacks++;
      originalUndo(frame);
    };
    customDocument.redo = (frame) => {
      redoCallbacks++;
      originalRedo(frame);
    };
    assert.strictEqual(customDocument.document.getText(), firstExpected);
    assert.strictEqual(
      sourceDocument.getText(),
      yamlSource,
      "editing through the registered Webview leaves the source buffer untouched",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      yamlSource,
      "editing through the registered Webview leaves disk untouched",
    );
    assert.strictEqual(
      findCustomTab(viewType, designerUri)?.isDirty,
      true,
      "native edit events dirty the custom tab",
    );

    await waitFor(
      () =>
        hostMessages.some(
          (message) =>
            message.type === "document" && message.version === customDocument!.document.version,
        ),
      "latest Designer preview before toolbar Undo",
    );
    const nativeMouse = await NativeDesignerMouse.connect();
    try {
      await nativeMouse.clickSelector(".history-undo");
    } finally {
      nativeMouse.close();
    }
    await waitFor(
      () => customDocument!.document.getText() === yamlSource,
      "the registered Custom Editor native Undo command",
    );
    assert.strictEqual(
      sourceDocument.getText(),
      yamlSource,
      "native Undo leaves the source buffer untouched",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      yamlSource,
      "native Undo leaves disk untouched",
    );
    assert.strictEqual(
      findCustomTab(viewType, designerUri)?.isDirty,
      false,
      "native Undo returns the custom tab to clean",
    );

    await waitFor(
      () => hostMessages.some((message) => message.type === "historyState" && message.canRedo),
      "Redo availability after native Undo",
    );
    await waitFor(
      () =>
        hostMessages.some(
          (message) =>
            message.type === "document" && message.version === customDocument!.document.version,
        ),
      "latest Designer preview before toolbar Redo",
    );
    const redoMouse = await NativeDesignerMouse.connect();
    try {
      await redoMouse.clickSelector(".history-redo");
    } finally {
      redoMouse.close();
    }
    await waitFor(
      () => customDocument!.document.getText() === firstExpected,
      "the registered Custom Editor native Redo command",
    );
    assert.strictEqual(
      sourceDocument.getText(),
      yamlSource,
      "native Redo leaves the source buffer untouched",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      yamlSource,
      "native Redo leaves disk untouched",
    );
    assert.strictEqual(
      findCustomTab(viewType, designerUri)?.isDirty,
      true,
      "native Redo restores the dirty draft",
    );

    // Keep the Designer visible but inactive. Hidden Webviews can suspend
    // message delivery, and host replies intentionally wait for visibility.
    await vscode.window.showTextDocument(sourceDocument, { viewColumn: vscode.ViewColumn.Beside });
    await waitFor(
      () => nativePanel!.visible && !nativePanel!.active,
      "source editor active beside Designer",
    );
    await transportPanel!.webview.postMessage({
      type: "xlflow-test-history",
      action: "undo",
      requestId: 699,
      version: customDocument.document.version,
    });
    await waitFor(
      () =>
        hostMessages.some(
          (message) =>
            message.type === "editResult" &&
            message.requestId === 699 &&
            message.error?.code === "historyUnavailable",
        ),
      "inactive Designer rejects history request",
    );
    assert.strictEqual(customDocument.document.getText(), firstExpected);
    assert.strictEqual(sourceDocument.getText(), yamlSource);
    const sourceTab = vscode.window.tabGroups.all
      .flatMap((group) => group.tabs)
      .find(
        (tab) =>
          tab.input instanceof vscode.TabInputText && tab.input.uri.toString() === uri.toString(),
      );
    if (sourceTab) await vscode.window.tabGroups.close(sourceTab, true);
    nativePanel!.reveal(nativePanel!.viewColumn);
    await waitFor(() => nativePanel!.active, "restore Designer focus before toolbar Undo");

    await vscode.commands.executeCommand("workbench.action.files.save");
    await waitForAsync(
      async () =>
        sourceDocument.getText() === yamlSource &&
        (await readFile(uri.fsPath, "utf8")) === yamlSource &&
        customDocument!.state.dirty &&
        findCustomTab(viewType, designerUri)?.isDirty === false,
      "generic native Save to checkpoint the draft and clear only the custom-tab dirty state",
    );
    assert.strictEqual(
      (await new DraftCheckpointStore(checkpointRoot).read(uri))?.state.text,
      firstExpected,
      "generic native Save persists the pending draft checkpoint",
    );
    assert.strictEqual(
      customDocument.state.dirty,
      true,
      "native checkpoint acknowledgement does not clear the Designer draft badge",
    );

    nativePanel!.reveal(nativePanel!.viewColumn);
    await waitFor(() => nativePanel!.active, "restore Designer focus before explicit publication");
    await vscode.commands.executeCommand(testSaveCommand);
    await waitForAsync(
      async () =>
        sourceDocument.getText() === firstExpected &&
        (await readFile(uri.fsPath, "utf8")) === firstExpected &&
        !customDocument!.state.dirty &&
        findCustomTab(viewType, designerUri)?.isDirty === false,
      "the explicit Designer Save command to publish while the native tab is already clean",
    );
    const publishedCheckpoint = await new DraftCheckpointStore(checkpointRoot).read(uri);
    assert.strictEqual(publishedCheckpoint?.state.text, firstExpected);
    assert.strictEqual(publishedCheckpoint?.state.baseText, firstExpected);

    undoCallbacks = 0;
    redoCallbacks = 0;
    const externalExpected = replaceCaption(firstExpected, "External");
    const externalEdit = new vscode.WorkspaceEdit();
    externalEdit.replace(
      sourceDocument.uri,
      captionRange(sourceDocument, "Native Save"),
      JSON.stringify("External"),
    );
    assert.ok(
      await vscode.workspace.applyEdit(externalEdit),
      "the external caption edit should apply",
    );
    await waitFor(
      () => !customDocument!.reconciling && customDocument!.document.getText() === externalExpected,
      "the external caption change to rebase the saved draft",
    );
    assert.ok(await sourceDocument.save(), "the external caption source should be saved");
    assert.strictEqual(
      sourceDocument.isDirty,
      false,
      "the external source starts native Undo clean",
    );
    const diskBeforeStaleUndo = await readFile(uri.fsPath, "utf8");
    assert.strictEqual(diskBeforeStaleUndo, externalExpected);
    nativePanel!.reveal(nativePanel!.viewColumn);
    await waitFor(
      () => findCustomTab(viewType, designerUri)?.isActive === true,
      "the registered custom editor to regain focus before native Undo",
    );
    assert.strictEqual(findCustomTab(viewType, designerUri)?.isActive, true);

    await vscode.commands.executeCommand("undo");
    try {
      await waitFor(
        () => undoCallbacks === 1,
        "native Undo to dispatch the trimmed history callback",
      );
    } catch (error) {
      const tab = findCustomTab(viewType, designerUri);
      const observation = {
        sourceText: sourceDocument.getText(),
        sourceDirty: sourceDocument.isDirty,
        diskText: await readFile(uri.fsPath, "utf8"),
        customText: customDocument.document.getText(),
        customDirty: tab?.isDirty,
        customActive: tab?.isActive,
        nativePanelVisible: nativePanel!.visible,
      };
      assert.fail(`${String(error)}; post-Undo state=${JSON.stringify(observation)}`);
    }
    const tabAfterStaleUndo = findCustomTab(viewType, designerUri);
    assert.strictEqual(
      sourceDocument.getText(),
      externalExpected,
      "native Undo leaves the source buffer unchanged",
    );
    assert.strictEqual(sourceDocument.isDirty, false, "native Undo leaves the source buffer clean");
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      diskBeforeStaleUndo,
      "native Undo leaves disk unchanged",
    );
    assert.strictEqual(
      customDocument.document.getText(),
      externalExpected,
      "stale native Undo leaves the draft baseline unchanged",
    );
    assert.strictEqual(
      customDocument.state.dirty,
      false,
      "a neutral historical callback leaves the draft model clean",
    );
    assert.strictEqual(
      tabAfterStaleUndo?.isActive,
      true,
      "native Undo remains focused on the Designer tab",
    );
    assert.strictEqual(
      tabAfterStaleUndo?.isDirty,
      true,
      "VS Code conservatively marks the native checkpoint dirty after the neutral historical Undo",
    );
    await vscode.commands.executeCommand("redo");
    await waitFor(
      () => redoCallbacks === 1 && findCustomTab(viewType, designerUri)?.isDirty === false,
      "native Redo to dispatch the trimmed history callback and return to the saved checkpoint",
    );
    assert.strictEqual(customDocument.document.getText(), externalExpected);
    assert.strictEqual(
      customDocument.state.dirty,
      false,
      "native Redo preserves the clean draft model",
    );
    assert.strictEqual(
      findCustomTab(viewType, designerUri)?.isDirty,
      false,
      "native Redo restores the saved checkpoint marker",
    );
    assert.strictEqual(
      sourceDocument.getText(),
      externalExpected,
      "stale callbacks preserve external buffer text",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      diskBeforeStaleUndo,
      "stale callbacks preserve disk",
    );
    assert.strictEqual(
      sourceDocument.isDirty,
      false,
      "stale callbacks preserve the clean external source",
    );

    const secondExpected = replaceCaption(externalExpected, "Webview Save");
    await postTestEdit(transportPanel!, customDocument, 2, "Webview Save");
    await waitFor(() => events.length === 2, "the second registered Webview edit");
    assert.strictEqual(findCustomTab(viewType, designerUri)?.isDirty, true);
    await transportPanel!.webview.postMessage({ type: "xlflow-test-save" });
    await waitForAsync(
      async () =>
        sourceDocument.getText() === secondExpected &&
        (await readFile(uri.fsPath, "utf8")) === secondExpected &&
        findCustomTab(viewType, designerUri)?.isDirty === false,
      "the Webview Save button to invoke native save and clear the custom-tab dirty state",
    );

    const diskBeforeConflict = await readFile(uri.fsPath, "utf8");
    const localConflict = replaceCaption(secondExpected, "Designer Conflict");
    await postTestEdit(transportPanel!, customDocument, 3, "Designer Conflict");
    await waitFor(() => events.length === 3, "the registered Webview conflict draft edit");
    assert.strictEqual(customDocument.state.dirty, true, "a pending conflict draft is dirty");
    assert.strictEqual(findCustomTab(viewType, designerUri)?.isDirty, true);

    const externalConflict = replaceCaption(secondExpected, "External Conflict");
    const conflictEdit = new vscode.WorkspaceEdit();
    conflictEdit.replace(
      sourceDocument.uri,
      captionRange(sourceDocument, "Webview Save"),
      JSON.stringify("External Conflict"),
    );
    assert.ok(
      await vscode.workspace.applyEdit(conflictEdit),
      "the external conflict edit should apply",
    );
    await waitFor(
      () => customDocument!.state.conflicted && !customDocument!.reconciling,
      "the registered editor to finish reconciling the source conflict",
    );
    assert.strictEqual(
      customDocument.state.text,
      localConflict,
      "the local Designer conflict draft is retained",
    );
    assert.strictEqual(sourceDocument.getText(), externalConflict);
    await provider.saveCustomDocument(customDocument, cancellation.token);
    const conflictedCheckpoint = await new DraftCheckpointStore(checkpointRoot).read(uri);
    assert.strictEqual(conflictedCheckpoint?.state.text, localConflict);
    await assert.rejects(
      provider.publishCustomDocument(customDocument, cancellation.token),
      /conflict/i,
      "publishing is blocked while the source conflict is unresolved",
    );
    assert.strictEqual(
      sourceDocument.getText(),
      externalConflict,
      "blocked save preserves source buffer text",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      diskBeforeConflict,
      "blocked save preserves disk",
    );
    assert.strictEqual(customDocument.state.dirty, true);

    await transportPanel!.webview.postMessage({ type: "xlflow-test-resolve", choice: "source" });
    await waitFor(
      () => events.length === 4 && !customDocument!.state.conflicted,
      "the registered Webview source-conflict resolution event",
    );
    assert.deepStrictEqual(confirmedResolutions, ["source"]);
    assert.strictEqual(customDocument.document.getText(), externalConflict);
    assert.strictEqual(customDocument.state.text, externalConflict);
    assert.strictEqual(
      customDocument.state.dirty,
      false,
      "the selected source becomes the clean baseline",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      diskBeforeConflict,
      "resolution does not save source text",
    );
    assert.strictEqual(
      findCustomTab(viewType, designerUri)?.isActive,
      true,
      "native Undo targets the registered custom editor",
    );

    await vscode.commands.executeCommand("undo");
    await waitFor(
      () => customDocument!.document.getText() === localConflict && customDocument!.state.dirty,
      "native Undo of conflict resolution to restore the prior Designer draft",
    );
    assert.strictEqual(
      findCustomTab(viewType, designerUri)?.isDirty,
      true,
      "the restored draft is dirty",
    );
    assert.strictEqual(
      sourceDocument.getText(),
      externalConflict,
      "Undo does not replace the source buffer",
    );
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      diskBeforeConflict,
      "Undo leaves disk unchanged",
    );

    await vscode.commands.executeCommand("undo");
    await waitFor(
      () => customDocument!.document.getText() === externalConflict && !customDocument!.state.dirty,
      "native Undo of the original edit to return to the latest source baseline",
    );
    assert.strictEqual(
      findCustomTab(viewType, designerUri)?.isDirty,
      false,
      "the latest source baseline is clean",
    );
    assert.strictEqual(sourceDocument.getText(), externalConflict);
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      diskBeforeConflict,
      "Undo never saves the external source",
    );
    assert.ok(
      await sourceDocument.save(),
      "save the external buffer after assertions to clean up the fixture",
    );
    const diskBeforeAutoSave = await readFile(uri.fsPath, "utf8");

    await waitFor(
      () =>
        hostMessages.some(
          (message) =>
            message.type === "document" && message.version === customDocument!.document.version,
        ),
      "the Webview to receive the latest native conflict-Undo baseline",
    );
    const workspaceFolder = vscode.workspace.getWorkspaceFolder(uri);
    assert.ok(workspaceFolder, "the autosave fixture must belong to the temporary test workspace");
    const filesConfiguration = vscode.workspace.getConfiguration("files", workspaceFolder.uri);
    const previousAutoSave = filesConfiguration.inspect<string>("autoSave")?.workspaceValue;
    const previousAutoSaveDelay =
      filesConfiguration.inspect<number>("autoSaveDelay")?.workspaceValue;
    const configurationTarget = vscode.ConfigurationTarget.Workspace;
    let autoSavedExpected = "";
    let sourceAfterAutoSave = externalConflict;
    let diskAfterAutoSave = diskBeforeAutoSave;
    try {
      await filesConfiguration.update("autoSaveDelay", 200, configurationTarget);
      await filesConfiguration.update("autoSave", "afterDelay", configurationTarget);
      autoSavedExpected = replaceCaption(externalConflict, "Auto Saved");
      await postTestEdit(transportPanel!, customDocument, 4, "Auto Saved");
      await waitFor(() => events.length === 5, "the final autosave draft edit event");
      await waitForAsync(
        async () =>
          sourceDocument.getText() === externalConflict &&
          (await readFile(uri.fsPath, "utf8")) === diskBeforeAutoSave &&
          !sourceDocument.isDirty &&
          customDocument!.state.dirty &&
          findCustomTab(viewType, designerUri)?.isDirty === false &&
          // The native dirty event can lag the draft edit. A clean tab alone
          // does not establish that Auto Save durably wrote this checkpoint.
          (await new DraftCheckpointStore(checkpointRoot).read(uri))?.state.text ===
            autoSavedExpected,
        "AutoSave afterDelay to checkpoint without publishing or clearing the Designer draft badge",
      );
      assert.strictEqual(
        (await new DraftCheckpointStore(checkpointRoot).read(uri))?.state.text,
        autoSavedExpected,
        "AutoSave afterDelay persists the latest draft checkpoint",
      );
      assert.ok(
        hostMessages.some((message) => message.type === "draftState" && message.dirty),
        "source-pending state remains visible after Auto Save",
      );

      await vscode.commands.executeCommand("undo");
      await waitFor(
        () =>
          customDocument!.state.text === externalConflict &&
          !customDocument!.state.dirty &&
          findCustomTab(viewType, designerUri)?.isDirty === false,
        "Undo followed by Auto Save to checkpoint the source baseline",
      );
      await vscode.commands.executeCommand("redo");
      await waitFor(
        () =>
          customDocument!.state.text === autoSavedExpected &&
          customDocument!.state.dirty &&
          findCustomTab(viewType, designerUri)?.isDirty === false,
        "Redo followed by Auto Save to preserve the pending draft",
      );
      assert.strictEqual(await readFile(uri.fsPath, "utf8"), diskBeforeAutoSave);

      const sourcePrefix = "# direct source Auto Save remains enabled\n";
      const sourceEdit = new vscode.WorkspaceEdit();
      sourceEdit.insert(uri, new vscode.Position(0, 0), sourcePrefix);
      assert.ok(await vscode.workspace.applyEdit(sourceEdit));
      sourceAfterAutoSave = sourcePrefix + externalConflict;
      autoSavedExpected = sourcePrefix + autoSavedExpected;
      await waitForAsync(
        async () =>
          !customDocument!.reconciling &&
          customDocument!.state.text === autoSavedExpected &&
          sourceDocument.getText() === sourceAfterAutoSave &&
          (await readFile(uri.fsPath, "utf8")) === sourceAfterAutoSave &&
          !sourceDocument.isDirty,
        "normal source Auto Save to publish only the direct YAML edit and update the Designer",
      );
      assert.ok(
        customDocument.state.dirty,
        "direct source Auto Save does not apply the Designer caption",
      );

      await vscode.commands.executeCommand(testSaveCommand);
      await waitForAsync(
        async () =>
          !customDocument!.state.dirty &&
          sourceDocument.getText() === autoSavedExpected &&
          (await readFile(uri.fsPath, "utf8")) === autoSavedExpected,
        "explicit Designer save while Auto Save is enabled and the native tab is clean",
      );
      sourceAfterAutoSave = autoSavedExpected;
      diskAfterAutoSave = autoSavedExpected;
      await waitFor(
        () =>
          hostMessages.some(
            (message) =>
              message.type === "document" && message.version === customDocument!.document.version,
          ),
        "the published baseline to reach the Webview",
      );
      autoSavedExpected = replaceCaption(autoSavedExpected, "Pending after Auto Save");
      await postTestEdit(transportPanel!, customDocument, 5, "Pending after Auto Save");
      await waitFor(
        () =>
          events.length === 6 &&
          customDocument!.state.dirty &&
          findCustomTab(viewType, designerUri)?.isDirty === false,
        "a fresh pending draft to be internally checkpointed before close",
      );
      assert.strictEqual(await readFile(uri.fsPath, "utf8"), diskAfterAutoSave);
    } finally {
      try {
        await filesConfiguration.update("autoSave", previousAutoSave, configurationTarget);
      } finally {
        await filesConfiguration.update(
          "autoSaveDelay",
          previousAutoSaveDelay,
          configurationTarget,
        );
      }
    }

    const closedDraft = customDocument!;
    const closedTab = findCustomTab(viewType, designerUri);
    assert.ok(closedTab, "the Designer tab should be open before close/reopen recovery");
    await vscode.window.tabGroups.close(closedTab, true);
    await waitFor(() => closedDraft.closed, "the native Designer tab to close its custom document");
    hostMessages.length = 0;
    await vscode.commands.executeCommand("vscode.openWith", uri, viewType);
    await waitFor(
      () =>
        customDocument !== closedDraft &&
        findCustomTab(viewType, designerUri) !== undefined &&
        hostMessages.some((message) => message.type === "document"),
      "reopening the Designer tab to restore its pending draft checkpoint",
    );
    assert.strictEqual(customDocument!.state.text, autoSavedExpected);
    assert.strictEqual(customDocument!.state.dirty, true);
    assert.strictEqual(customDocument!.state.conflicted, false);
    assert.strictEqual(sourceDocument.getText(), sourceAfterAutoSave);
    assert.strictEqual(await readFile(uri.fsPath, "utf8"), diskAfterAutoSave);
    nativePanel!.reveal(nativePanel!.viewColumn);
    await vscode.commands.executeCommand("workbench.action.files.save");
    await waitFor(
      () => findCustomTab(viewType, designerUri)?.isDirty === false,
      "checkpointing the recovered draft before closing the test tab",
    );
    assert.ok(customDocument!.state.dirty, "checkpointing recovery does not apply the source");
  } finally {
    editSubscription?.dispose();
    placeholderSubscriptions.forEach((subscription) => subscription.dispose());
    if (nativePanel) {
      for (const group of vscode.window.tabGroups.all) {
        for (const tab of group.tabs.slice()) {
          const input = tab.input;
          if (
            input instanceof vscode.TabInputCustom &&
            input.uri.toString() === designerUri.toString() &&
            input.viewType === viewType
          ) {
            nativePanel!.reveal(nativePanel!.viewColumn);
            await vscode.commands.executeCommand("workbench.action.files.save");
            await vscode.window.tabGroups.close(tab, true);
          }
        }
      }
    }
    registration?.dispose();
    saveCommandRegistration?.dispose();
    provider.dispose();
    client.dispose();
    cancellation.dispose();
    await rm(uri.fsPath, { force: true, maxRetries: 5, retryDelay: 100 });
    await rm(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

async function assertCheckpointRecovery(
  extensionUri: vscode.Uri,
  format: "yaml" | "json",
  initial: string,
): Promise<void> {
  const { directory, uri } = await createWorkspaceFixture(format);
  const checkpointRoot = vscode.Uri.file(path.join(directory, "checkpoints"));
  const store = new DraftCheckpointStore(checkpointRoot);
  const cancellation = new vscode.CancellationTokenSource();
  const clients: TestClient[] = [];
  const providers: UserFormEditorProvider[] = [];
  const disposedProviders = new Set<UserFormEditorProvider>();
  let document: UserFormCustomDocument | undefined;

  const createProvider = () => {
    const client = createTestClient();
    clients.push(client);
    const provider = new UserFormEditorProvider(
      extensionUri,
      client.client,
      undefined,
      undefined,
      checkpointRoot,
    );
    providers.push(provider);
    return provider;
  };
  const disposeProvider = (provider: UserFormEditorProvider) => {
    if (disposedProviders.has(provider)) return;
    disposedProviders.add(provider);
    provider.dispose();
  };
  const open = (provider: UserFormEditorProvider) =>
    provider.openCustomDocument(
      designerResourceUri(uri),
      { backupId: undefined, untitledDocumentData: undefined },
      cancellation.token,
    );

  try {
    await writeFile(uri.fsPath, initial, "utf8");
    const sourceDocument = await vscode.workspace.openTextDocument(uri);
    let provider = createProvider();
    document = await open(provider);
    const recoveredText = replaceCaption(initial, "Checkpointed draft");
    document.apply(
      [
        {
          range: captionRange(document.document, "Original"),
          newText: JSON.stringify("Checkpointed draft"),
        },
      ],
      "Checkpoint recovery fixture",
    );

    // No explicit checkpoint wait: open must read behind the queued write
    // emitted by the last edit, even when the panel closes immediately.
    document.dispose();
    document = await open(provider);
    assert.strictEqual(
      document.state.text,
      recoveredText,
      "immediate reopen waits for draft persistence",
    );
    assert.ok(document.state.dirty);

    let conflictSource: string | undefined;
    if (format === "json") {
      const parsed = JSON.parse(initial) as { controls: SpecControl[] };
      conflictSource =
        JSON.stringify({ ...parsed, controls: [...parsed.controls].reverse() }, null, 2) + "\n";
      assert.ok(
        await vscode.workspace.applyEdit(replaceWholeDocument(sourceDocument, conflictSource)),
        "the external canonical reorder should apply",
      );
      await waitFor(
        () => document!.state.conflicted && !document!.reconciling,
        "the canonical source identity conflict before checkpointing",
      );
      assert.strictEqual(document.state.text, recoveredText);
    }

    await provider.saveCustomDocument(document, cancellation.token);
    assert.strictEqual(document.state.dirty, true, "a checkpoint keeps the draft badge dirty");
    assert.strictEqual(sourceDocument.getText(), conflictSource ?? initial);
    assert.strictEqual(
      await readFile(uri.fsPath, "utf8"),
      initial,
      "checkpointing does not publish the source buffer",
    );
    const firstCheckpoint = await store.read(uri);
    assert.strictEqual(firstCheckpoint?.state.text, recoveredText);
    if (format === "json") {
      assert.strictEqual(firstCheckpoint?.state.conflicted, true);
      assert.strictEqual(firstCheckpoint?.state.pendingSourceText, conflictSource);
    }

    document.dispose();
    document = await open(provider);
    assert.strictEqual(document.state.text, recoveredText, "normal reopen recovers the draft");
    assert.strictEqual(document.state.dirty, true);
    assert.strictEqual(document.state.conflicted, format === "json");

    document.dispose();
    document = undefined;
    disposeProvider(provider);
    provider = createProvider();
    document = await open(provider);
    assert.strictEqual(document.state.text, recoveredText, "provider restart recovers the draft");
    assert.strictEqual(document.state.dirty, true);
    assert.strictEqual(
      document.state.conflicted,
      format === "json",
      "provider restart retains the canonical conflict",
    );
    if (format === "json") {
      assert.strictEqual(document.document.getText(), conflictSource);
      assert.strictEqual(document.state.serialize().pendingSourceText, conflictSource);
    }

    await provider.revertCustomDocument(document, cancellation.token);
    const revertedText = conflictSource ?? initial;
    assert.strictEqual(document.state.text, revertedText);
    assert.strictEqual(document.state.dirty, false);
    assert.strictEqual(document.state.conflicted, false);
    document.dispose();
    document = undefined;
    disposeProvider(provider);

    const afterRevertProvider = createProvider();
    document = await open(afterRevertProvider);
    assert.strictEqual(document.state.text, revertedText);
    assert.strictEqual(document.state.dirty, false, "revert clears pending draft recovery");
    assert.strictEqual(document.state.conflicted, false);
  } finally {
    document?.dispose();
    providers.forEach(disposeProvider);
    clients.forEach((client) => client.dispose());
    cancellation.dispose();
    await rm(uri.fsPath, { force: true, maxRetries: 5, retryDelay: 100 });
    await rm(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

async function assertMalformedCheckpointIsRetained(extensionUri: vscode.Uri): Promise<void> {
  const { directory, uri } = await createWorkspaceFixture("yaml");
  const checkpointRoot = vscode.Uri.file(path.join(directory, "checkpoints"));
  const store = new DraftCheckpointStore(checkpointRoot);
  const cancellation = new vscode.CancellationTokenSource();
  const client = createTestClient();
  const provider = new UserFormEditorProvider(
    extensionUri,
    client.client,
    undefined,
    undefined,
    checkpointRoot,
  );
  const malformed = "{ malformed checkpoint";
  try {
    await writeFile(uri.fsPath, yamlSource, "utf8");
    const location = store.location(uri);
    await mkdir(path.dirname(location.fsPath), { recursive: true });
    await writeFile(location.fsPath, malformed, "utf8");
    await assert.rejects(
      provider.openCustomDocument(
        designerResourceUri(uri),
        { backupId: undefined, untitledDocumentData: undefined },
        cancellation.token,
      ),
      (error: unknown) => error instanceof Error,
      "opening a malformed checkpoint must surface the recovery error",
    );
    assert.strictEqual(
      await readFile(location.fsPath, "utf8"),
      malformed,
      "a malformed recovery record is retained for diagnosis",
    );
  } finally {
    provider.dispose();
    client.dispose();
    cancellation.dispose();
    await rm(uri.fsPath, { force: true, maxRetries: 5, retryDelay: 100 });
    await rm(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

async function assertPrepareBuildSources(extensionUri: vscode.Uri): Promise<void> {
  for (const approve of [false, true]) {
    let confirmations = 0;
    await withTestEnvironment(
      extensionUri,
      "yaml",
      yamlSource,
      async (environment) => {
        const { provider, customDocument, uri, directory } = environment;
        await sendCaptionEdit(environment, 1900, "Build draft");
        const codeUri = vscode.Uri.file(path.join(directory, "Main.bas"));
        const unrelatedUri = vscode.Uri.file(path.join(directory, "Unrelated.bas"));
        await writeFile(codeUri.fsPath, "Option Explicit\n");
        await writeFile(unrelatedUri.fsPath, "Option Explicit\n");
        const code = await vscode.workspace.openTextDocument(codeUri);
        const unrelated = await vscode.workspace.openTextDocument(unrelatedUri);
        const edit = new vscode.WorkspaceEdit();
        edit.insert(codeUri, code.positionAt(code.getText().length), "' build\n");
        edit.insert(
          unrelatedUri,
          unrelated.positionAt(unrelated.getText().length),
          "' unrelated\n",
        );
        assert.ok(await vscode.workspace.applyEdit(edit));
        assert.strictEqual(await provider.prepareBuildSources([uri, codeUri, uri]), approve);
        assert.strictEqual(confirmations, 1, "one confirmation for all targeted sources");
        assert.strictEqual(
          await readFile(uri.fsPath, "utf8"),
          approve ? replaceCaption(yamlSource, "Build draft") : yamlSource,
        );
        assert.strictEqual(customDocument.state.dirty, !approve);
        assert.strictEqual(code.isDirty, !approve);
        assert.strictEqual(unrelated.isDirty, true, "unrelated text is never saved");
        assert.strictEqual(await readFile(unrelatedUri.fsPath, "utf8"), "Option Explicit\n");
        await code.save();
        await unrelated.save();
        if (approve) {
          assert.strictEqual(await provider.prepareBuildSources([uri, codeUri]), true);
          assert.strictEqual(confirmations, 1, "clean sources need no confirmation");
        }
      },
      async () => {
        confirmations++;
        return approve;
      },
    );
  }
  let confirmations = 0;
  await withTestEnvironment(
    extensionUri,
    "yaml",
    yamlSource,
    async (environment) => {
      const { provider, customDocument, uri, cancellation, checkpointRoot } = environment;
      await sendCaptionEdit(environment, 1901, "Closed draft");
      await provider.saveCustomDocument(customDocument, cancellation.token);
      customDocument.dispose();
      assert.strictEqual(await readFile(uri.fsPath, "utf8"), yamlSource);
      assert.strictEqual(await provider.prepareBuildSources([uri]), true);
      assert.strictEqual(confirmations, 1, "closed recoverable draft is included in confirmation");
      assert.strictEqual(
        await readFile(uri.fsPath, "utf8"),
        replaceCaption(yamlSource, "Closed draft"),
      );
      const checkpoint = await new DraftCheckpointStore(checkpointRoot).read(uri);
      assert.strictEqual(
        checkpoint?.state.text,
        checkpoint?.state.baseText,
        "recovered draft is marked published",
      );
    },
    async () => {
      confirmations++;
      return true;
    },
  );
  await withTestEnvironment(
    extensionUri,
    "yaml",
    yamlSource,
    async (environment) => {
      const { provider, customDocument, uri } = environment;
      await sendCaptionEdit(environment, 1902, "Conflicting draft");
      customDocument.state.retainConflict(replaceCaption(yamlSource, "External"));
      assert.strictEqual(await provider.prepareBuildSources([uri]), false);
      assert.strictEqual(await readFile(uri.fsPath, "utf8"), yamlSource);
      assert.strictEqual(
        customDocument.state.dirty,
        true,
        "a failed publication retains the draft",
      );
    },
    async () => true,
  );
}

async function assertReplacementPictureSelection(extensionUri: vscode.Uri): Promise<void> {
  const initial = JSON.stringify({
    ...JSON.parse(jsonSource),
    controls: [{ id: "image", name: "Image1", type: "Image", progId: "Forms.Image.1" }],
  });
  const selection = (token: string): PictureSelection => ({
    token,
    name: `${token}.bmp`,
    dataURI: "data:image/bmp;base64,Qk0=",
    width: 2,
    height: 3,
  });
  const originalOpen = PictureImportSession.prototype.open;
  const originalCancel = PictureImportSession.prototype.cancel;
  const originalCommit = PictureImportSession.prototype.commit;
  const staged = new WeakMap<PictureImportSession, PictureSelection>();
  let pick: () => Promise<PictureSelection | undefined>;
  PictureImportSession.prototype.open = async function () {
    const result = await pick();
    if (result) staged.set(this, result);
    return result;
  };
  PictureImportSession.prototype.cancel = function () {
    staged.delete(this);
    originalCancel.call(this);
  };
  PictureImportSession.prototype.commit = async function (token) {
    assert.strictEqual(staged.get(this)?.token, token, "commit uses the active staged selection");
    return { token, picture: { path: `src/forms/assets/${token}.bmp` } };
  };
  try {
    for (const outcome of ["cancel", "error", "replace", "close"] as const) {
      await withTestEnvironment(
        extensionUri,
        "json",
        initial,
        async (environment) => {
          const { panel, customDocument, client } = environment;
          client.client.requestUserFormEdit = async (document, version, text, operations) => {
            const operation = operations[0];
            assert.strictEqual(operation.type, "setControlProperty");
            if (operation.type !== "setControlProperty") throw new Error("Expected picture edit");
            assert.strictEqual(operation.field, "picture");
            const source = JSON.parse(text);
            source.controls[0].picture = operation.value;
            return {
              version,
              edits: [
                {
                  range: new vscode.Range(document.positionAt(0), document.positionAt(text.length)),
                  newText: JSON.stringify(source),
                },
              ],
            };
          };
          const sendPick = (requestId: number) =>
            panel.receive({
              type: "picturePick",
              requestId,
              controlId: "image",
              version: customDocument.document.version,
            });
          const reply = (requestId: number) =>
            waitFor(
              () =>
                panel.messages.find(
                  (message): message is Extract<HostMessage, { type: "pictureSelection" }> =>
                    message.type === "pictureSelection" && message.requestId === requestId,
                ),
              `picture selection ${requestId}`,
            );
          pick = async () => selection("first");
          sendPick(2100);
          assert.strictEqual((await reply(2100)).selection?.token, "first");
          let finishPick: ((value: PictureSelection) => void) | undefined;
          let started = false;
          pick = async () => {
            started = true;
            if (outcome === "error") throw new Error("invalid replacement");
            if (outcome === "close")
              return new Promise((resolve) => {
                finishPick = resolve;
              });
            return outcome === "replace" ? selection("second") : undefined;
          };
          sendPick(2101);
          await waitFor(() => started, "replacement picker started");
          if (outcome === "close") {
            panel.receive({ type: "pictureCancel" });
            finishPick!(selection("late"));
          }
          const replacement = await reply(2101);
          assert.strictEqual(!!replacement.error, outcome === "error" || outcome === "close");
          if (outcome === "close") return;
          const token = outcome === "replace" ? "second" : "first";
          panel.receive({
            type: "pictureCommit",
            requestId: 2102,
            controlId: "image",
            version: customDocument.document.version,
            token,
          });
          assert.strictEqual(
            (await waitForReply(panel, 2102)).error,
            undefined,
            `${outcome} replacement retains a committable displayed selection`,
          );
          assert.deepStrictEqual(JSON.parse(customDocument.state.text).controls[0].picture, {
            path: `src/forms/assets/${token}.bmp`,
          });
          assert.strictEqual(
            environment.sourceDocument.getText(),
            initial,
            "picture commit remains a draft edit",
          );
        },
        undefined,
        true,
      );
    }
  } finally {
    PictureImportSession.prototype.open = originalOpen;
    PictureImportSession.prototype.cancel = originalCancel;
    PictureImportSession.prototype.commit = originalCommit;
  }
}

async function withTestEnvironment(
  extensionUri: vscode.Uri,
  format: "yaml" | "json",
  initial: string,
  assertions: (environment: TestEnvironment) => Promise<void>,
  confirmBuildSave?: () => PromiseLike<boolean>,
  typedProperties = false,
): Promise<void> {
  const { directory, uri } = await createWorkspaceFixture(format);
  const cancellation = new vscode.CancellationTokenSource();
  const client = createTestClient(typedProperties);
  const checkpointRoot = vscode.Uri.file(path.join(directory, "checkpoints"));
  const provider = new UserFormEditorProvider(
    extensionUri,
    client.client,
    undefined,
    undefined,
    checkpointRoot,
    confirmBuildSave,
  );
  const panel = createFakePanel();
  try {
    await writeFile(uri.fsPath, initial, "utf8");
    const sourceDocument = await vscode.workspace.openTextDocument(uri);
    const customDocument = await provider.openCustomDocument(
      designerResourceUri(uri),
      { backupId: undefined, untitledDocumentData: undefined },
      cancellation.token,
    );
    assert.strictEqual(customDocument.uri.toString(), designerResourceUri(uri).toString());
    assert.strictEqual(customDocument.sourceUri.toString(), uri.toString());
    assert.strictEqual(customDocument.document.uri.toString(), uri.toString());
    const environment: TestEnvironment = {
      directory,
      checkpointRoot,
      uri,
      sourceDocument,
      customDocument,
      provider,
      client,
      cancellation,
      panel,
    };
    await provider.resolveCustomEditor(customDocument, panel.panel, cancellation.token);
    panel.receive({ type: "ready" });
    await waitForDocument(panel, customDocument.document, "the initial draft preview");
    await assertions(environment);
  } finally {
    panel.dispose();
    provider.dispose();
    client.dispose();
    cancellation.dispose();
    await rm(uri.fsPath, { force: true, maxRetries: 5, retryDelay: 100 });
    await rm(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

async function createWorkspaceFixture(
  extension: "yaml" | "json",
): Promise<{ directory: string; uri: vscode.Uri }> {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, "the VS Code test runner should open the temporary xlflow workspace");
  const projectRoot = folder.uri.fsPath;
  const configuration = await readFile(path.join(projectRoot, "xlflow.toml"), "utf8");
  assert.match(configuration, /forms\s*=\s*["']src\/forms["']/);
  const specsRoot = path.join(projectRoot, "src", "forms", "specs");
  await mkdir(specsRoot, { recursive: true });
  const directory = await mkdtemp(path.join(projectRoot, ".xlflow-userform-draft-"));
  const id = path.basename(directory).slice(".xlflow-userform-draft-".length);
  const uri = vscode.Uri.file(path.join(specsRoot, `ProviderDraft-${id}.${extension}`));
  return { directory, uri };
}

async function sendCaptionEdit(
  environment: TestEnvironment,
  requestId: number,
  caption: string,
): Promise<void> {
  const { customDocument, panel } = environment;
  panel.receive({
    type: "edit",
    requestId,
    version: customDocument.document.version,
    operations: [{ type: "setFormProperty", field: "caption", value: caption }],
  });
  const reply = await waitForReply(panel, requestId);
  assert.strictEqual(reply.error, undefined, `caption edit ${requestId} should succeed`);
}

async function postTestEdit(
  panel: vscode.WebviewPanel,
  document: UserFormCustomDocument,
  requestId: number,
  caption: string,
): Promise<void> {
  await panel.webview.postMessage({
    type: "xlflow-test-edit",
    requestId,
    version: document.document.version,
    operations: [{ type: "setFormProperty", field: "caption", value: caption }],
  });
}

function createTestClient(typedProperties = false): TestClient {
  const connection = new vscode.EventEmitter<void>();
  const client = {
    onDidChangePreviewConnection: connection.event,
    requestUserFormPreview: async (document: vscode.TextDocument) => {
      const text = document.getText();
      const source =
        document.languageId === "json"
          ? (JSON.parse(text) as { controls?: SpecControl[] })
          : undefined;
      return {
        version: document.version,
        editable: true,
        propertyEditable: true,
        ...(typedProperties ? { propertyEditorVersion: 1 } : {}),
        structuralEditable: true,
        parentEditable: true,
        pageEditable: true,
        tabEditable: true,
        productivityEditable: true,
        propertyGrid: {
          form: {
            descriptors: [
              { field: "caption", valueType: "string", required: false, nullable: true },
            ],
            values: { caption: { present: true, value: captionValue(text) } },
          },
          controls: {},
        },
        document: {
          form: { name: "Main", caption: captionValue(text) },
          controls: source?.controls ?? [],
        },
      };
    },
    requestUserFormEdit: async (
      document: vscode.TextDocument,
      version: number,
      text: string,
      operations: SemanticOperation[],
    ) => {
      const operation = operations[0];
      if (document.languageId === "json" && operation?.type === "addControl") {
        const source = JSON.parse(text) as { controls: SpecControl[] };
        const additions: SpecControl[] = operations.flatMap((op) =>
          op.type === "addControl" || op.type === "addPage" ? [{ ...op.control }] : [],
        );
        if (operation.control.type === "MultiPage") additions[0].selectedIndex = 0;
        source.controls.push(...additions);
        return {
          version,
          edits: [
            {
              range: new vscode.Range(document.positionAt(0), document.positionAt(text.length)),
              newText: JSON.stringify(source, null, 2) + "\n",
            },
          ],
        };
      }
      if (operation?.type !== "setFormProperty" || operation.field !== "caption")
        return {
          version,
          error: { code: "unexpectedOperation", message: "Expected caption edit." },
        };
      const range = captionRange(document, captionValue(text));
      return {
        version,
        edits: [{ range, newText: JSON.stringify(operation.value) }],
      };
    },
  } as unknown as XlflowLanguageClientManager;
  return { client, dispose: () => connection.dispose() };
}

function captionValue(text: string): string {
  const match = text.match(/caption"?:\s*"((?:[^"\\]|\\.)*)"/);
  assert.ok(match, "the fixture must contain a quoted caption field");
  return JSON.parse(`"${match[1]}"`) as string;
}

function captionRange(document: vscode.TextDocument, currentCaption: string): vscode.Range {
  return captionRangeInText(document, document.getText(), currentCaption);
}

function captionRangeInText(
  document: vscode.TextDocument,
  text: string,
  currentCaption: string,
): vscode.Range {
  const marker = /caption"?:\s*"((?:[^"\\]|\\.)*)"/;
  const match = marker.exec(text);
  assert.ok(match, "the source must contain a quoted caption field");
  assert.strictEqual(JSON.parse(`"${match[1]}"`), currentCaption);
  const start = match.index + match[0].lastIndexOf('"') - match[1].length - 1;
  const end = start + match[1].length + 2;
  return new vscode.Range(document.positionAt(start), document.positionAt(end));
}

function replaceCaption(text: string, caption: string): string {
  const match = /caption"?:\s*"((?:[^"\\]|\\.)*)"/.exec(text);
  assert.ok(match, "the fixture must contain a quoted caption field");
  const start = match.index + match[0].lastIndexOf('"') - match[1].length - 1;
  const end = start + match[1].length + 2;
  return text.slice(0, start) + JSON.stringify(caption) + text.slice(end);
}

function replaceWholeDocument(document: vscode.TextDocument, text: string): vscode.WorkspaceEdit {
  const edit = new vscode.WorkspaceEdit();
  edit.replace(
    document.uri,
    new vscode.Range(document.positionAt(0), document.positionAt(document.getText().length)),
    text,
  );
  return edit;
}

function createFakePanel(): FakePanel {
  const incoming = new vscode.EventEmitter<WebviewMessage>();
  const viewState = new vscode.EventEmitter<vscode.WebviewPanelOnDidChangeViewStateEvent>();
  const disposed = new vscode.EventEmitter<void>();
  const messages: HostMessage[] = [];
  const panel = {
    visible: true,
    onDidChangeViewState: viewState.event,
    onDidDispose: disposed.event,
    webview: {
      options: {},
      html: "",
      cspSource: "https://test.invalid",
      asWebviewUri: (uri: vscode.Uri) => uri,
      onDidReceiveMessage: incoming.event,
      postMessage: async (message: HostMessage) => {
        messages.push(message);
        return true;
      },
    },
  } as unknown as vscode.WebviewPanel;
  return {
    panel,
    messages,
    receive: (message) => incoming.fire(message),
    dispose: () => {
      disposed.fire();
      incoming.dispose();
      viewState.dispose();
      disposed.dispose();
    },
  };
}

function instrumentPanel(
  panel: vscode.WebviewPanel,
  hostMessages: HostMessage[],
): vscode.WebviewPanel {
  const webview = new Proxy(panel.webview, {
    get(target, key) {
      if (key === "postMessage")
        return (message: HostMessage | unknown) => {
          hostMessages.push(message as HostMessage);
          return target.postMessage(message);
        };
      const value = Reflect.get(target, key, target) as unknown;
      return typeof value === "function" ? value.bind(target) : value;
    },
    set(target, key, value) {
      const assigned = key === "html" ? addTestBridge(value as string) : value;
      return Reflect.set(target, key, assigned, target);
    },
  });
  return new Proxy(panel, {
    get(target, key) {
      if (key === "webview") return webview;
      const value = Reflect.get(target, key, target) as unknown;
      return typeof value === "function" ? value.bind(target) : value;
    },
  });
}

function addTestBridge(html: string): string {
  const nonce = html.match(/script-src 'nonce-([^']+)'/)?.[1];
  assert.ok(nonce, "the shipped Designer HTML should use a nonce for its scripts");
  const script = `<script nonce="${nonce}">
const testApi = acquireVsCodeApi();
window.acquireVsCodeApi = () => testApi;
window.addEventListener("message", event => {
  const message = event.data;
  if (message?.type === "xlflow-test-edit") {
    testApi.postMessage({ type: "edit", requestId: message.requestId, version: message.version, operations: message.operations });
  } else if (message?.type === "xlflow-test-history") {
    testApi.postMessage({ type: "history", action: message.action, requestId: message.requestId, version: message.version });
  } else if (message?.type === "xlflow-test-save") {
    testApi.postMessage({ type: "save" });
  } else if (message?.type === "xlflow-test-resolve") {
    testApi.postMessage({ type: "resolveConflict", choice: message.choice });
  }
});
</script>`;
  return html.replace("</head>", `${script}</head>`);
}

async function waitForDocument(
  panel: FakePanel,
  document: vscode.TextDocument,
  description: string,
): Promise<void> {
  const message = await waitFor(
    () =>
      [...panel.messages]
        .reverse()
        .find(
          (message) =>
            (message.type === "document" || message.type === "invalidDocument") &&
            message.version === document.version,
        ),
    description,
  );
  assert.strictEqual(message.type, "document", `${description} failed: ${JSON.stringify(message)}`);
}

async function waitForReply(panel: FakePanel, requestId: number) {
  return waitFor(
    () =>
      panel.messages.find(
        (message): message is Extract<HostMessage, { type: "editResult" }> =>
          message.type === "editResult" && message.requestId === requestId,
      ),
    `edit reply ${requestId}`,
  );
}

function findCustomTab(viewType: string, uri: vscode.Uri): vscode.Tab | undefined {
  for (const group of vscode.window.tabGroups.all)
    for (const tab of group.tabs) {
      const input = tab.input;
      if (
        input instanceof vscode.TabInputCustom &&
        input.viewType === viewType &&
        input.uri.toString() === uri.toString()
      )
        return tab;
    }
  return undefined;
}

async function waitFor<T>(read: () => T | undefined | false, description: string): Promise<T> {
  const deadline = Date.now() + 10_000;
  while (Date.now() < deadline) {
    const value = read();
    if (value !== undefined && value !== false) return value as T;
    await new Promise<void>((resolve) => setTimeout(resolve, 20));
  }
  assert.fail(`Timed out waiting for ${description}`);
}

async function waitForAsync(
  read: () => boolean | Promise<boolean>,
  description: string,
): Promise<void> {
  const deadline = Date.now() + 10_000;
  while (Date.now() < deadline) {
    if (await read()) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 20));
  }
  assert.fail(`Timed out waiting for ${description}`);
}
