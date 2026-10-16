import * as assert from "assert";
import { mkdtemp, mkdir, readFile, rm, stat, symlink, writeFile } from "fs/promises";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";
import type { XlflowLanguageClientManager } from "../../src/client";
import {
  UserFormEditorProvider,
  type ChooseUserFormEvent,
} from "../../src/userFormEditor/provider";
import { designerResourceUri } from "../../src/userFormEditor/resource";
import type {
  UserFormEventEditRequest,
  UserFormEventEditResult,
  UserFormEventsRequest,
  UserFormEventsResult,
  UserFormEventDescriptor,
} from "../../src/userFormEditor/events";
import type { HostMessage, SpecControl, WebviewMessage } from "../../src/userFormEditor/protocol";

interface EventFixture {
  root: string;
  formName: string;
  specUri: vscode.Uri;
  sourceDocument: vscode.TextDocument;
  customDocument: import("../../src/userFormEditor/customDocument").UserFormCustomDocument;
  provider: UserFormEditorProvider;
  panel: FakePanel;
  cancellation: vscode.CancellationTokenSource;
  eventRequests: UserFormEventsRequest[];
  editRequests: UserFormEventEditRequest[];
  handlerNames: Set<string>;
  setConfirmSave: (confirm: boolean) => void;
  setBeforeNextChooser: (action: () => Promise<void>) => void;
  refreshPreview: () => Promise<void>;
  notifyPreviewConnectionChanged: () => void;
}

interface FakePanel {
  panel: vscode.WebviewPanel;
  messages: HostMessage[];
  receive(message: WebviewMessage): void;
  dispose(): void;
}

const pageControls: SpecControl[] = [
  {
    id: "multipage-id",
    type: "MultiPage",
    progId: "Forms.MultiPage.1",
    name: "MultiPage1",
    left: 8,
    top: 8,
    width: 180,
    height: 120,
  },
  {
    id: "page-id",
    type: "Page",
    progId: "Forms.Page.1",
    name: "Page1",
    parentId: "multipage-id",
    left: 0,
    top: 0,
    width: 160,
    height: 90,
  },
];

/** Exercise event RPC routing and native WorkspaceEdit/Undo inside VS Code. */
export async function runUserFormEventsAssertions(extensionUri: vscode.Uri): Promise<void> {
  await assertCancelledSaveThenNewSidecarAndUndo(extensionUri);
  await assertSameVersionPreviewRefreshDuringChooser(extensionUri);
  await assertSourceChangeDuringChooserIsRejected(extensionUri);
  await assertPreviewConnectionChangeDuringChooserIsRejected(extensionUri);
  await assertExistingSidecarDirtyBufferAndPageOwner(extensionUri);
  await assertMissingFrmIsRejected(extensionUri);
  await assertExistingFrmCanReceiveNewHandler(extensionUri);
  await assertSidecarJunctionIsRejected(extensionUri);
  await assertSidecarFileLinkIsRejected(extensionUri);
  await assertFrmJunctionIsRejected(extensionUri);
  await assertFrmFileLinkIsRejected(extensionUri);
}

async function assertSameVersionPreviewRefreshDuringChooser(
  extensionUri: vscode.Uri,
): Promise<void> {
  await withEventFixture(extensionUri, "sidecar", pageControls, async (fixture) => {
    const codeDirectory = path.join(fixture.root, "Forms", "code");
    await mkdir(codeDirectory, { recursive: true });
    const codeUri = vscode.Uri.file(path.join(codeDirectory, `${fixture.formName}.bas`));
    await writeFile(codeUri.fsPath, "Option Explicit\n", "utf8");
    const codeDocument = await vscode.workspace.openTextDocument(codeUri);
    const sourceVersion = fixture.customDocument.document.version;

    fixture.setBeforeNextChooser(fixture.refreshPreview);
    const result = await openEvent(fixture, true, "page-id");

    assert.strictEqual(
      result.error,
      undefined,
      `same-version preview refresh cancelled event navigation: ${result.error}`,
    );
    assert.strictEqual(fixture.customDocument.document.version, sourceVersion);
    assert.strictEqual(fixture.eventRequests.length, 1);
    assert.strictEqual(fixture.eventRequests[0].targetId, "page-id");
    assert.strictEqual(fixture.editRequests.length, 1);
    assert.strictEqual(fixture.editRequests[0].targetId, "page-id");
    assert.ok(codeDocument.getText().includes("Private Sub MultiPage1_Enter()"));
  });
}

async function assertSourceChangeDuringChooserIsRejected(extensionUri: vscode.Uri): Promise<void> {
  await withEventFixture(extensionUri, "sidecar", pageControls, async (fixture) => {
    const codeDirectory = path.join(fixture.root, "Forms", "code");
    await mkdir(codeDirectory, { recursive: true });
    const codeUri = vscode.Uri.file(path.join(codeDirectory, `${fixture.formName}.bas`));
    const originalCode = "Option Explicit\n";
    await writeFile(codeUri.fsPath, originalCode, "utf8");
    const codeDocument = await vscode.workspace.openTextDocument(codeUri);

    fixture.setBeforeNextChooser(async () => {
      const sourceText = fixture.customDocument.state.text;
      const end = fixture.customDocument.document.positionAt(sourceText.length);
      fixture.customDocument.apply(
        [{ range: { start: end, end }, newText: "# changed during event selection\n" }],
        "Change Designer source during event selection",
      );
      await waitForCurrentPreview(fixture);
    });
    const result = await openEvent(fixture, true, "page-id");

    assert.ok(result.error, "a source revision change must invalidate event navigation");
    assert.strictEqual(fixture.eventRequests.length, 1);
    assert.strictEqual(fixture.editRequests.length, 0);
    assert.strictEqual(codeDocument.getText(), originalCode);
  });
}

async function assertPreviewConnectionChangeDuringChooserIsRejected(
  extensionUri: vscode.Uri,
): Promise<void> {
  await withEventFixture(extensionUri, "sidecar", pageControls, async (fixture) => {
    const codeDirectory = path.join(fixture.root, "Forms", "code");
    await mkdir(codeDirectory, { recursive: true });
    const codeUri = vscode.Uri.file(path.join(codeDirectory, `${fixture.formName}.bas`));
    const originalCode = "Option Explicit\n";
    await writeFile(codeUri.fsPath, originalCode, "utf8");
    const codeDocument = await vscode.workspace.openTextDocument(codeUri);

    fixture.setBeforeNextChooser(async () => {
      fixture.notifyPreviewConnectionChanged();
    });
    const result = await openEvent(fixture, true, "page-id");

    assert.ok(result.error, "a preview connection change must invalidate event navigation");
    assert.strictEqual(fixture.eventRequests.length, 1);
    assert.strictEqual(fixture.editRequests.length, 0);
    assert.strictEqual(codeDocument.getText(), originalCode);
  });
}

async function assertCancelledSaveThenNewSidecarAndUndo(extensionUri: vscode.Uri): Promise<void> {
  await withEventFixture(extensionUri, "sidecar", [], async (fixture) => {
    const codeUri = vscode.Uri.file(
      path.join(fixture.root, "Forms", "code", `${fixture.formName}.bas`),
    );
    const draftText = fixture.customDocument.state.text;
    const end = fixture.customDocument.document.positionAt(draftText.length);
    fixture.customDocument.apply(
      [
        {
          range: { start: end, end },
          newText: "# unsaved designer draft\n",
        },
      ],
      "Event save test draft",
    );
    await waitForCurrentPreview(fixture);

    fixture.setConfirmSave(false);
    const cancelled = await openEvent(fixture, false);
    assert.strictEqual(cancelled.error, undefined, "cancel is a quiet no-op");
    assert.strictEqual(
      fixture.eventRequests.length,
      0,
      "cancel must stop before the LSP event list",
    );
    assert.strictEqual(fixture.editRequests.length, 0);
    await assertFileMissing(codeUri.fsPath, "cancel must not create a code file");
    assert.strictEqual(fixture.customDocument.state.dirty, true);

    fixture.setConfirmSave(true);
    const unrelatedUri = vscode.Uri.file(path.join(fixture.root, "unrelated.txt"));
    const unrelatedDiskText = "unrelated source stays on disk\n";
    await writeFile(unrelatedUri.fsPath, unrelatedDiskText, "utf8");
    const unrelatedDocument = await vscode.workspace.openTextDocument(unrelatedUri);
    const unrelatedEdit = new vscode.WorkspaceEdit();
    unrelatedEdit.insert(
      unrelatedUri,
      unrelatedDocument.positionAt(unrelatedDocument.getText().length),
      "unsaved change\n",
    );
    assert.strictEqual(await vscode.workspace.applyEdit(unrelatedEdit), true);
    let participantCompleted = false;
    const participant = vscode.workspace.onWillSaveTextDocument((event) => {
      if (event.document.uri.toString() !== fixture.specUri.toString()) return;
      event.waitUntil(
        (async () => {
          await fixture.provider.saveCustomDocument(
            fixture.customDocument,
            fixture.cancellation.token,
          );
          participantCompleted = true;
          return [];
        })(),
      );
    });
    let result: Extract<HostMessage, { type: "eventResult" }> | undefined;
    let unrelatedDirtyAfterOpen = false;
    let unrelatedDiskAfterOpen = "";
    try {
      result = await openEvent(fixture, false);
    } finally {
      participant.dispose();
      unrelatedDirtyAfterOpen = unrelatedDocument.isDirty;
      unrelatedDiskAfterOpen = await readFile(unrelatedUri.fsPath, "utf8");
      if (unrelatedDocument.isDirty) await unrelatedDocument.save();
    }

    assert.strictEqual(result?.error, undefined, `save and open event failed: ${result?.error}`);
    assert.strictEqual(
      participantCompleted,
      true,
      "same-URI save participant must finish without queue deadlock",
    );
    assert.strictEqual(fixture.customDocument.state.dirty, false);
    assert.ok(unrelatedDirtyAfterOpen, "Save and Continue only saves the target UserForm");
    assert.strictEqual(unrelatedDiskAfterOpen, unrelatedDiskText);
    assert.ok(
      (await readFile(fixture.specUri.fsPath, "utf8")).includes("# unsaved designer draft"),
    );
    assert.strictEqual(fixture.eventRequests.length, 1);
    assert.strictEqual(fixture.eventRequests[0].targetId, undefined);
    assert.strictEqual(fixture.editRequests.length, 1);
    assert.strictEqual(fixture.editRequests[0].codeText, "");
    assert.strictEqual(fixture.editRequests[0].codeVersion, 0);
    assert.strictEqual(fixture.editRequests[0].codeSource, "sidecar");
    const editor = vscode.workspace.getConfiguration("editor", codeUri);
    const tabSize = editor.get<number>("tabSize", 4);
    const expectedIndent = editor.get<boolean>("insertSpaces", true)
      ? " ".repeat(Number.isSafeInteger(tabSize) && tabSize > 0 ? tabSize : 4)
      : "\t";
    assert.strictEqual(
      fixture.editRequests[0].indent,
      expectedIndent,
      "new handlers use the code editor's configured indentation",
    );
    const codeDocument = await vscode.workspace.openTextDocument(codeUri);
    assert.strictEqual(
      codeDocument.getText(),
      normalizeLineEndings(generatedText("UserForm_Initialize", "\n"), codeDocument.eol),
      "new file text is normalized to the active VS Code document EOL",
    );
    assert.strictEqual(codeDocument.isDirty, true, "new sidecar text remains unsaved");
    assert.strictEqual(
      await readFile(codeUri.fsPath, "utf8"),
      "",
      "event opening does not save code",
    );

    await vscode.commands.executeCommand("undo");
    await waitForFileMissing(codeUri.fsPath, "native Undo should remove the newly created sidecar");
  });
}

async function assertExistingSidecarDirtyBufferAndPageOwner(
  extensionUri: vscode.Uri,
): Promise<void> {
  await withEventFixture(extensionUri, "sidecar", pageControls, async (fixture) => {
    const codeDirectory = path.join(fixture.root, "Forms", "code");
    await mkdir(codeDirectory, { recursive: true });
    const codeUri = vscode.Uri.file(path.join(codeDirectory, `${fixture.formName}.bas`));
    const diskText = "Option Explicit\nPrivate Sub Existing()\n\tDebug.Print 1\nEnd Sub\n";
    await writeFile(codeUri.fsPath, diskText, "utf8");
    const codeDocument = await vscode.workspace.openTextDocument(codeUri);
    const userEdit = new vscode.WorkspaceEdit();
    userEdit.insert(
      codeUri,
      codeDocument.positionAt(codeDocument.getText().length),
      "' latest unsaved buffer\n",
    );
    assert.strictEqual(await vscode.workspace.applyEdit(userEdit), true);
    const dirtyText = codeDocument.getText();
    const dirtyVersion = codeDocument.version;
    assert.ok(codeDocument.isDirty);

    const first = await openEvent(fixture, true, "page-id");
    assert.strictEqual(first.error, undefined, `Page event operation failed: ${first.error}`);
    assert.strictEqual(
      fixture.eventRequests[0].targetId,
      "page-id",
      "the selected Page ID stays on the wire",
    );
    assert.strictEqual(fixture.editRequests[0].targetId, "page-id");
    assert.strictEqual(
      fixture.editRequests[0].eventName,
      "Enter",
      "the injected chooser's selection is used",
    );
    assert.strictEqual(
      fixture.editRequests[0].codeText,
      dirtyText,
      "latest dirty code text is sent verbatim",
    );
    assert.strictEqual(fixture.editRequests[0].codeVersion, dirtyVersion);
    assert.strictEqual(
      fixture.editRequests[0].indent,
      "\t",
      "existing body indentation is preserved",
    );
    assert.ok(codeDocument.getText().includes("Private Sub MultiPage1_Enter()"));
    assert.strictEqual(
      await readFile(codeUri.fsPath, "utf8"),
      diskText,
      "event opening never saves code",
    );

    const afterCreate = codeDocument.getText();
    const second = await openEvent(fixture, true, "page-id");
    assert.strictEqual(
      second.error,
      undefined,
      `existing handler navigation failed: ${second.error}`,
    );
    assert.strictEqual(fixture.editRequests.length, 2);
    assert.strictEqual(fixture.editRequests[1].codeText, afterCreate);
    assert.strictEqual(fixture.editRequests[1].codeVersion, codeDocument.version);
    assert.strictEqual(
      codeDocument.getText().match(/Private Sub MultiPage1_Enter\(\)/g)?.length,
      1,
      "repeating Open Event navigates to the existing procedure without duplicating it",
    );
    assert.strictEqual(await readFile(codeUri.fsPath, "utf8"), diskText);

    await vscode.commands.executeCommand("undo");
    assert.strictEqual(
      codeDocument.getText(),
      dirtyText,
      "Undo removes only the generated handler",
    );
    assert.ok(codeDocument.isDirty, "Undo preserves the user's preexisting dirty buffer");
    assert.strictEqual(await readFile(codeUri.fsPath, "utf8"), diskText);
  });
}

async function assertExistingFrmCanReceiveNewHandler(extensionUri: vscode.Uri): Promise<void> {
  await withEventFixture(extensionUri, "frm", [], async (fixture) => {
    const codeUri = vscode.Uri.file(path.join(fixture.root, "Forms", `${fixture.formName}.frm`));
    const diskText = [
      "VERSION 5.00",
      `Begin VB.UserForm ${fixture.formName}`,
      `   Caption = "Event fixture"`,
      "End",
      `Attribute VB_Name = "${fixture.formName}"`,
      "",
    ].join("\r");
    await writeFile(codeUri.fsPath, diskText, "utf8");
    const codeDocument = await vscode.workspace.openTextDocument(codeUri);
    const canonicalCodeText = codeDocument.getText();

    const result = await openEvent(fixture, false);
    assert.strictEqual(
      result.error,
      undefined,
      `new handler in existing .frm failed: ${result.error}`,
    );
    assert.strictEqual(fixture.editRequests.length, 1);
    assert.strictEqual(fixture.editRequests[0].codeSource, "frm");
    assert.strictEqual(fixture.editRequests[0].codeText, canonicalCodeText);
    assert.strictEqual(
      fixture.editRequests[0].newline,
      documentEol(codeDocument.eol),
      "CR-only sources use the canonical VS Code document EOL",
    );
    assert.ok(fixture.editRequests[0].codeURI.endsWith(`${fixture.formName}.frm`));
    assert.ok(codeDocument.getText().includes("Private Sub UserForm_Initialize()"));
    assert.strictEqual(
      await readFile(codeUri.fsPath, "utf8"),
      diskText,
      ".frm code remains unsaved",
    );
    assert.strictEqual(await stat(codeUri.fsPath).then(() => true), true);
  });
}

async function assertMissingFrmIsRejected(extensionUri: vscode.Uri): Promise<void> {
  await withEventFixture(extensionUri, "frm", [], async (fixture) => {
    const codeUri = vscode.Uri.file(path.join(fixture.root, "Forms", `${fixture.formName}.frm`));
    const result = await openEvent(fixture, false);

    assert.ok(
      result.error?.includes("configured .frm code source"),
      `missing .frm source should be explained to the user: ${result.error}`,
    );
    assert.strictEqual(fixture.eventRequests.length, 1);
    assert.strictEqual(fixture.editRequests.length, 0);
    await assertFileMissing(codeUri.fsPath, "event navigation must not create a missing .frm file");
  });
}

async function assertSidecarJunctionIsRejected(extensionUri: vscode.Uri): Promise<void> {
  await withEventFixture(extensionUri, "sidecar", [], async (fixture) => {
    const external = await mkdtemp(path.join(os.tmpdir(), "xlflow-userform-events-external-"));
    try {
      const externalCode = path.join(external, `${fixture.formName}.bas`);
      const localCodeDirectory = path.join(fixture.root, "Forms", "code");
      const sentinel = "Option Explicit\n' external sidecar sentinel\n";
      await writeFile(externalCode, sentinel, "utf8");
      await symlink(
        external,
        localCodeDirectory,
        process.platform === "win32" ? "junction" : "dir",
      );

      const result = await openEvent(fixture, false);
      assertLinkRejected(result.error, fixture.editRequests.length);
      assert.strictEqual(await readFile(externalCode, "utf8"), sentinel);
    } finally {
      await rm(external, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    }
  });
}

async function assertSidecarFileLinkIsRejected(extensionUri: vscode.Uri): Promise<void> {
  await withEventFixture(extensionUri, "sidecar", [], async (fixture) => {
    const external = await mkdtemp(path.join(os.tmpdir(), "xlflow-userform-events-external-"));
    try {
      const externalCode = path.join(external, `${fixture.formName}.bas`);
      const localCodeDirectory = path.join(fixture.root, "Forms", "code");
      const localCode = path.join(localCodeDirectory, `${fixture.formName}.bas`);
      const sentinel = "Option Explicit\n' external linked sidecar sentinel\n";
      await mkdir(localCodeDirectory, { recursive: true });
      await writeFile(externalCode, sentinel, "utf8");
      await symlink(externalCode, localCode, "file");

      const result = await openEvent(fixture, false);
      assertLinkRejected(result.error, fixture.editRequests.length);
      assert.strictEqual(await readFile(externalCode, "utf8"), sentinel);
    } finally {
      await rm(external, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    }
  });
}

async function assertFrmJunctionIsRejected(extensionUri: vscode.Uri): Promise<void> {
  await withEventFixture(extensionUri, "frm", [], async (fixture) => {
    const external = await mkdtemp(path.join(os.tmpdir(), "xlflow-userform-events-external-"));
    try {
      const externalFrm = path.join(external, `${fixture.formName}.frm`);
      const localFormsSubdirectory = path.join(fixture.root, "Forms", "ExternalForms");
      const sentinel = frmText(fixture.formName);
      await writeFile(externalFrm, sentinel, "utf8");
      await symlink(
        external,
        localFormsSubdirectory,
        process.platform === "win32" ? "junction" : "dir",
      );

      const result = await openEvent(fixture, false);
      assertLinkRejected(result.error, fixture.editRequests.length);
      assert.strictEqual(await readFile(externalFrm, "utf8"), sentinel);
    } finally {
      await rm(external, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    }
  });
}

async function assertFrmFileLinkIsRejected(extensionUri: vscode.Uri): Promise<void> {
  await withEventFixture(extensionUri, "frm", [], async (fixture) => {
    const external = await mkdtemp(path.join(os.tmpdir(), "xlflow-userform-events-external-"));
    try {
      const externalFrm = path.join(external, `${fixture.formName}.frm`);
      const localFormsSubdirectory = path.join(fixture.root, "Forms", "Nested");
      const localFrm = path.join(localFormsSubdirectory, `${fixture.formName}.frm`);
      const sentinel = frmText(fixture.formName);
      await mkdir(localFormsSubdirectory, { recursive: true });
      await writeFile(externalFrm, sentinel, "utf8");
      await symlink(externalFrm, localFrm, "file");

      const result = await openEvent(fixture, false);
      assertLinkRejected(result.error, fixture.editRequests.length);
      assert.strictEqual(await readFile(externalFrm, "utf8"), sentinel);
    } finally {
      await rm(external, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    }
  });
}

function assertLinkRejected(error: string | undefined, editRequestCount: number): void {
  assert.ok(
    error?.includes("symbolic links or junctions"),
    `expected localized link error, got: ${error}`,
  );
  assert.strictEqual(editRequestCount, 0, "unsafe external code must not reach the edit RPC");
}

function frmText(formName: string): string {
  return [
    "VERSION 5.00",
    `Begin VB.UserForm ${formName}`,
    '   Caption = "External fixture"',
    "End",
    `Attribute VB_Name = "${formName}"`,
    "",
  ].join("\n");
}

async function withEventFixture(
  extensionUri: vscode.Uri,
  codeSource: "sidecar" | "frm",
  controls: SpecControl[],
  assertions: (fixture: EventFixture) => Promise<void>,
): Promise<void> {
  const workspaceFolder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(workspaceFolder, "the VS Code test runner should open its temporary xlflow workspace");
  const workspaceRoot = workspaceFolder.uri.fsPath;
  const configPath = path.join(workspaceRoot, "xlflow.toml");
  const originalConfig = await readFile(configPath, "utf8");
  const root = await mkdtemp(path.join(workspaceRoot, ".xlflow-userform-events-"));
  const formsRoot = path.join(root, "Forms");
  const specsRoot = path.join(formsRoot, "specs");
  await mkdir(specsRoot, { recursive: true });
  const formName = `EventForm${path.basename(root).replace(/[^A-Za-z0-9]/g, "")}`;
  const specUri = vscode.Uri.file(path.join(specsRoot, `${formName}.yaml`));
  const configuredFormsRoot = path.relative(workspaceRoot, formsRoot).split(path.sep).join("/");
  const configured = originalConfig.replace(
    /(^\s*forms\s*=\s*)["'][^"']*["']/m,
    `$1"${configuredFormsRoot}"`,
  );
  const eventConfig = /(^\s*code_source\s*=\s*)["'][^"']*["']/m.test(configured)
    ? configured.replace(/(^\s*code_source\s*=\s*)["'][^"']*["']/m, `$1"${codeSource}"`)
    : `${configured.trimEnd()}\n\n[userform]\ncode_source = "${codeSource}"\n`;
  await writeFile(configPath, eventConfig, "utf8");
  await writeFile(
    specUri.fsPath,
    `schemaVersion: 1\nkind: xlflow.userform\nform:\n  name: ${formName}\ncontrols: []\n`,
    "utf8",
  );

  const connection = new vscode.EventEmitter<void>();
  const cancellation = new vscode.CancellationTokenSource();
  const eventRequests: UserFormEventsRequest[] = [];
  const editRequests: UserFormEventEditRequest[] = [];
  const handlerNames = new Set<string>();
  let confirmSave = false;
  let beforeNextChooser: (() => Promise<void>) | undefined;
  const setConfirmSave = (confirm: boolean) => (confirmSave = confirm);
  const setBeforeNextChooser = (action: () => Promise<void>) => (beforeNextChooser = action);
  const chooseEvent: ChooseUserFormEvent = async (events: readonly UserFormEventDescriptor[]) => {
    const action = beforeNextChooser;
    beforeNextChooser = undefined;
    await action?.();
    return events.find((event) => event.name === "Enter")?.name ?? events[0]?.name;
  };
  const client = {
    onDidChangePreviewConnection: connection.event,
    requestUserFormPreview: async (document: vscode.TextDocument) => ({
      version: document.version,
      eventEditable: true,
      editable: true,
      document: { form: { name: formName }, controls },
    }),
    requestUserFormEvents: async (
      _document: vscode.TextDocument,
      request: UserFormEventsRequest,
    ): Promise<UserFormEventsResult> => {
      eventRequests.push(request);
      const objectName =
        request.targetId === "page-id"
          ? "MultiPage1"
          : request.targetId === "button-id"
            ? "RunButton"
            : "UserForm";
      const isPage = request.targetId === "page-id";
      return {
        version: request.version,
        ...(request.targetId === undefined ? {} : { targetId: request.targetId }),
        objectName,
        events: isPage
          ? [
              { name: "Change", signature: "Private Sub MultiPage1_Change()", isDefault: true },
              { name: "Enter", signature: "Private Sub MultiPage1_Enter()", isDefault: false },
            ]
          : [
              {
                name: "Initialize",
                signature: "Private Sub UserForm_Initialize()",
                isDefault: true,
              },
            ],
      };
    },
    requestUserFormEventEdit: async (
      _document: vscode.TextDocument,
      request: UserFormEventEditRequest,
    ): Promise<UserFormEventEditResult> => {
      editRequests.push({ ...request });
      const objectName =
        request.targetId === "page-id"
          ? "MultiPage1"
          : request.targetId === "button-id"
            ? "RunButton"
            : "UserForm";
      const handler = `${objectName}_${request.eventName}`;
      const signature = `Private Sub ${handler}()`;
      if (handlerNames.has(handler)) {
        const bodyOffset = procedureBodyOffset(request.codeText, signature, request.newline);
        assert.notStrictEqual(
          bodyOffset,
          undefined,
          `existing procedure ${handler} must be present`,
        );
        const position = positionAt(request.codeText, bodyOffset!);
        return {
          version: request.version,
          codeVersion: request.codeVersion,
          edits: [],
          selection: { start: position, end: position },
          created: false,
        };
      }

      handlerNames.add(handler);
      const newline = request.newline;
      const prefix =
        request.codeText.length === 0
          ? `Option Explicit${newline}${newline}`
          : /(?:\r\n|\r|\n)$/.test(request.codeText)
            ? ""
            : newline;
      const inserted = `${prefix}${signature}${newline}${newline}End Sub${newline}`;
      const updated = request.codeText + inserted;
      const bodyOffset = procedureBodyOffset(updated, signature, newline)!;
      const position = positionAt(updated, bodyOffset);
      const end = positionAt(request.codeText, request.codeText.length);
      return {
        version: request.version,
        codeVersion: request.codeVersion,
        edits: [
          {
            range: { start: end, end },
            newText: inserted,
          },
        ],
        selection: { start: position, end: position },
        created: true,
      };
    },
  } as unknown as XlflowLanguageClientManager;
  const provider = new UserFormEditorProvider(
    extensionUri,
    client,
    undefined,
    undefined,
    vscode.Uri.file(path.join(root, ".draft-checkpoints")),
    undefined,
    async () => confirmSave,
    chooseEvent,
  );
  const panel = createFakePanel();
  let customDocument: EventFixture["customDocument"] | undefined;
  try {
    const sourceDocument = await vscode.workspace.openTextDocument(specUri);
    customDocument = await provider.openCustomDocument(
      designerResourceUri(sourceDocument.uri),
      { backupId: undefined, untitledDocumentData: undefined },
      cancellation.token,
    );
    await provider.resolveCustomEditor(customDocument, panel.panel, cancellation.token);
    panel.receive({ type: "ready" });
    const fixture: EventFixture = {
      root,
      formName,
      specUri,
      sourceDocument,
      customDocument,
      provider,
      panel,
      cancellation,
      eventRequests,
      editRequests,
      handlerNames,
      setConfirmSave,
      setBeforeNextChooser,
      refreshPreview: async () => {
        const documentVersion = customDocument!.document.version;
        const previewCount = panel.messages.filter(
          (message) => message.type === "document" && message.version === documentVersion,
        ).length;
        customDocument!.notify();
        await waitFor(
          () =>
            panel.messages.filter(
              (message) => message.type === "document" && message.version === documentVersion,
            ).length > previewCount,
          `refreshed preview for unchanged source version ${documentVersion}`,
        );
      },
      notifyPreviewConnectionChanged: () => connection.fire(),
    };
    await waitForCurrentPreview(fixture);
    await assertions(fixture);
  } finally {
    panel.dispose();
    provider.dispose();
    connection.dispose();
    cancellation.dispose();
    await closeFixtureEditors(root);
    await writeFile(configPath, originalConfig, "utf8");
    await rm(root, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

function createFakePanel(): FakePanel {
  const incoming = new vscode.EventEmitter<WebviewMessage>();
  const viewState = new vscode.EventEmitter<vscode.WebviewPanelOnDidChangeViewStateEvent>();
  const disposed = new vscode.EventEmitter<void>();
  const messages: HostMessage[] = [];
  const panel = {
    get visible() {
      return true;
    },
    get active() {
      return true;
    },
    viewColumn: vscode.ViewColumn.One,
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

async function openEvent(
  fixture: EventFixture,
  choose: boolean,
  targetId?: string,
): Promise<Extract<HostMessage, { type: "eventResult" }>> {
  await waitForCurrentPreview(fixture);
  const requestId =
    fixture.panel.messages.filter((message) => message.type === "eventResult").length + 1;
  fixture.panel.receive({
    type: "openEvent",
    requestId,
    version: fixture.customDocument.document.version,
    ...(targetId === undefined ? {} : { targetId }),
    choose,
  });
  return waitFor(
    () =>
      fixture.panel.messages.find(
        (message): message is Extract<HostMessage, { type: "eventResult" }> =>
          message.type === "eventResult" && message.requestId === requestId,
      ),
    `event result ${requestId}`,
  );
}

async function waitForCurrentPreview(fixture: EventFixture): Promise<void> {
  await waitFor(
    () =>
      [...fixture.panel.messages]
        .reverse()
        .find(
          (message) =>
            message.type === "document" &&
            message.version === fixture.customDocument.document.version &&
            message.document.eventEditable,
        ),
    `event-editable preview ${fixture.customDocument.document.version}`,
  );
}

async function waitFor<T>(read: () => T | undefined | false, description: string): Promise<T> {
  const deadline = Date.now() + 15_000;
  while (Date.now() < deadline) {
    const value = read();
    if (value !== undefined && value !== false) return value as T;
    await new Promise<void>((resolve) => setTimeout(resolve, 20));
  }
  assert.fail(`Timed out waiting for ${description}`);
}

function generatedText(handler: string, newline: string): string {
  return `Option Explicit${newline}${newline}Private Sub ${handler}()${newline}${newline}End Sub${newline}`;
}

function normalizeLineEndings(text: string, eol: vscode.EndOfLine): string {
  return text.replace(/\r\n|\r|\n/g, documentEol(eol));
}

function documentEol(eol: vscode.EndOfLine): string {
  return eol === vscode.EndOfLine.CRLF ? "\r\n" : "\n";
}

function procedureBodyOffset(text: string, signature: string, newline: string): number | undefined {
  const signatureOffset = text.indexOf(signature);
  if (signatureOffset < 0) return undefined;
  const bodyOffset = signatureOffset + signature.length + newline.length;
  return bodyOffset <= text.length ? bodyOffset : undefined;
}

function positionAt(text: string, offset: number): { line: number; character: number } {
  const preceding = text.slice(0, offset);
  const lines = preceding.split(/\r\n|\r|\n/);
  return { line: lines.length - 1, character: lines[lines.length - 1].length };
}

async function assertFileMissing(filePath: string, message: string): Promise<void> {
  await assert.rejects(stat(filePath), { code: "ENOENT" }, message);
}

async function waitForFileMissing(filePath: string, description: string): Promise<void> {
  const deadline = Date.now() + 15_000;
  while (Date.now() < deadline) {
    try {
      await stat(filePath);
    } catch {
      return;
    }
    await new Promise<void>((resolve) => setTimeout(resolve, 20));
  }
  assert.fail(`Timed out waiting for ${description}`);
}

async function closeFixtureEditors(root: string): Promise<void> {
  for (const group of vscode.window.tabGroups.all) {
    const fixtureTabs = group.tabs.filter((tab) => {
      const uri = (tab.input as { uri?: vscode.Uri }).uri;
      return (
        uri?.scheme === "file" && path.resolve(uri.fsPath).startsWith(path.resolve(root) + path.sep)
      );
    });
    if (fixtureTabs.length) await vscode.window.tabGroups.close(fixtureTabs, true);
  }
}
