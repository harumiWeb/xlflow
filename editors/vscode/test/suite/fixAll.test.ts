import * as assert from "assert";
import * as vscode from "vscode";
import { vscodeDebugExecutablePath } from "../../src/config";

const fixAllKind = "source.fixAll.xlflow.vb080";
const vb080Code = "VB080";
const testTimeoutMs = 30_000;

export async function run(): Promise<void> {
  const workspaceFolder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(workspaceFolder, "the Fix All test workspace should be open");

  const executablePath = vscodeDebugExecutablePath();
  assert.ok(executablePath, "the test runner should set XLFLOW_VSCODE_DEBUG_PATH");

  const extension =
    vscode.extensions.getExtension("ed2c27e6-6563-6407-a650-31eef08e0f25.xlflow-vscode") ??
    vscode.extensions.getExtension("harumiweb.xlflow-vscode");
  assert.ok(extension, "the xlflow extension should be loaded in the Extension Host");
  await extension.activate();

  const uri = vscode.Uri.joinPath(workspaceFolder.uri, "fix-all.bas");
  const original = [
    "Option Explicit",
    "Sub Work()",
    "    Dim total As Long",
    "    Let total = 1",
    "    Let total = 2",
    "End Sub",
    "",
  ].join("\n");
  const expected = original.replaceAll("Let total =", "total =");
  await vscode.workspace.fs.writeFile(uri, Buffer.from(original, "utf8"));

  const document = await vscode.workspace.openTextDocument(uri);
  assert.strictEqual(document.languageId, "vba", "the fixture should use the VBA language mode");
  await vscode.window.showTextDocument(document);

  try {
    const diagnostics = await waitForVB080Diagnostics(document);
    assert.strictEqual(diagnostics.length, 2, "both Let assignments should be diagnosed");

    const range = new vscode.Range(
      document.positionAt(0),
      document.positionAt(document.getText().length),
    );
    const actions = await vscode.commands.executeCommand<vscode.CodeAction[]>(
      "vscode.executeCodeActionProvider",
      document.uri,
      range,
      fixAllKind,
    );
    assert.ok(actions, "the extension host should return code actions");
    const fixAllActions = actions.filter((action) => action.kind?.value === fixAllKind);
    assert.strictEqual(
      fixAllActions.length,
      1,
      "the explicit VB080 source kind should return one action",
    );

    const [fixAll] = fixAllActions;
    assert.ok(fixAll.edit, "the source action should contain an edit");
    const edits = fixAll.edit.get(document.uri);
    assert.strictEqual(
      edits.length,
      2,
      "the edit should correct both occurrences in one workspace edit",
    );
    assert.ok(edits.every((edit) => !edit.newText.includes("Let")));
    console.log(
      `[fix-all] LSP action kind=${fixAll.kind?.value}; explicitOnly=${fixAllKind}; edits=${edits.length}`,
    );

    const applied = await vscode.workspace.applyEdit(fixAll.edit);
    assert.ok(applied, "VS Code should apply the LanguageClient workspace edit");
    assert.strictEqual(document.getText(), expected, "one apply should correct both assignments");
    assert.ok(document.isDirty, "applying the source action should create an undoable buffer edit");
    console.log("[fix-all] apply changed both Let assignments in the active editor buffer");

    await vscode.commands.executeCommand("undo");
    await waitForText(document, original);
    assert.strictEqual(
      document.getText(),
      original,
      "one native VS Code Undo should restore the source",
    );
    assert.strictEqual(
      document.isDirty,
      false,
      "Undo should return the buffer to its original saved state",
    );
    console.log("[fix-all] native VS Code Undo restored the exact original buffer in one command");
  } finally {
    await vscode.commands.executeCommand("workbench.action.closeActiveEditor");
  }
}

async function waitForVB080Diagnostics(
  document: vscode.TextDocument,
): Promise<vscode.Diagnostic[]> {
  const deadline = Date.now() + testTimeoutMs;
  while (Date.now() < deadline) {
    const matches = vscode.languages
      .getDiagnostics(document.uri)
      .filter((diagnostic) => diagnosticCode(diagnostic) === vb080Code);
    if (matches.length >= 2) return matches;
    await delay(100);
  }
  const codes = vscode.languages
    .getDiagnostics(document.uri)
    .map(diagnosticCode)
    .filter((code) => code !== undefined)
    .join(", ");
  throw new Error(
    `Timed out waiting for two ${vb080Code} diagnostics; received: ${codes || "none"}`,
  );
}

async function waitForText(document: vscode.TextDocument, expected: string): Promise<void> {
  const deadline = Date.now() + 5_000;
  while (Date.now() < deadline) {
    if (document.getText() === expected) return;
    await delay(25);
  }
  assert.strictEqual(
    document.getText(),
    expected,
    "the editor buffer did not reach the expected text",
  );
}

function diagnosticCode(diagnostic: vscode.Diagnostic): string | number | undefined {
  const code = diagnostic.code;
  return typeof code === "object" && code !== null ? code.value : code;
}

function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}
