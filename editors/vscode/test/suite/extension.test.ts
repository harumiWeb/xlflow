import * as assert from "assert";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";
import {
  backupIDFromPushEnvelope,
  backupQuickPickItems,
  compareBackupsNewestFirst,
  formatBackupTimestamp,
  formatBytes,
  initWorkbookExtensions,
  moduleInstallFailureMessage,
  moduleInstallPreflightBlocked,
  newProjectWorkbookPlaceholder,
  pruneSummary,
  validateKeepLastInput,
} from "../../src/commands";
import {
  cliNotificationSuppressionKey,
  normalizeAvailabilityFailure,
  normalizeAvailabilitySuccess,
} from "../../src/cliAvailability";
import {
  documentationSnippet,
  diagnosticActionKey,
  diagnosticRuleCode,
  disableLineSuffix,
  disableNextLineComment,
  parseProcedureDeclaration,
} from "../../src/codeActions";
import {
  isAnnotationCommentPrefix,
  isDocCommentSnippetPrefix,
  isProgIdStringPrefix,
  isStatementPrefix,
  lspCodeLensOptions,
  lspServerArgs,
} from "../../src/client";
import {
  isVbaSourcePath,
  mismatchedVbaLanguageDocuments,
  vbaFilesAssociationUpdate,
} from "../../src/languageAssociation";
import {
  sessionQuickPickItems,
  sessionStateFromEnvelope,
  sessionStatusText,
} from "../../src/session";
import {
  buildFormulaSnapshotNodes,
  buildUserFormModels,
  moduleContextValue,
  moduleGroups,
  readExcelPathFromToml,
  readFormsRootFromToml,
  readUserFormCodeSourceFromToml,
  saveRequiredProjectNode,
  userFormArtifactContextValue,
  userFormContextValue,
} from "../../src/sidebar";
import { sourceUri } from "../../src/testDiscovery";
import { discoveredTestDescription } from "../../src/testing";
import { buildTerminalCommandLine, containsVBAObjectModelAccessIssue } from "../../src/xlflow";
import {
  cliVersionSummary,
  lastCheckedKey,
  normalizeUpdateResult,
  shouldRunAutomaticCheck,
  updateDismissedKey,
  updateSummary,
} from "../../src/updateCheck";

export async function run(): Promise<void> {
  const config = vscode.workspace.getConfiguration("xlflow");
  const previousLspEnabled = config.get<boolean>("lsp.enabled");
  await config.update("lsp.enabled", false, vscode.ConfigurationTarget.Global);
  try {
    await runAssertions(config);
  } finally {
    await config.update("lsp.enabled", previousLspEnabled, vscode.ConfigurationTarget.Global);
  }
}

async function runAssertions(config: vscode.WorkspaceConfiguration): Promise<void> {
  const extension =
    vscode.extensions.getExtension("ed2c27e6-6563-6407-a650-31eef08e0f25.xlflow-vscode") ??
    vscode.extensions.getExtension("harumiweb.xlflow-vscode");
  assert.ok(extension, "extension should be discoverable");
  await extension.activate();
  assertLocalizationResources(extension.extensionPath);
  assert.deepStrictEqual([...initWorkbookExtensions], ["xlsm", "xlam", "xlsb"]);
  assert.ok(newProjectWorkbookPlaceholder.includes(".xlsb"));
  assert.ok(newProjectWorkbookPlaceholder.includes(".xlam"));

  const languages = await vscode.languages.getLanguages();
  assert.ok(languages.includes("vba"), "vba language should be registered");

  const commands = await vscode.commands.getCommands(true);
  for (const command of [
    "xlflow.restartLanguageServer",
    "xlflow.checkEnvironment",
    "xlflow.openInstallGuide",
    "xlflow.configurePath",
    "xlflow.retryCliDetection",
    "xlflow.checkForUpdates",
    "xlflow.newProject",
    "xlflow.initProject",
    "xlflow.skillInstall",
    "xlflow.moduleInstall",
    "xlflow.newModule",
    "xlflow.newStandardModule",
    "xlflow.newClassModule",
    "xlflow.newUserForm",
    "xlflow.pull",
    "xlflow.pullFormulas",
    "xlflow.push",
    "xlflow.rollbackWorkbook",
    "xlflow.pruneBackups",
    "xlflow.inspectWorkbook",
    "xlflow.runMacro",
    "xlflow.runProcedure",
    "xlflow.runTestProcedure",
    "xlflow.renameModule",
    "xlflow.deleteModule",
    "xlflow.revealSourceFile",
    "xlflow.copyModuleName",
    "xlflow.copyRelativePath",
    "xlflow.copyProcedureName",
    "xlflow.copyQualifiedName",
    "xlflow.renameUserForm",
    "xlflow.deleteUserForm",
    "xlflow.revealUserFormSource",
    "xlflow.copyUserFormName",
    "xlflow.copyUserFormRelativePath",
    "xlflow.test",
    "xlflow.lintWorkspace",
    "xlflow.formatDocument",
    "xlflow.formatProject",
    "xlflow.saveWorkbook",
    "xlflow.sessionStart",
    "xlflow.sessionAttach",
    "xlflow.sessionStatus",
    "xlflow.sessionRestart",
    "xlflow.sessionStop",
    "xlflow.sessionActions",
    "xlflow.openOutput",
    "xlflow.refreshProject",
    "xlflow.refreshModules",
    "xlflow.collapseModules",
    "xlflow.refreshUserForms",
    "xlflow.collapseUserForms",
    "xlflow.refreshTests",
    "xlflow.refreshFormulas",
    "xlflow.runAllTests",
    "xlflow.runDoctor",
    "xlflow.sessionToggle",
    "xlflow.setupActions",
    "xlflow.openDocumentation",
    "xlflow.openWorkbook",
    "xlflow.openModule",
    "xlflow.openProcedure",
    "xlflow.openUserFormArtifact",
    "xlflow.openFormulaSnapshotFile",
    "xlflow.revealFormulaSnapshotFile",
  ]) {
    assert.ok(commands.includes(command), `${command} should be registered`);
  }

  assert.strictEqual(config.get<string>("path"), "xlflow");
  assert.strictEqual(
    buildTerminalCommandLine("xlflow", ["run", "--interactive", "Sheet1.Sample"]),
    "xlflow run --interactive Sheet1.Sample",
  );
  assert.strictEqual(
    buildTerminalCommandLine("C:\\Program Files\\xlflow\\xlflow.exe", ["run", "Main.Hello World"]),
    '"C:\\Program Files\\xlflow\\xlflow.exe" run "Main.Hello World"',
  );
  assert.strictEqual(
    containsVBAObjectModelAccessIssue(
      "import_vba_components failed: get_vbproject failed: プログラミングによる Visual Basic プロジェクトへのアクセスは信頼性に欠けます",
    ),
    true,
  );
  assert.strictEqual(
    containsVBAObjectModelAccessIssue("warning: vba_object_model_access_disabled"),
    true,
  );
  assert.strictEqual(
    containsVBAObjectModelAccessIssue("[ok] VBIDE access - VBA project object model is available"),
    false,
  );
  assert.strictEqual(containsVBAObjectModelAccessIssue("ordinary xlflow output"), false);
  const backupFolder = {
    uri: vscode.Uri.file("C:/tmp/xlflow"),
    name: "xlflow",
    index: 0,
  };
  const backupDate = new Date("2026-07-12T13:42:01+09:00");
  assert.strictEqual(
    formatBackupTimestamp("2026-07-12T13:42:01+09:00"),
    `${backupDate.getFullYear()}-${pad2(backupDate.getMonth() + 1)}-${pad2(
      backupDate.getDate(),
    )} ${pad2(backupDate.getHours())}:${pad2(backupDate.getMinutes())}`,
  );
  assert.strictEqual(formatBackupTimestamp("not-a-date"), "not-a-date");
  assert.strictEqual(formatBytes(31.4 * 1024 * 1024), "31.4 MB");
  assert.strictEqual(formatBytes(42), "42 bytes");
  assert.strictEqual(validateKeepLastInput("3"), undefined);
  assert.strictEqual(validateKeepLastInput("0"), "Enter a positive whole number.");
  assert.strictEqual(validateKeepLastInput("1.5"), "Enter a positive whole number.");
  assert.strictEqual(
    backupIDFromPushEnvelope({ status: "failed", backup: { id: "20260712-push-a1b2c3" } }),
    "20260712-push-a1b2c3",
  );
  assert.strictEqual(
    backupIDFromPushEnvelope({ status: "failed", backup: { id: "  " } }),
    undefined,
  );
  const moduleInstallPreflight = {
    status: "failed",
    error: {
      code: "lint_failed",
      message: "1 source issue(s) must be fixed before pushing to Excel",
      phase: "preflight",
    },
  };
  assert.strictEqual(moduleInstallPreflightBlocked(moduleInstallPreflight, true), true);
  assert.strictEqual(moduleInstallPreflightBlocked(moduleInstallPreflight, false), false);
  assert.ok(
    moduleInstallFailureMessage(moduleInstallPreflight, "", 1, true).includes(
      "Helper modules were installed to source",
    ),
  );
  assert.strictEqual(
    moduleInstallFailureMessage(
      { status: "failed", error: { code: "module_install_failed", message: "already exists" } },
      "",
      1,
      false,
    ),
    "already exists",
  );
  assert.deepStrictEqual(
    [
      { id: "older", created_at: "2026-07-12T10:00:00+09:00" },
      { id: "newer", created_at: "2026-07-12T11:00:00+09:00" },
    ]
      .sort(compareBackupsNewestFirst)
      .map((record) => record.id),
    ["newer", "older"],
  );
  const backupItems = backupQuickPickItems(
    [
      {
        id: "20260712-134201-123-push-a1b2c3",
        created_at: "2026-07-12T13:42:01+09:00",
        reason: "before-push",
        workbook: "build/Book.xlsm",
        path: ".xlflow/backups/20260712-134201-123-push-a1b2c3/Book.xlsm",
        size_bytes: 31.4 * 1024 * 1024,
      },
      { id: "", created_at: "2026-07-12T14:00:00+09:00" },
      { id: "malformed", created_at: "bad", size_bytes: -1 },
    ],
    backupFolder,
  );
  assert.strictEqual(backupItems.length, 2);
  assert.strictEqual(backupItems[0].record.id, "20260712-134201-123-push-a1b2c3");
  assert.ok(backupItems[0].label.includes("before-push"));
  assert.ok(backupItems[0].detail?.includes("31.4 MB"));
  assert.deepStrictEqual(
    pruneSummary({
      backup_prune: {
        matched: 2,
        deleted: 1,
        failed: 1,
        freed_bytes: 2048,
        candidates: [{ id: "old", size_bytes: 2048 }],
      },
    }),
    {
      matched: 2,
      deleted: 1,
      failed: 1,
      freedBytes: 2048,
      candidateBytes: 2048,
      candidates: [{ id: "old", size_bytes: 2048 }],
    },
  );
  assert.deepStrictEqual(pruneSummary(undefined), {
    matched: 0,
    deleted: 0,
    failed: 0,
    freedBytes: 0,
    candidateBytes: 0,
    candidates: [],
  });
  assert.strictEqual(
    disableLineSuffix("    Dim staleValue As Long", "VB020"),
    " ' xlflow:disable-line VB020",
  );
  assert.strictEqual(
    disableLineSuffix("    Dim staleValue As Long ' xlflow:disable-line VB020", "VB021"),
    " VB021",
  );
  assert.strictEqual(
    disableLineSuffix("    Dim staleValue As Long ' xlflow:disable-line VB020", "VB020"),
    "",
  );
  assert.strictEqual(
    disableNextLineComment("    Dim staleValue As Long", "VB020", vscode.EndOfLine.LF),
    "    ' xlflow:disable-next-line VB020\n",
  );
  const suppressibleDiagnostic = new vscode.Diagnostic(
    new vscode.Range(0, 0, 0, 1),
    "unused local",
    vscode.DiagnosticSeverity.Warning,
  );
  suppressibleDiagnostic.source = "xlflow";
  suppressibleDiagnostic.code = "VB020";
  assert.strictEqual(diagnosticRuleCode(suppressibleDiagnostic), "VB020");
  const secondSuppressibleDiagnostic = new vscode.Diagnostic(
    new vscode.Range(1, 0, 1, 1),
    "unused local",
    vscode.DiagnosticSeverity.Warning,
  );
  secondSuppressibleDiagnostic.source = "xlflow";
  secondSuppressibleDiagnostic.code = "VB020";
  assert.notStrictEqual(
    diagnosticActionKey(suppressibleDiagnostic, "VB020"),
    diagnosticActionKey(secondSuppressibleDiagnostic, "VB020"),
  );
  suppressibleDiagnostic.code = "VB029";
  assert.strictEqual(diagnosticRuleCode(suppressibleDiagnostic), undefined);
  assert.deepStrictEqual(
    await lspServerArgs({ lspLogFile: ".xlflow/lsp.log", lspLogFileConfigured: false }, undefined),
    ["lsp", "--stdio"],
  );
  const codeLensConfig = {
    codeLensEnabled: true,
    codeLensRunProcedure: true,
    codeLensRunTests: true,
    codeLensUserFormEvents: false,
  };
  assert.deepStrictEqual(lspCodeLensOptions(codeLensConfig, false), {
    enabled: false,
    runProcedure: true,
    runTests: true,
    userFormEvents: false,
  });
  assert.deepStrictEqual(lspCodeLensOptions(codeLensConfig, true), {
    enabled: true,
    runProcedure: true,
    runTests: true,
    userFormEvents: false,
  });
  assert.deepStrictEqual(
    await lspServerArgs({ lspLogFile: ".xlflow/lsp.log", lspLogFileConfigured: true }, undefined),
    ["lsp", "--stdio", "--log-file", ".xlflow/lsp.log"],
  );
  assert.strictEqual(isStatementPrefix("Public Fu"), true);
  assert.strictEqual(isStatementPrefix("    Dim "), true);
  assert.strictEqual(isStatementPrefix("    ' comment"), false);
  assert.strictEqual(isDocCommentSnippetPrefix("'''"), true);
  assert.strictEqual(isDocCommentSnippetPrefix("    '''"), true);
  assert.strictEqual(isDocCommentSnippetPrefix("    ''' summary"), false);
  assert.strictEqual(isAnnotationCommentPrefix("'@"), true);
  assert.strictEqual(isAnnotationCommentPrefix("    ' @Ex"), true);
  assert.strictEqual(isAnnotationCommentPrefix("'@ExpectedError(\"text@"), false);
  assert.strictEqual(isAnnotationCommentPrefix('Debug.Print "@'), false);
  const parsedFunction = parseProcedureDeclaration(
    "Public Function FindCustomer(ByVal customerCode As String) As Customer",
  );
  assert.ok(parsedFunction, "function declaration should parse");
  assert.deepStrictEqual(parsedFunction, {
    name: "FindCustomer",
    kind: "function",
    parameters: ["customerCode"],
  });
  assert.ok(
    documentationSnippet(parsedFunction).includes("customerCode: ${2:Parameter description.}"),
    "documentation snippet should include argument tab stop",
  );
  const parsedArrayProcedure = parseProcedureDeclaration(
    "Public Sub AcceptItems(arr() As String, x As Long)",
  );
  assert.deepStrictEqual(parsedArrayProcedure?.parameters, ["arr", "x"]);
  const parsedProperty = parseProcedureDeclaration(
    "Public Property Get CurrentCustomer() As Customer",
  );
  assert.strictEqual(parsedProperty?.kind, "property_get");
  assert.ok(
    parsedProperty !== undefined && documentationSnippet(parsedProperty).includes("Returns:"),
    "property get documentation snippet should include Returns",
  );
  const parsedIndexedProperty = parseProcedureDeclaration(
    "Public Property Get Item(index As Long) As Variant",
  );
  assert.ok(
    parsedIndexedProperty !== undefined &&
      documentationSnippet(parsedIndexedProperty).includes("index: ${2:Parameter description.}"),
    "indexed property get documentation snippet should include Args",
  );
  assert.strictEqual(isProgIdStringPrefix('Set app = CreateObject("Excel.'), true);
  assert.strictEqual(isProgIdStringPrefix('Debug.Print "Excel.'), false);
  assert.strictEqual(isVbaSourcePath("src/modules/Main.bas"), true);
  assert.strictEqual(isVbaSourcePath("src/classes/Invoice.cls"), true);
  assert.strictEqual(isVbaSourcePath("src/forms/Customer.frm"), true);
  assert.strictEqual(isVbaSourcePath("README.md"), false);
  assert.deepStrictEqual(
    vbaFilesAssociationUpdate({ "*.txt": "plaintext", "*.bas": "other-vba" }),
    { "*.txt": "plaintext", "*.bas": "vba", "*.cls": "vba", "*.frm": "vba" },
  );
  const mismatchedDocs = mismatchedVbaLanguageDocuments([
    { uri: vscode.Uri.file("C:/tmp/Main.bas"), languageId: "vb" },
    { uri: vscode.Uri.file("C:/tmp/Class.cls"), languageId: "vba" },
    { uri: vscode.Uri.file("C:/tmp/readme.md"), languageId: "markdown" },
  ]);
  assert.strictEqual(mismatchedDocs.length, 1);
  assert.strictEqual(mismatchedDocs[0].uri.fsPath, vscode.Uri.file("C:/tmp/Main.bas").fsPath);
  const lspProjectDir = fs.mkdtempSync(path.join(os.tmpdir(), "xlflow-vscode-lsp-"));
  try {
    fs.writeFileSync(path.join(lspProjectDir, "xlflow.toml"), '[project]\nname = "test"\n');
    const folder: vscode.WorkspaceFolder = {
      uri: vscode.Uri.file(lspProjectDir),
      name: "xlflow-vscode-lsp",
      index: 0,
    };
    assert.deepStrictEqual(
      await lspServerArgs({ lspLogFile: ".xlflow/lsp.log", lspLogFileConfigured: false }, folder),
      ["lsp", "--stdio", "--log-file", ".xlflow/lsp.log"],
    );
  } finally {
    fs.rmSync(lspProjectDir, { recursive: true, force: true });
  }
  assert.strictEqual(sessionStatusText("inactive"), "$(circle-slash) xlflow: No Session");
  assert.strictEqual(sessionStatusText("active"), "$(check) xlflow: Session Active");
  assert.ok(
    sessionQuickPickItems("inactive").some((item) => item.action === "attach"),
    "inactive session QuickPick should offer attach",
  );
  assert.ok(
    sessionQuickPickItems("active").some((item) => item.action === "attach"),
    "active session QuickPick should offer attach",
  );
  assert.strictEqual(
    sessionStatusText("inactive", "notInitialized"),
    "$(circle-slash) xlflow: No Project",
  );
  assert.deepStrictEqual(normalizeAvailabilitySuccess("xlflow", "xlflow 0.1.0\n", ""), {
    ok: true,
    executable: "xlflow",
    version: "xlflow 0.1.0",
  });
  assert.strictEqual(
    cliVersionSummary(
      "OK xlflow version\n\nVersion:       dev\nCommit:        none\nDate:          unknown",
    ),
    "dev",
  );
  assert.strictEqual(cliVersionSummary("xlflow 0.1.0\n"), "0.1.0");
  assert.strictEqual(cliVersionSummary(undefined), undefined);
  assert.deepStrictEqual(
    normalizeUpdateResult({
      exitCode: 0,
      stdout:
        '{"status":"ok","command":"update check","update":{"current_version":"1.2.3","latest_version":"v1.2.4","update_available":true,"release_url":"https://example.com/v1.2.4"}}',
      stderr: "",
    }),
    {
      kind: "available",
      info: {
        currentVersion: "1.2.3",
        latestVersion: "v1.2.4",
        updateAvailable: true,
        releaseUrl: "https://example.com/v1.2.4",
      },
    },
  );
  assert.deepStrictEqual(
    normalizeUpdateResult({
      exitCode: 0,
      stdout:
        '{"status":"ok","command":"update check","update":{"current_version":"1.2.3","update_available":false}}',
      stderr: "",
    }),
    {
      kind: "upToDate",
      info: {
        currentVersion: "1.2.3",
        latestVersion: undefined,
        updateAvailable: false,
        releaseUrl: undefined,
      },
    },
  );
  assert.strictEqual(
    normalizeUpdateResult({
      exitCode: 3,
      stdout:
        '{"status":"failed","error":{"message":"network down"},"update":{"current_version":"1.2.3","update_available":false}}',
      stderr: "",
    }).kind,
    "error",
  );
  assert.strictEqual(lastCheckedKey("xlflow", "1.2.3"), "xlflow.update.lastChecked.xlflow.1.2.3");
  assert.strictEqual(
    updateDismissedKey("xlflow", "1.2.3", "v1.2.4"),
    "xlflow.update.dismissed.xlflow.1.2.3.v1.2.4",
  );
  assert.strictEqual(shouldRunAutomaticCheck(1000, undefined), true);
  assert.strictEqual(shouldRunAutomaticCheck(1000, 999), false);
  assert.strictEqual(shouldRunAutomaticCheck(24 * 60 * 60 * 1000 + 1, 0), true);
  assert.strictEqual(
    updateSummary("xlflow 1.2.3", {
      kind: "available",
      info: {
        currentVersion: "1.2.3",
        latestVersion: "v1.2.4",
        updateAvailable: true,
      },
    }),
    "1.2.3 -> v1.2.4",
  );
  const missingAvailability = normalizeAvailabilityFailure("xlflow", { code: "ENOENT" });
  assert.strictEqual(missingAvailability.ok, false);
  assert.strictEqual(missingAvailability.reason, "notFound");
  assert.strictEqual(missingAvailability.executable, "xlflow");
  assert.ok(missingAvailability.message.length > 0);
  const timeoutAvailability = normalizeAvailabilityFailure("xlflow", { timedOut: true });
  assert.strictEqual(timeoutAvailability.ok, false);
  assert.strictEqual(timeoutAvailability.reason, "failed");
  assert.strictEqual(timeoutAvailability.executable, "xlflow");
  assert.ok(timeoutAvailability.message.length > 0);
  assert.deepStrictEqual(normalizeAvailabilityFailure("xlflow", { stderr: "boom" }), {
    ok: false,
    reason: "failed",
    executable: "xlflow",
    message: "boom",
  });
  assert.strictEqual(
    sessionStateFromEnvelope({ status: "ok", session: { active: true } }),
    "active",
  );
  assert.strictEqual(
    sessionStateFromEnvelope({ status: "ok", session: { active: false } }),
    "inactive",
  );
  assert.strictEqual(sessionStateFromEnvelope({ status: "failed" }), "error");
  assert.strictEqual(saveRequiredProjectNode(true).command?.command, "xlflow.saveWorkbook");
  assert.strictEqual(saveRequiredProjectNode(false).command, undefined);
  assert.strictEqual(
    readExcelPathFromToml('[project]\nname = "sample"\n[excel]\npath = "build/Book.xlsm"\n'),
    "build/Book.xlsm",
  );
  assert.strictEqual(
    readFormsRootFromToml('[src]\nforms = "custom/forms"\n[userform]\ncode_source = "frm"\n'),
    "custom/forms",
  );
  assert.strictEqual(readFormsRootFromToml("[src]\nforms = 'custom/forms'\n"), "custom/forms");
  assert.strictEqual(
    readFormsRootFromToml('[src]\nforms = "custom/#forms" # comment\n'),
    "custom/#forms",
  );
  assert.strictEqual(readFormsRootFromToml('[src]\nmodules = "src/modules"\n'), "src/forms");
  assert.strictEqual(readUserFormCodeSourceFromToml('[userform]\ncode_source = "frm"\n'), "frm");
  assert.strictEqual(readUserFormCodeSourceFromToml('[project]\nname = "sample"\n'), "sidecar");

  assert.deepStrictEqual(
    buildUserFormModels("src/forms", "sidecar", [
      "src/forms/Calendar.frm",
      "src/forms/code/Calendar.bas",
      "src/forms/specs/Calendar.yml",
      "src/forms/specs/Calendar.yaml",
    ]),
    [
      {
        name: "Calendar",
        codeSource: "sidecar",
        artifacts: [
          {
            kind: "code",
            label: "Code: code/Calendar.bas",
            relativePath: "src/forms/code/Calendar.bas",
            missing: false,
          },
          {
            kind: "spec",
            label: "Spec: specs/Calendar.yaml",
            relativePath: "src/forms/specs/Calendar.yaml",
            missing: false,
          },
        ],
      },
    ],
  );
  assert.deepStrictEqual(buildUserFormModels("src/forms", "sidecar", ["src/forms/Only.frm"]), [
    {
      name: "Only",
      codeSource: "sidecar",
      artifacts: [
        { kind: "code", label: "Code", relativePath: undefined, missing: true },
        { kind: "spec", label: "Spec", relativePath: undefined, missing: true },
      ],
    },
  ]);
  assert.deepStrictEqual(
    buildUserFormModels("src/forms", "frm", [
      "src/forms/Legacy.frm",
      "src/forms/Sales/OrderForm.frm",
      "src/forms/code/Legacy.bas",
      "src/forms/specs/Legacy.yaml",
    ]),
    [
      {
        name: "Legacy",
        codeSource: "frm",
        artifacts: [
          {
            kind: "frm",
            label: "Legacy.frm",
            relativePath: "src/forms/Legacy.frm",
            missing: false,
          },
        ],
      },
      {
        name: "OrderForm",
        codeSource: "frm",
        artifacts: [
          {
            kind: "frm",
            label: "Sales/OrderForm.frm",
            relativePath: "src/forms/Sales/OrderForm.frm",
            missing: false,
          },
        ],
      },
    ],
  );
  const fakeFolder = {
    uri: vscode.Uri.file("C:/tmp/xlflow"),
    name: "xlflow",
    index: 0,
  };
  assert.strictEqual(
    cliNotificationSuppressionKey(fakeFolder.uri, {
      ok: false,
      reason: "notFound",
      executable: "xlflow",
      message: "missing",
    }),
    `${"xlflow.cliMissingNotice"}.${fakeFolder.uri.toString()}.xlflow`,
  );
  assert.strictEqual(
    comparableFsPath(sourceUri(fakeFolder, "src/modules/Main.bas")),
    comparableFsPath("C:/tmp/xlflow/src/modules/Main.bas"),
  );
  assert.strictEqual(
    comparableFsPath(sourceUri(fakeFolder, "src\\classes\\Invoice.cls")),
    comparableFsPath("C:/tmp/xlflow/src/classes/Invoice.cls"),
  );
  assert.strictEqual(
    comparableFsPath(sourceUri(fakeFolder, "C:\\work\\project\\src\\modules\\Main.bas")),
    comparableFsPath(vscode.Uri.file("C:\\work\\project\\src\\modules\\Main.bas")),
  );
  assert.strictEqual(
    discoveredTestDescription({
      status_hint: "skipped",
      skip: { reason: "Requires Access" },
      tags: ["integration"],
    }),
    "skipped: Requires Access | integration",
  );
  assert.strictEqual(
    discoveredTestDescription({
      status_hint: "todo",
      todo: {},
    }),
    "todo",
  );
  assert.deepStrictEqual(
    moduleGroups(fakeFolder, {
      inspect: {
        files: [
          { path: "src/forms/code/Form1.bas", moduleName: "Form1", moduleKind: "form" },
          { path: "src/modules/Main.bas", moduleName: "Main", moduleKind: "standard" },
          { path: "src/classes/Invoice.cls", moduleName: "Invoice", moduleKind: "class" },
          {
            path: "src/workbook/ThisWorkbook.cls",
            moduleName: "ThisWorkbook",
            moduleKind: "document",
          },
        ],
      },
    }).map((group) => group.label),
    ["Standard Modules", "Class Modules", "Document Modules"],
  );
  assert.strictEqual(moduleContextValue("standard"), "xlflow.module.standard");
  assert.strictEqual(moduleContextValue("class"), "xlflow.module.class");
  assert.strictEqual(moduleContextValue("document"), "xlflow.module.document");
  assert.strictEqual(userFormContextValue("sidecar"), "xlflow.userForm.sidecar");
  assert.strictEqual(userFormContextValue("frm"), "xlflow.userForm.frm");
  assert.strictEqual(
    userFormArtifactContextValue({ artifactKind: "code", missing: false }),
    "xlflow.userFormArtifact.code",
  );
  assert.strictEqual(
    userFormArtifactContextValue({ artifactKind: "spec", missing: true }),
    "xlflow.userFormMissingArtifact.spec",
  );
  const formulaManifest = JSON.stringify({
    version: 1,
    sheets: [
      { path: "sheets/001-Invoice.regions.jsonl" },
      { path: "sheets/002-Summary.regions.jsonl" },
    ],
  });
  const formulaNodes = buildFormulaSnapshotNodes(fakeFolder, formulaManifest, true);
  assert.deepStrictEqual(
    formulaNodes.map((node) => node.kind),
    ["formulaFile", "formulaGroup"],
  );
  assert.strictEqual(
    formulaNodes[0].kind === "formulaFile" ? formulaNodes[0].label : undefined,
    "names.jsonl",
  );
  assert.deepStrictEqual(
    formulaNodes[1].kind === "formulaGroup"
      ? formulaNodes[1].children.map((node) => node.label)
      : [],
    ["001-Invoice.regions.jsonl", "002-Summary.regions.jsonl"],
  );
  assert.deepStrictEqual(
    buildFormulaSnapshotNodes(fakeFolder, formulaManifest, false).map((node) => node.kind),
    ["formulaGroup"],
  );
  assert.deepStrictEqual(
    buildFormulaSnapshotNodes(fakeFolder, undefined, false).map((node) => node.kind),
    ["formulaEmpty"],
  );
  assert.deepStrictEqual(
    buildFormulaSnapshotNodes(fakeFolder, "{not-json", false).map((node) => node.kind),
    ["formulaEmpty"],
  );
  assert.deepStrictEqual(
    buildFormulaSnapshotNodes(fakeFolder, JSON.stringify({ version: 2, sheets: [] }), false).map(
      (node) => node.kind,
    ),
    ["formulaEmpty"],
  );
}

function assertLocalizationResources(extensionPath: string): void {
  const manifest = readJson<Record<string, unknown>>(path.join(extensionPath, "package.json"));
  const packageNls = readJson<Record<string, string>>(path.join(extensionPath, "package.nls.json"));
  const packageNlsJa = readJson<Record<string, string>>(
    path.join(extensionPath, "package.nls.ja.json"),
  );
  assert.strictEqual(manifest.l10n, "./l10n");
  assert.strictEqual(manifest.displayName, "%extension.displayName%");
  assert.strictEqual(manifest.description, "%extension.description%");
  assert.strictEqual(
    readPath(manifest, ["contributes", "commands", 0, "title"]),
    "%command.restartLanguageServer.title%",
  );
  assert.strictEqual(
    readPath(manifest, ["contributes", "views", "xlflow", 0, "name"]),
    "%view.setup.name%",
  );
  assert.strictEqual(
    menuWhen(manifest, "view/title", "xlflow.sessionStart"),
    "view == xlflow.project && xlflow.sessionStartEnabled",
  );
  assert.strictEqual(
    menuWhen(manifest, "view/title", "xlflow.sessionStop"),
    "view == xlflow.project && xlflow.sessionStopEnabled",
  );
  assert.ok(
    hasMenuItem(manifest, "view/title", "xlflow.saveWorkbook", {
      when: "view == xlflow.project && xlflow.saveRequired",
      group: "navigation@6",
    }),
    "project view title menu should contribute xlflow.saveWorkbook only when save is required",
  );
  assert.ok(
    hasMenuItem(manifest, "view/title", "xlflow.rollbackWorkbook", {
      when: "view == xlflow.project",
      group: "navigation@4",
    }),
    "project view title menu should contribute xlflow.rollbackWorkbook",
  );
  assert.ok(
    hasMenuItem(manifest, "view/title", "xlflow.pruneBackups", {
      when: "view == xlflow.project",
      group: "navigation@5",
    }),
    "project view title menu should contribute xlflow.pruneBackups",
  );
  assert.ok(
    hasView(manifest, "xlflow.formulas", "%view.formulas.name%"),
    "xlflow.formulas view should be contributed",
  );
  assert.ok(
    hasMenuItem(manifest, "view/title", "xlflow.pullFormulas", {
      when: "view == xlflow.formulas",
      group: "navigation@1",
    }),
    "formulas view title menu should contribute xlflow.pullFormulas",
  );
  assert.ok(
    hasMenuItem(manifest, "view/title", "xlflow.refreshFormulas", {
      when: "view == xlflow.formulas",
      group: "navigation@2",
    }),
    "formulas view title menu should contribute xlflow.refreshFormulas",
  );
  assert.ok(
    hasMenuItem(manifest, "view/item/context", "xlflow.formatDocument", {
      when: "view == xlflow.modules && viewItem =~ /^xlflow\\.module\\.(standard|class|document)$/",
      group: "2_workspace@0",
    }),
    "module context menu should contribute xlflow.formatDocument",
  );
  assert.ok(
    hasMenuItem(manifest, "view/item/context", "xlflow.openFormulaSnapshotFile", {
      when: "view == xlflow.formulas && viewItem == xlflow.formulaFile",
      group: "navigation@1",
    }),
    "formula file context menu should contribute open action",
  );
  assert.ok(
    hasMenuItem(manifest, "view/item/context", "xlflow.revealFormulaSnapshotFile", {
      when: "view == xlflow.formulas && viewItem == xlflow.formulaFile",
      group: "3_file@1",
    }),
    "formula file context menu should contribute reveal action",
  );
  const placeholders = collectManifestPlaceholders(manifest);
  for (const key of placeholders) {
    assert.ok(packageNls[key], `package.nls.json should define ${key}`);
    assert.ok(packageNlsJa[key], `package.nls.ja.json should define ${key}`);
  }
  assert.deepStrictEqual(
    Object.keys(packageNlsJa).sort(),
    Object.keys(packageNls).sort(),
    "package nls key sets should match",
  );

  const bundle = readJson<Record<string, string>>(
    path.join(extensionPath, "l10n", "bundle.l10n.json"),
  );
  const bundleJa = readJson<Record<string, string>>(
    path.join(extensionPath, "l10n", "bundle.l10n.ja.json"),
  );
  assert.deepStrictEqual(
    Object.keys(bundleJa).sort(),
    Object.keys(bundle).sort(),
    "runtime l10n bundle key sets should match",
  );
}

function readJson<T>(filePath: string): T {
  return JSON.parse(fs.readFileSync(filePath, "utf8")) as T;
}

function collectManifestPlaceholders(value: unknown): string[] {
  const keys = new Set<string>();
  const visit = (candidate: unknown): void => {
    if (typeof candidate === "string") {
      const match = candidate.match(/^%([^%]+)%$/);
      if (match !== null) {
        keys.add(match[1]);
      }
      return;
    }
    if (Array.isArray(candidate)) {
      for (const item of candidate) {
        visit(item);
      }
      return;
    }
    if (typeof candidate === "object" && candidate !== null) {
      for (const item of Object.values(candidate)) {
        visit(item);
      }
    }
  };
  visit(value);
  return [...keys].sort();
}

function readPath(value: unknown, parts: Array<string | number>): unknown {
  let current = value;
  for (const part of parts) {
    if (typeof current !== "object" || current === null) {
      return undefined;
    }
    current = (current as Record<string | number, unknown>)[part];
  }
  return current;
}

function menuWhen(manifest: Record<string, unknown>, menu: string, command: string): unknown {
  const items = readPath(manifest, ["contributes", "menus", menu]);
  assert.ok(Array.isArray(items), `${menu} should be an array`);
  const item = items.find(
    (candidate) =>
      typeof candidate === "object" &&
      candidate !== null &&
      (candidate as Record<string, unknown>).command === command,
  );
  assert.ok(item, `${command} should be contributed to ${menu}`);
  return (item as Record<string, unknown>).when;
}

function hasView(manifest: Record<string, unknown>, id: string, name: string): boolean {
  const views = readPath(manifest, ["contributes", "views", "xlflow"]);
  assert.ok(Array.isArray(views), "xlflow views should be an array");
  return views.some((candidate) => {
    if (typeof candidate !== "object" || candidate === null) {
      return false;
    }
    const view = candidate as Record<string, unknown>;
    return view.id === id && view.name === name;
  });
}

function hasMenuItem(
  manifest: Record<string, unknown>,
  menu: string,
  command: string,
  expected: Record<string, unknown>,
): boolean {
  const items = readPath(manifest, ["contributes", "menus", menu]);
  assert.ok(Array.isArray(items), `${menu} should be an array`);
  return items.some((candidate) => {
    if (typeof candidate !== "object" || candidate === null) {
      return false;
    }
    const item = candidate as Record<string, unknown>;
    return (
      item.command === command &&
      Object.entries(expected).every(([key, value]) => item[key] === value)
    );
  });
}

function comparableFsPath(value: vscode.Uri | string | undefined): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  const fsPath =
    typeof value === "string" && /^[A-Za-z]:[\\/]/.test(value)
      ? vscode.Uri.file(value).fsPath
      : typeof value === "string"
        ? value
        : value.fsPath;
  return path.normalize(fsPath).toLowerCase();
}

function pad2(value: number): string {
  return String(value).padStart(2, "0");
}
