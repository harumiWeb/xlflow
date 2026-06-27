import * as assert from "assert";
import * as vscode from "vscode";
import { sessionStateFromEnvelope, sessionStatusText } from "../../src/session";
import {
  buildUserFormModels,
  moduleGroups,
  readExcelPathFromToml,
  readFormsRootFromToml,
  readUserFormCodeSourceFromToml,
} from "../../src/sidebar";

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

  const languages = await vscode.languages.getLanguages();
  assert.ok(languages.includes("vba"), "vba language should be registered");

  const commands = await vscode.commands.getCommands(true);
  for (const command of [
    "xlflow.restartLanguageServer",
    "xlflow.checkEnvironment",
    "xlflow.newProject",
    "xlflow.initProject",
    "xlflow.skillInstall",
    "xlflow.moduleInstall",
    "xlflow.newModule",
    "xlflow.newStandardModule",
    "xlflow.newClassModule",
    "xlflow.newUserForm",
    "xlflow.pull",
    "xlflow.push",
    "xlflow.runMacro",
    "xlflow.runProcedure",
    "xlflow.runTestProcedure",
    "xlflow.test",
    "xlflow.lintWorkspace",
    "xlflow.formatDocument",
    "xlflow.formatProject",
    "xlflow.saveWorkbook",
    "xlflow.sessionStart",
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
    "xlflow.runAllTests",
    "xlflow.runDoctor",
    "xlflow.sessionToggle",
    "xlflow.setupActions",
    "xlflow.openDocumentation",
    "xlflow.openModule",
    "xlflow.openProcedure",
    "xlflow.openUserFormArtifact",
  ]) {
    assert.ok(commands.includes(command), `${command} should be registered`);
  }

  assert.strictEqual(config.get<string>("path"), "xlflow");
  assert.strictEqual(sessionStatusText("inactive"), "$(circle-slash) xlflow: No Session");
  assert.strictEqual(sessionStatusText("active"), "$(check) xlflow: Session Active");
  assert.strictEqual(
    sessionStatusText("inactive", "notInitialized"),
    "$(circle-slash) xlflow: No Project",
  );
  assert.strictEqual(
    sessionStateFromEnvelope({ status: "ok", session: { active: true } }),
    "active",
  );
  assert.strictEqual(
    sessionStateFromEnvelope({ status: "ok", session: { active: false } }),
    "inactive",
  );
  assert.strictEqual(sessionStateFromEnvelope({ status: "failed" }), "error");
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
  assert.deepStrictEqual(
    moduleGroups(fakeFolder, {
      inspect: {
        files: [
          { path: "src/forms/code/Form1.bas", moduleName: "Form1", moduleKind: "form" },
          { path: "src/modules/Main.bas", moduleName: "Main", moduleKind: "standard" },
        ],
      },
    }).map((group) => group.label),
    ["Standard Modules"],
  );
}
