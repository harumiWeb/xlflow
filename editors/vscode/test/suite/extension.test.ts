import * as assert from "assert";
import * as fs from "fs";
import * as path from "path";
import * as vscode from "vscode";
import {
  cliNotificationSuppressionKey,
  normalizeAvailabilityFailure,
  normalizeAvailabilitySuccess,
} from "../../src/cliAvailability";
import { sessionStateFromEnvelope, sessionStatusText } from "../../src/session";
import {
  buildUserFormModels,
  cliVersionSummary,
  moduleContextValue,
  moduleGroups,
  readExcelPathFromToml,
  readFormsRootFromToml,
  readUserFormCodeSourceFromToml,
  userFormArtifactContextValue,
  userFormContextValue,
} from "../../src/sidebar";
import { sourceUri } from "../../src/testDiscovery";

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

  const languages = await vscode.languages.getLanguages();
  assert.ok(languages.includes("vba"), "vba language should be registered");

  const commands = await vscode.commands.getCommands(true);
  for (const command of [
    "xlflow.restartLanguageServer",
    "xlflow.checkEnvironment",
    "xlflow.openInstallGuide",
    "xlflow.configurePath",
    "xlflow.retryCliDetection",
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
  assert.deepStrictEqual(normalizeAvailabilityFailure("xlflow", { code: "ENOENT" }), {
    ok: false,
    reason: "notFound",
    executable: "xlflow",
    message: "xlflow executable was not found.",
  });
  assert.deepStrictEqual(normalizeAvailabilityFailure("xlflow", { timedOut: true }), {
    ok: false,
    reason: "failed",
    executable: "xlflow",
    message: "xlflow version timed out.",
  });
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
    comparableFsPath("C:\\work\\project\\src\\modules\\Main.bas"),
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
    hasMenuItem(manifest, "view/item/context", "xlflow.formatDocument", {
      when: "view == xlflow.modules && viewItem =~ /^xlflow\\.module\\.(standard|class|document)$/",
      group: "2_workspace@0",
    }),
    "module context menu should contribute xlflow.formatDocument",
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
  const fsPath = typeof value === "string" ? value : value.fsPath;
  return path.normalize(fsPath).toLowerCase();
}
