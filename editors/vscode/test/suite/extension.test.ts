import * as assert from "assert";
import * as vscode from "vscode";
import { sessionStateFromEnvelope, sessionStatusText } from "../../src/session";

export async function run(): Promise<void> {
  const config = vscode.workspace.getConfiguration("xlflow");
  await config.update("lsp.enabled", false, vscode.ConfigurationTarget.Global);

  const extension = vscode.extensions.getExtension("harumiweb.xlflow-vscode");
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
    "xlflow.pull",
    "xlflow.push",
    "xlflow.runMacro",
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
  ]) {
    assert.ok(commands.includes(command), `${command} should be registered`);
  }

  assert.strictEqual(config.get<string>("path"), "xlflow");
  assert.strictEqual(sessionStatusText("inactive"), "$(circle-slash) xlflow");
  assert.strictEqual(sessionStatusText("active"), "$(check) xlflow: Session");
  assert.strictEqual(
    sessionStateFromEnvelope({ status: "ok", session: { active: true } }),
    "active",
  );
  assert.strictEqual(
    sessionStateFromEnvelope({ status: "ok", session: { active: false } }),
    "inactive",
  );
  assert.strictEqual(sessionStateFromEnvelope({ status: "failed" }), "error");
}
