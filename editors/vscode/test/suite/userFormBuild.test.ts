import * as assert from "assert";
import * as fs from "fs/promises";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";
import { discoverFormBuildInputs } from "../../src/userFormBuild";
import { buildUserFormModels, userFormContextValue } from "../../src/sidebar";

export async function runUserFormBuildAssertions(): Promise<void> {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "xlflow-form-build-"));
  const folder: vscode.WorkspaceFolder = {
    uri: vscode.Uri.file(root),
    name: "form-build",
    index: 0,
  };
  const projectRoot = folder.uri.fsPath;
  try {
    const formsRoot = vscode.Uri.joinPath(folder.uri, "custom", "forms").fsPath;
    await fs.mkdir(path.join(formsRoot, "specs", "nested"), { recursive: true });
    await fs.mkdir(path.join(formsRoot, "code"), { recursive: true });
    await fs.mkdir(path.join(formsRoot, "legacy"), { recursive: true });
    await fs.mkdir(path.join(projectRoot, "src", "forms", "specs"), { recursive: true });
    const config = '[excel]\npath = "build/Target.xlsm"\n[src]\nforms = "custom/forms"\n';
    await fs.writeFile(path.join(projectRoot, "xlflow.toml"), config);
    for (const spec of ["Alpha.yaml", "Beta.json", "Gamma.yml"]) {
      await fs.writeFile(path.join(formsRoot, "specs", spec), "form: {}\n");
    }
    await fs.writeFile(path.join(formsRoot, "specs", "nested", "Ignore.yaml"), "form: {}\n");
    await fs.writeFile(
      path.join(projectRoot, "src", "forms", "specs", "WrongRoot.yaml"),
      "form: {}\n",
    );
    await fs.writeFile(path.join(formsRoot, "code", "Alpha.bas"), "' sidecar code");
    await fs.writeFile(path.join(formsRoot, "code", "NoSpec.bas"), "' no spec");
    await fs.writeFile(
      path.join(formsRoot, "legacy", "Alpha.frm"),
      "' form build never consumes this code",
    );
    await fs.writeFile(path.join(formsRoot, "NoSpec.frm"), "' no spec");
    const sidecar = await discoverFormBuildInputs(folder);
    assert.strictEqual(sidecar.workbook, path.join(projectRoot, "build", "Target.xlsm"));
    assert.strictEqual(sidecar.formsRoot.fsPath, formsRoot);
    assert.deepStrictEqual(
      sidecar.sources.map((source) => source.name),
      ["Alpha", "Beta", "Gamma"],
    );
    assert.deepStrictEqual(sidecar.duplicates, []);
    assert.strictEqual(sidecar.excluded, 1);
    assert.deepStrictEqual(sidecar.codeByName.alpha, [path.join(formsRoot, "code", "Alpha.bas")]);
    assert.deepStrictEqual(sidecar.codeByName.beta, []);
    await fs.writeFile(
      path.join(projectRoot, "xlflow.toml"),
      config + '[userform]\ncode_source = "frm"\n',
    );
    const legacy = await discoverFormBuildInputs(folder);
    assert.deepStrictEqual(legacy.codeByName.alpha, [path.join(formsRoot, "code", "Alpha.bas")]);
    assert.deepStrictEqual(legacy.codeByName.beta, []);
    assert.strictEqual(legacy.excluded, 1);
    assert.deepStrictEqual(
      legacy.sources.map((source) => source.name),
      ["Alpha", "Beta", "Gamma"],
    );
    // Rename through an intermediate path to work on case-insensitive Windows filesystems.
    await fs.rename(
      path.join(formsRoot, "code", "Alpha.bas"),
      path.join(formsRoot, "code", "case-change.tmp"),
    );
    await fs.rename(
      path.join(formsRoot, "code", "case-change.tmp"),
      path.join(formsRoot, "code", "ALPHA.BAS"),
    );
    const uppercase = await discoverFormBuildInputs(folder);
    assert.deepStrictEqual(uppercase.codeByName.alpha, [path.join(formsRoot, "code", "ALPHA.BAS")]);
    assert.deepStrictEqual(uppercase.codeByName.beta, []);
    await fs.unlink(path.join(formsRoot, "code", "ALPHA.BAS"));
    const withoutSidecar = await discoverFormBuildInputs(folder);
    assert.deepStrictEqual(
      withoutSidecar.codeByName.alpha,
      [],
      "An existing .frm must never become the build's code save target",
    );
    await fs.unlink(path.join(formsRoot, "code", "NoSpec.bas"));
    await fs.rmdir(path.join(formsRoot, "code"));
    const withoutCodeDirectory = await discoverFormBuildInputs(folder);
    assert.deepStrictEqual(withoutCodeDirectory.codeByName.alpha, []);
    assert.deepStrictEqual(withoutCodeDirectory.codeByName.beta, []);
    await fs.writeFile(path.join(formsRoot, "specs", "ALPHA.JSON"), "{}");
    const duplicate = await discoverFormBuildInputs(folder);
    assert.deepStrictEqual(duplicate.duplicates[0].map((file) => file.toLowerCase()).sort(), [
      "alpha.json",
      "alpha.yaml",
    ]);
    assert.deepStrictEqual(
      duplicate.sources.map((source) => source.name),
      ["Beta", "Gamma"],
    );
    await fs.writeFile(path.join(projectRoot, "xlflow.toml"), "[src]\nforms = 'custom/forms'\n");
    await assert.rejects(discoverFormBuildInputs(folder), /\[excel\]\.path/);
    assert.strictEqual(userFormContextValue("sidecar", true), "xlflow.userForm.sidecar.buildable");
    assert.strictEqual(userFormContextValue("sidecar", false), "xlflow.userForm.sidecar");
    const json = buildUserFormModels("custom/forms", "sidecar", [
      "custom/forms/specs/OnlyJson.json",
    ]);
    assert.strictEqual(json.length, 1);
    assert.strictEqual(json[0].name, "OnlyJson");
    assert.strictEqual(
      json[0].artifacts.find((artifact) => artifact.kind === "spec")?.missing,
      false,
    );
  } finally {
    // mkdtemp returns an owned directory; verify its resolved boundary before recursive removal.
    const relative = path.relative(path.resolve(os.tmpdir()), path.resolve(root));
    assert.ok(relative && !relative.startsWith("..") && !path.isAbsolute(relative));
    assert.ok(path.basename(root).startsWith("xlflow-form-build-"));
    await fs.rm(root, { recursive: true, force: true });
  }
}
