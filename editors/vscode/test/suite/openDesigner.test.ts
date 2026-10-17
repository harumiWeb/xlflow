import * as assert from "assert";
import * as vscode from "vscode";
import { isDesignerSource, openDesignerCommand } from "../../src/userFormEditor/openDesigner";
import { designerResourceUri, designerSourceUri } from "../../src/userFormEditor/resource";

export async function runOpenDesignerTests(): Promise<void> {
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder, "the native test workspace is required");
  const created: vscode.Uri[] = [];
  const configUri = vscode.Uri.joinPath(folder.uri, "xlflow.toml");
  const originalConfig = await vscode.workspace.fs.readFile(configUri);
  const customSpecsDirectory = vscode.Uri.joinPath(folder.uri, "custom", "forms", "specs");
  let customSpecsDirectoryCreated = false;
  try {
    for (const extension of ["yaml", "yml", "json"]) {
      const uri = vscode.Uri.joinPath(
        folder.uri,
        "src",
        "forms",
        "specs",
        `ContextMenu.${extension}`,
      );
      created.push(uri);
      const text =
        extension === "json"
          ? '{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"ContextMenu"},"controls":[]}\n'
          : "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: ContextMenu\ncontrols: []\n";
      await vscode.workspace.fs.writeFile(uri, Buffer.from(text));
      const document = await vscode.workspace.openTextDocument(uri);
      await vscode.window.showTextDocument(document);
      const edit = new vscode.WorkspaceEdit();
      edit.insert(
        uri,
        document.positionAt(text.length),
        extension === "json" ? " " : "# unsaved source\n",
      );
      assert.ok(await vscode.workspace.applyEdit(edit));
      const dirtyText = document.getText();
      assert.ok(await isDesignerSource(uri));
      assert.strictEqual(
        designerSourceUri(designerResourceUri(document.uri)).toString(),
        document.uri.toString(),
      );
      await vscode.commands.executeCommand(openDesignerCommand, uri);
      const matches = () =>
        vscode.window.tabGroups.all
          .flatMap((group) => group.tabs)
          .filter(
            (tab) =>
              tab.input instanceof vscode.TabInputCustom &&
              tab.input.uri.toString() === designerResourceUri(document.uri).toString(),
          );
      assert.strictEqual(
        matches().length,
        1,
        "the context command opens the dedicated Designer resource",
      );
      await vscode.commands.executeCommand(openDesignerCommand, uri);
      assert.strictEqual(matches().length, 1, "repeated opening reuses the Designer");
      const tab = matches()[0];
      assert.ok(tab.input instanceof vscode.TabInputCustom);
      assert.strictEqual(designerSourceUri(tab.input.uri).toString(), document.uri.toString());
      assert.strictEqual(document.getText(), dirtyText);
      assert.ok(document.isDirty, "opening must not save source changes");
      assert.strictEqual(
        Buffer.from(await vscode.workspace.fs.readFile(uri)).toString("utf8"),
        text,
      );
      // Restore the text without publishing, then clean up only this test's tabs.
      const restore = new vscode.WorkspaceEdit();
      restore.replace(
        uri,
        new vscode.Range(document.positionAt(0), document.positionAt(document.getText().length)),
        text,
      );
      await vscode.workspace.applyEdit(restore);
      await document.save();
    }
    const defaultUri = vscode.Uri.joinPath(folder.uri, "src", "forms", "specs", "DefaultOpen.yaml");
    created.push(defaultUri);
    await vscode.workspace.fs.writeFile(
      defaultUri,
      Buffer.from(
        "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: DefaultOpen\ncontrols: []\n",
      ),
    );
    await vscode.commands.executeCommand("vscode.open", defaultUri);
    const defaultTab = vscode.window.tabGroups.activeTabGroup.activeTab;
    assert.ok(defaultTab?.input instanceof vscode.TabInputCustom);
    assert.strictEqual(
      defaultTab.input.viewType,
      "xlflow.userFormDesigner",
      "opening a standard FormSpec uses the Designer by default",
    );

    const unrelatedUri = vscode.Uri.joinPath(folder.uri, "other.yaml");
    created.push(unrelatedUri);
    await vscode.workspace.fs.writeFile(unrelatedUri, Buffer.from("name: unrelated\n"));
    await vscode.commands.executeCommand("vscode.open", unrelatedUri);
    assert.strictEqual(
      vscode.window.activeTextEditor?.document.uri.toString(),
      unrelatedUri.toString(),
      "an unrelated YAML file keeps the text editor as its default",
    );
    assert.strictEqual(
      await isDesignerSource(vscode.Uri.joinPath(folder.uri, "other.yaml")),
      false,
    );
    assert.strictEqual(
      await isDesignerSource(
        vscode.Uri.joinPath(folder.uri, "src", "forms", "specs", "nested", "Other.yaml"),
      ),
      false,
    );
    assert.strictEqual(await isDesignerSource(vscode.Uri.parse("untitled:Main.yaml")), false);

    await vscode.workspace.fs.writeFile(
      configUri,
      Buffer.from('[project]\nname = "vscode-clipboard-test"\n\n[src]\nforms = "custom/forms"\n'),
    );
    await vscode.workspace.fs.createDirectory(customSpecsDirectory);
    customSpecsDirectoryCreated = true;
    const customRootUri = vscode.Uri.joinPath(customSpecsDirectory, "CustomRoot.yaml");
    created.push(customRootUri);
    await vscode.workspace.fs.writeFile(
      customRootUri,
      Buffer.from(
        "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: CustomRoot\ncontrols: []\n",
      ),
    );
    assert.ok(await isDesignerSource(customRootUri));
    await vscode.commands.executeCommand("vscode.open", customRootUri);
    assert.strictEqual(
      vscode.window.activeTextEditor?.document.uri.toString(),
      customRootUri.toString(),
      "a configured nonstandard forms root keeps the text editor as its default",
    );
    await vscode.commands.executeCommand(
      "vscode.openWith",
      customRootUri,
      "xlflow.userFormDesigner",
    );
    const customRootTab = vscode.window.tabGroups.activeTabGroup.activeTab;
    assert.ok(customRootTab?.input instanceof vscode.TabInputCustom);
    assert.strictEqual(
      customRootTab.input.viewType,
      "xlflow.userFormDesigner",
      "the Designer remains available for a configured nonstandard forms root",
    );
  } finally {
    const sourceKeys = new Set(
      created.flatMap((uri) => [
        uri.toString(),
        designerResourceUri(vscode.Uri.parse(uri.toString())).toString(),
      ]),
    );
    for (const group of vscode.window.tabGroups.all) {
      const tabs = group.tabs.filter((tab) => {
        if (
          (tab.input instanceof vscode.TabInputCustom ||
            tab.input instanceof vscode.TabInputText) &&
          sourceKeys.has(tab.input.uri.toString())
        )
          return true;
        return false;
      });
      if (tabs.length > 0) {
        await vscode.window.tabGroups.close(tabs, true);
      }
    }
    for (const uri of created) await vscode.workspace.fs.delete(uri);
    if (customSpecsDirectoryCreated)
      await vscode.workspace.fs.delete(customSpecsDirectory, { recursive: true });
    await vscode.workspace.fs.writeFile(configUri, originalConfig);
  }
}
