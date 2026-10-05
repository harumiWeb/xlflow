import * as assert from "assert";
import * as vscode from "vscode";
import { mkdtemp, readFile, unlink, rmdir, writeFile } from "fs/promises";
import * as os from "os";
import * as path from "path";
import { designerHTML, designerViewType } from "../../src/userFormEditor/provider";
import { designerDocument } from "../../src/userFormEditor/model";
import { designerStrings } from "../../src/userFormEditor/protocol";
import type { HostMessage } from "../../src/userFormEditor/protocol";

// Exercise the shipped renderer inside a real VS Code Webview, including CSP
// enforcement and CSSOM geometry. The harness is only added to test HTML.
export async function runUserFormWebviewAssertions(extensionUri: vscode.Uri): Promise<void> {
  const panel = vscode.window.createWebviewPanel(
    "xlflow.designerTest",
    "Designer test",
    vscode.ViewColumn.One,
    {},
  );
  const resources = vscode.Uri.joinPath(extensionUri, "dist", "userFormDesigner");
  panel.webview.options = { enableScripts: true, localResourceRoots: [resources] };
  const html = designerHTML(panel.webview, resources);
  assert.ok(html.includes("default-src 'none'"));
  assert.ok(!html.includes("unsafe-inline"));
  const nonce = html.match(/script-src 'nonce-([^']+)'/)![1];
  const harness = `<script nonce="${nonce}">
const testApi = acquireVsCodeApi();
window.acquireVsCodeApi = () => testApi;
const violations = [];
window.addEventListener("securitypolicyviolation", e => violations.push(e.violatedDirective));
window.addEventListener("message", e => {
  if (e.data.type !== "document" && e.data.type !== "invalidDocument") return;
  setTimeout(() => {
    const label = document.querySelector(".control.label");
    const buttonCenters = [...document.querySelectorAll(".commandbutton, .togglebutton")].map(button => {
      const bounds = button.getBoundingClientRect();
      const caption = button.firstElementChild.getBoundingClientRect();
      return {
        type: button.className,
        delta: (caption.top + caption.bottom - bounds.top - bounds.bottom) / 2,
        shadow: getComputedStyle(button).boxShadow,
        topColor: getComputedStyle(button).borderTopColor,
        bottomColor: getComputedStyle(button).borderBottomColor
      };
    });
    const combo = document.querySelector(".combobox");
    const arrow = document.querySelector(".combo-arrow");
    const comboBounds = combo?.getBoundingClientRect();
    const arrowBounds = arrow?.getBoundingClientRect();
    testApi.postMessage({
      type: "testSnapshot", version: e.data.version,
      count: document.querySelectorAll(".control").length,
      title: document.querySelector(".form-title")?.textContent,
      header: document.querySelector("header")?.textContent,
      error: document.querySelector('[role="alert"]')?.textContent,
      left: label && getComputedStyle(label).left,
      width: label && getComputedStyle(label).width,
      images: document.querySelectorAll("img").length,
      buttonCenters,
      combo: comboBounds && arrowBounds && {
        height: comboBounds.height,
        arrowHeight: arrowBounds.height,
        arrowWidth: arrowBounds.width,
        rightInset: comboBounds.right - arrowBounds.right,
        shadow: getComputedStyle(combo).boxShadow,
        arrowShadow: getComputedStyle(arrow).boxShadow
      },
      violations
    });
  }, 100);
});
</script>`;
  const pending = new Map<number, (snapshot: Record<string, unknown>) => void>();
  let ready!: () => void;
  const readiness = new Promise<void>((resolve) => {
    ready = resolve;
  });
  const subscription = panel.webview.onDidReceiveMessage((message) => {
    if (message.type === "ready") ready();
    if (message.type === "testSnapshot") pending.get(message.version)?.(message);
  });
  const timeout = <T>(promise: Promise<T>): Promise<T> =>
    new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error("Designer Webview test timed out")), 15000);
      promise.then(
        (value) => {
          clearTimeout(timer);
          resolve(value);
        },
        (error) => {
          clearTimeout(timer);
          reject(error);
        },
      );
    });
  const snapshot = async (message: Extract<HostMessage, { version: number }>) => {
    const response = new Promise<Record<string, unknown>>((resolve) =>
      pending.set(message.version, resolve),
    );
    await panel.webview.postMessage(message);
    const result = await timeout(response);
    pending.delete(message.version);
    return result;
  };
  try {
    panel.webview.html = html.replace("<body>", "<body>" + harness);
    await timeout(readiness);
    await panel.webview.postMessage({
      type: "localization",
      strings: {
        ...designerStrings,
        header: "読み取り専用プレビュー",
        openText: "テキストエディターを開く",
        lastValid: "最後の有効な表示です。",
      },
    });
    const document = designerDocument({
      form: {
        name: "Main",
        caption: "<img src=x onerror=bad()>",
        build: { clientWidth: 240, clientHeight: 180 },
      },
      controls: [
        {
          id: "label",
          type: "Label",
          name: "Label1",
          caption: "<script>bad()</script>",
          left: 0.75,
          width: 72,
          height: 18,
        },
        {
          id: "button",
          type: "CommandButton",
          name: "Today",
          caption: "今日",
          top: 24,
          width: 90,
          height: 30,
        },
        {
          id: "toggle",
          type: "ToggleButton",
          name: "Toggle",
          caption: "First line\nSecond line",
          top: 60,
          width: 90,
          height: 45,
        },
        {
          id: "combo",
          type: "ComboBox",
          name: "Year",
          text: "2026",
          top: 110,
          width: 96,
          height: 30,
        },
      ],
    });
    const valid = await snapshot({ type: "document", version: 1, document });
    assert.strictEqual(valid.count, 4);
    assert.ok(String(valid.header).includes("読み取り専用プレビュー"));
    assert.ok(String(valid.header).includes("テキストエディターを開く"));
    const centers = valid.buttonCenters as {
      type: string;
      delta: number;
      shadow: string;
      topColor: string;
      bottomColor: string;
    }[];
    assert.strictEqual(centers.length, 2);
    for (const center of centers) {
      assert.ok(Math.abs(center.delta) < 1, `${center.type} caption should be vertically centered`);
      assert.notStrictEqual(center.shadow, "none", "raised button has inset bevel shading");
      assert.notStrictEqual(
        center.topColor,
        center.bottomColor,
        "button edges have highlight and shadow",
      );
    }
    const combo = valid.combo as {
      height: number;
      arrowHeight: number;
      arrowWidth: number;
      rightInset: number;
      shadow: string;
      arrowShadow: string;
    };
    assert.ok(combo.arrowHeight >= combo.height - 4, "dropdown button spans the field height");
    assert.strictEqual(combo.arrowWidth, 18);
    assert.ok(combo.rightInset <= 2, "dropdown button aligns with the right edge");
    assert.notStrictEqual(combo.shadow, "none", "combo field is recessed");
    assert.notStrictEqual(combo.arrowShadow, "none", "dropdown button is raised");
    assert.strictEqual(valid.title, document.caption);
    assert.strictEqual(valid.images, 0);
    assert.strictEqual(valid.left, "1px");
    assert.strictEqual(valid.width, "96px");
    assert.deepStrictEqual(valid.violations, []);
    const invalid = await snapshot({
      type: "invalidDocument",
      version: 2,
      error: { code: "syntax", message: "Invalid YAML" },
    });
    assert.strictEqual(invalid.count, 4);
    assert.ok(String(invalid.error).includes("Invalid YAML"));
    assert.ok(String(invalid.error).includes("最後の有効な表示です。"));
    const recovered = await snapshot({
      type: "document",
      version: 3,
      document: { ...document, caption: "Recovered" },
    });
    assert.strictEqual(recovered.title, "Recovered");
    assert.strictEqual(recovered.error, undefined);
    assert.deepStrictEqual(recovered.violations, []);
  } finally {
    subscription.dispose();
    panel.dispose();
  }
}

export async function runCustomEditorRegistrationAssertions(): Promise<void> {
  const directory = await mkdtemp(path.join(os.tmpdir(), "xlflow-designer-"));
  const file = path.join(directory, "Main.yaml");
  const source =
    "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols: []\n";
  const uri = vscode.Uri.file(file);
  try {
    await writeFile(file, source, "utf8");
    const document = await vscode.workspace.openTextDocument(uri);
    await vscode.commands.executeCommand("vscode.openWith", uri, designerViewType);
    const tab = vscode.window.tabGroups.activeTabGroup.activeTab;
    assert.ok(tab?.input instanceof vscode.TabInputCustom);
    assert.strictEqual(tab.input.viewType, designerViewType);
    assert.strictEqual(document.getText(), source);
    assert.strictEqual(document.isDirty, false, "opening a designer must not dirty the source");
    await vscode.commands.executeCommand("vscode.openWith", uri, "default");
    assert.strictEqual(vscode.window.activeTextEditor?.document.uri.toString(), uri.toString());
    assert.strictEqual(document.getText(), source);
    assert.strictEqual(document.isDirty, false);
    assert.strictEqual(await readFile(file, "utf8"), source);
  } finally {
    for (const group of vscode.window.tabGroups.all) {
      for (const tab of group.tabs) {
        const input = tab.input;
        if (
          (input instanceof vscode.TabInputText || input instanceof vscode.TabInputCustom) &&
          input.uri.toString() === uri.toString()
        )
          await vscode.window.tabGroups.close(tab, true);
      }
    }
    await unlink(file);
    await rmdir(directory);
  }
}
