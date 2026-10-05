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
const editMessages = [];
window.acquireVsCodeApi = () => ({
  postMessage(message) { if (message.type === "edit") editMessages.push(message); testApi.postMessage(message); },
  getState: () => testApi.getState(), setState: (state) => testApi.setState(state)
});
const violations = [];
window.addEventListener("securitypolicyviolation", e => violations.push(e.violatedDirective));
window.addEventListener("message", e => {
  if (!["document", "invalidDocument", "testAction"].includes(e.data.type)) return;
  setTimeout(async () => {
    for (const action of e.data.actions ?? []) {
      const target = document.querySelector(action.selector);
      if (!target) { testApi.postMessage({type: "testSnapshot", version: e.data.version, missingTarget: action.selector}); return; }
      const props = { bubbles: true, cancelable: true, ...action.props };
      if (action.event === "change") { if ("value" in props) target.value = props.value; if ("checked" in props) target.checked = props.checked; }
      target.dispatchEvent(action.event === "change" ? new Event("change", props) : action.event.startsWith("key") ? new KeyboardEvent(action.event, props) : new PointerEvent(action.event, props));
      await new Promise(resolve => setTimeout(resolve, 30));
    }
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
      top: label && getComputedStyle(label).top,
      height: label && getComputedStyle(label).height,
      handles: document.querySelectorAll(".resize-handle").length,
      editMessages,
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
  const snapshot = async (
    message:
      | Extract<HostMessage, { version: number }>
      | {
          type: "testAction";
          version: number;
          actions: { selector: string; event: string; props?: Record<string, unknown> }[];
        },
  ) => {
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

    const interactive = {
      ...document,
      controls: document.controls.map((control) => ({ ...control })),
    };
    interactive.controls[0] = {
      ...interactive.controls[0],
      left: 12,
      top: 12,
      width: 72,
      height: 18,
    };
    await snapshot({ type: "document", version: 4, document: interactive, editable: true });
    const pointer = (
      selector: string,
      event: string,
      clientX: number,
      clientY: number,
      pointerId = 1,
    ) => ({
      selector,
      event,
      props: { pointerId, button: 0, clientX, clientY },
    });
    const dragged = await snapshot({
      type: "testAction",
      version: 5,
      actions: [
        pointer(".control.label", "pointerdown", 30, 30),
        pointer(".designer-viewport", "pointermove", 46, 38),
      ],
    });
    assert.strictEqual(dragged.left, "32px");
    assert.strictEqual(dragged.top, "24px");
    assert.strictEqual(dragged.handles, 8);
    assert.deepStrictEqual(dragged.editMessages, [], "drag preview sends no edits");
    const moved = await snapshot({
      type: "testAction",
      version: 6,
      actions: [pointer(".designer-viewport", "pointerup", 46, 38)],
    });
    const moveMessages = moved.editMessages as { version: number; operations: unknown[] }[];
    assert.strictEqual(moveMessages.length, 1);
    assert.strictEqual(moveMessages[0].version, 4);
    assert.deepStrictEqual(moveMessages[0].operations, [
      { type: "moveControl", controlId: "label", left: 24, top: 18 },
    ]);
    interactive.controls[0] = { ...interactive.controls[0], left: 24, top: 18 };
    await snapshot({ type: "document", version: 7, document: interactive, editable: true });
    await snapshot({
      type: "testAction",
      version: 8,
      actions: [
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointerup", 40, 40),
      ],
    });
    const resizing = await snapshot({
      type: "testAction",
      version: 9,
      actions: [
        pointer(".handle-nw", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 48, 44),
      ],
    });
    assert.strictEqual((resizing.editMessages as unknown[]).length, 1);
    const resized = await snapshot({
      type: "testAction",
      version: 10,
      actions: [pointer(".designer-viewport", "pointerup", 48, 44)],
    });
    assert.deepStrictEqual((resized.editMessages as { operations: unknown[] }[])[1].operations, [
      { type: "moveControl", controlId: "label", left: 30, top: 21 },
      { type: "resizeControl", controlId: "label", width: 66, height: 15 },
    ]);
    interactive.controls[0] = {
      ...interactive.controls[0],
      left: 30,
      top: 21,
      width: 66,
      height: 15,
    };
    await snapshot({ type: "document", version: 11, document: interactive, editable: true });
    const key = (event: string, key: string, shiftKey = false) => ({
      selector: ".designer-viewport",
      event,
      props: { key, shiftKey },
    });
    const repeating = await snapshot({
      type: "testAction",
      version: 12,
      actions: [
        key("keydown", "ArrowRight"),
        key("keydown", "ArrowRight"),
        key("keydown", "ArrowDown", true),
      ],
    });
    assert.strictEqual((repeating.editMessages as unknown[]).length, 2);
    const nudged = await snapshot({
      type: "testAction",
      version: 13,
      actions: [key("keyup", "ArrowRight"), key("keyup", "ArrowDown")],
    });
    assert.deepStrictEqual((nudged.editMessages as { operations: unknown[] }[])[2].operations, [
      { type: "moveControl", controlId: "label", left: 32, top: 31 },
    ]);
    await snapshot({ type: "document", version: 14, document: interactive, editable: true });
    const cancelled = await snapshot({
      type: "testAction",
      version: 15,
      actions: [
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 100, 100),
        key("keydown", "Escape"),
        pointer(".designer-viewport", "pointerup", 100, 100),
      ],
    });
    assert.strictEqual((cancelled.editMessages as unknown[]).length, 3);
    assert.strictEqual(cancelled.left, "40px");
    assert.deepStrictEqual(cancelled.violations, []);

    await snapshot({ type: "document", version: 16, document: interactive, editable: true });
    const zoomed = await snapshot({
      type: "testAction",
      version: 17,
      actions: [
        { selector: ".designer-toolbar select", event: "change", props: { value: "2" } },
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 56, 48),
        pointer(".designer-viewport", "pointerup", 56, 48),
      ],
    });
    assert.deepStrictEqual(
      (zoomed.editMessages as { operations: unknown[] }[])[3].operations,
      [{ type: "moveControl", controlId: "label", left: 36, top: 24 }],
      "200% zoom uses inverse point transform",
    );
    await snapshot({ type: "document", version: 18, document: interactive, editable: true });
    const snapped = await snapshot({
      type: "testAction",
      version: 19,
      actions: [
        {
          selector: ".designer-toolbar input:nth-of-type(1)",
          event: "change",
          props: { checked: true },
        },
      ],
    });
    assert.strictEqual(
      (snapped.editMessages as unknown[]).length,
      4,
      "grid display is not an edit",
    );
    const gridSnap = await snapshot({
      type: "testAction",
      version: 20,
      actions: [
        {
          selector: ".designer-toolbar label:nth-child(2) input",
          event: "change",
          props: { checked: true },
        },
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 48, 40),
        pointer(".designer-viewport", "pointerup", 48, 40),
      ],
    });
    assert.deepStrictEqual((gridSnap.editMessages as { operations: unknown[] }[])[4].operations, [
      { type: "moveControl", controlId: "label", left: 32, top: 21 },
    ]);
    await snapshot({ type: "document", version: 21, document: interactive, editable: true });
    const pointerCancelled = await snapshot({
      type: "testAction",
      version: 22,
      actions: [
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 80, 80),
        pointer(".designer-viewport", "pointercancel", 80, 80),
        pointer(".designer-viewport", "pointerup", 80, 80),
      ],
    });
    assert.strictEqual((pointerCancelled.editMessages as unknown[]).length, 5);
    assert.strictEqual(pointerCancelled.left, "40px");
    interactive.controls[0] = { ...interactive.controls[0], width: 0.5, height: 0.5 };
    await snapshot({ type: "document", version: 23, document: interactive, editable: true });
    const tinyMove = await snapshot({
      type: "testAction",
      version: 24,
      actions: [
        { selector: ".designer-toolbar select", event: "change", props: { value: "1" } },
        {
          selector: ".designer-toolbar label:nth-child(2) input",
          event: "change",
          props: { checked: false },
        },
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 48, 48),
        pointer(".designer-viewport", "pointerup", 48, 48),
        key("keydown", "ArrowRight"),
        key("keyup", "ArrowRight"),
      ],
    });
    assert.strictEqual((tinyMove.editMessages as unknown[]).length, 5, "tiny control cannot move");
    assert.strictEqual(tinyMove.left, "40px");
    const tinyResize = await snapshot({
      type: "testAction",
      version: 25,
      actions: [
        pointer(".handle-se", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 48, 48),
        pointer(".designer-viewport", "pointerup", 48, 48),
      ],
    });
    assert.deepStrictEqual((tinyResize.editMessages as { operations: unknown[] }[])[5].operations, [
      { type: "resizeControl", controlId: "label", width: 6.5, height: 6.5 },
    ]);
    interactive.controls[0] = { ...interactive.controls[0], width: 66, height: 15 };
    await snapshot({ type: "document", version: 26, document: interactive, editable: true });
    const secondaryCancelled = await snapshot({
      type: "testAction",
      version: 27,
      actions: [
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 48, 48),
        pointer(".control.label", "pointerdown", 80, 80, 2),
        pointer(".designer-viewport", "pointercancel", 80, 80, 2),
        pointer(".designer-viewport", "lostpointercapture", 80, 80, 2),
      ],
    });
    assert.strictEqual(secondaryCancelled.left, "48px", "secondary pointer leaves preview intact");
    assert.strictEqual((secondaryCancelled.editMessages as unknown[]).length, 6);
    const primaryReleased = await snapshot({
      type: "testAction",
      version: 28,
      actions: [pointer(".designer-viewport", "pointerup", 48, 48)],
    });
    assert.deepStrictEqual(
      (primaryReleased.editMessages as { operations: unknown[] }[])[6].operations,
      [{ type: "moveControl", controlId: "label", left: 36, top: 27 }],
    );
    await snapshot({ type: "document", version: 29, document: interactive, editable: true });
    const primaryCaptureLost = await snapshot({
      type: "testAction",
      version: 30,
      actions: [
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointermove", 48, 48),
        pointer(".designer-viewport", "lostpointercapture", 48, 48),
        pointer(".designer-viewport", "pointerup", 48, 48),
      ],
    });
    assert.strictEqual((primaryCaptureLost.editMessages as unknown[]).length, 7);
    assert.strictEqual(primaryCaptureLost.left, "40px", "active pointer capture loss cancels");
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
