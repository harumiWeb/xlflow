import * as assert from "assert";
import { NativeDesignerMouse } from "./designerNativeMouse";
import * as vscode from "vscode";
import { mkdtemp, readFile, unlink, rm, writeFile } from "fs/promises";
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
const clipboardMessages = [];
const openEventMessages = [];
let hitTestControlId;
window.acquireVsCodeApi = () => ({
  postMessage(message) { if (message.type === "edit") editMessages.push(message); if (message.type === "clipboard") clipboardMessages.push(message); if (message.type === "openEvent") openEventMessages.push(message); testApi.postMessage(message); },
  getState: () => testApi.getState(), setState: (state) => testApi.setState(state)
});
const violations = [];
const runtimeErrors = [];
let pointerFormOrigin;
let pointerHandleOrigin;
window.addEventListener("error", e => runtimeErrors.push(e.message));
window.addEventListener("unhandledrejection", e => runtimeErrors.push(String(e.reason)));
window.addEventListener("securitypolicyviolation", e => violations.push(e.violatedDirective));
window.addEventListener("message", e => {
  if (!["document", "invalidDocument", "testAction"].includes(e.data.type)) return;
  setTimeout(async () => {
    if (e.data.clearEditMessages) editMessages.length = 0;
    for (const action of e.data.actions ?? []) {
      let target = document.querySelector(action.selector);
      if (!target) { testApi.postMessage({type: "testSnapshot", version: e.data.version, missingTarget: action.selector}); return; }
      const props = { bubbles: true, cancelable: true, ...action.props };
      if (props.expandIfCollapsed && target.closest('[role="treeitem"]')?.getAttribute('aria-expanded') === 'true') continue;
      if (props.formPoints) {
        // Establish focus/scroll before computing synthetic pointer coordinates;
        // the real pointer-down handler focuses the canvas and may scroll it.
        if (action.event === 'pointerdown' && !props.preserveScroll) {
          const canvas = document.querySelector('.designer-viewport');
          canvas.focus({ preventScroll: true });
          canvas.scrollIntoView({ block: 'start', inline: 'start', behavior: 'instant' });
          await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        }
        const currentBounds = document.querySelector(".form").getBoundingClientRect();
        if (action.event === 'pointerdown') pointerFormOrigin = currentBounds;
        // Keep a gesture in one screen coordinate system even when selection
        // breadcrumbs or notices change the form's position during the drag.
        const bounds = pointerFormOrigin ?? currentBounds;
        const zoom = Number(document.querySelector(".designer-toolbar select").value);
        if (props.handleCenter) {
          const handle = target.getBoundingClientRect();
          props.formPoints = [(handle.left + handle.right - 2 * bounds.left) * 72 / (2 * 96 * zoom), (handle.top + handle.bottom - 2 * bounds.top) * 72 / (2 * 96 * zoom)];
        }
        props.clientX = bounds.left + props.formPoints[0] * 96 * zoom / 72;
        props.clientY = bounds.top + props.formPoints[1] * 96 * zoom / 72;
        if (props.handleCenter) pointerHandleOrigin = { x: props.clientX, y: props.clientY };
        if (props.handleDelta && pointerHandleOrigin) {
          props.clientX = pointerHandleOrigin.x + props.handleDelta[0] * 96 * zoom / 72;
          props.clientY = pointerHandleOrigin.y + props.handleDelta[1] * 96 * zoom / 72;
        }
      }
      if (props.centerHitTest) {
        const bounds = target.getBoundingClientRect();
        props.clientX = (bounds.left + bounds.right) / 2;
        props.clientY = (bounds.top + bounds.bottom) / 2;
        props.hitTest = true;
      }
      if (props.hitTest) {
        target = document.elementFromPoint(props.clientX, props.clientY);
        if (!target) { testApi.postMessage({type: "testSnapshot", version: e.data.version, missingTarget: "hit test " + action.selector}); return; }
        if (props.expectControlHit) hitTestControlId = target.closest('[data-control-id]')?.dataset.controlId;
      }
      if (action.event === "scroll") { if ("scrollTop" in props) target.scrollTop = props.scrollTop; if ("scrollLeft" in props) target.scrollLeft = props.scrollLeft; }
      if (action.event === "change" || action.event === "input") { if ("value" in props) target.value = props.value; if ("checked" in props) target.checked = props.checked; }
      target.dispatchEvent(["change", "input"].includes(action.event) ? new Event(action.event, props) : action.event.startsWith("key") ? new KeyboardEvent(action.event, props) : action.event === "focusout" ? new FocusEvent("blur", props) : action.event === "dblclick" ? new MouseEvent(action.event, { ...props, detail: props.detail ?? 2 }) : new PointerEvent(action.event, props));
      if (action.event === 'pointerup' || action.event === 'pointercancel') pointerFormOrigin = undefined;
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
      formColor: document.querySelector('.form') && getComputedStyle(document.querySelector('.form')).color,
      controlColors: Object.fromEntries([...document.querySelectorAll('.control[data-control-id]')].map(control => [control.dataset.controlId, getComputedStyle(control).color])),
      controlFonts: Object.fromEntries([...document.querySelectorAll('.control[data-control-id]')].map(control => {
        const style = getComputedStyle(control);
        return [control.dataset.controlId, { weight: style.fontWeight, style: style.fontStyle, family: style.fontFamily, size: style.fontSize, decoration: style.textDecorationLine }];
      })),
      count: document.querySelectorAll(".control").length,
      title: document.querySelector(".form-title")?.textContent,
      header: document.querySelector("header")?.textContent,
      error: document.querySelector('[role="alert"]')?.textContent,
      left: label && getComputedStyle(label).left,
      width: label && getComputedStyle(label).width,
      top: label && getComputedStyle(label).top,
      height: label && getComputedStyle(label).height,
      formWidth: document.querySelector(".form") && getComputedStyle(document.querySelector(".form")).width,
      formHeight: document.querySelector(".form") && getComputedStyle(document.querySelector(".form")).height,
      formHandles: document.querySelectorAll("[data-form-handle]").length,
      handles: document.querySelectorAll(".resize-handle").length,
      toolboxDisabled: document.querySelector('[data-toolbox-type="Label"]')?.disabled,
      placing: !!document.querySelector(".placing-control"),
      selectedId: document.querySelector(".selection-outline")?.dataset.controlId,
      selectedIds: [...document.querySelectorAll(".selection-outline")].map(c => c.dataset.controlId),
      groupSelection: !!document.querySelector(".group-selection-outline"),
      canvasFocused: document.activeElement === document.querySelector(".designer-viewport"),
      dropParentId: document.querySelector("[data-drop-parent-id]")?.dataset.dropParentId,
      floating: !!document.querySelector(".drag-preview"),
      selectionPath: [...document.querySelectorAll(".selection-path button")].map(b => b.textContent),
      worldPositions: [...document.querySelectorAll(".control[data-control-id]")].map(c => {
        const bounds = c.getBoundingClientRect();
        const form = document.querySelector(".form").getBoundingClientRect();
        const scale = 96 * Number(document.querySelector(".designer-toolbar select").value) / 72;
        return { id: c.dataset.controlId, left: (bounds.left - form.left) / scale, top: (bounds.top - form.top) / scale };
      }),
      overflowHits: (() => {
        const child = document.querySelector('[data-control-id="overflow-child"]');
        if (!child) return undefined;
        // Hit testing uses viewport coordinates. Wrapped productivity controls
        // can push the canvas below the smaller Linux CI Webview viewport.
        // Scroll the canvas, not the clipped child or its Frame content.
        document.querySelector('.designer-viewport').scrollIntoView({ block: 'start', inline: 'start', behavior: 'instant' });
        const bounds = child.getBoundingClientRect();
        const content = child.parentElement.getBoundingClientRect();
        const y = (bounds.top + bounds.bottom) / 2;
        return {
          inside: document.elementFromPoint((bounds.left + content.right) / 2, y)?.closest('[data-control-id]')?.dataset.controlId,
          outside: document.elementFromPoint((content.right + bounds.right) / 2, y)?.closest('[data-control-id]')?.dataset.controlId,
        };
      })(),
      editMessages,
      clipboardMessages,
      openEventMessages,
      hitTestControlId,
      tabOrderItems: [...document.querySelectorAll(".tab-order-editor li > span")].map(item => item.textContent),
      activeRightTab: document.querySelector('[role="tab"][aria-selected="true"]')?.id,
      canvasScroll: (() => { const c = document.querySelector(".canvas-scroll"); return c && {left:c.scrollLeft, top:c.scrollTop, width:c.clientWidth, height:c.clientHeight}; })(),
      codiconLoaded: document.fonts.check('16px codicon'),
      images: document.querySelectorAll("img").length,
      propertyTitle: document.querySelector(".property-grid h2")?.textContent,
      propertyRows: [...document.querySelectorAll("[data-property-field]")].map(row => ({
        field: row.dataset.propertyField,
        value: row.querySelector("input")?.value,
        disabled: row.querySelector("input")?.disabled,
        text: row.textContent,
        error: row.querySelector('[role="alert"]')?.textContent
      })),
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
      , runtimeErrors
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
          clearEditMessages?: boolean;
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
        width: 240,
        height: 180,
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
      eventEditable: true,
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
    const nativeMouse = await NativeDesignerMouse.connect();
    try {
      const initialNative = await snapshot({ type: "testAction", version: 99, actions: [] });
      const initialOpenEvents = (initialNative.openEventMessages as unknown[]).length;
      const controlPoint = await nativeMouse.doubleClick("control");
      assert.strictEqual(controlPoint.hitInForm, true);
      assert.strictEqual(controlPoint.hitControlId, "label");
      const openedControl = await snapshot({ type: "testAction", version: 100, actions: [] });
      const controlOpenEvents = openedControl.openEventMessages as {
        type: string;
        requestId: number;
        version: number;
        targetId?: string;
        choose: boolean;
      }[];
      assert.strictEqual(controlOpenEvents.length, initialOpenEvents + 1);
      assert.deepStrictEqual(controlOpenEvents.at(-1), {
        type: "openEvent",
        requestId: controlOpenEvents.at(-1)!.requestId,
        version: 4,
        targetId: "label",
        choose: false,
      });
      await panel.webview.postMessage({
        type: "eventResult",
        requestId: controlOpenEvents.at(-1)!.requestId,
      });
      await snapshot({ type: "testAction", version: 101, actions: [] });

      const formPoint = await nativeMouse.doubleClick("form");
      assert.strictEqual(formPoint.hitInForm, true);
      assert.strictEqual(formPoint.hitControlId, undefined);
      const openedForm = await snapshot({ type: "testAction", version: 102, actions: [] });
      const formOpenEvents = openedForm.openEventMessages as {
        type: string;
        requestId: number;
        version: number;
        targetId?: string;
        choose: boolean;
      }[];
      assert.strictEqual(formOpenEvents.length, initialOpenEvents + 2);
      assert.deepStrictEqual(formOpenEvents.at(-1), {
        type: "openEvent",
        requestId: formOpenEvents.at(-1)!.requestId,
        version: 4,
        choose: false,
      });
      await panel.webview.postMessage({
        type: "eventResult",
        requestId: formOpenEvents.at(-1)!.requestId,
      });
      await snapshot({ type: "testAction", version: 103, actions: [] });

      await nativeMouse.dragControl();
      const afterNativeDrag = await snapshot({ type: "testAction", version: 104, actions: [] });
      assert.strictEqual(
        (afterNativeDrag.openEventMessages as unknown[]).length,
        initialOpenEvents + 2,
        "native control drag does not open an event",
      );
      const dragEdits = afterNativeDrag.editMessages as { requestId: number }[];
      assert.strictEqual(dragEdits.length, 1, "native drag still commits its control move");
      await panel.webview.postMessage({ type: "editResult", requestId: dragEdits[0].requestId });
      await snapshot({ type: "document", version: 4, document: interactive, editable: true });
      await snapshot({
        type: "testAction",
        version: 105,
        actions: [],
        clearEditMessages: true,
      });
    } finally {
      nativeMouse.close();
    }
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
          selector: ".grid-toggle input",
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
          selector: ".snap-toggle input",
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
          selector: ".snap-toggle input",
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

    const propertyGrid = {
      form: {
        descriptors: [
          { field: "caption", valueType: "string" as const, required: false, nullable: true },
          { field: "width", valueType: "number" as const, required: false, nullable: true },
          {
            field: "height",
            valueType: "number" as const,
            required: false,
            nullable: true,
          },
        ],
        values: {
          caption: { present: true, value: "Authored caption" },
          width: { present: true, value: 240 },
          height: { present: false, value: null },
        },
      },
      controls: {
        label: {
          descriptors: [
            { field: "caption", valueType: "string" as const, required: false, nullable: true },
          ],
          values: { caption: { present: true, value: "Authored label" } },
        },
      },
    };
    const propertyDocument = () => ({
      type: "document" as const,
      version: 40,
      document: interactive,
      editable: true,
      propertyEditable: true,
      propertyGrid,
    });
    await snapshot(propertyDocument());
    const formSelected = await snapshot({
      type: "testAction",
      version: 41,
      actions: [
        pointer(".form", "pointerdown", 0, 0),
        pointer(".designer-viewport", "pointerup", 0, 0),
      ],
    });
    assert.ok(String(formSelected.propertyTitle).includes("Main"));
    const field = (name: string) => `[data-property-field="${name}"] input`;
    const changedCaption = await snapshot({
      type: "testAction",
      version: 42,
      actions: [
        { selector: field("caption"), event: "input", props: { value: "新しいキャプション 😀" } },
        { selector: field("caption"), event: "keydown", props: { key: "Enter" } },
        { selector: field("caption"), event: "focusout" },
      ],
    });
    const propertyEdits = changedCaption.editMessages as {
      requestId: number;
      operations: unknown[];
    }[];
    assert.strictEqual(propertyEdits.length, 8, "Enter and blur create exactly one property edit");
    assert.deepStrictEqual(propertyEdits.at(-1)!.operations, [
      { type: "setFormProperty", field: "caption", value: "新しいキャプション 😀" },
    ]);
    await panel.webview.postMessage({
      type: "editResult",
      requestId: propertyEdits.at(-1)!.requestId,
      error: {
        code: "invalid",
        message: "Rejected property",
        diagnostics: [
          { operationIndex: 0, field: "caption", code: "UFV001", message: "Invalid caption" },
        ],
      },
    });
    const rejected = await snapshot({
      type: "testAction",
      version: 43,
      actions: [{ selector: field("caption"), event: "focusout" }],
    });
    const rejectedRows = rejected.propertyRows as {
      field: string;
      value: string;
      error?: string;
    }[];
    assert.strictEqual(
      rejectedRows.find((r) => r.field === "caption")!.value,
      "新しいキャプション 😀",
      "rejection retains the input",
    );
    assert.ok(rejectedRows.find((r) => r.field === "caption")!.error?.includes("UFV001"));
    assert.strictEqual(
      (rejected.editMessages as unknown[]).length,
      8,
      "rejected draft is not resent on blur",
    );
    const invalidNumber = await snapshot({
      type: "testAction",
      version: 44,
      actions: [
        { selector: field("width"), event: "input", props: { value: "invalid" } },
        { selector: field("width"), event: "keydown", props: { key: "Enter" } },
      ],
    });
    assert.strictEqual((invalidNumber.editMessages as unknown[]).length, 8);
    assert.ok(
      (invalidNumber.propertyRows as { field: string; error?: string }[]).find(
        (r) => r.field === "width",
      )!.error,
    );
    const staleDraft = await snapshot({
      type: "document",
      version: 45,
      document: interactive,
      editable: true,
      propertyEditable: true,
      propertyGrid,
    });
    assert.ok(
      (staleDraft.propertyRows as { field: string; error?: string }[])
        .find((r) => r.field === "caption")!
        .error?.includes(designerStrings.staleProperty),
    );
    const retry = await snapshot({
      type: "testAction",
      version: 46,
      actions: [
        { selector: field("caption"), event: "input", props: { value: "Retry caption" } },
        { selector: field("caption"), event: "focusout" },
      ],
    });
    const retryEdits = retry.editMessages as { requestId: number; operations: unknown[] }[];
    assert.strictEqual(retryEdits.length, 9);
    await panel.webview.postMessage({
      type: "editResult",
      requestId: retryEdits.at(-1)!.requestId,
    });
    await snapshot({
      type: "document",
      version: 47,
      document: interactive,
      editable: true,
      propertyEditable: true,
      propertyGrid,
    });
    const controlSelected = await snapshot({
      type: "testAction",
      version: 48,
      actions: [
        pointer(".control.label", "pointerdown", 40, 40),
        pointer(".designer-viewport", "pointerup", 40, 40),
      ],
    });
    assert.ok(String(controlSelected.propertyTitle).includes("Label1"));
    const controlEdit = await snapshot({
      type: "testAction",
      version: 49,
      actions: [
        { selector: field("caption"), event: "input", props: { value: "Control caption" } },
        { selector: field("caption"), event: "keydown", props: { key: "Enter" } },
      ],
    });
    assert.deepStrictEqual(
      (controlEdit.editMessages as { operations: unknown[] }[]).at(-1)!.operations,
      [
        {
          type: "setControlProperty",
          controlId: "label",
          field: "caption",
          value: "Control caption",
        },
      ],
    );
    assert.deepStrictEqual(controlEdit.violations, []);

    const allEdits = controlEdit.editMessages as { requestId: number; operations: unknown[] }[];
    await panel.webview.postMessage({ type: "editResult", requestId: allEdits.at(-1)!.requestId });
    const structuralDocument = {
      type: "document" as const,
      version: 50,
      document: interactive,
      structuralEditable: true,
    };
    const structuralReady = await snapshot(structuralDocument);
    assert.strictEqual(
      structuralReady.toolboxDisabled,
      false,
      "structural editing is independent of geometry and properties",
    );
    const added = await snapshot({
      type: "testAction",
      version: 51,
      actions: [
        { selector: '[data-toolbox-type="TextBox"]', event: "click" },
        { selector: ".form", event: "pointerdown", props: { button: 0, formPoints: [12, 16] } },
      ],
    });
    const structuralEdits = added.editMessages as {
      requestId: number;
      version: number;
      operations: {
        type: string;
        control?: {
          id: string;
          name: string;
          type: string;
          left: number;
          top: number;
          width: number;
          height: number;
        };
      }[];
    }[];
    assert.strictEqual(structuralEdits.length, allEdits.length + 1);
    const addition = structuralEdits.at(-1)!;
    assert.strictEqual(addition.version, 50);
    const newControl = addition.operations[0].control!;
    assert.match(newControl.id, /^control-[0-9a-f-]{36}$/);
    assert.strictEqual(newControl.name, "TextBox1");
    assert.deepStrictEqual(
      [newControl.left, newControl.top, newControl.width, newControl.height],
      [12, 16, 120, 18],
    );
    assert.strictEqual(
      added.count,
      4,
      "pending creation does not replace canonical document state",
    );
    interactive.controls.push({ ...newControl, approximate: false });
    const addedDocument = await snapshot({ ...structuralDocument, version: 52 });
    await panel.webview.postMessage({ type: "editResult", requestId: addition.requestId });
    assert.strictEqual(addedDocument.selectedId, newControl.id);
    const deleting = await snapshot({
      type: "testAction",
      version: 53,
      actions: [{ selector: ".delete-control", event: "click" }],
    });
    const deleteEdits = deleting.editMessages as { requestId: number; operations: unknown[] }[];
    assert.deepStrictEqual(deleteEdits.at(-1)!.operations, [
      { type: "removeControl", controlId: newControl.id, cascade: true },
    ]);
    await panel.webview.postMessage({
      type: "editResult",
      requestId: deleteEdits.at(-1)!.requestId,
      error: { code: "editCancelled", message: "Cancelled" },
    });
    const cancelledDeletion = await snapshot({ type: "testAction", version: 54, actions: [] });
    assert.strictEqual(cancelledDeletion.count, 5);
    assert.strictEqual(cancelledDeletion.error, undefined);
    const deleted = await snapshot({
      type: "testAction",
      version: 55,
      actions: [key("keydown", "Delete")],
    });
    const deletedEdits = deleted.editMessages as { requestId: number; operations: unknown[] }[];
    assert.strictEqual(deletedEdits.length, deleteEdits.length + 1);
    interactive.controls = interactive.controls.filter((control) => control.id !== newControl.id);
    await panel.webview.postMessage({
      type: "editResult",
      requestId: deletedEdits.at(-1)!.requestId,
    });
    const afterDeletion = await snapshot({ ...structuralDocument, version: 56 });
    assert.strictEqual(afterDeletion.count, 4);
    assert.strictEqual(afterDeletion.selectedId, undefined);
    assert.deepStrictEqual(afterDeletion.violations, []);

    const parenting = designerDocument({
      form: { name: "Main", width: 600, height: 400 },
      controls: [
        { id: "label", name: "Label1", type: "Label", left: 10, top: 10, width: 20, height: 10 },
        {
          id: "frame",
          name: "Frame1",
          type: "Frame",
          left: 200,
          top: 100,
          width: 180,
          height: 160,
        },
        {
          id: "nested",
          parentId: "frame",
          name: "Frame2",
          type: "Frame",
          left: 20,
          top: 30,
          width: 70,
          height: 60,
        },
        {
          id: "child",
          parentId: "nested",
          name: "Child",
          type: "Label",
          left: 5,
          top: 6,
          width: 20,
          height: 10,
        },
      ],
    });
    const parentingDocument = (version: number) => ({
      type: "document" as const,
      version,
      document: parenting,
      editable: true,
      structuralEditable: true,
      parentEditable: true,
    });
    const parentPointer = (selector: string, event: string, left: number, top: number) => ({
      selector,
      event,
      props: { button: 0, pointerId: 1, formPoints: [left, top] },
    });
    const position = (result: Record<string, unknown>, id: string) =>
      (result.worldPositions as { id: string; left: number; top: number }[]).find(
        (c) => c.id === id,
      )!;
    await snapshot(parentingDocument(60));
    const parentPreview = await snapshot({
      type: "testAction",
      version: 61,
      actions: [
        { selector: ".designer-toolbar select", event: "change", props: { value: "1" } },
        { selector: ".snap-toggle input", event: "change", props: { checked: false } },
        parentPointer('.control[data-control-id="label"]', "pointerdown", 12, 12),
        parentPointer(".designer-viewport", "pointermove", 220, 120),
      ],
    });
    assert.strictEqual(parentPreview.dropParentId, "frame");
    assert.strictEqual(parentPreview.floating, true);
    assert.ok(Math.abs(position(parentPreview, "label").left - 218) < 0.1);
    const parentDrop = await snapshot({
      type: "testAction",
      version: 62,
      actions: [parentPointer(".designer-viewport", "pointerup", 220, 120)],
    });
    const parentEdits = parentDrop.editMessages as { requestId: number; operations: unknown[] }[];
    assert.deepStrictEqual(parentEdits.at(-1)!.operations, [
      { type: "setParent", controlId: "label", parentId: "frame", left: 17.25, top: 17.25 },
    ]);
    parenting.controls[0] = {
      ...parenting.controls[0],
      parentId: "frame",
      left: 17.25,
      top: 17.25,
    };
    await panel.webview.postMessage({
      type: "editResult",
      requestId: parentEdits.at(-1)!.requestId,
    });
    const parented = await snapshot(parentingDocument(63));
    assert.ok(
      Math.abs(position(parented, "label").left - position(parentPreview, "label").left) < 0.1,
      "preview and nested canonical rendering agree",
    );
    assert.deepStrictEqual(parented.selectionPath, ["Main", "Frame1", "Label1"]);
    const unparent = await snapshot({
      type: "testAction",
      version: 64,
      actions: [
        parentPointer('.control[data-control-id="label"]', "pointerdown", 220, 120),
        parentPointer(".designer-viewport", "pointerup", 100, 50),
      ],
    });
    const unparentEdits = unparent.editMessages as { requestId: number; operations: unknown[] }[];
    assert.deepStrictEqual(unparentEdits.at(-1)!.operations, [
      { type: "setParent", controlId: "label", parentId: "", left: 98, top: 48 },
    ]);
    parenting.controls[0] = { ...parenting.controls[0], parentId: undefined, left: 98, top: 48 };
    await panel.webview.postMessage({
      type: "editResult",
      requestId: unparentEdits.at(-1)!.requestId,
    });
    await snapshot(parentingDocument(65));
    const nestedPreview = await snapshot({
      type: "testAction",
      version: 66,
      actions: [
        parentPointer('.control[data-control-id="nested"]', "pointerdown", 222, 132),
        parentPointer(".designer-viewport", "pointermove", 100, 50),
      ],
    });
    assert.strictEqual(nestedPreview.floating, true);
    assert.ok(
      Math.abs(
        position(nestedPreview, "child").left - position(nestedPreview, "nested").left - 5.75,
      ) < 0.1,
      "floating Frame preserves descendant local coordinates",
    );
    const cancelNested = await snapshot({
      type: "testAction",
      version: 67,
      actions: [
        key("keydown", "Escape"),
        parentPointer(".designer-viewport", "pointerup", 100, 50),
      ],
    });
    assert.strictEqual((cancelNested.editMessages as unknown[]).length, unparentEdits.length);
    assert.strictEqual(cancelNested.floating, false);
    const nestedSelected = await snapshot({
      type: "testAction",
      version: 68,
      actions: [
        {
          ...parentPointer('.control[data-control-id="child"]', "pointerdown", 228.5, 139.5),
          props: { pointerId: 1, button: 0, formPoints: [228.5, 139.5], hitTest: true },
        },
        parentPointer(".designer-viewport", "pointerup", 228.5, 139.5),
        { selector: ".selection-path > span button", event: "click" },
      ],
    });
    assert.strictEqual(nestedSelected.selectedId, "frame");
    assert.strictEqual(nestedSelected.canvasFocused, true);
    const breadcrumbMove = await snapshot({
      type: "testAction",
      version: 69,
      actions: [key("keydown", "ArrowRight"), key("keyup", "ArrowRight")],
    });
    const breadcrumbEdits = breadcrumbMove.editMessages as {
      requestId: number;
      operations: unknown[];
    }[];
    assert.deepStrictEqual(breadcrumbEdits.at(-1)!.operations, [
      { type: "moveControl", controlId: "frame", left: 201, top: 100 },
    ]);
    parenting.controls[1] = { ...parenting.controls[1], left: 201 };
    await panel.webview.postMessage({
      type: "editResult",
      requestId: breadcrumbEdits.at(-1)!.requestId,
    });
    await snapshot(parentingDocument(70));
    const frameAddition = await snapshot({
      type: "testAction",
      version: 71,
      actions: [
        { selector: '[data-toolbox-type="TextBox"]', event: "click" },
        parentPointer(".form", "pointerdown", 211, 110),
      ],
    });
    const frameAddEdits = frameAddition.editMessages as {
      requestId: number;
      operations: { type: string; control: { parentId: string; left: number; top: number } }[];
    }[];
    assert.strictEqual(frameAddEdits.at(-1)!.operations[0].type, "addControl");
    assert.deepStrictEqual(
      [
        frameAddEdits.at(-1)!.operations[0].control.parentId,
        frameAddEdits.at(-1)!.operations[0].control.left,
        frameAddEdits.at(-1)!.operations[0].control.top,
      ],
      ["frame", 9.25, 9.25],
    );
    assert.deepStrictEqual(frameAddition.violations, []);
    await panel.webview.postMessage({
      type: "editResult",
      requestId: (frameAddition.editMessages as { requestId: number }[]).at(-1)!.requestId,
      error: { code: "editCancelled", message: "Cancelled" },
    });
    const overflow = designerDocument({
      form: { name: "Overflow", width: 400, height: 300 },
      controls: [
        { id: "frame", name: "Frame1", type: "Frame", left: 20, top: 20, width: 100, height: 80 },
        {
          id: "overflow-child",
          parentId: "frame",
          name: "Child",
          type: "Label",
          left: 90,
          top: 10,
          width: 20,
          height: 18,
        },
      ],
    });
    const clipped = await snapshot({ ...parentingDocument(72), document: overflow });
    assert.deepStrictEqual(clipped.overflowHits, { inside: "overflow-child" });
    assert.strictEqual(overflow.controls[1].left, 90, "load preserves overflow source geometry");
    const adjusted = await snapshot({
      type: "testAction",
      version: 73,
      actions: [
        key("keydown", "Escape"),
        parentPointer('.control[data-control-id="overflow-child"]', "pointerdown", 115, 31),
        parentPointer(".designer-viewport", "pointerup", 116, 31),
      ],
    });
    const overflowEdits = adjusted.editMessages as { operations: unknown[] }[];
    assert.deepStrictEqual(
      overflowEdits.at(-1)!.operations,
      [{ type: "moveControl", controlId: "overflow-child", left: 78.5, top: 10 }],
      "a short drag from the visible part retains the Frame parent and clamps geometry",
    );
    assert.deepStrictEqual(adjusted.violations, []);
    const collections = designerDocument({
      form: { name: "Collections", width: 600, height: 400 },
      controls: [
        {
          id: "multi",
          type: "MultiPage",
          name: "Multi",
          left: 40,
          top: 40,
          width: 240,
          height: 180,
          selectedIndex: 0,
        },
        { id: "page-a", type: "Page", name: "PageA", parentId: "multi", zIndex: 0 },
        { id: "page-b", type: "Page", name: "PageB", parentId: "multi", zIndex: 1 },
        { id: "child-a", type: "Label", name: "ChildA", parentId: "page-a", left: 10, top: 10 },
        { id: "child-b", type: "Label", name: "ChildB", parentId: "page-b", left: 10, top: 10 },
        {
          id: "tabs",
          type: "TabStrip",
          name: "Tabs",
          left: 300,
          top: 40,
          tabs: [{ name: "General" }],
        },
      ],
    });
    const collectionDocument = (version: number): Extract<HostMessage, { type: "document" }> => ({
      type: "document",
      version,
      document: collections,
      editable: true,
      parentEditable: true,
      structuralEditable: true,
      pageEditable: true,
      tabEditable: true,
      propertyGrid: {
        form: { descriptors: [], values: {} },
        controls: {},
        tabs: {
          tabs: [
            {
              descriptors: [
                { field: "caption", valueType: "string", nullable: true, required: false },
              ],
              values: { caption: { present: false, value: null } },
            },
          ],
        },
      },
    });
    const beforePageSwitch = await snapshot(collectionDocument(80));
    assert.strictEqual(
      beforePageSwitch.title,
      "Collections",
      "source replacement cancels a preview whose control was removed",
    );
    assert.deepStrictEqual(beforePageSwitch.runtimeErrors, []);
    const pageSwitch = await snapshot({
      type: "testAction",
      version: 81,
      actions: [{ selector: '[data-page-tab="page-b"]', event: "click" }],
    });
    assert.strictEqual(pageSwitch.missingTarget, undefined);
    assert.strictEqual(
      (pageSwitch.editMessages as unknown[]).length,
      (beforePageSwitch.editMessages as unknown[]).length,
    );
    const shownIds = (pageSwitch.worldPositions as { id: string }[]).map((c) => c.id);
    assert.ok(shownIds.includes("child-b"));
    assert.ok(!shownIds.includes("child-a"));
    assert.strictEqual(collections.controls[0].selectedIndex, 0);
    const pageDeleteProtected = await snapshot({
      type: "testAction",
      version: 81,
      actions: [
        { selector: '.selection-path button[aria-current="true"]', event: "click" },
        key("keydown", "Delete"),
        { selector: ".delete-control", event: "click" },
      ],
    });
    assert.strictEqual(pageDeleteProtected.missingTarget, undefined);
    assert.strictEqual(
      (pageDeleteProtected.editMessages as unknown[]).length,
      (pageSwitch.editMessages as unknown[]).length,
      "canvas and toolbar Delete leave the selected Page intact",
    );
    const reorderedPage = await snapshot({
      type: "testAction",
      version: 82,
      actions: [
        { selector: '.collection-editor button[aria-label="Move up: PageB"]', event: "click" },
      ],
    });
    const pageRequest = (
      reorderedPage.editMessages as { requestId: number; operations: unknown[] }[]
    ).at(-1)!;
    assert.deepStrictEqual(pageRequest.operations, [
      { type: "reorderPage", controlId: "page-b", index: 0 },
    ]);
    await panel.webview.postMessage({ type: "editResult", requestId: pageRequest.requestId });
    await snapshot(collectionDocument(83));
    const tabSelected = await snapshot({
      type: "testAction",
      version: 84,
      actions: [
        parentPointer('.control[data-control-id="tabs"]', "pointerdown", 310, 45),
        parentPointer(".designer-viewport", "pointerup", 310, 45),
        {
          selector: '.collection-editor[aria-label="Tabs"] [data-property-field="caption"] input',
          event: "input",
          props: { value: "<script>caption</script>" },
        },
        {
          selector: '.collection-editor[aria-label="Tabs"] [data-property-field="caption"] input',
          event: "keydown",
          props: { key: "Enter" },
        },
      ],
    });
    const tabRequest = (
      tabSelected.editMessages as { requestId: number; operations: unknown[] }[]
    ).at(-1)!;
    assert.deepStrictEqual(tabRequest.operations, [
      {
        type: "setTabProperty",
        controlId: "tabs",
        index: 0,
        field: "caption",
        value: "<script>caption</script>",
      },
    ]);
    assert.deepStrictEqual(tabSelected.violations, []);
    await panel.webview.postMessage({ type: "editResult", requestId: tabRequest.requestId });
    const treeSelected = await snapshot({
      type: "testAction",
      version: 90,
      actions: [
        { selector: "#designer-tab-structure", event: "click" },
        {
          selector:
            '[data-structure-id="multi"] > .structure-tree-row > .structure-tree-disclosure',
          event: "click",
          props: { expandIfCollapsed: true },
        },
        {
          selector:
            '[data-structure-id="page-b"] > .structure-tree-row > .structure-tree-disclosure',
          event: "click",
          props: { expandIfCollapsed: true },
        },
        { selector: '[data-structure-id="child-b"] > .structure-tree-row', event: "click" },
      ],
    });
    assert.strictEqual(treeSelected.missingTarget, undefined);
    assert.strictEqual(treeSelected.activeRightTab, "designer-tab-structure");
    assert.deepStrictEqual(treeSelected.selectedIds, ["child-b"]);
    assert.ok((treeSelected.worldPositions as { id: string }[]).some((c) => c.id === "child-b"));
    assert.ok(!(treeSelected.worldPositions as { id: string }[]).some((c) => c.id === "child-a"));
    assert.strictEqual(
      (treeSelected.editMessages as unknown[]).length,
      (tabSelected.editMessages as unknown[]).length,
      "Structure selection changes no source",
    );
    assert.strictEqual(treeSelected.codiconLoaded, true);
    const productivityDocument = designerDocument({
      form: { name: "Productivity", width: 400, height: 240 },
      controls: [
        { id: "a", name: "Label1", type: "Label", left: 20, top: 20, width: 60, height: 18 },
        { id: "b", name: "TextBox2", type: "TextBox", left: 110, top: 20, width: 60, height: 18 },
      ],
    });
    const productivityMessage = (version: number): Extract<HostMessage, { type: "document" }> => ({
      type: "document",
      version,
      document: productivityDocument,
      editable: true,
      productivityEditable: true,
      structuralEditable: true,
      propertyEditable: true,
      propertyGrid: {
        form: { descriptors: [], values: {} },
        controls: {
          a: {
            descriptors: [
              { field: "tabIndex", valueType: "integer", nullable: true, required: false },
            ],
            values: { tabIndex: { present: false, value: null } },
          },
          b: {
            descriptors: [
              { field: "tabIndex", valueType: "integer", nullable: true, required: false },
            ],
            values: { tabIndex: { present: true, value: 8 } },
          },
        },
      },
    });
    await snapshot(productivityMessage(95));
    const selectedInterior = await snapshot({
      type: "testAction",
      version: 96,
      actions: [
        { selector: ".designer-toolbar select", event: "change", props: { value: "1" } },
        { selector: ".snap-toggle input", event: "change", props: { checked: false } },
        {
          selector: '.control[data-control-id="a"]',
          event: "pointerdown",
          props: { pointerId: 1, button: 0, formPoints: [50, 29], hitTest: true },
        },
        parentPointer(".designer-viewport", "pointerup", 50, 29),
      ],
    });
    assert.deepStrictEqual(
      selectedInterior.selectedIds,
      ["a"],
      "interior click selects the control",
    );
    assert.strictEqual(selectedInterior.canvasFocused, true);
    const pressedInterior = await snapshot({
      type: "testAction",
      version: 97,
      actions: [
        {
          selector: '.control[data-control-id="a"]',
          event: "pointerdown",
          props: { pointerId: 1, button: 0, formPoints: [50, 29], hitTest: true },
        },
      ],
    });
    assert.deepStrictEqual(
      pressedInterior.selectedIds,
      ["a"],
      "holding the interior retains selection",
    );
    assert.strictEqual(pressedInterior.canvasFocused, true);
    assert.deepStrictEqual(pressedInterior.editMessages, selectedInterior.editMessages);
    const draggedInterior = await snapshot({
      type: "testAction",
      version: 98,
      actions: [
        parentPointer(".designer-viewport", "pointermove", 62, 35),
        parentPointer(".designer-viewport", "pointerup", 62, 35),
      ],
    });
    assert.deepStrictEqual(draggedInterior.selectedIds, ["a"]);
    const interiorEdits = draggedInterior.editMessages as {
      requestId: number;
      operations: unknown[];
    }[];
    assert.strictEqual(
      interiorEdits.length,
      (selectedInterior.editMessages as unknown[]).length + 1,
    );
    assert.deepStrictEqual(interiorEdits.at(-1)!.operations, [
      {
        type: "moveControl",
        controlId: "a",
        left: 32,
        top: 26,
      },
    ]);
    await panel.webview.postMessage({
      type: "editResult",
      requestId: interiorEdits.at(-1)!.requestId,
    });
    const beforeGroup = await snapshot(productivityMessage(100));
    const selectedGroup = await snapshot({
      type: "testAction",
      version: 101,
      actions: [
        { selector: ".designer-toolbar select", event: "change", props: { value: "1" } },
        { selector: ".snap-toggle input", event: "change", props: { checked: false } },
        {
          ...parentPointer('.control[data-control-id="a"]', "pointerdown", 25, 25),
          props: { pointerId: 1, button: 0, formPoints: [25, 25], hitTest: true },
        },
        parentPointer(".designer-viewport", "pointerup", 25, 25),
        {
          ...parentPointer('.control[data-control-id="b"]', "pointerdown", 115, 25),
          props: { pointerId: 1, button: 0, formPoints: [115, 25], ctrlKey: true, hitTest: true },
        },
        parentPointer(".designer-viewport", "pointerup", 115, 25),
      ],
    });
    assert.deepStrictEqual((selectedGroup.selectedIds as string[]).sort(), ["a", "b"]);
    assert.strictEqual(selectedGroup.groupSelection, true);
    assert.strictEqual(selectedGroup.handles, 0, "group selection has no group resize handles");
    const groupDragged = await snapshot({
      type: "testAction",
      version: 102,
      actions: [
        {
          ...parentPointer('.control[data-control-id="a"]', "pointerdown", 25, 25),
          props: { pointerId: 1, button: 0, formPoints: [25, 25], hitTest: true },
        },
        parentPointer(".designer-viewport", "pointermove", 37, 31),
        parentPointer(".designer-viewport", "pointerup", 37, 31),
      ],
    });
    const groupRequests = groupDragged.editMessages as {
      requestId: number;
      operations: unknown[];
    }[];
    assert.strictEqual(groupRequests.length, (beforeGroup.editMessages as unknown[]).length + 1);
    const groupOperation = groupRequests.at(-1)!.operations[0] as {
      type: string;
      moves: { controlId: string; left: number; top: number }[];
    };
    assert.strictEqual(groupRequests.at(-1)!.operations.length, 1);
    assert.strictEqual(groupOperation.type, "moveControls", JSON.stringify(groupRequests.at(-1)));
    assert.deepStrictEqual(
      [
        {
          ...groupOperation,
          moves: [...groupOperation.moves].sort((a, b) => a.controlId.localeCompare(b.controlId)),
        },
      ],
      [
        {
          type: "moveControls",
          moves: [
            { controlId: "a", left: 32, top: 26 },
            { controlId: "b", left: 122, top: 26 },
          ],
        },
      ],
    );
    await panel.webview.postMessage({
      type: "editResult",
      requestId: groupRequests.at(-1)!.requestId,
    });
    await snapshot(productivityMessage(103));
    const copiedGroup = await snapshot({
      type: "testAction",
      version: 104,
      actions: [
        { selector: ".designer-viewport", event: "keydown", props: { key: "c", ctrlKey: true } },
      ],
    });
    const copyRequests = copiedGroup.clipboardMessages as {
      requestId: number;
      action: string;
      controlIds: string[];
    }[];
    assert.strictEqual(copyRequests.at(-1)!.action, "copy");
    assert.deepStrictEqual([...copyRequests.at(-1)!.controlIds].sort(), ["a", "b"]);
    assert.strictEqual(
      (copiedGroup.editMessages as unknown[]).length,
      groupRequests.length,
      "copy emits no source edit",
    );
    await panel.webview.postMessage({
      type: "editResult",
      requestId: copyRequests.at(-1)!.requestId,
    });
    const reorderedLocalTabs = await snapshot({
      type: "testAction",
      version: 104,
      actions: [
        { selector: "#designer-tab-tabOrder", event: "click" },
        { selector: '.tab-order-editor button[aria-label="Move down: TextBox2"]', event: "click" },
        { selector: "#designer-tab-structure", event: "click" },
        { selector: "#designer-tab-tabOrder", event: "click" },
      ],
    });
    assert.deepStrictEqual(reorderedLocalTabs.tabOrderItems, ["0. Label1", "1. TextBox2"]);
    assert.strictEqual(
      (reorderedLocalTabs.editMessages as unknown[]).length,
      groupRequests.length,
      "right tabs retain local order without applying",
    );
    // Restore the original order before testing the established normalization.
    await snapshot({
      type: "testAction",
      version: 104,
      actions: [
        { selector: '.tab-order-editor button[aria-label="Move down: Label1"]', event: "click" },
      ],
    });
    const assignedTabs = await snapshot({
      type: "testAction",
      version: 105,
      actions: [
        { selector: "#designer-tab-tabOrder", event: "click" },
        { selector: ".tab-order-editor > button", event: "click" },
      ],
    });
    const tabOrderRequest = (
      assignedTabs.editMessages as { requestId: number; operations: unknown[] }[]
    ).at(-1)!;
    assert.deepStrictEqual(tabOrderRequest.operations, [
      { type: "setTabOrder", parentId: "", controlIds: ["b", "a"] },
    ]);
    assert.deepStrictEqual(assignedTabs.violations, []);
    assert.deepStrictEqual(assignedTabs.runtimeErrors, []);
    await panel.webview.postMessage({ type: "editResult", requestId: tabOrderRequest.requestId });
    const root = await snapshot({
      type: "document",
      version: 110,
      document: interactive,
      editable: true,
      propertyEditable: true,
      propertyGrid,
    });
    const priorRootEdits = (root.editMessages as unknown[]).length;
    const rootPreview = await snapshot({
      type: "testAction",
      version: 111,
      actions: [
        { selector: ".designer-toolbar select", event: "change", props: { value: "1.5" } },
        { selector: ".selection-path button", event: "click" },
        {
          selector: '[data-form-handle="se"]',
          event: "pointerdown",
          props: { pointerId: 1, button: 0, formPoints: [0, 0], handleCenter: true, hitTest: true },
        },
        {
          selector: ".designer-viewport",
          event: "pointermove",
          props: { pointerId: 1, formPoints: [0, 0], handleDelta: [20, 15] },
        },
      ],
    });
    assert.strictEqual(rootPreview.formHandles, 3);
    assert.strictEqual(
      (rootPreview.editMessages as unknown[]).length,
      priorRootEdits,
      "form preview does not edit source",
    );
    assert.ok(
      Math.abs(parseFloat(String(rootPreview.formWidth)) - ((interactive.width + 20) * 96) / 72) <
        0.05,
    );
    assert.ok(
      Math.abs(parseFloat(String(rootPreview.formHeight)) - ((interactive.height + 15) * 96) / 72) <
        0.05,
    );
    const rootResized = await snapshot({
      type: "testAction",
      version: 112,
      actions: [
        {
          selector: ".designer-viewport",
          event: "pointerup",
          props: { pointerId: 1, formPoints: [0, 0], handleDelta: [20, 15] },
        },
      ],
    });
    const rootRequests = rootResized.editMessages as { operations: unknown[] }[];
    assert.strictEqual(rootRequests.length, priorRootEdits + 1);
    assert.deepStrictEqual(rootRequests.at(-1)!.operations, [
      { type: "setFormProperty", field: "width", value: interactive.width + 20 },
      { type: "setFormProperty", field: "height", value: interactive.height + 15 },
    ]);
    assert.deepStrictEqual(rootResized.violations, []);
    assert.deepStrictEqual(rootResized.runtimeErrors, []);
    const resizeRequest = (rootResized.editMessages as { requestId: number }[]).at(-1)!;
    await panel.webview.postMessage({ type: "editResult", requestId: resizeRequest.requestId });
    const scrolledDocument = designerDocument({
      form: { name: "ScrollCheck", width: 1600, height: 1000 },
      controls: [
        {
          id: "scrolled",
          name: "ScrolledLabel",
          type: "Label",
          left: 180,
          top: 180,
          width: 72,
          height: 24,
        },
      ],
    });
    const beforeScroll = await snapshot({
      type: "document",
      version: 120,
      document: scrolledDocument,
      editable: true,
    });
    const scrolled = await snapshot({
      type: "testAction",
      version: 121,
      actions: [
        { selector: ".designer-toolbar select", event: "change", props: { value: "1.5" } },
        { selector: ".snap-toggle input", event: "change", props: { checked: false } },
        { selector: ".canvas-scroll", event: "scroll", props: { scrollLeft: 160, scrollTop: 180 } },
        {
          selector: '.control[data-control-id="scrolled"]',
          event: "pointerdown",
          props: {
            pointerId: 1,
            button: 0,
            formPoints: [190, 190],
            hitTest: true,
            preserveScroll: true,
          },
        },
        {
          selector: ".designer-viewport",
          event: "pointermove",
          props: { pointerId: 1, formPoints: [198, 196] },
        },
        {
          selector: ".designer-viewport",
          event: "pointerup",
          props: { pointerId: 1, formPoints: [198, 196] },
        },
      ],
    });
    assert.strictEqual(scrolled.missingTarget, undefined);
    const scrollState = scrolled.canvasScroll as { left: number; top: number };
    assert.strictEqual(scrollState.left, 160, "canvas focus preserves horizontal scroll");
    assert.strictEqual(scrollState.top, 180, "canvas focus preserves vertical scroll");
    assert.strictEqual(
      (scrolled.editMessages as unknown[]).length,
      (beforeScroll.editMessages as unknown[]).length + 1,
    );
    assert.deepStrictEqual(
      (scrolled.editMessages as { operations: unknown[] }[]).at(-1)!.operations,
      [{ type: "moveControl", controlId: "scrolled", left: 188, top: 186 }],
    );
    assert.deepStrictEqual(scrolled.violations, []);
    assert.deepStrictEqual(scrolled.runtimeErrors, []);
    const colorModel = designerDocument({
      form: {
        name: "Colors",
        width: 300,
        height: 220,
        foreColor: { kind: "rgb", value: "#FF0000" },
      },
      controls: [
        { id: "defaultLabel", name: "DefaultLabel", type: "Label" },
        {
          id: "explicitLabel",
          name: "ExplicitLabel",
          type: "Label",
          foreColor: { kind: "rgb", value: "#0000FF" },
        },
        {
          id: "frame",
          name: "Frame1",
          type: "Frame",
          left: 0,
          top: 40,
          width: 200,
          height: 140,
          foreColor: { kind: "rgb", value: "#008000" },
        },
        { id: "nestedDefault", name: "NestedDefault", type: "Label", parentId: "frame" },
        {
          id: "nestedExplicit",
          name: "NestedExplicit",
          type: "Label",
          parentId: "frame",
          top: 30,
          foreColor: { kind: "rgb", value: "#0000FF" },
        },
      ],
    });
    const colorSource = JSON.stringify(colorModel);
    const colors = await snapshot({ type: "document", version: 122, document: colorModel });
    assert.strictEqual(colors.formColor, "rgb(255, 0, 0)");
    assert.deepStrictEqual(
      colors.controlColors,
      {
        defaultLabel: "rgb(17, 17, 17)",
        explicitLabel: "rgb(0, 0, 255)",
        frame: "rgb(0, 128, 0)",
        nestedDefault: "rgb(17, 17, 17)",
        nestedExplicit: "rgb(0, 0, 255)",
      },
      "Form/Frame ForeColor must not recolor existing children without their own authored color",
    );
    colorModel.foreColor = { kind: "rgb", value: "#FF00FF" };
    const changedColors = await snapshot({ type: "document", version: 123, document: colorModel });
    assert.strictEqual(changedColors.formColor, "rgb(255, 0, 255)");
    assert.deepStrictEqual(
      changedColors.controlColors,
      colors.controlColors,
      "changing only Form.ForeColor preserves child colors",
    );
    colorModel.foreColor = { kind: "rgb", value: "#FF0000" };
    assert.strictEqual(
      JSON.stringify(colorModel),
      colorSource,
      "rendering must not author default ForeColor values",
    );
    assert.deepStrictEqual(changedColors.violations, []);
    assert.deepStrictEqual(changedColors.runtimeErrors, []);
    const defaultFonts = changedColors.controlFonts as Record<string, Record<string, string>>;
    colorModel.font = {
      name: "Arial",
      size: 24,
      bold: true,
      italic: true,
      underline: true,
      strikeThrough: true,
    };
    colorModel.controls.find((control) => control.id === "frame")!.font = {
      name: "Courier New",
      size: 18,
      bold: true,
      italic: true,
      underline: true,
      strikeThrough: true,
    };
    colorModel.controls.find((control) => control.id === "explicitLabel")!.font = {
      name: "Georgia",
      size: 15,
      bold: true,
      italic: false,
      underline: true,
      strikeThrough: false,
    };
    colorModel.controls.find((control) => control.id === "nestedExplicit")!.font = { bold: false };
    const fontSource = JSON.stringify(colorModel);
    const fonts = await snapshot({ type: "document", version: 124, document: colorModel });
    const renderedFonts = fonts.controlFonts as Record<string, Record<string, string>>;
    assert.deepStrictEqual(
      renderedFonts.defaultLabel,
      defaultFonts.defaultLabel,
      "Form Font must not change an existing child's default font",
    );
    assert.deepStrictEqual(
      renderedFonts.nestedDefault,
      defaultFonts.nestedDefault,
      "Frame Font must not change nested child fonts",
    );
    assert.deepStrictEqual(
      renderedFonts.nestedExplicit,
      defaultFonts.nestedExplicit,
      "a partial Font must not inherit unspecified members from its parent",
    );
    assert.deepStrictEqual(
      renderedFonts.explicitLabel,
      {
        weight: "700",
        style: "normal",
        family: "Georgia, sans-serif",
        size: "20px",
        decoration: "underline",
      },
      "explicit child Font remains effective independently",
    );
    assert.strictEqual(renderedFonts.frame.weight, "700", "container's own Font remains effective");
    assert.strictEqual(renderedFonts.frame.style, "italic");
    assert.strictEqual(renderedFonts.frame.decoration, "underline line-through");
    assert.strictEqual(
      JSON.stringify(colorModel),
      fontSource,
      "font preview defaults must never be authored",
    );
    assert.deepStrictEqual(fonts.violations, []);
    assert.deepStrictEqual(fonts.runtimeErrors, []);
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
    // VS Code/Windows may release its directory watcher just after tab close.
    await rm(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}
