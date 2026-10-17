import * as assert from "node:assert/strict";
import { render } from "preact";
import { act } from "preact/test-utils";
import { designerDocument, pointsToPixels } from "../src/userFormEditor/model";
import { controlMetadata, createControl, subtreeIds } from "../src/userFormEditor/toolbox";
import { initialControlOperations } from "../src/userFormEditor/collections";
import { designerStrings, isWebviewMessage } from "../src/userFormEditor/protocol";
import type { DesignerDocument, HostMessage, WebviewMessage } from "../src/userFormEditor/protocol";

const { JSDOM } = require("jsdom");
const dom = new JSDOM('<!doctype html><html><body><div id="app"></div></body></html>', {
  url: "https://designer.test/",
});
Object.assign(globalThis, { window: dom.window, document: dom.window.document });
dom.window.scrollTo = () => {};
Object.defineProperty(dom.window.HTMLElement.prototype, "onpointerdown", {
  configurable: true,
  value: null,
});
const messages: WebviewMessage[] = [];
Object.assign(globalThis, {
  acquireVsCodeApi: () => ({
    postMessage: (m: WebviewMessage) => messages.push(m),
    getState: () => undefined,
    setState: () => {},
  }),
});
const root = document.getElementById("app")!;
const edits = () =>
  messages.filter((m): m is Extract<WebviewMessage, { type: "edit" }> => m.type === "edit");
let model: DesignerDocument = designerDocument({
  form: { name: "Main", width: 480, height: 360 },
  controls: [
    { id: "label", name: "lAbEl1", type: "Label", left: 10, top: 10 },
    { id: "frame", name: "Frame1", type: "Frame", left: 200, top: 50 },
    { id: "child", parentId: "frame", name: "TextBox1", type: "TextBox" },
    { id: "multi", name: "MultiPage1", type: "MultiPage", selectedIndex: 0 },
    { id: "page", parentId: "multi", name: "Page1", type: "Page" },
  ],
});

function testGeneration() {
  const colliding = {
    ...model,
    name: "Label2",
    controls: [...model.controls, { ...model.controls[0], id: "control-taken" }],
  };
  const uuids = ["taken", "fresh"];
  const control = createControl(colliding, "Label", 13, 17, true, () => uuids.shift()!);
  assert.ok(control);
  assert.equal(control.id, "control-fresh");
  assert.equal(control.name, "Label3", "form and controls reserve case-insensitive names");
  assert.deepEqual([control.left, control.top], [16, 16]);
  assert.equal(
    createControl(colliding, "Label", 0, 0, false, () => "taken"),
    undefined,
    "a broken UUID source cannot hang",
  );
  for (const metadata of controlMetadata) {
    const created = createControl(model, metadata.type, 10000, -10, false, () => metadata.type)!;
    assert.ok(created);
    assert.equal(created.type, metadata.type);
    assert.match(created.name, /^[A-Za-z][A-Za-z0-9_]*$/);
    assert.equal("parentId" in created, false);
    assert.equal(created.top, 0);
    assert.equal(created.left, model.width - created.width);
    const rendered = designerDocument({
      form: { name: "Main" },
      controls: [{ id: "fallback", name: "Fallback", type: metadata.type }],
    });
    assert.deepEqual(
      [created.width, created.height],
      [rendered.controls[0].width, rendered.controls[0].height],
    );
    assert.equal(
      isWebviewMessage({
        type: "edit",
        requestId: 1,
        version: 1,
        operations: [{ type: "addControl", control: created }],
      }),
      metadata.type !== "MultiPage",
    );
    if (metadata.type === "MultiPage") assert.equal(created.selectedIndex, -1);
    if (metadata.type === "TabStrip") {
      assert.equal(created.selectedIndex, 0);
      assert.deepEqual(created.tabs, [
        { name: "Tab1", caption: "Tab1" },
        { name: "Tab2", caption: "Tab2" },
      ]);
    }
    const renamed = { ...created, name: "Renamed" };
    assert.equal(renamed.id, created.id);
  }
  assert.equal(createControl({ ...model, width: 1 }, "Label", 0, 0, false), undefined);
  assert.equal(createControl(model, "Label", Infinity, 0, false), undefined);
  const multi = createControl(model, "MultiPage", 0, 0, false, () => "owner")!;
  const pageIds = ["owner", "first", "first", "second"];
  const creation = initialControlOperations(model, multi, () => pageIds.shift()!)!;
  assert.equal(creation.length, 3);
  assert.ok(isWebviewMessage({ type: "edit", requestId: 1, version: 1, operations: creation }));
  const pages = creation.slice(1).map((op) => {
    assert.equal(op.type, "addPage");
    if (op.type !== "addPage") throw new Error("expected Page addition");
    assert.equal(op.control.parentId, multi.id);
    return op.control;
  });
  assert.deepEqual(
    pages.map((page) => page.name),
    ["Page2", "Page3"],
  );
  assert.deepEqual(
    pages.map((page) => page.id),
    ["control-first", "control-second"],
  );
  assert.equal(
    initialControlOperations(model, multi, () => "owner"),
    undefined,
  );
  const initialMessage = (operations: unknown[]) => ({
    type: "edit",
    requestId: 1,
    version: 1,
    operations,
  });
  const tabStrip = createControl(model, "TabStrip", 0, 0, false)!;
  assert.equal(
    isWebviewMessage(
      initialMessage([
        { type: "addControl", control: { ...tabStrip, tabs: [], selectedIndex: -1 } },
      ]),
    ),
    false,
  );
  for (const operations of [
    creation.slice(0, 2),
    [...creation, creation[1]],
    [creation[0], creation[1], { type: "addPage", control: { ...pages[1], parentId: "other" } }],
    [creation[0], creation[1], { type: "addPage", control: { ...pages[1], id: pages[0].id } }],
    [
      creation[0],
      creation[1],
      { type: "addPage", control: { ...pages[1], name: pages[0].name.toLowerCase() } },
    ],
  ])
    assert.equal(isWebviewMessage(initialMessage(operations)), false);
  assert.deepEqual(
    [...subtreeIds([...model.controls, { id: "grandchild", parentId: "child" }], "frame")],
    ["frame", "child", "grandchild"],
  );
  const valid = { type: "addControl", control };
  const message = (operations: unknown[]) => ({
    type: "edit",
    requestId: 1,
    version: 1,
    operations,
  });
  for (const payload of [
    { ...control, parentId: null },
    { ...control, id: "" },
    { ...control, name: "1Invalid" },
    { ...control, type: "Page" },
    { ...control, width: Infinity },
    { ...control, width: 0 },
    { ...control, tabs: [{}] },
    { ...control, selectedIndex: 1 },
  ])
    assert.equal(isWebviewMessage(message([{ type: "addControl", control: payload }])), false);
  assert.equal(isWebviewMessage(message([valid, valid])), false);
  assert.equal(
    isWebviewMessage(
      message([{ type: "removeControl", controlId: "frame", cascade: true, extra: true }]),
    ),
    false,
  );
  assert.equal(
    isWebviewMessage(message([{ type: "removeControl", controlId: "frame", cascade: "true" }])),
    false,
  );
}

async function host(message: HostMessage) {
  await act(() => window.dispatchEvent(new dom.window.MessageEvent("message", { data: message })));
}
async function refresh(version: number, structuralEditable = true, collectionEditable = true) {
  await host({
    type: "document",
    version,
    document: model,
    editable: false,
    structuralEditable,
    pageEditable: collectionEditable,
    tabEditable: collectionEditable,
  });
}
async function click(element: Element) {
  await act(() =>
    element.dispatchEvent(new dom.window.MouseEvent("click", { bubbles: true, button: 0 })),
  );
}
async function pointer(element: Element, left = 0, top = 0) {
  await act(() =>
    element.dispatchEvent(
      new dom.window.MouseEvent("pointerdown", {
        bubbles: true,
        button: 0,
        clientX: left,
        clientY: top,
      }),
    ),
  );
}
async function key(element: Element, value: string) {
  await act(() =>
    element.dispatchEvent(new dom.window.KeyboardEvent("keydown", { bubbles: true, key: value })),
  );
}

async function run() {
  testGeneration();
  const { App } = require("../webview/userFormDesigner/App");
  await act(() => {
    render(null, root);
    render(<App />, root);
  });
  await refresh(1, false);
  assert.equal(root.querySelectorAll("[data-toolbox-type]").length, 14);
  assert.equal(root.querySelector('[data-toolbox-type="Page"]'), null);
  assert.ok(root.textContent?.includes(designerStrings.structuralReadOnly));
  assert.equal(
    (root.querySelector('[data-toolbox-type="Label"]') as HTMLButtonElement).disabled,
    true,
  );
  assert.equal(edits().length, 0);
  await refresh(1);
  const undoButton = root.querySelector<HTMLButtonElement>(".history-undo")!;
  const redoButton = root.querySelector<HTMLButtonElement>(".history-redo")!;
  assert.equal(undoButton.disabled, true);
  assert.equal(redoButton.disabled, true);
  await host({ type: "historyState", canUndo: true, canRedo: false });
  assert.equal(undoButton.disabled, false);
  assert.equal(redoButton.disabled, true);
  for (const state of [
    { saving: true, conflicted: false },
    { saving: false, conflicted: true },
  ]) {
    await host({ type: "draftState", dirty: true, ...state });
    assert.equal(undoButton.disabled, true, "history is disabled while saving or conflicted");
  }
  await host({ type: "draftState", dirty: true, saving: false, conflicted: false });
  await click(undoButton);
  const historyRequest = messages.at(-1)!;
  assert.equal(historyRequest.type, "history");
  if (historyRequest.type !== "history") throw new Error("Expected history request");
  assert.equal(isWebviewMessage(historyRequest), true);
  assert.equal(isWebviewMessage({ ...historyRequest, action: "delete" }), false);
  assert.equal(isWebviewMessage({ ...historyRequest, requestId: -1 }), false);
  assert.equal(isWebviewMessage({ ...historyRequest, version: -1 }), false);
  assert.equal(undoButton.disabled, true, "only one history request can be in flight");
  await host({ type: "editResult", requestId: historyRequest.requestId });
  await host({ type: "historyState", canUndo: false, canRedo: true });
  assert.equal(redoButton.disabled, false);
  assert.ok(
    !root.textContent?.includes(designerStrings.readOnly),
    "structural-only servers are editable",
  );
  // The structural capability does not require geometry/property editing.
  let version = 1;
  await refresh(version, true, false);
  for (const type of ["MultiPage", "TabStrip"])
    assert.equal(
      (root.querySelector(`[data-toolbox-type="${type}"]`) as HTMLButtonElement).disabled,
      true,
    );
  assert.equal(
    (root.querySelector('[data-toolbox-type="Label"]') as HTMLButtonElement).disabled,
    false,
  );
  await refresh(version);
  for (const metadata of controlMetadata) {
    await click(root.querySelector(`[data-toolbox-type="${metadata.type}"]`)!);
    assert.ok(root.querySelector(".placing-control"));
    // Clicking an existing control still creates a root control, never parents
    // by overlap. Two zoom levels verify screen -> point conversion.
    const zoom = metadata.type === "Label" ? 2 : 1;
    const zoomSelect = root.querySelector<HTMLSelectElement>(".designer-toolbar select")!;
    await act(() => {
      zoomSelect.value = String(zoom);
      zoomSelect.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
    });
    const before = edits().length;
    await pointer(
      root.querySelector('[data-control-id="frame"]')!,
      pointsToPixels(24, zoom),
      pointsToPixels(32, zoom),
    );
    assert.equal(edits().length, before + 1);
    const operation = edits().at(-1)!.operations[0];
    assert.equal(operation.type, "addControl");
    if (operation.type !== "addControl") throw new Error("expected add");
    assert.equal(operation.control.left, 24);
    assert.equal(operation.control.top, 32);
    assert.equal(operation.control.type, metadata.type);
    assert.equal("parentId" in operation.control, false);
    assert.equal(edits().at(-1)!.version, version);
    // Disable placement during the pending request and acknowledge in either
    // response/preview order. Success selects the authoritative new control.
    await pointer(root.querySelector(".form")!);
    assert.equal(edits().length, before + 1);
    model = {
      ...model,
      controls: [
        ...model.controls,
        {
          ...operation.control,
          selectedIndex: metadata.type === "MultiPage" ? 0 : operation.control.selectedIndex,
          approximate: false,
        },
        ...edits()
          .at(-1)!
          .operations.slice(1)
          .flatMap((op) =>
            op.type === "addPage"
              ? [{ ...op.control, left: 0, top: 0, width: 0, height: 0, approximate: true }]
              : [],
          ),
      ],
    };
    version++;
    if (metadata.type === "Label") {
      await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
      await refresh(version);
    } else {
      await refresh(version);
      await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
    }
    assert.ok(root.querySelector(`.selection-outline[data-control-id="${operation.control.id}"]`));
    assert.equal(root.querySelector(".placing-control"), null);
    if (metadata.type === "MultiPage") {
      assert.equal(edits().at(-1)!.operations.length, 3);
      assert.equal(
        root.querySelectorAll(`[data-control-id="${operation.control.id}"] .tabs button`).length,
        2,
      );
    }
    if (metadata.type === "TabStrip")
      assert.deepEqual(
        Array.from(
          root.querySelectorAll(`[data-control-id="${operation.control.id}"] .tabs span`),
        ).map((element) => element.textContent),
        ["Tab1", "Tab2"],
      );
  }
  const beforeCancel = edits().length;
  await click(root.querySelector('[data-toolbox-type="Label"]')!);
  await key(root.querySelector('[data-toolbox-type="Label"]')!, "Escape");
  assert.equal(root.querySelector(".placing-control"), null);
  assert.equal(edits().length, beforeCancel);
  await click(root.querySelector('[data-toolbox-type="Label"]')!);
  await key(root.querySelector(".designer-viewport")!, "Escape");
  assert.equal(root.querySelector(".placing-control"), null);

  await pointer(root.querySelector('[data-control-id="label"]')!);
  const pageTab = root.querySelector<HTMLButtonElement>('[data-page-tab="page"]')!;
  await act(() => pageTab.focus());
  assert.equal(document.activeElement, pageTab);
  assert.ok(root.querySelector('.selection-outline[data-control-id="label"]'));
  const beforeFocusedPageDelete = edits().length;
  await key(pageTab, "Delete");
  assert.equal(
    edits().length,
    beforeFocusedPageDelete,
    "Delete from a focused Page tab must not delete the previously selected control",
  );
  await act(() => (root.querySelector(".designer-viewport") as HTMLElement).focus());
  const viewport = root.querySelector(".designer-viewport")!;
  for (const guard of [
    { repeat: true },
    { isComposing: true },
    { ctrlKey: true },
    { metaKey: true },
    { altKey: true },
    { shiftKey: true },
  ]) {
    await act(() =>
      viewport.dispatchEvent(
        new dom.window.KeyboardEvent("keydown", { key: "Backspace", bubbles: true, ...guard }),
      ),
    );
    assert.equal(edits().length, beforeFocusedPageDelete, "guarded Backspace never deletes");
  }
  await key(viewport, "Backspace");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "removeControl", controlId: "label", cascade: true },
  ]);
  await host({
    type: "editResult",
    requestId: edits().at(-1)!.requestId,
    error: { code: "editCancelled", message: "Cancelled" },
  });
  assert.equal(root.querySelector('[role="alert"]'), null);
  assert.ok(root.querySelector('[data-control-id="label"]'));
  await pointer(root.querySelector('[data-control-id="frame"]')!);
  await click(root.querySelector(".delete-control")!);
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "removeControl", controlId: "frame", cascade: true },
  ]);
  const beforeStaleReply = edits().length;
  await refresh(++version);
  assert.equal(
    (root.querySelector(".delete-control") as HTMLButtonElement).disabled,
    true,
    "an external preview cannot acknowledge a pending structural confirmation",
  );
  await click(root.querySelector(".delete-control")!);
  assert.equal(edits().length, beforeStaleReply);
  await host({
    type: "editResult",
    requestId: edits().at(-1)!.requestId,
    error: { code: "staleDocument", message: "Source changed" },
  });
  assert.ok(root.querySelector('[data-control-id="frame"]'));
  await click(root.querySelector(".delete-control")!);
  assert.equal(edits().length, beforeStaleReply + 1);
  model = { ...model, controls: model.controls.filter((c) => !["frame", "child"].includes(c.id)) };
  await refresh(++version);
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
  assert.equal(root.querySelector('[data-control-id="frame"]'), null);
  assert.equal(root.querySelector('[data-control-id="child"]'), null);
  assert.equal((root.querySelector(".delete-control") as HTMLButtonElement).disabled, true);
  await click(root.querySelector('[data-page-tab="page"]')!);
  assert.equal((root.querySelector(".delete-control") as HTMLButtonElement).disabled, true);
  const beforePage = edits().length;
  await key(root.querySelector(".designer-viewport")!, "Delete");
  assert.equal(edits().length, beforePage);

  await pointer(root.querySelector('[data-control-id="label"]')!);
  await host({
    type: "document",
    version,
    document: model,
    structuralEditable: true,
    propertyEditable: true,
    propertyGrid: {
      form: { descriptors: [], values: {} },
      controls: {
        label: {
          descriptors: [{ field: "name", valueType: "string", required: true, nullable: false }],
          values: { name: { present: true, value: "lAbEl1" } },
        },
      },
    },
  });
  const nameInput = root.querySelector<HTMLInputElement>('[data-property-field="name"] input')!;
  await key(nameInput, "Delete");
  await key(nameInput, "Backspace");
  assert.equal(edits().length, beforePage, "Delete in a property input cannot delete a control");
  await act(() => {
    nameInput.value = "RenamedLabel";
    nameInput.dispatchEvent(new dom.window.Event("input", { bubbles: true }));
  });
  await key(nameInput, "Enter");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setControlProperty", controlId: "label", field: "name", value: "RenamedLabel" },
  ]);
  model = {
    ...model,
    controls: model.controls.map((c) => (c.id === "label" ? { ...c, name: "RenamedLabel" } : c)),
  };
  await refresh(++version);
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
  assert.ok(
    root.querySelector('.selection-outline[data-control-id="label"]'),
    "rename retains stable identity and selection",
  );
  const beforeInvalid = edits().length;

  await host({
    type: "invalidDocument",
    version: ++version,
    error: { code: "invalid", message: "Invalid source" },
  });
  assert.equal(
    (root.querySelector('[data-toolbox-type="Label"]') as HTMLButtonElement).disabled,
    true,
  );
  assert.ok(root.textContent?.includes(designerStrings.lastValid));
  await refresh(version);
  model = { ...model, width: 1, height: 1 };
  await refresh(++version);
  await click(root.querySelector('[data-toolbox-type="Label"]')!);
  await pointer(root.querySelector(".form")!);
  assert.ok(root.textContent?.includes(designerStrings.cannotInsert));
  assert.equal(edits().length, beforeInvalid);
  await act(() => render(null, root));
  process.stdout.write("UserForm structural generation and DOM assertions passed\n");
}
run().catch((error) => {
  process.stderr.write(`${error.stack ?? error}\n`);
  process.exitCode = 1;
});
