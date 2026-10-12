import * as assert from "node:assert/strict";
import { render } from "preact";
import { act } from "preact/test-utils";
import { designerDocument } from "../src/userFormEditor/model";
import { displayPages, createPage, moveToPage } from "../src/userFormEditor/collections";
import { dropTarget, isControlVisible } from "../src/userFormEditor/geometry";
import { createControl } from "../src/userFormEditor/toolbox";
import { isWebviewMessage } from "../src/userFormEditor/protocol";
import type { HostMessage, PropertyTarget, WebviewMessage } from "../src/userFormEditor/protocol";

const { JSDOM } = require("jsdom");
const dom = new JSDOM('<!doctype html><html><body><div id="app"></div></body></html>', {
  url: "https://designer.test/",
});
Object.assign(globalThis, { window: dom.window, document: dom.window.document });
dom.window.scrollTo = () => {};
for (const name of ["onpointerdown", "onpointermove", "onpointerup", "onpointercancel"])
  Object.defineProperty(dom.window.HTMLElement.prototype, name, {
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
const model = designerDocument({
  form: { name: "Main", width: 600, height: 400 },
  controls: [
    {
      id: "multi",
      name: "Multi",
      type: "MultiPage",
      left: 40,
      top: 40,
      width: 240,
      height: 180,
      selectedIndex: 0,
    },
    { id: "first", name: "Page1", type: "Page", parentId: "multi", zIndex: 0 },
    { id: "second", name: "Page2", type: "Page", parentId: "multi", zIndex: 1 },
    {
      id: "first-child",
      name: "FirstChild",
      type: "Label",
      parentId: "first",
      left: 10,
      top: 10,
      width: 20,
      height: 10,
      caption: "First child",
    },
    {
      id: "second-child",
      name: "SecondChild",
      type: "Label",
      parentId: "second",
      left: 10,
      top: 10,
      width: 20,
      height: 10,
      caption: "Second child",
    },
    {
      id: "tabs",
      name: "Tabs",
      type: "TabStrip",
      left: 300,
      top: 40,
      tabs: [
        { name: "First", caption: "First tab" },
        { name: "Second", caption: "Second tab" },
      ],
    },
  ],
});
const tab = (name: string): PropertyTarget => ({
  descriptors: [
    { field: "name", valueType: "string", required: true, nullable: false },
    { field: "caption", valueType: "string", required: false, nullable: true },
    { field: "enabled", valueType: "boolean", required: false, nullable: true },
  ],
  values: {
    name: { present: true, value: name },
    caption: { present: false, value: null },
    enabled: { present: true, value: false },
  },
});
let version = 1;
let tabNames = ["First", "Second"];
const send = async (message: HostMessage) =>
  act(() => {
    window.dispatchEvent(new dom.window.MessageEvent("message", { data: message }));
  });
const load = async (pageEditable = true, tabEditable = true) =>
  send({
    type: "document",
    version: version++,
    document: model,
    editable: true,
    parentEditable: true,
    structuralEditable: true,
    pageEditable,
    tabEditable,
    propertyEditable: true,
    propertyGrid: {
      form: { descriptors: [], values: {} },
      controls: {
        first: { descriptors: [], values: {} },
        second: { descriptors: [], values: {} },
        tabs: {
          descriptors: [
            { field: "selectedIndex", valueType: "integer", required: false, nullable: true },
          ],
          values: { selectedIndex: { present: true, value: 0 } },
        },
      },
      tabs: { tabs: tabNames.map(tab) },
    },
  });
const edits = () =>
  messages.filter((m): m is Extract<WebviewMessage, { type: "edit" }> => m.type === "edit");
async function click(selector: string) {
  const el = root.querySelector<HTMLElement>(selector);
  assert.ok(el, selector);
  await act(() => el.click());
}
async function selectControl(id: string) {
  const el = root.querySelector(`[data-control-id="${id}"]`);
  assert.ok(el);
  await act(() => {
    el.dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
}
async function acknowledge() {
  const edit = edits().at(-1)!;
  await send({ type: "editResult", requestId: edit.requestId });
}
async function run() {
  const { App } = require("../webview/userFormDesigner/App");
  await act(() => {
    render(null, root);
    render(<App />, root);
  });
  await load();
  await click('[data-page-tab="second"]');
  assert.ok(root.querySelector('[data-control-id="second-child"]'));
  assert.equal(root.querySelector('[data-control-id="first-child"]'), null);
  assert.equal(edits().length, 0, "page switching is transient");
  assert.equal(model.controls[0].selectedIndex, 0, "source model is unchanged");
  assert.equal(root.querySelectorAll(".resize-handle").length, 0, "Page has no geometry handles");
  const pageDeleteCount = edits().length;
  for (const selector of [
    '[data-page-tab="second"]',
    '.collection-editor[aria-label="Pages"] li:nth-child(2) > button:first-child',
    '.selection-path button[aria-current="true"]',
  ]) {
    await click(selector);
    assert.equal(root.querySelector<HTMLButtonElement>(".delete-control")!.disabled, true);
    await click(".delete-control");
    await act(() => {
      const canvas = root.querySelector<HTMLElement>(".designer-viewport")!;
      canvas.focus();
      canvas.dispatchEvent(
        new dom.window.KeyboardEvent("keydown", { bubbles: true, key: "Delete" }),
      );
    });
    assert.equal(edits().length, pageDeleteCount, "Page deletion is only available in its list");
  }
  const projected = displayPages(model, { multi: "second" });
  assert.equal(isControlVisible(projected, "first-child"), false);
  assert.equal(dropTarget(projected, { left: 70, top: 80 }, new Set(), true)?.parentId, "second");
  assert.equal(
    dropTarget(projected, { left: 70, top: 80 }),
    undefined,
    "legacy Page parenting remains unavailable",
  );
  assert.equal(
    createControl(model, "TextBox", 10, 10, false, () => "fresh", "second", true)?.parentId,
    "second",
  );
  assert.equal(
    createControl(model, "TextBox", 10, 10, false, () => "fresh", "second"),
    undefined,
  );
  const fresh = createPage(model, "multi", () => "fresh")!;
  assert.equal(fresh.name, "Page3");
  assert.equal("width" in fresh, false);
  assert.equal(createPage(model, "tabs"), undefined);
  assert.deepEqual(moveToPage(model, "first-child", "second"), {
    type: "setParent",
    controlId: "first-child",
    parentId: "second",
    left: 10,
    top: 10,
  });
  assert.equal(moveToPage(model, "multi", "second"), undefined, "cycle prevention");
  await click('.collection-editor button[aria-label="Move up: Page2"]');
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "reorderPage", controlId: "second", index: 0 },
  ]);
  assert.ok(isWebviewMessage(edits().at(-1)));
  await acknowledge();
  await click('.collection-editor button[aria-label="Delete: Page2"]');
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "removePage", controlId: "second", cascade: true },
  ]);
  await acknowledge();
  for (const previewFirst of [true, false]) {
    model.controls[3].parentId = "first";
    await load();
    await click('[data-page-tab="first"]');
    await selectControl("first-child");
    const picker = root.querySelector<HTMLSelectElement>('select[aria-label="Move to Page"]')!;
    await act(() => {
      picker.value = "second";
      picker.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
    });
    assert.equal(edits().at(-1)!.operations[0].type, "setParent");
    model.controls[3].parentId = "second";
    if (previewFirst) {
      await load();
      await acknowledge();
    } else {
      await acknowledge();
      await load();
    }
    assert.ok(
      root.querySelector('.selection-outline[data-control-id="first-child"]'),
      "moving between Pages keeps selection in either response order",
    );
    assert.ok(root.querySelector('[data-control-id="first-child"]'));
    assert.equal(model.controls[0].selectedIndex, 0);
  }
  await selectControl("tabs");
  const tabSection = root.querySelector('.collection-editor[aria-label="Tabs"]')!;
  assert.ok(tabSection);
  await act(() => (tabSection.querySelector("button") as HTMLButtonElement).click());
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "addTab", controlId: "tabs", tab: { name: "Tab1", caption: "Tab1" } },
  ]);
  await acknowledge();
  const caption = tabSection.querySelector<HTMLInputElement>(
    '[data-property-field="caption"] input',
  )!;
  await act(() => {
    caption.value = "Updated";
    caption.dispatchEvent(new dom.window.Event("input", { bubbles: true }));
  });
  await act(() => {
    caption.dispatchEvent(new dom.window.KeyboardEvent("keydown", { bubbles: true, key: "Enter" }));
  });
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setTabProperty", controlId: "tabs", index: 0, field: "caption", value: "Updated" },
  ]);
  await acknowledge();
  const inputCaption = async (value: string) => {
    const input = root.querySelector<HTMLInputElement>(
      '.collection-editor[aria-label="Tabs"] [data-property-field="caption"] input',
    )!;
    await act(() => {
      input.value = value;
      input.dispatchEvent(new dom.window.Event("input", { bubbles: true }));
    });
  };
  const chooseTab = async (index: number) => {
    const button = root.querySelectorAll<HTMLButtonElement>(
      '.collection-editor[aria-label="Tabs"] li > button:first-child',
    )[index];
    await act(() => button.click());
  };
  await inputCaption("Draft First");
  await chooseTab(1);
  await inputCaption("Draft Second");
  await chooseTab(0);
  await click('.collection-editor button[aria-label="Move down: First"]');
  tabNames = ["Second", "First"];
  model.controls[5].tabs = [{ name: "Second" }, { name: "First" }];
  await load();
  await acknowledge();
  assert.equal(
    root.querySelector<HTMLInputElement>(
      '.collection-editor[aria-label="Tabs"] [data-property-field="caption"] input',
    )!.value,
    "Draft First",
    "reordering preserves the selected tab's own draft",
  );
  await chooseTab(0);
  assert.equal(
    root.querySelector<HTMLInputElement>(
      '.collection-editor[aria-label="Tabs"] [data-property-field="caption"] input',
    )!.value,
    "Draft Second",
    "the other tab also retains its own draft",
  );
  await inputCaption("Updated Second");
  const secondCaption = root.querySelector<HTMLInputElement>(
    '.collection-editor[aria-label="Tabs"] [data-property-field="caption"] input',
  )!;
  await act(() => {
    secondCaption.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", { bubbles: true, key: "Enter" }),
    );
  });
  assert.deepEqual(edits().at(-1)!.operations, [
    {
      type: "setTabProperty",
      controlId: "tabs",
      index: 0,
      field: "caption",
      value: "Updated Second",
    },
  ]);
  await acknowledge();
  const count = edits().length;
  await load(false, false);
  assert.equal(
    (root.querySelector('.collection-editor[aria-label="Tabs"] > button') as HTMLButtonElement)
      .disabled,
    true,
  );
  await click('[data-page-tab="first"]');
  assert.equal(
    (root.querySelector('.collection-editor[aria-label="Pages"] > button') as HTMLButtonElement)
      .disabled,
    true,
  );
  assert.equal(edits().length, count);
  for (const operation of [
    { type: "addPage", control: { ...fresh, width: 10 } },
    { type: "setTabProperty", controlId: "tabs", index: 0, field: "caption" },
    { type: "reorderTab", controlId: "tabs", index: 0, toIndex: -1 },
    { type: "addTab", controlId: "tabs", tab: { name: "First", parentId: "fake" } },
  ])
    assert.equal(
      isWebviewMessage({ type: "edit", requestId: 1, version: 1, operations: [operation] }),
      false,
    );
  await act(() => render(null, root));
  process.stdout.write("UserForm Page and TabStrip collection assertions passed\n");
}
run().catch((error) => {
  process.stderr.write(`${error.stack ?? error}\n`);
  process.exitCode = 1;
});
