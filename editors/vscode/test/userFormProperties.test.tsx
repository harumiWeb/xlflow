import * as assert from "node:assert/strict";
import { render } from "preact";
import { act } from "preact/test-utils";
import type {
  HostMessage,
  PropertyGridData,
  PropertyTarget,
  WebviewMessage,
} from "../src/userFormEditor/protocol";
import { designerStrings } from "../src/userFormEditor/protocol";
import { designerDocument } from "../src/userFormEditor/model";

// Keep DOM assertions in the same standalone esbuild runner style as rendering tests.
const { JSDOM } = require("jsdom");
const dom = new JSDOM('<!doctype html><html><body><div id="app"></div></body></html>', {
  url: "https://designer.test/",
});
Object.assign(globalThis, { window: dom.window, document: dom.window.document });
dom.window.scrollTo = () => {};
// jsdom 26 has no PointerEvent handler property. Expose browser event casing
// so Preact registers the same pointerdown event as the real Webview.
Object.defineProperty(dom.window.HTMLElement.prototype, "onpointerdown", {
  configurable: true,
  value: null,
});
const messages: WebviewMessage[] = [];
Object.assign(globalThis, {
  acquireVsCodeApi: () => ({
    postMessage: (message: WebviewMessage) => messages.push(message),
    getState: () => undefined,
    setState: () => {},
  }),
});
const root = document.getElementById("app")!;
const target = (): PropertyTarget => ({
  descriptors: [
    { field: "caption", valueType: "string", required: false, nullable: true },
    { field: "width", valueType: "number", required: false, nullable: false },
    { field: "tabIndex", valueType: "integer", required: false, nullable: false },
    { field: "visible", valueType: "boolean", required: false, nullable: false },
    { field: "value", valueType: "any", required: false, nullable: true },
    {
      field: "orientation",
      valueType: "string",
      required: false,
      nullable: false,
      allowedValues: ["horizontal", "vertical"],
    },
    {
      field: "selectedIndex",
      valueType: "integer",
      required: false,
      nullable: false,
      allowedValues: ["0", "1"],
    },
    {
      field: "enabled",
      valueType: "boolean",
      required: false,
      nullable: false,
      allowedValues: ["true", "false"],
    },
  ],
  values: {
    caption: { present: true, value: "Source caption" },
    width: { present: true, value: 100 },
    tabIndex: { present: true, value: 0 },
    visible: { present: true, value: false },
    value: { present: true, value: "false" },
    orientation: { present: false, value: null },
    selectedIndex: { present: true, value: 0 },
    enabled: { present: true, value: false },
  },
});
const model = designerDocument({
  form: { name: "Main", width: 240, height: 180 },
  controls: [
    { id: "form", name: "FormIDLabel", type: "Label", caption: "Control render" },
    { id: "label", name: "Label1", type: "Label", caption: "Render caption" },
    { id: "multi", name: "Multi", type: "MultiPage", selectedIndex: 0 },
    {
      id: "first",
      parentId: "multi",
      name: "FirstPage",
      type: "Page",
      caption: "First",
      zIndex: 2,
    },
    {
      id: "second",
      parentId: "multi",
      name: "SecondPage",
      type: "Page",
      caption: "Second",
      zIndex: 1,
    },
    {
      id: "firstChild",
      parentId: "first",
      name: "FirstChild",
      type: "Label",
      caption: "First child",
    },
    {
      id: "secondChild",
      parentId: "second",
      name: "SecondChild",
      type: "Label",
      caption: "Second child",
    },
  ],
});
const pageTarget = (): PropertyTarget => ({
  ...target(),
  descriptors: target().descriptors.filter((d) => d.field === "caption"),
});
let grid: PropertyGridData = {
  form: target(),
  controls: {
    label: target(),
    first: pageTarget(),
    second: pageTarget(),
    form: { ...pageTarget(), values: { caption: { present: true, value: "Control source" } } },
  },
};
const edits = () =>
  messages.filter((m): m is Extract<WebviewMessage, { type: "edit" }> => m.type === "edit");
async function host(message: HostMessage) {
  await act(() => {
    window.dispatchEvent(new dom.window.MessageEvent("message", { data: message }));
  });
}
async function refresh(
  version: number,
  propertyEditable = true,
  propertyGrid: PropertyGridData | undefined = grid,
) {
  await host({
    type: "document",
    version,
    document: model,
    editable: false,
    propertyEditable,
    propertyGrid,
  });
}
function row(field: string): HTMLElement {
  const result = root.querySelector<HTMLElement>(`[data-property-field="${field}"]`);
  assert.ok(result, `property row ${field}`);
  return result;
}
function input(field: string): HTMLInputElement {
  const result = row(field).querySelector("input");
  assert.ok(result);
  return result;
}
async function type(field: string, value: string) {
  await act(() => {
    const element = input(field);
    element.value = value;
    element.dispatchEvent(new dom.window.Event("input", { bubbles: true }));
  });
}
async function key(field: string, value: string) {
  await act(() => {
    input(field).dispatchEvent(
      new dom.window.KeyboardEvent("keydown", { key: value, bubbles: true }),
    );
  });
}
async function blur(field: string) {
  await act(() => {
    input(field).dispatchEvent(new dom.window.FocusEvent("blur", { bubbles: true }));
  });
}
async function choose(field: string, value: string, index = 0) {
  await act(() => {
    const select = row(field).querySelectorAll("select")[index];
    assert.ok(select);
    select.value = value;
    select.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
  });
}
async function acknowledge(error?: Extract<HostMessage, { type: "editResult" }>["error"]) {
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId, error });
}

async function run() {
  const { App } = require("../webview/userFormDesigner/App");
  await act(() => {
    render(null, root);
    render(<App />, root);
  });
  await host({ type: "document", version: 1, document: model, editable: false });
  assert.ok(root.textContent?.includes(designerStrings.propertyReadOnly));
  assert.equal(
    root.querySelectorAll("[data-property-field]").length,
    0,
    "old server must not get guessed descriptors",
  );
  await refresh(1);
  assert.equal(input("caption").value, "Source caption", "authoring values, not render defaults");
  assert.equal(input("width").disabled, false, "property capability is independent of geometry");
  assert.equal(row("visible").querySelector("select")!.value, "false");
  assert.ok(row("orientation").textContent?.includes(designerStrings.unset));
  assert.equal(row("value").querySelector("select")!.value, "string", "string false retains type");
  assert.equal(input("value").value, "false");
  assert.equal(edits().length, 0, "rendering does not write source");

  await type("caption", "日本語 <script>literal</script>");
  await key("caption", "Enter");
  await blur("caption");
  assert.equal(edits().length, 1, "Enter plus blur commits exactly once");
  assert.deepEqual(edits()[0].operations, [
    { type: "setFormProperty", field: "caption", value: "日本語 <script>literal</script>" },
  ]);
  await host({ type: "editResult", requestId: edits()[0].requestId + 100 });
  assert.equal(input("width").disabled, true, "unrelated acknowledgement cannot unblock request");
  await acknowledge();
  assert.equal(
    input("width").disabled,
    false,
    "no-op acknowledgement unblocks without document refresh",
  );
  assert.equal(input("caption").value, "Source caption", "source remains authoritative");

  await type("width", "not a number");
  await key("width", "Enter");
  assert.equal(edits().length, 1);
  assert.equal(input("width").value, "not a number");
  assert.ok(row("width").textContent?.includes(designerStrings.invalidNumber));
  await type("tabIndex", "1.5");
  await blur("tabIndex");
  assert.ok(row("tabIndex").textContent?.includes(designerStrings.invalidInteger));
  await type("tabIndex", "9007199254740992");
  await key("tabIndex", "Enter");
  assert.ok(row("tabIndex").textContent?.includes(designerStrings.invalidInteger));
  await key("width", "Escape");
  await blur("width");
  assert.equal(input("width").value, "100");
  assert.equal(edits().length, 1);

  await type("width", "200");
  await key("width", "Enter");
  await acknowledge({
    code: "invalid",
    message: "Invalid geometry",
    diagnostics: [
      {
        operationIndex: 0,
        field: "form.width",
        code: "out_of_range",
        message: "Too wide",
        suggestion: "Use a smaller width",
      },
    ],
  });
  await refresh(1);
  assert.equal(input("width").value, "200");
  assert.ok(row("width").textContent?.includes("out_of_range: Too wide Use a smaller width"));
  await blur("width");
  assert.equal(edits().length, 2, "failed draft is not resent by delayed blur");
  await key("width", "Escape");
  await key("tabIndex", "Escape");

  await type("caption", "Stale draft");
  grid = {
    ...grid,
    form: {
      ...grid.form,
      values: { ...grid.form.values, caption: { present: true, value: "External caption" } },
    },
  };
  await refresh(2);
  assert.equal(input("caption").value, "Stale draft");
  assert.ok(row("caption").textContent?.includes(designerStrings.staleProperty));
  await key("caption", "Enter");
  await blur("caption");
  assert.equal(edits().length, 2);
  await type("caption", "Retry");
  await blur("caption");
  assert.equal(edits().at(-1)!.version, 2);
  grid.form.values.caption = { present: true, value: "Retry" };
  await refresh(3);
  await acknowledge();
  assert.equal(input("caption").value, "Retry");
  assert.ok(
    !row("caption").textContent?.includes(designerStrings.staleProperty),
    "own source refresh is not a stale draft",
  );
  await type("caption", "Discard");
  await key("caption", "Escape");
  await blur("caption");
  assert.equal(input("caption").value, "Retry");

  await choose("caption", "null");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setFormProperty", field: "caption", value: null },
  ]);
  await acknowledge();
  await choose("value", "boolean");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setFormProperty", field: "value", value: false },
  ]);
  await acknowledge();
  await choose("orientation", "vertical");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setFormProperty", field: "orientation", value: "vertical" },
  ]);
  await acknowledge();
  await choose("selectedIndex", "1");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setFormProperty", field: "selectedIndex", value: 1 },
  ]);
  await acknowledge();
  await choose("enabled", "true");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setFormProperty", field: "enabled", value: true },
  ]);
  await acknowledge();

  await act(() => {
    root
      .querySelector('[data-control-id="label"]')!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  assert.ok(root.querySelector(".property-grid h2")!.textContent?.includes("Label1"));
  await type("caption", "Control caption");
  await key("caption", "Enter");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setControlProperty", controlId: "label", field: "caption", value: "Control caption" },
  ]);
  await acknowledge({
    code: "invalid",
    message: "Rejected caption",
    diagnostics: [
      {
        operationIndex: 0,
        field: "controls[0].caption",
        code: "invalid_caption",
        message: "Cannot use this caption",
      },
    ],
  });
  await refresh(3);
  assert.equal(
    input("caption").value,
    "Control caption",
    "same-version refresh after rejection preserves draft",
  );
  assert.ok(row("caption").textContent?.includes("invalid_caption: Cannot use this caption"));
  await key("caption", "Escape");
  const beforeSelection = edits().length;
  await host({
    type: "document",
    version: 3,
    document: model,
    editable: true,
    propertyEditable: true,
    propertyGrid: grid,
  });
  await act(() => {
    (root.querySelector('[data-page-tab="first"]') as HTMLButtonElement).click();
  });
  assert.ok(root.querySelector(".property-grid h2")!.textContent?.includes("FirstPage"));
  assert.equal(
    root.querySelectorAll(".resize-handle").length,
    0,
    "Page tab selection never enables Page geometry",
  );
  assert.equal(
    root.querySelectorAll('[data-property-field="width"]').length,
    0,
    "Page properties follow server descriptors",
  );
  assert.ok(
    root.textContent?.includes("Second child"),
    "selecting a tab preserves compiler page rendering order/index",
  );
  assert.ok(!root.textContent?.includes("First child"));
  assert.equal(edits().length, beforeSelection, "selection is not a semantic edit");
  await act(() => {
    root
      .querySelector(".form")!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  assert.ok(root.querySelector(".property-grid h2")!.textContent?.includes("Main"));
  await refresh(4, false);
  assert.equal(input("caption").disabled, true);
  assert.equal(edits().length, beforeSelection);
  grid = {
    ...grid,
    form: {
      ...target(),
      values: {
        ...target().values,
        caption: { present: true, value: null },
        width: { present: true, value: 0 },
        value: { present: true, value: "" },
      },
    },
  };
  await refresh(5);
  assert.equal(row("caption").querySelector("select")!.value, "null");
  assert.ok(row("caption").textContent?.includes(designerStrings.nullValue));
  assert.ok(!row("caption").textContent?.includes(designerStrings.unset));
  assert.equal(input("width").value, "0");
  assert.equal(input("value").value, "");
  assert.ok(!row("value").textContent?.includes(designerStrings.unset));
  await host({
    type: "localization",
    strings: { ...designerStrings, properties: "プロパティ", propertyWidth: "幅", layout: "配置" },
  });
  assert.ok(root.querySelector(".property-grid h2")!.textContent?.includes("プロパティ"));
  assert.equal(row("width").querySelector("label")!.textContent, "幅");
  assert.equal(
    edits().length,
    beforeSelection,
    "source refresh and localization never create edits",
  );
  await act(() => {
    root
      .querySelector('[data-control-id="label"]')!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  await host({
    type: "document",
    version: 5,
    document: model,
    editable: true,
    propertyEditable: true,
    propertyGrid: grid,
  });
  const viewport = root.querySelector(".designer-viewport")!;
  const nudge = async () => {
    await act(() => {
      viewport.dispatchEvent(
        new dom.window.KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }),
      );
      viewport.dispatchEvent(
        new dom.window.KeyboardEvent("keyup", { key: "ArrowRight", bubbles: true }),
      );
    });
  };
  await nudge();
  assert.equal(edits().length, beforeSelection + 1);
  await host({
    type: "document",
    version: 6,
    document: {
      ...model,
      controls: model.controls.map((c) => (c.id === "label" ? { ...c, left: c.left + 1 } : c)),
    },
    editable: true,
    propertyEditable: true,
    propertyGrid: grid,
  });
  await nudge();
  assert.equal(
    edits().length,
    beforeSelection + 2,
    "new document alone acknowledges geometry and unblocks next gesture",
  );
  await acknowledge();
  await type("caption", "Concurrent draft");
  await key("caption", "Enter");
  grid = {
    ...grid,
    controls: {
      ...grid.controls,
      label: {
        ...grid.controls.label,
        values: {
          ...grid.controls.label.values,
          caption: { present: true, value: "Concurrent external caption" },
        },
      },
    },
  };
  await refresh(7);
  assert.equal(
    input("caption").value,
    "Concurrent draft",
    "preview before edit reply does not discard submitted draft",
  );
  assert.equal(
    input("caption").disabled,
    true,
    "property request waits for its response even after new preview",
  );
  await acknowledge({ code: "stale", message: "The edit used an older version" });
  assert.equal(input("caption").value, "Concurrent draft");
  assert.ok(row("caption").textContent?.includes(designerStrings.staleProperty));
  assert.ok(row("caption").textContent?.includes("stale: The edit used an older version"));
  await key("caption", "Escape");
  assert.equal(input("caption").value, "Concurrent external caption");
  await act(() => input("caption").focus());
  await type("caption", "Draft across selection");
  const beforeBlurSelection = edits().length;
  await act(() => {
    root
      .querySelector(".form")!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  assert.equal(edits().length, beforeBlurSelection + 1, "selection focus commits the old row");
  assert.deepEqual(edits().at(-1)!.operations, [
    {
      type: "setControlProperty",
      controlId: "label",
      field: "caption",
      value: "Draft across selection",
    },
  ]);
  await acknowledge({
    code: "invalid",
    message: "Rejected caption",
    diagnostics: [
      {
        operationIndex: 0,
        field: "controls[0].caption",
        code: "invalid_caption",
        message: "Cannot use this caption",
      },
    ],
  });
  // An unrelated edit reply must not replace the deferred response for Label1.
  await choose("caption", "string");
  await type("caption", "Other target edit");
  await key("caption", "Enter");
  await acknowledge();
  await act(() => {
    root
      .querySelector('[data-control-id="label"]')!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  assert.equal(input("caption").value, "Draft across selection");
  assert.ok(row("caption").textContent?.includes("invalid_caption: Cannot use this caption"));
  const afterRejection = edits().length;
  await blur("caption");
  assert.equal(edits().length, afterRejection, "returning to rejected draft never retries it");
  await key("caption", "Escape");
  assert.equal(input("caption").value, "Concurrent external caption");
  for (const field of ["width", "tabIndex"]) {
    const before = edits().length;
    await type(field, "9007199254740993");
    await key(field, "Enter");
    assert.equal(edits().length, before, "typed numeric descriptors cannot round input");
    assert.ok(row(field).querySelector('[role="alert"]'));
    await key(field, "Escape");
  }
  await choose("value", "number");
  for (const value of ["9007199254740993", "0.100000000000000000001", "1e-400"]) {
    const before = edits().length;
    await type("value", value);
    await key("value", "Enter");
    await blur("value");
    assert.equal(edits().length, before, "precision loss cannot write rounded source");
    assert.equal(input("value").value, value, "exact rejected draft is retained");
    assert.ok(row("value").querySelector('[role="alert"]'));
  }
  for (const value of ["0.1", "+01.00", ".5", "1e2", "9007199254740992"]) {
    const before = edits().length;
    await type("value", value);
    await key("value", "Enter");
    assert.equal(edits().length, before + 1);
    assert.equal((edits().at(-1)!.operations[0] as { value: number }).value, Number(value));
    await acknowledge();
    await choose("value", "number");
  }
  await key("value", "Escape");
  await type("caption", "Hidden panel draft");
  await key("caption", "Enter");
  // Showing a retained Webview replays its current document before the
  // provider's queued editResult; property pending must survive this replay.
  await refresh(7);
  assert.equal(input("caption").disabled, true);
  await acknowledge({ code: "staleDocument", message: "Hidden edit is stale" });
  assert.equal(input("caption").disabled, false, "replayed reply releases the busy state");
  assert.equal(input("caption").value, "Hidden panel draft");
  assert.ok(row("caption").textContent?.includes("staleDocument: Hidden edit is stale"));
  await key("caption", "Escape");
  await act(() => {
    root
      .querySelector(".form")!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  await choose("caption", "string");
  await type("caption", "Form-only draft");
  const beforeIdentity = edits().length;
  await act(() => {
    root
      .querySelector('[data-control-id="form"]')!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  assert.ok(root.querySelector(".property-grid h2")!.textContent?.includes("FormIDLabel"));
  assert.equal(
    input("caption").value,
    "Control source",
    "form draft cannot populate control ID form",
  );
  await key("caption", "Enter");
  assert.equal(edits().length, beforeIdentity, "selection cannot submit another target's draft");
  await type("caption", "Control-only draft");
  await key("caption", "Enter");
  assert.deepEqual(edits().at(-1)!.operations, [
    {
      type: "setControlProperty",
      controlId: "form",
      field: "caption",
      value: "Control-only draft",
    },
  ]);
  await acknowledge({ code: "invalid", message: "Control rejected" });
  await act(() => {
    root
      .querySelector(".form")!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  assert.equal(input("caption").value, "Form-only draft");
  await key("caption", "Enter");
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setFormProperty", field: "caption", value: "Form-only draft" },
  ]);
  await acknowledge({ code: "invalid", message: "Form rejected" });
  await act(() => {
    root
      .querySelector('[data-control-id="form"]')!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 }));
  });
  assert.equal(input("caption").value, "Control-only draft");
  assert.ok(row("caption").textContent?.includes("Control rejected"));
  assert.ok(!row("caption").textContent?.includes("Form rejected"));
  await act(() => {
    render(null, root);
  });
  dom.window.close();
  console.log("UserForm Property Grid DOM assertions passed.");
}
void run().catch((error) => {
  console.error(error);
  dom.window.close();
  process.exitCode = 1;
});
