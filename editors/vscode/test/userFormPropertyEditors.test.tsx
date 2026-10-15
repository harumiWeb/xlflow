import * as assert from "node:assert/strict";
import { render } from "preact";
import { act } from "preact/test-utils";
import { designerStrings } from "../src/userFormEditor/protocol";
import type {
  HostMessage,
  PropertyDescriptor,
  PropertyTarget,
  WebviewMessage,
} from "../src/userFormEditor/protocol";
import { designerDocument } from "../src/userFormEditor/model";
import {
  colorStyle,
  previewDocument,
} from "../webview/userFormDesigner/propertyEditors/appearance";
import { editorStrings } from "../webview/userFormDesigner/propertyEditors/strings";

const { JSDOM } = require("jsdom");
const dom = new JSDOM('<!doctype html><html><body><div id="app"></div></body></html>', {
  url: "https://designer.test/",
});
Object.assign(globalThis, { window: dom.window, document: dom.window.document });
for (const name of [
  "onpointerdown",
  "onpointermove",
  "onpointerup",
  "onpointercancel",
  "onlostpointercapture",
])
  Object.defineProperty(dom.window.HTMLElement.prototype, name, {
    configurable: true,
    value: null,
  });
dom.window.HTMLElement.prototype.scrollIntoView = () => {};
const messages: WebviewMessage[] = [];
Object.assign(globalThis, {
  acquireVsCodeApi: () => ({
    postMessage: (message: WebviewMessage) => messages.push(message),
    getState: () => undefined,
    setState: () => {},
  }),
});
const root = document.getElementById("app")!;
const descriptor = (
  field: string,
  valueType: PropertyDescriptor["valueType"],
  extra: Partial<PropertyDescriptor> = {},
): PropertyDescriptor => ({ field, valueType, required: false, nullable: true, ...extra });
const font = descriptor("font", "object", {
  editor: "font",
  children: [
    descriptor("name", "string"),
    descriptor("size", "number", { minimum: 0, exclusiveMinimum: true, maximum: 3276.75 }),
    descriptor("bold", "boolean"),
    descriptor("italic", "boolean"),
    descriptor("underline", "boolean"),
    descriptor("strikeThrough", "boolean"),
    descriptor("charset", "integer", {
      minimum: 0,
      maximum: 255,
      suggestions: [{ value: 128, label: "Shift JIS" }],
    }),
  ],
});
const color = descriptor("backColor", "object", {
  editor: "color",
  children: [
    descriptor("kind", "string", {
      required: true,
      nullable: false,
      allowedValues: ["rgb", "system", "ole"],
    }),
    descriptor("value", "string", {
      required: true,
      nullable: false,
      suggestions: [
        { value: "buttonFace", label: "buttonFace" },
        { value: "windowText", label: "windowText" },
      ],
    }),
  ],
});
const target = (): PropertyTarget => ({
  descriptors: [
    descriptor("caption", "string"),
    descriptor("width", "number", { minimum: 0 }),
    descriptor("height", "number", { minimum: 0 }),
    descriptor("enabled", "boolean", { editor: "boolean" }),
    font,
    color,
  ],
  values: {
    caption: { present: true, value: "Main" },
    width: { present: true, value: 240 },
    height: { present: true, value: 180 },
    enabled: { present: false, value: null },
    font: {
      present: true,
      value: { name: "My unavailable font", size: 12, bold: false, underline: null, charset: 17 },
    },
    backColor: { present: true, value: { kind: "ole", value: "0x81234567" } },
  },
});
const source = designerDocument({
  form: {
    name: "Main",
    width: 240,
    height: 180,
    font: { name: "My unavailable font", size: 12, bold: false },
  },
  controls: [
    {
      id: "label",
      name: "Label1",
      type: "Label",
      left: 12,
      top: 24,
      width: 100,
      height: 24,
      caption: "Sample",
    },
    { id: "image", name: "Image1", type: "Image", left: 120, top: 24, width: 80, height: 80 },
  ],
});
const grid = {
  form: target(),
  controls: {
    label: {
      ...target(),
      descriptors: [
        ...target().descriptors,
        descriptor("left", "number"),
        descriptor("top", "number"),
      ],
      values: {
        ...target().values,
        left: { present: true, value: 12 },
        top: { present: true, value: 24 },
        width: { present: true, value: 100 },
        height: { present: true, value: 24 },
      },
    },
    image: {
      descriptors: [
        descriptor("picture", "object", {
          editor: "picture",
          nullable: false,
          children: [descriptor("path", "string"), descriptor("remove", "boolean")],
        }),
      ],
      values: {},
    },
  },
};
async function host(message: HostMessage) {
  await act(() => window.dispatchEvent(new dom.window.MessageEvent("message", { data: message })));
}
async function refresh(version = 1) {
  await host({
    type: "document",
    version,
    document: source,
    propertyEditable: true,
    propertyEditorVersion: 1,
    propertyGrid: grid,
  });
}
const edits = () =>
  messages.filter(
    (message): message is Extract<WebviewMessage, { type: "edit" }> => message.type === "edit",
  );
const dialog = () => root.querySelector<HTMLElement>('[role="dialog"]')!;
function button(text: string): HTMLButtonElement {
  const result = Array.from(dialog().querySelectorAll("button")).find(
    (element) => element.textContent === text,
  );
  assert.ok(result, `dialog button ${text}`);
  return result;
}
function field(name: string): HTMLInputElement {
  const label = Array.from(dialog().querySelectorAll("label")).find(
    (element) => element.querySelector("span")?.textContent === name,
  );
  assert.ok(label, `editor field ${name}`);
  const input = label.querySelector("input");
  assert.ok(input);
  return input;
}
async function input(element: HTMLInputElement, text: string) {
  await act(() => {
    element.value = text;
    element.dispatchEvent(new dom.window.Event("input", { bubbles: true }));
  });
}
async function click(element: HTMLElement) {
  await act(() => element.click());
}
async function key(element: HTMLElement, name: string, extra: Record<string, unknown> = {}) {
  await act(() =>
    element.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", { key: name, bubbles: true, ...extra }),
    ),
  );
}
async function open(name: string) {
  const launcher = root.querySelector<HTMLButtonElement>(
    `[data-property-field="${name}"] .property-editor-launch`,
  )!;
  assert.ok(launcher && !launcher.disabled);
  launcher.focus();
  await click(launcher);
  assert.ok(dialog());
  return launcher;
}
async function select(id: string) {
  await act(() =>
    root
      .querySelector(`[data-control-id="${id}"]`)!
      .dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true, button: 0 })),
  );
}

async function run() {
  const { App } = require("../webview/userFormDesigner/App");
  await act(() => {
    render(null, root);
    render(<App />, root);
  });
  await refresh();
  assert.equal(edits().length, 0);
  assert.equal(
    root.querySelector('[data-property-field="caption"] .codicon-ellipsis'),
    null,
    "scalar state menu has no editor ellipsis",
  );
  const unsetBoolean = root.querySelector<HTMLInputElement>(
    '[data-property-field="enabled"] input[type="checkbox"]',
  )!;
  assert.ok(unsetBoolean.indeterminate, "unset boolean is distinct from false");

  await host({ type: "historyState", canUndo: true, canRedo: true });
  assert.equal(root.querySelector<HTMLButtonElement>(".history-undo")!.disabled, false);
  const launcher = await open("font");
  assert.equal(
    root.querySelector<HTMLButtonElement>(".history-undo")!.disabled,
    true,
    "modal disables Undo",
  );
  assert.equal(
    root.querySelector<HTMLButtonElement>(".history-redo")!.disabled,
    true,
    "modal disables Redo",
  );
  assert.equal(field("Name").value, "My unavailable font");
  assert.equal(field("Charset").value, "17", "arbitrary numeric charset remains editable");
  assert.ok(
    dialog().querySelector('option[value="128"]'),
    "charset named hints come from metadata",
  );
  assert.ok(field("Underline").indeterminate, "explicit null is not coerced to false");
  assert.equal(edits().length, 0, "opening never authors defaults");
  await input(field("Size"), "24");
  assert.equal(
    root.querySelector<HTMLElement>(".form")!.style.fontSize,
    "24pt",
    "local live Canvas preview",
  );
  assert.equal(edits().length, 0);
  await click(button("Cancel"));
  assert.equal(dialog(), null);
  assert.equal(
    root.querySelector<HTMLButtonElement>(".history-undo")!.disabled,
    false,
    "closing modal restores history availability",
  );
  assert.equal(root.querySelector<HTMLElement>(".form")!.style.fontSize, "12pt");
  assert.equal(document.activeElement, launcher, "focus returns to originating editor launcher");
  assert.equal(edits().length, 0, "Cancel emits no semantic operation");

  await open("font");
  await input(field("Size"), "18");
  await key(field("Size"), "Enter", { isComposing: true });
  assert.equal(edits().length, 0, "IME confirmation does not submit a dialog");
  const controls = dialog().querySelectorAll<HTMLElement>("input, select, button");
  controls[controls.length - 1].focus();
  await key(controls[controls.length - 1], "Tab");
  assert.equal(document.activeElement, controls[0], "Tab wraps within the modal");
  controls[0].focus();
  await key(controls[0], "Tab", { shiftKey: true });
  assert.equal(
    document.activeElement,
    controls[controls.length - 1],
    "Shift Tab wraps within the modal",
  );
  await key(field("Size"), "Enter");
  assert.equal(edits().length, 1, "OK submits one transaction");
  assert.deepEqual(edits()[0].operations, [
    {
      type: "setFormProperty",
      field: "font",
      value: { name: "My unavailable font", size: 18, bold: false, underline: null, charset: 17 },
    },
  ]);
  await host({
    type: "editResult",
    requestId: edits()[0].requestId,
    error: { code: "rejected", message: "Keep this draft" },
  });
  assert.equal(field("Size").value, "18");
  assert.ok(dialog().textContent?.includes("Keep this draft"));
  await refresh(2);
  assert.equal(field("Size").value, "18", "source change retains modal draft");
  assert.equal(button("OK").disabled, true, "source change blocks stale confirmation");
  assert.equal(
    root.querySelector<HTMLElement>(".form")!.style.fontSize,
    "12pt",
    "source change removes temporary preview",
  );
  await key(dialog(), "Escape");
  assert.equal(dialog(), null);
  await open("font");
  await input(field("Charset"), "257");
  assert.equal(button("OK").disabled, true);
  await input(field("Charset"), "42");
  assert.equal(button("OK").disabled, false);
  await input(field("Size"), "0.100000000000000000001");
  assert.equal(button("OK").disabled, true, "Font input preserves exact numeric precision");
  await input(field("Size"), "0");
  assert.equal(button("OK").disabled, true, "exclusive minimum comes from metadata");
  await input(field("Size"), "14");
  await click(button("OK"));
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
  assert.equal(dialog(), null, "matching success closes editor");
  await refresh(3);

  await open("backColor");
  assert.equal(dialog().querySelector<HTMLInputElement>('input[type="text"]')!.value, "0x81234567");
  const beforePalette = edits().length;
  const swatches = dialog().querySelectorAll<HTMLButtonElement>(".color-palette-swatch");
  assert.equal(swatches.length, 48, "classic palette offers 48 colors");
  await key(swatches[0], "ArrowDown");
  assert.equal(document.activeElement, swatches[8], "arrow navigation follows eight columns");
  assert.equal(edits().length, beforePalette, "browsing palette does not author a value");
  await click(swatches[18]);
  assert.equal(dialog().querySelector<HTMLInputElement>('input[type="text"]')!.value, "#FF8000");
  assert.equal(swatches[18].getAttribute("aria-pressed"), "true");
  assert.equal(root.querySelector<HTMLElement>(".form")!.style.backgroundColor, "rgb(255, 128, 0)");
  assert.equal(edits().length, beforePalette, "swatches update the preview until confirmation");
  await click(button("Cancel"));
  assert.equal(edits().length, beforePalette, "Cancel discards the palette selection");
  await open("backColor");
  await click(
    dialog().querySelector<HTMLButtonElement>('.color-palette-swatch[aria-label="#FF8000"]')!,
  );
  await click(button("OK"));
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setFormProperty", field: "backColor", value: { kind: "rgb", value: "#FF8000" } },
  ]);
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
  await open("backColor");
  const systemTab = dialog().querySelector<HTMLButtonElement>('[role="tab"][id$="system-tab"]')!;
  const beforeTabs = edits().length;
  await click(systemTab);
  assert.equal(
    dialog().querySelector<HTMLInputElement>('input[type="text"]')!.value,
    "0x81234567",
    "tab changes preserve unknown OLE draft",
  );
  assert.equal(edits().length, beforeTabs);
  assert.equal(
    dialog().querySelectorAll(".color-system-choice").length,
    2,
    "system colors are supplied by metadata",
  );
  await click(dialog().querySelector<HTMLButtonElement>(".color-system-choice")!);
  assert.equal(dialog().querySelectorAll("select")[1].value, "buttonFace");
  await click(button("Cancel"));
  await open("backColor");
  assert.equal(
    colorStyle({ kind: "ole", value: "0x81234567" }),
    undefined,
    "unknown OLE is never converted to guessed RGB",
  );
  const kind = dialog().querySelector("select")!;
  await act(() => {
    kind.value = "system";
    kind.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
  });
  const system = dialog().querySelectorAll("select")[1];
  assert.ok(system);
  await act(() => {
    system.value = "windowText";
    system.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
  });
  const colorSample = dialog().querySelector<HTMLElement>(".color-editor-sample")!;
  assert.equal(
    colorSample.style.backgroundColor,
    "",
    "selected foreground color cannot replace the identifier's themed background",
  );
  assert.equal(colorSample.querySelector(".color-identifier")!.textContent, "windowText");
  assert.equal(
    colorSample.querySelector<HTMLElement>(".color-swatch")!.style.backgroundColor.toLowerCase(),
    "canvastext",
    "selected color appears on the swatch only",
  );
  await act(() => {
    system.value = "buttonFace";
    system.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
  });
  await click(button("OK"));
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setFormProperty", field: "backColor", value: { kind: "system", value: "buttonFace" } },
  ]);
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });

  await click(root.querySelector<HTMLButtonElement>(".position-editor-launch")!);
  assert.equal(dialog().querySelectorAll("input").length, 2, "form exposes only Width and Height");
  await input(field("Width"), "260");
  await click(button("OK"));
  assert.deepEqual(
    edits().at(-1)!.operations,
    [{ type: "setGeometry", width: 260 }],
    "Position sends changed axes only",
  );
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
  await select("label");
  await click(root.querySelector<HTMLButtonElement>(".position-editor-launch")!);
  await input(field("Left"), "32");
  await input(field("Height"), "28");
  await click(button("OK"));
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setGeometry", controlId: "label", left: 32, height: 28 },
  ]);
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });

  await open("font");
  await input(field("Size"), "30");
  await host({ type: "editingUnavailable" });
  assert.equal(field("Size").value, "30");
  assert.equal(button("OK").disabled, true, "LSP restart invalidates a same-version editor");
  assert.equal(
    root.querySelector<HTMLElement>('[data-control-id="label"]')!.style.fontSize,
    "12pt",
    "LSP restart restores canonical authoring appearance, including legacy metadata",
  );
  await click(button("Cancel"));
  await refresh(4);
  await select("image");
  await open("picture");
  const choose = Array.from(dialog().querySelectorAll("button")).find((element) =>
    element.textContent?.startsWith("Choose BMP"),
  )!;
  const beforePicture = edits().length;
  await click(choose);
  const pick = messages.findLast((message) => message.type === "picturePick")!;
  assert.equal(pick.type, "picturePick");
  const selection = {
    token: "image-token",
    name: "sample.bmp",
    dataURI: "data:image/bmp;base64,Qk0=",
    width: 2,
    height: 3,
  };
  if (pick.type === "picturePick")
    await host({ type: "pictureSelection", requestId: pick.requestId, selection });
  assert.ok(
    root.querySelector('.control.image img[src="data:image/bmp;base64,Qk0="]'),
    "native selection previews on Canvas without source edits",
  );
  assert.equal(edits().length, beforePicture);
  await click(button("Cancel"));
  assert.equal(messages.at(-1)?.type, "pictureCancel");
  assert.equal(root.querySelector(".control.image img"), null);
  await open("picture");
  await click(
    Array.from(dialog().querySelectorAll("button")).find((element) =>
      element.textContent?.startsWith("Choose BMP"),
    )!,
  );
  const secondPick = messages.findLast((message) => message.type === "picturePick")!;
  if (secondPick.type === "picturePick")
    await host({
      type: "pictureSelection",
      requestId: secondPick.requestId,
      selection: { ...selection, token: "commit-token" },
    });
  await click(button("OK"));
  const pictureCommit = messages.findLast((message) => message.type === "pictureCommit")!;
  assert.equal(pictureCommit.type, "pictureCommit");
  assert.equal(edits().length, beforePicture, "picture commit uses host asset transaction only");
  if (pictureCommit.type === "pictureCommit")
    await host({
      type: "editResult",
      requestId: pictureCommit.requestId,
      error: { code: "import_failed", message: "Retain selected picture" },
    });
  assert.ok(dialog().querySelector("img"), "import failure retains selected preview and draft");
  await click(button("Remove picture"));
  await click(button("OK"));
  assert.deepEqual(edits().at(-1)!.operations, [
    { type: "setControlProperty", controlId: "image", field: "picture", value: { remove: true } },
  ]);
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
  assert.equal(dialog(), null);
  const before = JSON.stringify(source);
  const styled = previewDocument(source, { controlId: "label", values: { font: { size: 99 } } });
  assert.notEqual(styled, source);
  assert.equal(JSON.stringify(source), before, "preview never mutates authored model");
  const nullGrid = {
    ...grid,
    controls: {
      ...grid.controls,
      label: {
        ...grid.controls.label,
        values: {
          ...grid.controls.label.values,
          font: { present: true, value: null },
          backColor: { present: true, value: null },
        },
      },
    },
  };
  const nullSnapshot = JSON.stringify(nullGrid);
  await host({
    type: "document",
    version: 4,
    document: source,
    propertyEditable: true,
    propertyEditorVersion: 1,
    propertyGrid: nullGrid,
  });
  await select("label");
  const beforeNullObjects = edits().length;
  assert.equal(
    root.querySelector(".property-object-state, .property-value-menu"),
    null,
    "Designer exposes no type/null selectors",
  );
  await open("font");
  assert.equal(field("Name").value, "");
  await click(button("OK"));
  assert.equal(
    edits().length,
    beforeNullObjects,
    "confirming an untouched null object preserves source",
  );
  await open("backColor");
  await click(button("Cancel"));
  assert.equal(edits().length, beforeNullObjects, "opening/cancelling null Color leaves it intact");
  await open("font");
  await input(field("Size"), "10");
  await click(button("OK"));
  assert.deepEqual(
    edits().at(-1)!.operations,
    [{ type: "setControlProperty", controlId: "label", field: "font", value: { size: 10 } }],
    "confirming a concrete Font replaces null without filling defaults",
  );
  await host({ type: "editResult", requestId: edits().at(-1)!.requestId });
  assert.equal(JSON.stringify(nullGrid), nullSnapshot, "source null metadata is never mutated");
  await refresh(4);
  const localized = {
    ...designerStrings,
    ...editorStrings,
    properties: "プロパティ",
    editorCancel: "キャンセル",
    editorOK: "確定",
    editorStale: "ソースが変更されました。開き直してください。",
  };
  await host({ type: "localization", strings: localized });
  await select("label");
  await open("font");
  assert.equal(
    field("Name").value,
    "My unavailable font",
    "property names remain English with Japanese chrome",
  );
  assert.ok(button("キャンセル"));
  assert.ok(button("確定"));
  await refresh(5);
  assert.ok(dialog().textContent?.includes("ソースが変更されました。開き直してください。"));
  await click(button("キャンセル"));
  await host({
    type: "document",
    version: 5,
    document: source,
    propertyEditable: true,
    propertyGrid: grid,
  });
  assert.equal(
    root.querySelector<HTMLButtonElement>('[data-property-field="font"] .property-editor-launch')!
      .disabled,
    true,
    "old server cannot enable typed editors through scalar capability",
  );
  await host({
    type: "document",
    version: 6,
    document: source,
    editable: true,
    propertyEditable: true,
    productivityEditable: true,
    propertyEditorVersion: 1,
    propertyGrid: grid,
  });
  const pointer = async (
    selector: string,
    type: string,
    left: number,
    top: number,
    ctrlKey = false,
  ) => {
    await act(() => {
      const event = new dom.window.MouseEvent(type, {
        bubbles: true,
        cancelable: true,
        button: 0,
        clientX: (left * 96) / 72,
        clientY: (top * 96) / 72,
        ctrlKey,
      });
      Object.defineProperty(event, "pointerId", { value: 1 });
      root.querySelector(selector)!.dispatchEvent(event);
    });
  };
  await pointer('.control[data-control-id="label"]', "pointerdown", 20, 30);
  await pointer(".designer-viewport", "pointerup", 20, 30);
  await pointer('.control[data-control-id="image"]', "pointerdown", 125, 30, true);
  await pointer(".designer-viewport", "pointerup", 125, 30);
  assert.deepEqual(
    Array.from(root.querySelectorAll<HTMLElement>(".selection-outline")).map(
      (element) => element.dataset.controlId,
    ),
    ["label", "image"],
    "prepare sibling group",
  );
  const beforeGroup = edits().length;
  // Pressing the non-primary selected control promotes it and re-renders App.
  // That selection-only render must not look like a new Canvas document and
  // cancel the gesture before pointermove/pointerup arrive.
  await pointer('.control[data-control-id="label"]', "pointerdown", 20, 30);
  await pointer(".designer-viewport", "pointermove", 32, 36);
  await pointer(".designer-viewport", "pointerup", 32, 36);
  assert.equal(
    edits().length,
    beforeGroup + 1,
    "selection promotion retains the active group drag",
  );
  assert.deepEqual(edits().at(-1)!.operations, [
    {
      type: "moveControls",
      moves: [
        { controlId: "image", left: 132, top: 30 },
        { controlId: "label", left: 24, top: 30 },
      ],
    },
  ]);
  await act(() => render(null, root));
  dom.window.close();
  console.log(
    "UserForm typed property editors, modal, stale preview, and image bridge assertions passed",
  );
}
void run().catch((error) => {
  console.error(error);
  dom.window.close();
  process.exitCode = 1;
});
