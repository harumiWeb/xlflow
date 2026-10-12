import * as assert from "node:assert/strict";
import { render } from "preact";
import { act } from "preact/test-utils";
import { DesignerCanvas } from "../webview/userFormDesigner/DesignerCanvas";
import { designerDocument, pointsToPixels } from "../src/userFormEditor/model";
import {
  dropTarget,
  parentOrigin,
  placeGeometry,
  selectionPath,
  worldGeometry,
  zoomLevels,
} from "../src/userFormEditor/geometry";
import { createControl, subtreeIds } from "../src/userFormEditor/toolbox";
import type { GeometryOperation, ParentOperation } from "../src/userFormEditor/protocol";

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
const root = document.getElementById("app")!;
const model = designerDocument({
  form: { name: "Main", build: { clientWidth: 600, clientHeight: 400 } },
  controls: [
    { id: "label", name: "Label1", type: "Label", left: 10, top: 10, width: 20, height: 10 },
    { id: "frame", name: "Frame1", type: "Frame", left: 200, top: 100, width: 180, height: 160 },
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
    { id: "multi", name: "Multi", type: "MultiPage", left: 400, top: 100, width: 180, height: 160 },
    { id: "page", parentId: "multi", name: "Page1", type: "Page" },
    {
      id: "pagechild",
      parentId: "page",
      name: "PageChild",
      type: "Label",
      left: 10,
      top: 10,
      width: 20,
      height: 10,
    },
  ],
});

function geometryAssertions() {
  assert.deepEqual(parentOrigin(model, "nested"), { left: 221.5, top: 131.5 });
  assert.deepEqual(worldGeometry(model, model.controls[3]), {
    left: 226.5,
    top: 137.5,
    width: 20,
    height: 10,
  });
  assert.deepEqual(parentOrigin(model, "page"), { left: 401.5, top: 118 });
  assert.deepEqual(
    selectionPath(model, "child").map((c) => c.id),
    ["frame", "nested", "child"],
  );
  assert.equal(
    dropTarget(model, { left: 230, top: 140 })?.parentId,
    "nested",
    "leaf inherits the nearest Frame",
  );
  assert.equal(dropTarget(model, { left: 210, top: 110 })?.parentId, "frame");
  assert.equal(dropTarget(model, { left: 100, top: 50 })?.parentId, "");
  assert.equal(
    dropTarget(model, { left: 420, top: 140 }),
    undefined,
    "Page does not fall back to root",
  );
  assert.equal(dropTarget(model, { left: 700, top: 50 }), undefined);
  assert.equal(
    dropTarget(model, { left: 230, top: 140 }, subtreeIds(model.controls, "frame"))?.parentId,
    "",
    "moving subtree cannot target itself",
  );
  const overlap = {
    ...model,
    controls: [
      ...model.controls,
      { ...model.controls[1], id: "front", name: "Front", width: 40, height: 40, zIndex: 100 },
    ],
  };
  assert.equal(
    dropTarget(overlap, { left: 230, top: 140 })?.parentId,
    "nested",
    "the front Frame's far border is outside its bounds",
  );
  assert.equal(
    dropTarget(overlap, { left: 230, top: 135 })?.parentId,
    "front",
    "visual stacking wins over another deeper Frame",
  );
  assert.equal(
    dropTarget(
      {
        ...overlap,
        controls: overlap.controls.map((c) => (c.id === "front" ? { ...c, visible: false } : c)),
      },
      { left: 230, top: 135 },
    )?.parentId,
    "nested",
  );
  assert.deepEqual(
    placeGeometry({ left: -3, top: 99, width: 20, height: 10 }, { width: 50, height: 40 }, true),
    { left: 0, top: 30, width: 20, height: 10 },
  );
  assert.equal(
    placeGeometry({ left: 0, top: 0, width: 51, height: 10 }, { width: 50, height: 40 }),
    undefined,
  );
  const created = createControl(model, "Label", 99, 99, false, () => "fresh", "nested")!;
  assert.equal(created, undefined, "default Label does not fit the nested Frame");
  const outer = createControl(model, "Label", 999, -9, false, () => "fresh", "frame")!;
  assert.deepEqual([outer.parentId, outer.left, outer.top], ["frame", 106.5, 0]);
  assert.equal(
    createControl(model, "Label", 0, 0, false, () => "fresh", "page"),
    undefined,
  );
}

async function pointer(
  selector: string,
  type: string,
  left: number,
  top: number,
  zoom = 1,
  pointerId = 1,
) {
  const element = root.querySelector(selector)!;
  assert.ok(element, selector);
  const event = new dom.window.MouseEvent(type, {
    bubbles: true,
    button: 0,
    clientX: pointsToPixels(left, zoom),
    clientY: pointsToPixels(top, zoom),
  });
  Object.defineProperty(event, "pointerId", { value: pointerId });
  await act(() => element.dispatchEvent(event));
}

async function run() {
  geometryAssertions();
  for (const zoom of zoomLevels) {
    const operations: (GeometryOperation | ParentOperation)[][] = [];
    await act(() =>
      render(
        <DesignerCanvas
          document={model}
          editable
          parentEditable
          zoom={zoom}
          onCommit={(ops) => operations.push(ops)}
        />,
        root,
      ),
    );
    await pointer('.control[data-control-id="label"]', "pointerdown", 12, 12, zoom);
    assert.ok(
      root.querySelector('.selection-outline[data-control-id="label"]'),
      "pointerdown selects the control",
    );
    await pointer(".designer-viewport", "pointermove", 220, 120, zoom);
    assert.equal(
      root.querySelector("[data-drop-parent-id]")?.getAttribute("data-drop-parent-id"),
      "frame",
    );
    assert.ok(root.querySelector('.drag-preview [data-control-id="label"]'));
    assert.equal(operations.length, 0, "preview cannot edit source");
    await pointer(".designer-viewport", "pointerup", 220, 120, zoom);
    assert.deepEqual(
      operations,
      [[{ type: "setParent", controlId: "label", parentId: "frame", left: 17.25, top: 17.25 }]],
      `reparent at zoom ${zoom}`,
    );
    await act(() => render(null, root));
  }
  const operations: (GeometryOperation | ParentOperation)[][] = [];
  const mount = async (parentEditable = true, snapping = false) => {
    await act(() => {
      render(null, root);
      render(
        <DesignerCanvas
          document={model}
          editable
          parentEditable={parentEditable}
          snapping={snapping}
          onCommit={(ops) => operations.push(ops)}
        />,
        root,
      );
    });
  };
  await mount();
  await pointer('.control[data-control-id="child"]', "pointerdown", 228.5, 139.5);
  await pointer(".designer-viewport", "pointerup", 100, 50);
  assert.deepEqual(operations.at(-1), [
    { type: "setParent", controlId: "child", parentId: "", left: 98, top: 48 },
  ]);
  await mount();
  await pointer('.control[data-control-id="nested"]', "pointerdown", 222, 132);
  await pointer(".designer-viewport", "pointermove", 100, 50);
  assert.ok(
    root.querySelector('.drag-preview [data-control-id="child"]'),
    "Frame preview carries descendants",
  );
  const child = root.querySelector<HTMLElement>('.drag-preview .control[data-control-id="child"]')!;
  assert.equal(child.style.left, `${pointsToPixels(5)}px`);
  await pointer(".designer-viewport", "pointerup", 100, 50);
  assert.equal(operations.at(-1)?.[0].type, "setParent");
  assert.equal(operations.at(-1)?.length, 1, "children require no separate edits");
  await mount();
  const before = operations.length;
  await pointer('.control[data-control-id="label"]', "pointerdown", 12, 12);
  await pointer(".designer-viewport", "pointerup", 12, 12);
  assert.equal(operations.length, before, "a click cannot reparent");
  await pointer('.control[data-control-id="label"]', "pointerdown", 12, 12);
  await pointer(".designer-viewport", "pointermove", 420, 140);
  assert.ok(root.textContent?.includes("cannot be placed"));
  await pointer(".designer-viewport", "pointerup", 420, 140);
  assert.equal(operations.length, before, "Page drop cannot commit");
  await pointer('.control[data-control-id="label"]', "pointerdown", 12, 12);
  await pointer(".designer-viewport", "pointermove", 220, 120);
  await pointer(".designer-viewport", "pointercancel", 220, 120, 1, 2);
  assert.ok(root.querySelector(".drag-preview"), "secondary pointer cannot cancel");
  await pointer(".designer-viewport", "pointercancel", 220, 120);
  assert.equal(root.querySelector(".drag-preview"), null);
  assert.equal(operations.length, before);
  await mount(false);
  await pointer('.control[data-control-id="label"]', "pointerdown", 12, 12);
  await pointer(".designer-viewport", "pointerup", 220, 120);
  assert.equal(operations.at(-1)?.[0].type, "moveControl", "old servers keep root movement");
  await mount();
  await pointer('.control[data-control-id="child"]', "pointerdown", 228.5, 139.5);
  await pointer(".designer-viewport", "pointerup", 228.5, 139.5);
  const buttons = [...root.querySelectorAll<HTMLButtonElement>(".selection-path button")];
  assert.deepEqual(
    buttons.map((b) => b.textContent),
    ["Main", "Frame1", "Frame2", "Child"],
  );
  await act(() => buttons[1].click());
  assert.ok(root.querySelector('.selection-outline[data-control-id="frame"]'));
  assert.equal(document.activeElement, root.querySelector(".designer-viewport"));
  await act(() =>
    document.activeElement!.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", { bubbles: true, key: "ArrowRight" }),
    ),
  );
  await act(() =>
    document.activeElement!.dispatchEvent(
      new dom.window.KeyboardEvent("keyup", { bubbles: true, key: "ArrowRight" }),
    ),
  );
  assert.deepEqual(
    operations.at(-1),
    [{ type: "moveControl", controlId: "frame", left: 201, top: 100 }],
    "breadcrumb-selected ancestors retain keyboard movement",
  );
  await act(() => render(null, root));
  process.stdout.write("UserForm parenting geometry and DOM assertions passed\n");
}
run().catch((error) => {
  process.stderr.write(`${error.stack ?? error}\n`);
  process.exitCode = 1;
});
