import * as assert from "assert";
import { render } from "preact";
import { act } from "preact/test-utils";
import { renderToString } from "preact-render-to-string";
import { DesignerCanvas, ControlGlyph } from "../webview/userFormDesigner/DesignerCanvas";
import { designerDocument, pointsToPixels } from "../src/userFormEditor/model";
import type { PropertyOperation, SpecControl } from "../src/userFormEditor/protocol";
import { runUserFormEditorAssertions } from "./suite/userFormEditor.test";
import { runUserFormGeometryAssertions } from "./suite/userFormGeometry.test";

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
  await act(() => {
    element.dispatchEvent(event);
  });
}

async function rootResizeAssertions() {
  const model = designerDocument({
    form: { name: "Resize", width: 600, height: 400 },
    controls: [
      { id: "child", name: "Child", type: "Label", left: 50, top: 40, width: 30, height: 20 },
    ],
  });
  const operations: PropertyOperation[][] = [];
  const mount = async (
    options: {
      zoom?: number;
      snapping?: boolean;
      busy?: boolean;
      editable?: boolean;
      formResizable?: boolean;
      placing?: boolean;
      selectedId?: string;
      document?: typeof model;
    } = {},
  ) => {
    await act(() => {
      render(null, root);
      render(
        <DesignerCanvas
          document={options.document ?? model}
          editable={options.editable ?? true}
          formResizable={options.formResizable ?? true}
          busy={options.busy}
          placing={options.placing}
          selectedId={options.selectedId}
          onSelectionChange={() => {}}
          zoom={options.zoom}
          snapping={options.snapping}
          onFormResize={(ops) => operations.push(ops)}
        />,
        root,
      );
    });
  };
  const width = () => root.querySelector<HTMLElement>(".form")!.style.width;
  const height = () => root.querySelector<HTMLElement>(".form")!.style.height;

  await mount({ formResizable: false });
  assert.equal(
    root.querySelectorAll("[data-form-handle]").length,
    0,
    "read-only metadata hides root handles",
  );
  await mount({ editable: false });
  assert.equal(
    root.querySelectorAll("[data-form-handle]").length,
    3,
    "form property capability is independent of control geometry",
  );
  await pointer('[data-form-handle="e"]', "pointerdown", 600, 200);
  await pointer(".designer-viewport", "pointerup", 608, 200);
  assert.deepEqual(
    operations.at(-1),
    [{ type: "setFormProperty", field: "width", value: 608 }],
    "property-only editors can resize the root",
  );
  await mount({ placing: true });
  assert.equal(
    root.querySelectorAll("[data-form-handle]").length,
    0,
    "placement hides root handles",
  );
  await mount({ busy: true });
  assert.equal(
    root.querySelectorAll("[data-form-handle]").length,
    0,
    "busy mode hides root handles",
  );
  await mount({ selectedId: "child" });
  assert.equal(
    root.querySelectorAll("[data-form-handle]").length,
    0,
    "root handles require root selection",
  );

  for (const zoom of [0.5, 1, 1.5, 2]) {
    await mount({ zoom });
    const handles = Array.from(root.querySelectorAll<HTMLElement>(".form-resize-handle"));
    assert.deepEqual(
      handles.map((handle) => handle.dataset.formHandle).sort(),
      ["e", "s", "se"],
      "root uses distinct east/south/corner handles",
    );
    assert.equal(
      root.querySelectorAll(".resize-handle").length,
      0,
      "root handles do not change control handle selectors",
    );
    assert.ok(handles.every((handle) => handle.closest(".form-client-area")));
    assert.equal(
      root.querySelector(".form-title [data-form-handle]"),
      null,
      "title bar has no resize hit area",
    );

    const previous = operations.length;
    await pointer('[data-form-handle="e"]', "pointerdown", 600, 200, zoom);
    await pointer(".designer-viewport", "pointerup", 632, 200, zoom);
    assert.deepEqual(
      operations.at(-1),
      [{ type: "setFormProperty", field: "width", value: 632 }],
      `east resize uses client points at ${zoom}x zoom`,
    );
    assert.equal(operations.length, previous + 1, "pointerup commits exactly once");
    assert.equal(
      width(),
      `${pointsToPixels(632)}px`,
      "preview uses client width without shell borders",
    );
    assert.equal(height(), `${pointsToPixels(400)}px`, "east handle preserves client height");
    await pointer(".designer-viewport", "pointerup", 632, 200, zoom);
    assert.equal(operations.length, previous + 1, "duplicate pointerup cannot repeat the edit");
  }

  await mount({ snapping: true });
  await pointer('[data-form-handle="e"]', "pointerdown", 600, 200);
  await pointer(".designer-viewport", "pointerup", 629, 200);
  assert.deepEqual(
    operations.at(-1),
    [{ type: "setFormProperty", field: "width", value: 632 }],
    "root size snaps to the 8pt grid",
  );

  let before = operations.length;
  await mount();
  await pointer('[data-form-handle="e"]', "pointerdown", 600, 200);
  await pointer(".designer-viewport", "pointerup", 600, 200);
  assert.equal(operations.length, before, "no-op resize creates no operation");

  await mount();
  await pointer('[data-form-handle="se"]', "pointerdown", 600, 400);
  await pointer(".designer-viewport", "pointerup", 0, 0);
  assert.deepEqual(
    operations.at(-1),
    [
      { type: "setFormProperty", field: "width", value: 80 },
      { type: "setFormProperty", field: "height", value: 60 },
    ],
    "root cannot shrink below immediate child extents",
  );
  assert.deepEqual(
    [
      model.controls[0].left,
      model.controls[0].top,
      model.controls[0].width,
      model.controls[0].height,
    ],
    [50, 40, 30, 20],
    "root resize preserves child geometry",
  );

  const cancelGesture = async (cancel: "escape" | "pointercancel" | "lostcapture" | "blur") => {
    before = operations.length;
    await mount();
    await pointer('[data-form-handle="se"]', "pointerdown", 600, 400);
    await pointer(".designer-viewport", "pointermove", 632, 424);
    if (cancel === "escape") {
      await act(() => {
        root
          .querySelector(".designer-viewport")!
          .dispatchEvent(new dom.window.KeyboardEvent("keydown", { bubbles: true, key: "Escape" }));
      });
    } else if (cancel === "blur") {
      await act(() => {
        dom.window.dispatchEvent(new dom.window.Event("blur"));
      });
    } else {
      await pointer(
        ".designer-viewport",
        cancel === "pointercancel" ? "pointercancel" : "lostpointercapture",
        632,
        424,
      );
    }
    await pointer(".designer-viewport", "pointerup", 632, 424);
    assert.equal(operations.length, before, `${cancel} cancels without committing`);
    assert.equal(width(), `${pointsToPixels(600)}px`, `${cancel} restores the client width`);
    assert.equal(height(), `${pointsToPixels(400)}px`, `${cancel} restores the client height`);
  };
  await cancelGesture("escape");
  await cancelGesture("pointercancel");
  await cancelGesture("lostcapture");
  await cancelGesture("blur");

  for (const gate of ["busy", "formResizable"] as const) {
    before = operations.length;
    await mount();
    await pointer('[data-form-handle="e"]', "pointerdown", 600, 200);
    await pointer(".designer-viewport", "pointermove", 632, 200);
    await act(() =>
      render(
        <DesignerCanvas
          document={model}
          editable
          formResizable={gate === "busy"}
          busy={gate === "busy"}
          onFormResize={(ops) => operations.push(ops)}
        />,
        root,
      ),
    );
    await pointer(".designer-viewport", "pointerup", 632, 200);
    assert.equal(operations.length, before, `${gate} becoming unavailable cancels active resize`);
    assert.equal(width(), `${pointsToPixels(600)}px`, `${gate} cancellation clears preview`);
  }

  before = operations.length;
  await mount();
  await pointer('[data-form-handle="e"]', "pointerdown", 600, 200);
  await pointer(".designer-viewport", "pointerup", 632, 200);
  assert.equal(operations.length, before + 1, "pointerup submits a changed client size");
  assert.equal(width(), `${pointsToPixels(632)}px`, "preview remains until the edit reply");
  await act(() => {
    render(
      <DesignerCanvas document={model} editable formResizable busy onFormResize={() => {}} />,
      root,
    );
  });
  assert.equal(width(), `${pointsToPixels(632)}px`, "busy reply keeps the pending preview");
  await act(() => {
    render(
      <DesignerCanvas document={model} editable formResizable onFormResize={() => {}} />,
      root,
    );
  });
  assert.equal(width(), `${pointsToPixels(600)}px`, "failed reply clears preview when busy ends");

  await mount();
  await pointer('[data-form-handle="e"]', "pointerdown", 600, 200);
  await pointer(".designer-viewport", "pointermove", 632, 200);
  const updated = designerDocument({
    form: { name: "Resize", width: 620, height: 410 },
    controls: model.controls,
  });
  await act(() => {
    render(
      <DesignerCanvas document={updated} editable formResizable onFormResize={() => {}} />,
      root,
    );
  });
  assert.equal(width(), `${pointsToPixels(620)}px`, "document updates replace pending preview");
  await act(() => render(null, root));
}
async function run() {
  await runUserFormEditorAssertions();
  runUserFormGeometryAssertions();
  const mixed = {
    name: "Mixed",
    observed: { insideWidth: 280, insideHeight: 210 },
  };
  const measured = designerDocument({ form: { ...mixed, width: 300, height: 240 }, controls: [] });
  assert.deepStrictEqual(
    [measured.width, measured.height, measured.approximate],
    [300, 240, false],
  );
  const explicit = designerDocument({
    form: { ...mixed, width: 260, height: 190 },
    controls: [],
  });
  assert.deepStrictEqual(
    [explicit.width, explicit.height, explicit.approximate],
    [260, 190, false],
  );
  assert.strictEqual(
    designerDocument({
      form: { name: "ObservedOnly", observed: { clientWidth: 260, height: 200 } },
      controls: [],
    }).approximate,
    true,
  );
  const zeroSize = designerDocument({
    form: { name: "ZeroSize", width: 0, height: 0, observed: { width: 240, height: 180 } },
    controls: [],
  });
  assert.deepStrictEqual(
    [zeroSize.width, zeroSize.height, zeroSize.approximate],
    [0, 0, false],
    "canonical zero dimensions remain authored and observed dimensions are ignored",
  );
  const observedControl = designerDocument({
    form: { name: "ObservedControl" },
    controls: [
      {
        id: "observed-label",
        type: "Label",
        name: "ObservedLabel",
        observed: {
          left: 90,
          top: 70,
          width: 190,
          height: 80,
          caption: "Snapshot caption",
          enabled: false,
        },
      },
    ],
  }).controls[0];
  assert.deepStrictEqual(
    [
      observedControl.left,
      observedControl.top,
      observedControl.width,
      observedControl.height,
      observedControl.approximate,
      observedControl.caption,
      observedControl.enabled,
    ],
    [0, 0, 72, 18, true, undefined, undefined],
    "snapshot-only control values do not become preview authoring intent",
  );
  for (const [fields, expected] of [
    [{ value: "West", list: ["East", "West"] }, "West"],
    [{ value: "West" }, "West"],
    [{ value: "", list: ["East"] }, ""],
    [{ text: "", value: "West" }, ""],
    [{ value: 0, list: ["East"] }, "0"],
  ] as [Partial<SpecControl>, string][]) {
    const control = {
      id: "combo",
      type: "ComboBox",
      name: "Combo",
      left: 0,
      top: 0,
      width: 120,
      height: 18,
      approximate: false,
      ...fields,
    };
    const rendered = renderToString(<ControlGlyph control={control} />);
    assert.ok(rendered.includes(`<span class="combo-text">${expected}</span>`));
  }
  const types = [
    "Label",
    "TextBox",
    "CommandButton",
    "CheckBox",
    "OptionButton",
    "ComboBox",
    "ListBox",
    "ToggleButton",
    "SpinButton",
    "ScrollBar",
    "Image",
    "Frame",
    "MultiPage",
    "TabStrip",
  ];
  const model = designerDocument({
    form: { name: "Main", caption: "<script>bad()</script>", width: 240, height: 180 },
    controls: [
      ...types.map((type, index) => ({
        id: type,
        type,
        name: type,
        left: 0.75,
        top: index * 10,
        caption: "<img src=x onerror=bad()>",
        tabs: [{ name: "Tab", caption: "Details" }],
      })),
      { id: "page", type: "Page", name: "Page1", parentId: "MultiPage" },
      { id: "child", type: "Label", name: "PageChild", parentId: "page", caption: "Page child" },
      { id: "nested", type: "Frame", name: "Nested", parentId: "Frame" },
      {
        id: "nestedLabel",
        type: "Label",
        name: "NestedLabel",
        parentId: "nested",
        caption: "Nested child",
      },
    ],
  });
  const html = renderToString(<DesignerCanvas document={model} />);
  for (const type of [...types, "Page"])
    assert.ok(html.includes(`control ${type.toLowerCase()}`), type);
  assert.ok(html.includes("left:1px"), "fractional point geometry");
  assert.ok(html.includes("Page child"));
  assert.ok(html.includes("Nested child"));
  assert.ok(
    html.includes("bad()") && /&(?:lt|#60|#x3c);script/i.test(html),
    "caption must be escaped",
  );
  assert.ok(!html.includes("<script>"));
  assert.ok(!html.includes("<img"));
  const empty = renderToString(
    <DesignerCanvas
      document={designerDocument({
        form: { name: "Empty" },
        controls: [{ id: "empty", type: "MultiPage", name: "EmptyPage", selectedIndex: -1 }],
      })}
    />,
  );
  assert.ok(!empty.includes("page-content"), "empty MultiPage has no selected Page");
  const hidden = designerDocument({
    form: { name: "Visibility" },
    controls: [
      { id: "multi", type: "MultiPage", name: "Multi", selectedIndex: 1 },
      {
        id: "hiddenPage",
        type: "Page",
        name: "HiddenPage",
        parentId: "multi",
        caption: "Hidden tab",
        visible: false,
      },
      {
        id: "visiblePage",
        type: "Page",
        name: "VisiblePage",
        parentId: "multi",
        caption: "Visible tab",
      },
      {
        id: "hiddenPageLabel",
        type: "Label",
        name: "HiddenPageLabel",
        parentId: "hiddenPage",
        caption: "Hidden page child",
      },
      {
        id: "visiblePageLabel",
        type: "Label",
        name: "VisiblePageLabel",
        parentId: "visiblePage",
        caption: "Visible page child",
      },
      {
        id: "hiddenLabel",
        type: "Label",
        name: "HiddenLabel",
        caption: "Hidden control",
        visible: false,
      },
      { id: "hiddenFrame", type: "Frame", name: "HiddenFrame", visible: false },
      {
        id: "frameChild",
        type: "Label",
        name: "FrameChild",
        parentId: "hiddenFrame",
        caption: "Hidden frame child",
      },
    ],
  });
  const hiddenHTML = renderToString(<DesignerCanvas document={hidden} />);
  assert.ok(
    hiddenHTML.includes("Visible page child"),
    "selectedIndex uses original Page collection",
  );
  for (const text of ["Hidden tab", "Hidden page child", "Hidden control", "Hidden frame child"]) {
    assert.ok(!hiddenHTML.includes(text), text);
  }
  hidden.controls[0].selectedIndex = 0;
  const hiddenSelected = renderToString(<DesignerCanvas document={hidden} />);
  assert.ok(
    !hiddenSelected.includes("page-content"),
    "hidden selected Page must not select a different Page",
  );
  hidden.controls[1].visible = true;
  hidden.controls[1].zIndex = 1;
  hidden.controls[2].zIndex = 0;
  const reordered = renderToString(<DesignerCanvas document={hidden} />);
  assert.ok(
    reordered.includes("Visible page child"),
    "selection matches compiler zIndex-ordered Page collection",
  );
  assert.ok(!reordered.includes("Hidden page child"));
  await rootResizeAssertions();
  console.log("UserForm synchronization and rendering assertions passed.");
}
void run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
