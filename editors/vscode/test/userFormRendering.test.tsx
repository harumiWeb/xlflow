import * as assert from "assert";
import { renderToString } from "preact-render-to-string";
import { DesignerCanvas, ControlGlyph } from "../webview/userFormDesigner/DesignerCanvas";
import { designerDocument } from "../src/userFormEditor/model";
import type { SpecControl } from "../src/userFormEditor/protocol";
import { runUserFormEditorAssertions } from "./suite/userFormEditor.test";

async function run() {
  await runUserFormEditorAssertions();
  const mixed = {
    name: "Mixed",
    build: { width: 300, height: 240 },
    observed: { insideWidth: 280, insideHeight: 210 },
  };
  const measured = designerDocument({ form: mixed, controls: [] });
  assert.deepStrictEqual(
    [measured.width, measured.height, measured.approximate],
    [280, 210, false],
  );
  const explicit = designerDocument({
    form: { ...mixed, build: { ...mixed.build, clientWidth: 260, clientHeight: 190 } },
    controls: [],
  });
  assert.deepStrictEqual(
    [explicit.width, explicit.height, explicit.approximate],
    [260, 190, false],
  );
  assert.strictEqual(
    designerDocument({
      form: { name: "OuterHeight", build: { clientWidth: 260, height: 200 } },
      controls: [],
    }).approximate,
    true,
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
  console.log("UserForm synchronization and rendering assertions passed.");
}
void run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
