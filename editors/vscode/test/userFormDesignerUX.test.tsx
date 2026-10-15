import * as assert from "node:assert/strict";
import { render } from "preact";
import { act } from "preact/test-utils";
import { designerDocument } from "../src/userFormEditor/model";
import { designerStrings } from "../src/userFormEditor/protocol";
import type { DesignerDocument, DesignerStrings } from "../src/userFormEditor/protocol";
import { StructureTree } from "../webview/userFormDesigner/StructureTree";

const { JSDOM } = require("jsdom");
const dom = new JSDOM('<!doctype html><html><body><div id="app"></div></body></html>', {
  url: "https://designer.test/",
});
Object.assign(globalThis, { window: dom.window, document: dom.window.document });

const root = document.getElementById("app")!;
const strings: DesignerStrings = { ...designerStrings, structure: "Structure" };
const model = designerDocument({
  form: { name: "MainForm", caption: "Main form", width: 600, height: 400 },
  controls: [
    { id: "frame", name: "Frame1", type: "Frame", zIndex: 1 },
    { id: "nested", name: "CustomWidget", type: "Vendor.Custom", parentId: "frame", zIndex: 1 },
    {
      id: "hidden",
      name: "HiddenLabel",
      type: "Label",
      parentId: "frame",
      visible: false,
      zIndex: 0,
    },
    { id: "multi", name: "MultiPage1", type: "MultiPage", zIndex: 2 },
    { id: "page", name: "Page1", type: "Page", parentId: "multi" },
    { id: "page-child", name: "PageChild", type: "Label", parentId: "page" },
    {
      id: "tabs",
      name: "TabStrip1",
      type: "TabStrip",
      zIndex: 3,
      tabs: [
        { name: "FirstTab", caption: "First tab" },
        { name: "SecondTab", caption: "Second tab" },
      ],
    },
    { id: "alpha", name: "Alpha", type: "Label", zIndex: 0 },
  ],
});

function item(id: string): HTMLElement {
  const result = root.querySelector<HTMLElement>(`[data-structure-id="${id}"]`);
  assert.ok(result, `missing structure item ${id}`);
  return result;
}

async function renderTree(
  document: DesignerDocument,
  selectionIds: string[],
  onSelect: (id?: string) => void = () => {},
  busy = false,
) {
  await act(() =>
    render(
      <StructureTree
        document={document}
        selectionIds={selectionIds}
        strings={strings}
        busy={busy}
        onSelect={onSelect}
      />,
      root,
    ),
  );
}

async function press(id: string, key: string) {
  await act(() => {
    const target = item(id);
    target.focus();
    target.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }),
    );
  });
}

async function structureIncludesAllControlsAndOrderedHierarchy() {
  await renderTree(model, ["nested"], () => {}, true);
  const tree = root.querySelector<HTMLElement>("[role='tree']")!;
  assert.equal(tree.getAttribute("aria-label"), "Structure");
  assert.equal(tree.getAttribute("aria-busy"), "true");
  assert.equal(item("").getAttribute("data-structure-id"), "");
  assert.equal(item("").getAttribute("aria-label"), "MainForm");
  assert.equal(item("nested").getAttribute("aria-label"), "CustomWidget");
  assert.equal(item("").getAttribute("aria-expanded"), "true");
  assert.equal(item("frame").getAttribute("aria-expanded"), "true");
  assert.equal(item("nested").getAttribute("aria-selected"), "true");

  const directChildren = [
    ...item("").querySelectorAll(":scope > [role='group'] > [role='treeitem']"),
  ].map((entry) => (entry as HTMLElement).dataset.structureId);
  assert.deepEqual(directChildren, ["alpha", "frame", "multi", "tabs"]);
  assert.ok(item("hidden"), "invisible controls remain listed");
  assert.ok(item("nested"), "custom controls remain listed");

  await act(() => item("multi").querySelector<HTMLElement>(".structure-tree-disclosure")!.click());
  assert.ok(item("page"), "Page controls remain listed");

  await act(() => item("page").querySelector<HTMLElement>(".structure-tree-disclosure")!.click());
  assert.ok(item("page-child"), "controls under a Page remain listed");
  assert.equal(
    root.textContent?.includes("First tab"),
    false,
    "TabStrip tabs are not tree entries",
  );
  assert.equal(
    root.textContent?.includes("Second tab"),
    false,
    "TabStrip tabs are not tree entries",
  );
}

async function clickSelectionAndExpansionStaySeparate() {
  const selected: (string | undefined)[] = [];
  await renderTree(model, [], (id) => selected.push(id));

  await act(() => {
    item("").querySelector<HTMLElement>(".structure-tree-disclosure")!.click();
  });
  assert.equal(selected.length, 0, "expanding the form root does not select it");
  assert.equal(
    root.querySelector("[data-structure-id='alpha']"),
    null,
    "root disclosure collapses children",
  );

  await act(() => {
    item("").querySelector<HTMLElement>(".structure-tree-row")!.click();
  });
  assert.deepEqual(selected, [undefined], "the form root can be selected without a control id");

  await act(() => item("").querySelector<HTMLElement>(".structure-tree-disclosure")!.click());
  await act(() => item("alpha").querySelector<HTMLElement>(".structure-tree-row")!.click());
  assert.deepEqual(selected, [undefined, "alpha"]);
}

async function keyboardNavigationUsesRovingFocusAndExplicitSelection() {
  const selected: (string | undefined)[] = [];
  await renderTree(model, [], (id) => selected.push(id));
  assert.equal(item("").tabIndex, 0);
  assert.equal(item("alpha").tabIndex, -1);

  await press("", "ArrowDown");
  assert.equal(item("alpha").tabIndex, 0);
  assert.equal(document.activeElement, item("alpha"));
  assert.deepEqual(selected, [], "arrow navigation moves focus without selecting");

  await press("alpha", "ArrowDown");
  assert.equal(item("frame").tabIndex, 0);
  await press("frame", "ArrowRight");
  assert.equal(item("frame").getAttribute("aria-expanded"), "true");
  assert.equal(item("frame").tabIndex, 0);
  await press("frame", "ArrowRight");
  assert.equal(item("hidden").tabIndex, 0, "children follow orderedChildren order");
  await press("hidden", "Enter");
  await press("hidden", " ");
  assert.deepEqual(selected, ["hidden", "hidden"]);

  await press("hidden", "ArrowLeft");
  assert.equal(item("frame").tabIndex, 0);
  await press("frame", "ArrowLeft");
  assert.equal(item("frame").getAttribute("aria-expanded"), "false");
  await press("frame", "Home");
  assert.equal(item("").tabIndex, 0);
  await press("", "End");
  assert.equal(item("tabs").tabIndex, 0);
  await press("tabs", "ArrowUp");
  assert.equal(item("multi").tabIndex, 0);
}

async function removedSelectionLeavesAValidRovingItem() {
  await renderTree(model, ["page-child"]);
  assert.equal(item("multi").getAttribute("aria-expanded"), "true");
  assert.equal(item("page").getAttribute("aria-expanded"), "true");
  assert.equal(item("page-child").tabIndex, 0);
  item("page-child").focus();
  assert.equal(document.activeElement, item("page-child"));

  const afterDelete: DesignerDocument = {
    ...model,
    controls: model.controls.filter((control) => control.id !== "page-child"),
  };
  await renderTree(afterDelete, []);

  assert.equal(root.querySelector("[data-structure-id='page-child']"), null);
  assert.equal(item("").tabIndex, 0);
  assert.equal(item("").getAttribute("aria-selected"), "true");
  assert.equal(
    document.activeElement,
    item(""),
    "focus moves to the surviving root after deletion",
  );
  await press("", "ArrowDown");
  assert.equal(item("alpha").tabIndex, 0);
}

async function run() {
  await structureIncludesAllControlsAndOrderedHierarchy();
  await act(() => render(null, root));
  await clickSelectionAndExpansionStaySeparate();
  await act(() => render(null, root));
  await keyboardNavigationUsesRovingFocusAndExplicitSelection();
  await act(() => render(null, root));
  await removedSelectionLeavesAValidRovingItem();
  await act(() => render(null, root));
  await structureDeletionGuards();
  await act(() => render(null, root));
}

async function structureDeletionGuards() {
  let deletions = 0;
  const show = async (busy = false) =>
    act(() =>
      render(
        <StructureTree
          document={model}
          selectionIds={["nested"]}
          strings={strings}
          busy={busy}
          onSelect={() => {}}
          onDelete={() => {
            deletions++;
            return true;
          }}
        />,
        root,
      ),
    );
  await show();
  const send = async (id: string, options: KeyboardEventInit = {}) => {
    const event = new dom.window.KeyboardEvent("keydown", {
      key: "Backspace",
      bubbles: true,
      cancelable: true,
      ...options,
    });
    await act(() => item(id).dispatchEvent(event));
    return event;
  };
  for (const guard of [
    { repeat: true },
    { isComposing: true },
    { ctrlKey: true },
    { altKey: true },
    { metaKey: true },
    { shiftKey: true },
  ])
    await send("nested", guard);
  await send("");
  await send("frame");
  assert.equal(deletions, 0, "unselected focus, root, repeats, IME and modifiers cannot delete");
  const removed = await send("nested");
  assert.equal(deletions, 1);
  assert.equal(removed.defaultPrevented, true);
  await show(true);
  await send("nested");
  assert.equal(deletions, 1, "busy structure never deletes");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
