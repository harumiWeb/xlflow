import * as assert from "node:assert/strict";
import { render } from "preact";
import type { ComponentProps } from "preact";
import { TabOrderEditor } from "../webview/userFormDesigner/TabOrderEditor";
import { act } from "preact/test-utils";
import { designerDocument, pointsToPixels } from "../src/userFormEditor/model";
import {
  defaultDuplicateParent,
  defaultPasteParent,
  eligibleTabOrder,
  fullyEnclosedChildren,
  isProductivityControl,
  moveSiblingGroup,
  promoteSelection,
  siblingSelection,
  updateSelection,
} from "../src/userFormEditor/productivity";
import { isWebviewMessage, designerStrings } from "../src/userFormEditor/protocol";
import type {
  PropertyGridData,
  ProductivityOperation,
  SemanticOperation,
} from "../src/userFormEditor/protocol";
import { DesignerCanvas } from "../webview/userFormDesigner/DesignerCanvas";
import { ProductivityToolbar as CommandToolbar } from "../webview/userFormDesigner/ProductivityToolbar";

// Keep the established operation regressions exercising both extracted UI
// components, without coupling tab-order authoring to toolbar placement.
function ProductivityToolbar({
  propertyGrid,
  tabEnabled,
  ...props
}: ComponentProps<typeof CommandToolbar> & {
  propertyGrid?: PropertyGridData;
  tabEnabled: boolean;
}) {
  return (
    <>
      <CommandToolbar {...props} />
      <TabOrderEditor
        document={props.document}
        selectionIds={props.selectionIds}
        propertyGrid={propertyGrid}
        strings={props.strings}
        tabEnabled={tabEnabled}
        busy={props.busy}
        onOperation={props.onOperation}
      />
    </>
  );
}
async function openCommands(label: string) {
  const matches = [...root.querySelectorAll<HTMLButtonElement>('button[aria-haspopup="menu"]')];
  const button = matches.find((button) => button.textContent === label);
  assert.ok(button, `missing menu ${label}`);
  await act(() => button.click());
}
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
  form: { name: "Main", width: 300, height: 180 },
  controls: [
    { id: "a", name: "LabelA", type: "Label", left: 10, top: 10, width: 20, height: 10 },
    { id: "b", name: "LabelB", type: "Label", left: 40, top: 10, width: 20, height: 10 },
    {
      id: "c",
      name: "LabelC",
      type: "Label",
      left: 80,
      top: 10,
      width: 20,
      height: 10,
      visible: false,
      enabled: false,
    },
    { id: "frame", name: "Frame1", type: "Frame", left: 120, top: 20, width: 100, height: 90 },
    {
      id: "nested",
      name: "Nested",
      type: "Label",
      parentId: "frame",
      left: 5,
      top: 5,
      width: 20,
      height: 10,
    },
    { id: "multi", name: "Pages", type: "MultiPage", left: 10, top: 50, width: 100, height: 80 },
    { id: "page", name: "Page1", type: "Page", parentId: "multi" },
  ],
});

const propertyGrid: PropertyGridData = {
  form: { descriptors: [], values: {} },
  controls: {
    a: {
      descriptors: [{ field: "tabIndex", valueType: "integer", required: false, nullable: false }],
      values: { tabIndex: { present: true, value: 0 } },
    },
    b: {
      descriptors: [{ field: "tabIndex", valueType: "integer", required: false, nullable: false }],
      values: { tabIndex: { present: true, value: 1 } },
    },
    c: {
      descriptors: [{ field: "tabIndex", valueType: "integer", required: false, nullable: false }],
      values: { tabIndex: { present: true, value: 2 } },
    },
    page: {
      descriptors: [{ field: "tabIndex", valueType: "integer", required: false, nullable: false }],
      values: { tabIndex: { present: true, value: 3 } },
    },
  },
};

async function pointer(
  selector: string,
  type: string,
  x: number,
  y: number,
  options: { pointerId?: number; ctrlKey?: boolean; metaKey?: boolean; shiftKey?: boolean } = {},
) {
  const element = root.querySelector(selector);
  assert.ok(element, selector);
  const event = new dom.window.MouseEvent(type, {
    bubbles: true,
    button: 0,
    clientX: x,
    clientY: y,
    ctrlKey: options.ctrlKey,
    metaKey: options.metaKey,
    shiftKey: options.shiftKey,
  });
  Object.defineProperty(event, "pointerId", { value: options.pointerId ?? 1 });
  await act(() => element.dispatchEvent(event));
}

async function click(selector: string) {
  const element = root.querySelector(selector);
  assert.ok(element, selector);
  await act(() => element.dispatchEvent(new dom.window.MouseEvent("click", { bubbles: true })));
}

async function run() {
  assert.deepEqual(updateSelection(model, ["a"], "b", true), ["a", "b"]);
  assert.deepEqual(updateSelection(model, ["a", "b"], "a", true), ["b"]);
  assert.deepEqual(updateSelection(model, ["a", "b"], "nested", true), ["nested"]);
  assert.deepEqual(updateSelection(model, ["a", "b"], "page", true), ["page"]);
  assert.deepEqual(siblingSelection(model, ["a", "nested"]), []);
  assert.deepEqual(promoteSelection(model, ["a", "b"], "a"), ["b", "a"]);

  const customProgIdDocument = {
    ...model,
    controls: model.controls.map((control) =>
      control.id === "b" ? { ...control, progId: "Vendor.Label" } : control,
    ),
  };
  assert.equal(isProductivityControl(customProgIdDocument.controls[1]), false);
  assert.deepEqual(updateSelection(customProgIdDocument, ["a"], "b", true), ["b"]);
  assert.deepEqual(siblingSelection(customProgIdDocument, ["b"]), []);
  assert.equal(moveSiblingGroup(customProgIdDocument, ["a", "b"], 3, 4), undefined);
  const nestedCustomDocument = {
    ...model,
    controls: model.controls.map((control) =>
      control.id === "nested" ? { ...control, progId: "Vendor.Label" } : control,
    ),
  };
  assert.equal(defaultPasteParent(nestedCustomDocument, ["nested"]), "frame");
  const customFrameDocument = {
    ...model,
    controls: model.controls.map((control) =>
      control.id === "frame" ? { ...control, progId: "Vendor.Frame.1" } : control,
    ),
  };
  assert.equal(defaultPasteParent(customFrameDocument, ["frame"]), "");
  const customOwnerDocument = {
    ...model,
    controls: model.controls.map((control) =>
      control.id === "multi" ? { ...control, progId: "Vendor.MultiPage.1" } : control,
    ),
  };
  assert.equal(defaultPasteParent(customOwnerDocument, ["page"]), "multi");

  const fractionalDocument = {
    ...model,
    controls: model.controls.map((control) =>
      control.id === "a"
        ? { ...control, left: 10.123456789, top: 10.456789123 }
        : control.id === "b"
          ? { ...control, left: 40.123456789, top: 10.456789123 }
          : control,
    ),
  };
  const verticalOnlyMove = moveSiblingGroup(fractionalDocument, ["a", "b"], 0, 1, true);
  assert.equal(verticalOnlyMove?.type, "moveControls");
  if (verticalOnlyMove?.type === "moveControls")
    assert.deepEqual(
      verticalOnlyMove.moves.map((move) => move.left),
      [10.123456789, 40.123456789],
      "zero displacement keeps original double coordinates without rounding or snapping",
    );
  assert.equal(moveSiblingGroup(fractionalDocument, ["a", "b"], 0, 0, true), undefined);
  const translated = moveSiblingGroup(fractionalDocument, ["a", "b"], 1, 2);
  assert.equal(translated?.type, "moveControls");
  if (translated?.type === "moveControls") {
    assert.deepEqual(
      translated.moves.map(({ left, top }) => ({ left, top })),
      [
        { left: 11.123456789, top: 12.456789123 },
        { left: 41.123456789, top: 12.456789123 },
      ],
      "a shared translation preserves fractional coordinates on changed axes",
    );
  }
  assert.equal(
    moveSiblingGroup({ ...model, width: 55 }, ["a", "b"], 0, 5),
    undefined,
    "a group that collectively exceeds its parent is rejected",
  );
  assert.equal(
    moveSiblingGroup(
      {
        ...model,
        controls: model.controls.map((control) =>
          control.id === "a" ? { ...control, top: -0.25 } : control,
        ),
      },
      ["a", "b"],
      5,
      0,
    ),
    undefined,
    "movement on one axis cannot silently correct an invalid untouched axis",
  );

  assert.deepEqual(
    moveSiblingGroup(model, ["a", "b"], 15, 8),
    {
      type: "moveControls",
      moves: [
        { controlId: "a", left: 25, top: 18 },
        { controlId: "b", left: 55, top: 18 },
      ],
    },
    "group movement preserves relative positions and emits one operation",
  );
  const snappedGroup = moveSiblingGroup(model, ["a", "b"], 7, 7, true);
  assert.equal(snappedGroup?.type, "moveControls");
  if (snappedGroup?.type === "moveControls")
    assert.deepEqual(
      [
        snappedGroup.moves[1].left - snappedGroup.moves[0].left,
        snappedGroup.moves[1].top - snappedGroup.moves[0].top,
      ],
      [30, 0],
      "snapping uses a shared delta for every sibling",
    );
  assert.equal(defaultPasteParent(model, ["frame"]), "frame");
  assert.equal(defaultPasteParent(model, ["page"]), "page");
  assert.equal(defaultPasteParent(model, ["a"]), "");
  assert.equal(defaultDuplicateParent(model, ["frame"]), "");
  assert.equal(defaultDuplicateParent(model, ["nested"]), "frame");
  assert.equal(
    defaultPasteParent(model, ["a", "frame"]),
    "",
    "a Frame in a multi-selection is not treated as the paste container",
  );
  assert.deepEqual(eligibleTabOrder(model, propertyGrid, ""), ["a", "b", "c"]);

  const enclosed = fullyEnclosedChildren(
    model,
    "",
    { left: 0, top: 0, right: 250, bottom: 140 },
    (id) =>
      ({
        a: { left: 10, top: 10, right: 30, bottom: 20 },
        b: { left: 40, top: 10, right: 60, bottom: 20 },
        c: { left: 80, top: 10, right: 100, bottom: 20 },
        frame: { left: 120, top: 20, right: 220, bottom: 110 },
      })[id],
  );
  assert.deepEqual(enclosed, ["a", "b", "frame"], "marquee selects visible direct children only");
  assert.deepEqual(
    fullyEnclosedChildren(
      customProgIdDocument,
      "",
      { left: 0, top: 0, right: 100, bottom: 30 },
      (id) => (["a", "b"].includes(id) ? { left: 10, top: 10, right: 60, bottom: 20 } : undefined),
    ),
    ["a"],
    "marquee excludes custom ProgIDs as multi-selection targets",
  );

  const clipboard = {
    type: "clipboard",
    requestId: 1,
    version: 2,
    action: "paste",
    controlIds: [],
    parentId: "frame",
  } as const;
  assert.equal(isWebviewMessage(clipboard), true);
  assert.equal(
    isWebviewMessage({
      type: "edit",
      requestId: 2,
      version: 2,
      operations: [
        {
          type: "pasteControls",
          snapshot: { format: "yaml", text: "controls: []", controlIds: ["a"] },
          allocations: [{ sourceId: "a", id: "new-a", name: "LabelA2" }],
          parentId: "",
          offset: { left: 8, top: 8 },
        },
      ],
    }),
    false,
    "Webview must use the host clipboard action instead of sending raw paste snapshots",
  );

  const emitted: ProductivityOperation[][] = [];
  const shortcutActions: string[] = [];
  await act(() =>
    render(
      <DesignerCanvas
        document={model}
        editable
        parentEditable
        productivityEditable
        onCommit={(operations) => emitted.push(operations as ProductivityOperation[])}
        onShortcut={(action) => {
          shortcutActions.push(action);
          return true;
        }}
      />,
      root,
    ),
  );
  await pointer(
    '.control[data-control-id="a"]',
    "pointerdown",
    pointsToPixels(12),
    pointsToPixels(12),
  );
  await pointer(
    '.control[data-control-id="a"]',
    "pointerup",
    pointsToPixels(12),
    pointsToPixels(12),
  );
  await pointer(
    '.control[data-control-id="b"]',
    "pointerdown",
    pointsToPixels(42),
    pointsToPixels(12),
    {
      shiftKey: true,
    },
  );
  assert.ok(root.querySelector('.selection-outline[data-control-id="a"]'));
  assert.ok(root.querySelector('.selection-outline[data-control-id="b"]'));
  assert.ok(root.querySelector(".group-selection-outline"));
  await pointer(
    '.control[data-control-id="b"]',
    "pointerdown",
    pointsToPixels(42),
    pointsToPixels(12),
    {
      shiftKey: true,
    },
  );
  assert.equal(root.querySelector('.selection-outline[data-control-id="b"]'), null);
  await pointer(
    '.control[data-control-id="b"]',
    "pointerdown",
    pointsToPixels(42),
    pointsToPixels(12),
    {
      ctrlKey: true,
    },
  );
  await pointer(
    '.control[data-control-id="b"]',
    "pointerdown",
    pointsToPixels(42),
    pointsToPixels(12),
    {
      ctrlKey: true,
    },
  );
  assert.equal(root.querySelector('.selection-outline[data-control-id="b"]'), null);
  await pointer(
    '.control[data-control-id="b"]',
    "pointerdown",
    pointsToPixels(42),
    pointsToPixels(12),
    {
      metaKey: true,
    },
  );
  assert.ok(root.querySelector('.selection-outline[data-control-id="b"]'));

  await pointer(
    '.control[data-control-id="b"]',
    "pointerdown",
    pointsToPixels(42),
    pointsToPixels(12),
  );
  await pointer(".designer-viewport", "pointermove", pointsToPixels(57), pointsToPixels(20));
  const movingBoundary = root.querySelector<HTMLElement>(".group-selection-outline")!;
  assert.equal(movingBoundary.style.left, `${pointsToPixels(25)}px`);
  assert.equal(movingBoundary.style.top, `${pointsToPixels(18)}px`);
  assert.equal(movingBoundary.style.width, `${pointsToPixels(50)}px`);
  await pointer(".designer-viewport", "pointerup", pointsToPixels(57), pointsToPixels(20));
  assert.deepEqual(emitted.at(-1), [
    {
      type: "moveControls",
      moves: [
        { controlId: "a", left: 25, top: 18 },
        { controlId: "b", left: 55, top: 18 },
      ],
    },
  ]);

  const viewport = root.querySelector<HTMLElement>(".designer-viewport")!;
  await act(() =>
    viewport.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", { bubbles: true, key: "ArrowRight" }),
    ),
  );
  await act(() => dom.window.dispatchEvent(new dom.window.Event("blur")));
  assert.deepEqual(
    emitted.at(-1),
    [
      {
        type: "moveControls",
        moves: [
          { controlId: "a", left: 11, top: 10 },
          { controlId: "b", left: 41, top: 10 },
        ],
      },
    ],
    "arrow-key movement commits the complete sibling group",
  );

  await act(() =>
    viewport.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", {
        bubbles: true,
        key: "c",
        ctrlKey: true,
      }),
    ),
  );
  const input = document.createElement("input");
  viewport.append(input);
  await act(() =>
    input.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", {
        bubbles: true,
        key: "c",
        ctrlKey: true,
      }),
    ),
  );
  assert.deepEqual(shortcutActions, ["copy"], "canvas shortcuts do not capture input shortcuts");
  input.remove();

  const committedCount = emitted.length;
  await pointer('.control[data-control-id="a"]', "pointerdown", 12, 12);
  await pointer(".designer-viewport", "pointermove", 30, 30);
  await pointer(".designer-viewport", "pointercancel", 30, 30, { pointerId: 7 });
  assert.ok(
    root.querySelector(".group-selection-outline"),
    "another pointer cannot cancel the drag",
  );
  await pointer(".designer-viewport", "pointercancel", 30, 30);
  assert.equal(emitted.length, committedCount, "pointer cancellation discards the group move");
  assert.equal(
    root.querySelector<HTMLElement>(".group-selection-outline")!.style.left,
    `${pointsToPixels(10)}px`,
    "cancel restores the source selection boundary",
  );

  const beforeExternalUpdate = emitted.length;
  await pointer(
    '.control[data-control-id="a"]',
    "pointerdown",
    pointsToPixels(12),
    pointsToPixels(12),
  );
  await pointer(".designer-viewport", "pointermove", pointsToPixels(22), pointsToPixels(12));
  const externalUpdate = { ...model, caption: "Updated externally" };
  await act(() =>
    render(
      <DesignerCanvas
        document={externalUpdate}
        editable
        parentEditable
        productivityEditable
        onCommit={(operations) => emitted.push(operations as ProductivityOperation[])}
      />,
      root,
    ),
  );
  await pointer(".designer-viewport", "pointerup", pointsToPixels(22), pointsToPixels(12));
  assert.equal(
    emitted.length,
    beforeExternalUpdate,
    "a source update cancels the active group gesture",
  );
  assert.equal(root.querySelector(".form-title")?.textContent, "Updated externally");

  await act(() => render(null, root));
  let contextTarget = "";
  await act(() =>
    render(
      <DesignerCanvas
        document={model}
        productivityEditable
        onContextMenu={(id) => (contextTarget = id)}
      />,
      root,
    ),
  );
  const rightClick = new dom.window.MouseEvent("contextmenu", {
    bubbles: true,
    button: 2,
    clientX: 45,
    clientY: 15,
  });
  await act(() => root.querySelector('.control[data-control-id="b"]')!.dispatchEvent(rightClick));
  assert.equal(contextTarget, "b");
  assert.ok(root.querySelector('.selection-outline[data-control-id="b"]'));

  await act(() => render(null, root));
  await act(() => render(<DesignerCanvas document={model} productivityEditable />, root));
  const setRect = (selector: string, left: number, top: number, right: number, bottom: number) => {
    const element = root.querySelector<HTMLElement>(selector);
    assert.ok(element, `marquee fixture has no ${selector} element`);
    Object.defineProperty(element, "getBoundingClientRect", {
      configurable: true,
      value: () => ({ left, top, right, bottom, width: right - left, height: bottom - top }),
    });
  };
  setRect(".form", 0, 0, 300, 180);
  setRect('.control[data-control-id="a"]', 10, 10, 30, 20);
  setRect('.control[data-control-id="b"]', 40, 10, 60, 20);
  setRect('.control[data-control-id="frame"]', 120, 20, 220, 110);
  await pointer(".form", "pointerdown", 0, 0);
  await pointer(".designer-viewport", "pointermove", 70, 40);
  assert.ok(root.querySelector(".marquee-selection"));
  await pointer(".designer-viewport", "pointerup", 70, 40);
  assert.ok(root.querySelector('.selection-outline[data-control-id="a"]'));
  assert.ok(root.querySelector('.selection-outline[data-control-id="b"]'));
  assert.equal(root.querySelector('.selection-outline[data-control-id="c"]'), null);
  assert.ok(root.querySelector(".group-selection-outline"));

  await act(() => render(null, root));
  const resizeDocument = {
    ...model,
    controls: model.controls.map((control) =>
      control.id === "a" ? { ...control, width: 100, height: 40 } : control,
    ),
  };
  const resizeOperations: unknown[][] = [];
  await act(() =>
    render(
      <DesignerCanvas
        document={resizeDocument}
        editable
        productivityEditable
        onCommit={(operations) => resizeOperations.push(operations)}
      />,
      root,
    ),
  );
  await pointer('.control[data-control-id="a"]', "pointerdown", 40, 40);
  await pointer(".designer-viewport", "pointerup", 40, 40);
  await pointer(".handle-nw", "pointerdown", 40, 40);
  await pointer(".designer-viewport", "pointermove", 48, 44);
  await pointer(".designer-viewport", "pointerup", 48, 44);
  assert.deepEqual(
    resizeOperations.at(-1),
    [
      { type: "moveControl", controlId: "a", left: 16, top: 13 },
      { type: "resizeControl", controlId: "a", width: 94, height: 37 },
    ],
    "handles on the sibling selection outline retain single-control resize",
  );

  await act(() => render(null, root));
  const operations: SemanticOperation[] = [];
  const clipboardActions: { action: string; ids: string[]; parentId: string }[] = [];
  await act(() =>
    render(
      <ProductivityToolbar
        document={model}
        selectionIds={["a", "b", "c"]}
        propertyGrid={propertyGrid}
        strings={designerStrings}
        enabled
        structuralEnabled
        tabEnabled
        clipboardAvailable
        canPaste
        busy={false}
        onOperation={(operation) => operations.push(operation)}
        onClipboard={(action, ids, parentId) => clipboardActions.push({ action, ids, parentId })}
        onProperties={() => {}}
        onCloseMenu={() => {}}
      />,
      root,
    ),
  );
  await openCommands(designerStrings.align);
  const alignLeft = root.querySelectorAll<HTMLButtonElement>("button");
  const alignButton = [...alignLeft].find(
    (button) => button.textContent === designerStrings.alignLeft,
  )!;
  await act(() => alignButton.click());
  assert.deepEqual(operations.at(-1), {
    type: "alignControls",
    controlIds: ["a", "b", "c"],
    anchorId: "c",
    alignment: "left",
  });
  await openCommands(designerStrings.distribute);
  const distribute = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.distributeHorizontal,
  )!;
  await act(() => distribute.click());
  assert.deepEqual(operations.at(-1), {
    type: "distributeControls",
    controlIds: ["a", "b", "c"],
    axis: "horizontal",
  });

  const apply = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.applyTabOrder,
  )!;
  assert.equal(apply.disabled, true, "an already contiguous source order is a no-op");
  await click('button[aria-label="Move down: LabelA"]');
  assert.equal(apply.disabled, false);
  await act(() => apply.click());
  assert.deepEqual(operations.at(-1), {
    type: "setTabOrder",
    parentId: "",
    controlIds: ["b", "a", "c"],
  });

  const copy = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.copy,
  )!;
  await act(() => copy.click());
  const paste = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.paste,
  )!;
  await act(() => paste.click());
  assert.deepEqual(clipboardActions, [
    { action: "copy", ids: ["a", "b", "c"], parentId: "" },
    { action: "paste", ids: [], parentId: "" },
  ]);

  await act(() =>
    render(
      <ProductivityToolbar
        document={model}
        selectionIds={["frame"]}
        propertyGrid={propertyGrid}
        strings={designerStrings}
        enabled
        structuralEnabled
        tabEnabled
        clipboardAvailable
        canPaste
        busy={false}
        onOperation={(operation) => operations.push(operation)}
        onClipboard={(action, ids, parentId) => clipboardActions.push({ action, ids, parentId })}
        onProperties={() => {}}
        onCloseMenu={() => {}}
      />,
      root,
    ),
  );
  await openCommands(designerStrings.moreActions);
  const duplicateFrame = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.duplicate,
  )!;
  await act(() => duplicateFrame.click());
  const pasteIntoFrame = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.paste,
  )!;
  await act(() => pasteIntoFrame.click());
  assert.deepEqual(clipboardActions.slice(-2), [
    { action: "duplicate", ids: ["frame"], parentId: "" },
    { action: "paste", ids: [], parentId: "frame" },
  ]);

  await click('button[aria-label="Move down: LabelA"]');
  const externallyUpdatedGrid = {
    ...propertyGrid,
    controls: { ...propertyGrid.controls },
  };
  await act(() =>
    render(
      <ProductivityToolbar
        document={{ ...model }}
        selectionIds={["a", "b", "c"]}
        propertyGrid={externallyUpdatedGrid}
        strings={designerStrings}
        enabled
        structuralEnabled
        tabEnabled
        clipboardAvailable
        canPaste
        busy={false}
        onOperation={(operation) => operations.push(operation)}
        onClipboard={(action, ids, parentId) => clipboardActions.push({ action, ids, parentId })}
        onProperties={() => {}}
        onCloseMenu={() => {}}
      />,
      root,
    ),
  );
  const externallyResetOrder = [...root.querySelectorAll(".tab-order-editor li span")].map(
    (element) => element.textContent,
  );
  assert.deepEqual(externallyResetOrder, ["0. LabelA", "1. LabelB", "2. LabelC"]);
  const externallyResetApply = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.applyTabOrder,
  )!;
  assert.equal(externallyResetApply.disabled, true, "a new source version discards local ordering");

  const gappedPropertyGrid: PropertyGridData = {
    ...propertyGrid,
    controls: {
      ...propertyGrid.controls,
      b: {
        ...propertyGrid.controls.b,
        values: { tabIndex: { present: true, value: 4 } },
      },
      c: {
        ...propertyGrid.controls.c,
        values: {},
      },
    },
  };
  await act(() =>
    render(
      <ProductivityToolbar
        document={model}
        selectionIds={["a", "b", "c"]}
        propertyGrid={gappedPropertyGrid}
        strings={designerStrings}
        enabled={false}
        structuralEnabled
        tabEnabled
        clipboardAvailable
        canPaste={false}
        busy={false}
        onOperation={(operation) => operations.push(operation)}
        onClipboard={(action, ids, parentId) => clipboardActions.push({ action, ids, parentId })}
        onProperties={() => {}}
        onCloseMenu={() => {}}
      />,
      root,
    ),
  );
  const normalizeTabOrder = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.applyTabOrder,
  )!;
  assert.equal(normalizeTabOrder.disabled, false, "missing and gapped indexes need normalization");
  await act(() => normalizeTabOrder.click());
  assert.deepEqual(operations.at(-1), {
    type: "setTabOrder",
    parentId: "",
    controlIds: ["a", "b", "c"],
  });
  const deleteWithGeometryDenied = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.deleteItem,
  )!;
  assert.equal(
    deleteWithGeometryDenied.disabled,
    false,
    "Delete requires structural, not geometry, edit",
  );
  await act(() => deleteWithGeometryDenied.click());
  assert.deepEqual(operations.at(-1), {
    type: "removeControls",
    controlIds: ["a", "b", "c"],
    cascade: true,
  });
  const copyWithoutGeometryEdit = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.copy,
  )!;
  assert.equal(
    copyWithoutGeometryEdit.disabled,
    false,
    "copy needs productivity capability, not geometry edit",
  );
  await act(() => copyWithoutGeometryEdit.click());
  assert.equal(clipboardActions.at(-1)?.action, "copy");

  let propertiesOpened = false;
  await act(() =>
    render(
      <ProductivityToolbar
        document={customProgIdDocument}
        selectionIds={["b"]}
        strings={designerStrings}
        enabled={false}
        structuralEnabled
        tabEnabled={false}
        clipboardAvailable={false}
        canPaste={false}
        busy={false}
        onOperation={(operation) => operations.push(operation)}
        onClipboard={(action, ids, parentId) => clipboardActions.push({ action, ids, parentId })}
        onProperties={() => (propertiesOpened = true)}
        onCloseMenu={() => {}}
      />,
      root,
    ),
  );
  const customDelete = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.deleteItem,
  )!;
  assert.equal(customDelete.disabled, false, "single custom controls retain structural deletion");
  await act(() => customDelete.click());
  assert.deepEqual(operations.at(-1), {
    type: "removeControl",
    controlId: "b",
    cascade: true,
  });
  await openCommands(designerStrings.moreActions);
  const customProperties = [...root.querySelectorAll<HTMLButtonElement>("button")].find(
    (button) => button.textContent === designerStrings.propertiesAction,
  )!;
  assert.equal(customProperties.disabled, false, "single custom controls retain Properties access");
  await act(() => customProperties.click());
  assert.equal(propertiesOpened, true);

  const toolbarProps = {
    document: model,
    selectionIds: ["page"],
    strings: designerStrings,
    enabled: true,
    structuralEnabled: true,
    tabEnabled: true,
    clipboardAvailable: true,
    canPaste: true,
    busy: false,
    onOperation: (operation: SemanticOperation) => operations.push(operation),
    onClipboard: () => {},
    onProperties: () => {},
    onCloseMenu: () => {},
  };
  await act(() =>
    render(<ProductivityToolbar {...toolbarProps} contextMenu={{ x: 2, y: 3 }} />, root),
  );
  const pageDeleteButtons = [...root.querySelectorAll<HTMLButtonElement>("button")].filter(
    (button) => button.textContent === designerStrings.deleteItem,
  );
  assert.equal(pageDeleteButtons.length, 2);
  assert.ok(
    pageDeleteButtons.every((button) => button.disabled),
    "Page Delete disabled in toolbar and menu",
  );

  const childMetadata = propertyGrid.controls.a;
  const nestedMetadata: PropertyGridData = {
    ...propertyGrid,
    controls: { ...propertyGrid.controls, nested: childMetadata, "page-child": childMetadata },
  };
  const parentModel = {
    ...model,
    controls: [
      ...model.controls,
      { ...model.controls[0], id: "page-child", name: "PageChild", parentId: "page" },
    ],
  };
  for (const invalid of [false, true]) {
    const currentModel = invalid
      ? {
          ...parentModel,
          controls: parentModel.controls.map((control) =>
            ["frame", "multi"].includes(control.id)
              ? { ...control, progId: `Vendor.${control.type}.1` }
              : control,
          ),
        }
      : parentModel;
    await act(() =>
      render(
        <ProductivityToolbar
          {...toolbarProps}
          document={currentModel}
          selectionIds={["nested"]}
          propertyGrid={nestedMetadata}
        />,
        root,
      ),
    );
    assert.deepEqual(
      [...root.querySelectorAll<HTMLOptionElement>(".tab-order-editor select option")].map(
        (option) => option.value,
      ),
      invalid ? [""] : ["", "frame", "page"],
      "tab-order parents use server-compatible containers",
    );
    if (invalid)
      assert.equal(root.querySelector<HTMLSelectElement>(".tab-order-editor select")!.value, "");
  }

  let menuClosures = 0;
  const menuProps = {
    ...toolbarProps,
    selectionIds: ["a"],
    structuralEnabled: false,
    onCloseMenu: () => {
      menuClosures++;
    },
    contextMenu: { x: 3, y: 4 },
  };
  let paneClosures = 0;
  const paneKeyDown = (event: KeyboardEvent) => {
    if (event.key === "Escape" && !event.defaultPrevented) paneClosures++;
  };
  await act(() =>
    render(
      <div onKeyDown={paneKeyDown}>
        <ProductivityToolbar {...menuProps} />
      </div>,
      root,
    ),
  );
  const enabledMenuItems = [
    ...root.querySelectorAll<HTMLButtonElement>('[role="menuitem"]:not(:disabled)'),
  ];
  assert.equal(
    document.activeElement,
    enabledMenuItems[0],
    "opening menu skips disabled items and focuses first enabled action",
  );
  await act(() =>
    enabledMenuItems[0].dispatchEvent(
      new dom.window.KeyboardEvent("keydown", { bubbles: true, key: "ArrowDown" }),
    ),
  );
  assert.equal(document.activeElement, enabledMenuItems[1]);
  await pointer('[role="menu"]', "pointerdown", 0, 0);
  assert.equal(menuClosures, 0, "inside pointer-down keeps menu open");
  await act(() =>
    enabledMenuItems[0].dispatchEvent(
      new dom.window.KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
    ),
  );
  assert.equal(menuClosures, 1);
  assert.equal(paneClosures, 0, "Escape closes the context menu before the enclosing pane");
  await act(() =>
    document.body.dispatchEvent(new dom.window.MouseEvent("pointerdown", { bubbles: true })),
  );
  assert.equal(menuClosures, 2, "outside pointer-down closes menu");
  await act(() => render(null, root));
  await act(() =>
    dom.window.dispatchEvent(new dom.window.KeyboardEvent("keydown", { key: "Escape" })),
  );
  assert.equal(menuClosures, 2, "unmount removes menu listeners");

  for (const [selection, groups] of [
    [[], [designerStrings.moreActions]],
    [["a"], [designerStrings.order, designerStrings.moreActions]],
    [
      ["a", "b"],
      [
        designerStrings.align,
        designerStrings.sameSize,
        designerStrings.order,
        designerStrings.moreActions,
      ],
    ],
    [
      ["a", "b", "c"],
      [
        designerStrings.align,
        designerStrings.distribute,
        designerStrings.sameSize,
        designerStrings.order,
        designerStrings.moreActions,
      ],
    ],
    [["a", "nested"], [designerStrings.moreActions]],
  ] as [string[], string[]][]) {
    await act(() => render(<CommandToolbar {...toolbarProps} selectionIds={selection} />, root));
    assert.deepEqual(
      [...root.querySelectorAll('[aria-haspopup="menu"]')].map((button) => button.textContent),
      groups,
      "command groups reflect eligible sibling selection count",
    );
  }
  await act(() =>
    render(
      <div onKeyDown={paneKeyDown}>
        <CommandToolbar {...toolbarProps} selectionIds={["a", "b", "c"]} busy />
      </div>,
      root,
    ),
  );
  await openCommands(designerStrings.align);
  assert.equal(
    root.querySelectorAll('[role="menuitem"]:not(:disabled)').length,
    0,
    "busy layout commands stay disabled",
  );
  const alignTrigger = [...root.querySelectorAll<HTMLButtonElement>('[aria-haspopup="menu"]')].find(
    (button) => button.textContent === designerStrings.align,
  )!;
  await act(() =>
    root
      .querySelector('[role="menu"]')!
      .dispatchEvent(
        new dom.window.KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
      ),
  );
  assert.equal(root.querySelector('[role="menu"]'), null);
  assert.equal(document.activeElement, alignTrigger, "Escape returns focus to the group trigger");
  assert.equal(paneClosures, 0, "Escape closes the command menu before the enclosing pane");
  await act(() => render(null, root));
  process.stdout.write(
    "UserForm productivity selection, transaction, and toolbar assertions passed\n",
  );
}

run().catch((error) => {
  process.stderr.write(`${error.stack ?? error}\n`);
  process.exitCode = 1;
});
