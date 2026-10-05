import * as assert from "assert";
import { designerDocument, pixelsToPoints, pointsToPixels } from "../../src/userFormEditor/model";
import {
  isControlVisible,
  geometryOperations,
  moveGeometry,
  parentBounds,
  requiredSize,
  resizeGeometry,
  resizeHandles,
  zoomLevels,
} from "../../src/userFormEditor/geometry";
import { isWebviewMessage } from "../../src/userFormEditor/protocol";

export function runUserFormGeometryAssertions(): void {
  for (const zoom of zoomLevels) {
    for (const point of [0, 0.123456789, 0.75, 99.875, 240])
      assert.ok(Math.abs(pixelsToPoints(pointsToPixels(point, zoom), zoom) - point) < 1e-12);
  }
  const start = { left: 10.123456789, top: 12.5, width: 50, height: 30 };
  const bounds = { width: 200, height: 100 };
  assert.deepStrictEqual(moveGeometry(start, -100, 100, bounds), { ...start, left: 0, top: 70 });
  assert.deepStrictEqual(moveGeometry(start, 3, 4, bounds, true), { ...start, left: 16, top: 16 });
  assert.strictEqual(moveGeometry({ ...start, width: 201 }, 1, 1, bounds), undefined);
  const overflowX = { left: 50, top: 10, width: 100, height: 20 };
  const square = { width: 100, height: 100 };
  const minimum = { width: 1, height: 1 };
  for (const handle of ["n", "s"] as const)
    assert.strictEqual(
      resizeGeometry(overflowX, handle, 0, 10, square, minimum),
      undefined,
      "vertical resize cannot silently correct horizontal overflow",
    );
  const overflowY = { left: 10, top: 50, width: 20, height: 100 };
  for (const handle of ["e", "w"] as const)
    assert.strictEqual(
      resizeGeometry(overflowY, handle, 10, 0, square, minimum),
      undefined,
      "horizontal resize cannot silently correct vertical overflow",
    );
  assert.deepStrictEqual(
    resizeGeometry(overflowX, "e", -75, 0, square, minimum),
    { ...overflowX, width: 25 },
    "an explicit width resize can repair horizontal overflow",
  );
  assert.deepStrictEqual(resizeGeometry(overflowY, "s", 0, -75, square, minimum), {
    ...overflowY,
    height: 25,
  });
  for (const handle of resizeHandles) {
    const end = resizeGeometry(start, handle, 4, 3, bounds, { width: 1, height: 1 })!;
    assert.ok(end.width >= 1 && end.height >= 1, handle);
    assert.ok(
      end.left >= 0 &&
        end.top >= 0 &&
        end.left + end.width <= bounds.width &&
        end.top + end.height <= bounds.height,
      handle,
    );
    assert.strictEqual(end.left, start.left + (handle.includes("w") ? 4 : 0));
    assert.strictEqual(end.top, start.top + (handle.includes("n") ? 3 : 0));
  }
  const nw = resizeGeometry(start, "nw", 4, 3, bounds, { width: 1, height: 1 })!;
  assert.deepStrictEqual(
    geometryOperations("button", start, nw).map((op) => op.type),
    ["moveControl", "resizeControl"],
  );
  assert.deepStrictEqual(geometryOperations("button", start, start), []);
  let current = start;
  for (let index = 0; index < 500; index++) {
    current = moveGeometry(current, pixelsToPoints(13, 1.25), pixelsToPoints(7, 1.25), bounds)!;
    current = moveGeometry(current, -pixelsToPoints(13, 1.25), -pixelsToPoints(7, 1.25), bounds)!;
  }
  assert.ok(
    Math.abs(current.left - start.left) < 1e-10 && Math.abs(current.top - start.top) < 1e-10,
    "no accumulated round-trip drift",
  );
  let resizedCurrent = start;
  for (let index = 0; index < 500; index++) {
    resizedCurrent = resizeGeometry(
      resizedCurrent,
      "se",
      pixelsToPoints(7, 1.25),
      pixelsToPoints(3, 1.25),
      bounds,
      minimum,
    )!;
    resizedCurrent = resizeGeometry(
      resizedCurrent,
      "se",
      -pixelsToPoints(7, 1.25),
      -pixelsToPoints(3, 1.25),
      bounds,
      minimum,
    )!;
  }
  assert.ok(
    Math.abs(resizedCurrent.width - start.width) < 1e-10 &&
      Math.abs(resizedCurrent.height - start.height) < 1e-10,
    "resize round trips do not accumulate coordinate drift",
  );
  const model = designerDocument({
    form: { name: "Main", build: { clientWidth: 200, clientHeight: 100 } },
    controls: [
      { id: "frame", name: "Frame", type: "Frame", width: 100, height: 80 },
      {
        id: "nested",
        name: "Nested",
        type: "Frame",
        parentId: "frame",
        left: 5,
        top: 7,
        width: 50,
        height: 40,
      },
      {
        id: "child",
        name: "Child",
        type: "TextBox",
        parentId: "nested",
        left: 10,
        top: 5,
        width: 20,
        height: 18,
      },
      { id: "multi", name: "Multi", type: "MultiPage", width: 100, height: 80 },
      { id: "page", name: "Page", type: "Page", parentId: "multi" },
      { id: "hidden", name: "Hidden", type: "Page", parentId: "multi", visible: false },
      {
        id: "pagechild",
        name: "PageChild",
        type: "TextBox",
        parentId: "page",
        left: 10,
        top: 10,
        width: 20,
        height: 20,
      },
      {
        id: "hiddenchild",
        name: "HiddenChild",
        type: "TextBox",
        parentId: "hidden",
        left: 50,
        top: 30,
        width: 40,
        height: 30,
      },
    ],
  });
  assert.deepStrictEqual(parentBounds(model, model.controls[2]), { width: 48.5, height: 38.5 });
  assert.deepStrictEqual(parentBounds(model, model.controls[6]), { width: 97, height: 60.5 });
  assert.deepStrictEqual(requiredSize(model, model.controls[0]), { width: 56.5, height: 48.5 });
  assert.deepStrictEqual(
    requiredSize(model, model.controls[3]),
    { width: 93, height: 79.5 },
    "hidden Page protects its children",
  );
  assert.ok(isControlVisible(model, "pagechild"));
  assert.ok(!isControlVisible(model, "hiddenchild"));
  model.controls[3].selectedIndex = 1;
  assert.ok(
    !isControlVisible(model, "pagechild"),
    "non-selected Page descendants cannot retain keyboard selection",
  );
  const resized = resizeGeometry(
    model.controls[0],
    "se",
    -1000,
    -1000,
    bounds,
    requiredSize(model, model.controls[0]),
  )!;
  assert.deepStrictEqual([resized.width, resized.height], [56.5, 48.5]);
  assert.strictEqual(
    resizeGeometry(start, "se", 1, 1, bounds, { width: 201, height: 1 }),
    undefined,
  );
  const message = {
    type: "edit",
    requestId: 1,
    version: 2,
    operations: [{ type: "moveControl", controlId: "child", left: 0, top: 0 }],
  };
  assert.ok(isWebviewMessage(message));
  assert.ok(
    !isWebviewMessage({ ...message, operations: [{ ...message.operations[0], left: NaN }] }),
  );
  assert.ok(
    !isWebviewMessage({
      ...message,
      operations: [{ ...message.operations[0], parentId: "other" }],
    }),
  );
  assert.ok(
    !isWebviewMessage({ ...message, operations: [message.operations[0], message.operations[0]] }),
  );
}
