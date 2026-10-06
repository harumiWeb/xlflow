import * as assert from "assert";
import * as path from "path";
import { isFormSpecPath, PreviewSynchronizer } from "../../src/userFormEditor/document";
import { designerDocument, pointsToPixels } from "../../src/userFormEditor/model";
import { isWebviewMessage } from "../../src/userFormEditor/protocol";
import type { HostMessage, PreviewResult } from "../../src/userFormEditor/protocol";

export async function runUserFormEditorAssertions(): Promise<void> {
  const root = path.resolve("preview-project");
  assert.ok(isFormSpecPath(path.join(root, "custom/forms/specs/Main.yaml"), root, "custom/forms"));
  assert.ok(isFormSpecPath(path.join(root, "custom/forms/specs/Main.json"), root, "custom/forms"));
  assert.ok(!isFormSpecPath(path.join(root, "src/forms/specs/Main.yaml"), root, "custom/forms"));
  assert.ok(
    !isFormSpecPath(path.join(root, "custom/forms/specs/nested/Main.yaml"), root, "custom/forms"),
  );
  assert.ok(!isFormSpecPath(path.join(root, "other.yaml"), root, "custom/forms"));
  assert.ok(isWebviewMessage({ type: "ready" }));
  assert.ok(!isWebviewMessage({ type: "moveControl" }));
  assert.ok(!isWebviewMessage(null));
  const property = { type: "setFormProperty", field: "caption", value: "" };
  const edit = { type: "edit", requestId: 1, version: 1, operations: [property] };
  for (const value of [null, "", 0, false, "false", 12.5]) {
    assert.ok(isWebviewMessage({ ...edit, operations: [{ ...property, value }] }));
  }
  assert.ok(
    isWebviewMessage({
      ...edit,
      operations: [
        { type: "setControlProperty", controlId: "label", field: "caption", value: "hi" },
      ],
    }),
  );
  for (const value of [undefined, NaN, Infinity, {}, []]) {
    assert.ok(!isWebviewMessage({ ...edit, operations: [{ ...property, value }] }));
  }
  assert.ok(
    !isWebviewMessage({ ...edit, operations: [{ type: "setFormProperty", field: "caption" }] }),
  );
  assert.ok(!isWebviewMessage({ ...edit, operations: [{ ...property, controlId: "label" }] }));
  assert.ok(!isWebviewMessage({ ...edit, operations: [property, property] }));
  assert.ok(
    !isWebviewMessage({
      ...edit,
      operations: [property, { type: "moveControl", controlId: "label", left: 1, top: 2 }],
    }),
  );
  assert.strictEqual(pointsToPixels(0.75), 1);
  const document = {
    form: { name: "Main", build: { clientWidth: 123.5, clientHeight: 100 } },
    controls: [
      {
        id: "label",
        type: "Label",
        name: "Label1",
        caption: "<script>alert(1)</script>",
        left: 1.5,
      },
    ],
  };
  const model = designerDocument(document);
  assert.strictEqual(model.width, 123.5);
  assert.strictEqual(model.controls[0].left, 1.5);
  assert.strictEqual(model.controls[0].width, 72);
  assert.strictEqual(model.controls[0].caption, document.controls[0].caption);
  assert.ok(model.controls[0].approximate);
  assert.ok(!("width" in document.controls[0]), "normalization must not mutate input");

  const sync = new PreviewSynchronizer();
  const messages: HostMessage[] = [];
  const send = (message: HostMessage) => messages.push(message);
  let version = 1;
  await sync.update(
    1,
    async () => ({ version: 1, document }),
    () => version,
    send,
  );
  assert.strictEqual(messages.at(-1)?.type, "document");
  const propertyGrid = {
    form: { descriptors: [], values: { caption: { present: true, value: "" } } },
    controls: {},
  };
  await sync.update(
    1,
    async () => ({ version: 1, document, editable: true, propertyEditable: true, propertyGrid }),
    () => version,
    send,
  );
  const propertyMessage = messages.at(-1);
  assert.ok(propertyMessage?.type === "document");
  assert.strictEqual(propertyMessage.propertyEditable, true);
  assert.deepStrictEqual(
    propertyMessage.propertyGrid,
    propertyGrid,
    "authored values survive renderer normalization",
  );
  version = 2;
  await sync.update(
    2,
    async () => ({ version: 2, error: { code: "syntax", message: "Invalid YAML" } }),
    () => version,
    send,
  );
  assert.strictEqual(messages.at(-1)?.type, "invalidDocument");
  const replay: HostMessage[] = [];
  sync.replay((m) => replay.push(m));
  assert.deepStrictEqual(
    replay.map((m) => m.type),
    ["document", "invalidDocument"],
  );

  let resolveOld!: (result: PreviewResult) => void;
  const pending = sync.update(
    2,
    () =>
      new Promise((resolve) => {
        resolveOld = resolve;
      }),
    () => version,
    send,
  );
  version = 3;
  await sync.update(
    3,
    async () => ({ version: 3, document }),
    () => version,
    send,
  );
  const length = messages.length;
  resolveOld({ version: 2, error: { code: "old", message: "stale" } });
  await pending;
  assert.strictEqual(messages.length, length, "stale result must not overwrite recovery");
  const restored: HostMessage[] = [];
  sync.replay((m) => restored.push(m));
  assert.deepStrictEqual(
    restored.map((m) => m.type),
    ["document"],
  );
  const pendingConnection = sync.update(
    3,
    () =>
      new Promise((resolve) => {
        resolveOld = resolve;
      }),
    () => version,
    send,
  );
  sync.invalidate();
  resolveOld({ version: 3, document });
  await pendingConnection;
  assert.strictEqual(messages.length, length, "connection change must retire pending result");
  const pendingDispose = sync.update(
    3,
    () =>
      new Promise((resolve) => {
        resolveOld = resolve;
      }),
    () => version,
    send,
  );
  sync.dispose();
  resolveOld({ version: 3, document });
  await pendingDispose;
  assert.strictEqual(messages.length, length, "disposed editor must not receive results");
}
