import * as assert from "node:assert/strict";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { createHash } from "node:crypto";
import type * as vscode from "vscode";
import { PictureImportSession } from "../src/userFormEditor/pictures";
import type {
  PictureValidationRequest,
  PictureValidationResult,
} from "../src/userFormEditor/pictures";

// This validator stub checks host transactions; actual decode/root checks have
// Go coverage in userform_picture_test.go and picture's decoder tests.
async function run(): Promise<void> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "xlflow-picture-host-"));
  const root = path.join(directory, "project");
  await fs.mkdir(root);
  const bytes = Buffer.from("validated image bytes");
  const source = path.join(directory, "external.jpeg");
  await fs.writeFile(source, bytes);
  const hash = createHash("sha256").update(bytes).digest("hex");
  let selected: string | undefined = source;
  const validate = async (request: PictureValidationRequest): Promise<PictureValidationResult> => {
    const data = request.path
      ? await fs.readFile(path.join(root, request.path))
      : Buffer.from(request.data!, "base64");
    return {
      format: "jpeg",
      sha256: createHash("sha256").update(data).digest("hex"),
      width: 2,
      height: 3,
      dataURI: `data:image/jpeg;base64,${data.toString("base64")}`,
    };
  };
  const session = (formsRoot = "custom/forms") =>
    new PictureImportSession({
      projectRoot: root,
      formsRoot,
      sourceURI: "file:///Main.yaml",
      validate,
      pick: async () =>
        selected ? ({ scheme: "file", fsPath: selected } as vscode.Uri) : undefined,
    });
  const asset = path.join(root, "custom/forms/assets", `${hash}.jpg`);
  try {
    const cancel = session();
    const selection = await cancel.open();
    assert.ok(selection?.dataURI.startsWith("data:image/jpeg;base64,"));
    await assert.rejects(fs.stat(path.dirname(asset)), { code: "ENOENT" });
    cancel.cancel();
    await assert.rejects(cancel.commit(selection!.token), /expired/);
    selected = undefined;
    assert.equal(await cancel.open(), undefined);

    selected = source;
    const first = session();
    const firstSelection = (await first.open())!;
    await assert.rejects(first.commit("untrusted"), /expired/);
    const receipt = await first.commit(firstSelection.token);
    assert.equal(receipt.picture.path, `custom/forms/assets/${hash}.jpg`);
    assert.deepEqual(await fs.readFile(asset), bytes);
    await first.rollback(receipt, async () => true);
    assert.deepEqual(await fs.readFile(asset), bytes, "referenced assets survive failed edits");
    await first.rollback(receipt, async () => false);
    await assert.rejects(fs.stat(asset), { code: "ENOENT" });
    // A failed edit retains its selection for retry.
    const retry = await first.commit(firstSelection.token);
    first.accept(retry);
    await first.rollback(retry, async () => false);
    assert.deepEqual(await fs.readFile(asset), bytes, "successful assets survive Undo/Redo");

    const reuse = session();
    const reuseSelection = (await reuse.open())!;
    const reused = await reuse.commit(reuseSelection.token);
    await reuse.rollback(reused, async () => false);
    assert.deepEqual(
      await fs.readFile(asset),
      bytes,
      "preexisting assets never belong to rollback",
    );
    await fs.writeFile(asset, "different content");
    await assert.rejects(reuse.commit(reuseSelection.token), /different asset/);
    assert.equal(await fs.readFile(asset, "utf8"), "different content");
    await fs.writeFile(asset, bytes);

    const partialSession = session("partial/forms");
    const partialSelection = (await partialSession.open())!;
    const probe = await fs.open(source, "r");
    const prototype = Object.getPrototypeOf(probe) as { writeFile: typeof probe.writeFile };
    const writeFile = prototype.writeFile;
    await probe.close();
    try {
      prototype.writeFile = async function (this: fs.FileHandle): Promise<void> {
        await this.write(bytes, 0, 5, 0);
        throw new Error("simulated partial write");
      };
      await assert.rejects(partialSession.commit(partialSelection.token), /partial write/);
    } finally {
      prototype.writeFile = writeFile;
    }
    const partialAsset = path.join(root, "partial/forms/assets", `${hash}.jpg`);
    assert.deepEqual(await fs.readFile(partialAsset), bytes.subarray(0, 5));
    await partialSession.rollback({ token: partialSelection.token }, async () => false);
    await assert.rejects(fs.stat(partialAsset), { code: "ENOENT" });

    const modifiedSession = session("modified/forms");
    const modifiedSelection = (await modifiedSession.open())!;
    const modifiedReceipt = await modifiedSession.commit(modifiedSelection.token);
    const modifiedAsset = path.join(root, modifiedReceipt.picture.path);
    await fs.writeFile(modifiedAsset, "unrelated modification");
    await assert.rejects(
      modifiedSession.rollback(modifiedReceipt, async () => false),
      /changed/,
    );
    assert.equal(await fs.readFile(modifiedAsset, "utf8"), "unrelated modification");

    selected = path.join(root, "existing.jpeg");
    await fs.writeFile(selected, bytes);
    const existing = session();
    const existingSelection = (await existing.open())!;
    const existingReceipt = await existing.commit(existingSelection.token);
    assert.equal(existingReceipt.picture.path, "existing.jpeg");
    await existing.rollback(existingReceipt, async () => false);
    assert.deepEqual(await fs.readFile(selected), bytes);
    await fs.writeFile(selected, "modified");
    await assert.rejects(existing.commit(existingSelection.token), /changed/);

    selected = source;
    const escape = session("../outside");
    const escapeSelection = (await escape.open())!;
    await assert.rejects(escape.commit(escapeSelection.token), /inside the project/);
    // Windows junction creation requires no symlink privilege.
    const linked = path.join(root, "linked");
    await fs.symlink(directory, linked, "junction");
    const linkedSession = session("linked/forms");
    const linkedSelection = (await linkedSession.open())!;
    await assert.rejects(linkedSession.commit(linkedSelection.token), /symbolic links|junctions/);
    selected = path.join(linked, "external.jpeg");
    await assert.rejects(session().open(), /links must stay/);
    await fs.unlink(linked);
    console.log("UserForm picture host assertions passed");
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
}

void run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
