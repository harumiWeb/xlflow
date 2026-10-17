import * as assert from "node:assert/strict";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { pictureAssetReferenced } from "../src/userFormEditor/pictureReferences";
import type { PictureReferenceSources } from "../src/userFormEditor/pictureReferences";

async function run(): Promise<void> {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "xlflow-picture-references-"));
  const project = path.join(temporary, "project");
  const checkpoints = path.join(temporary, "external-checkpoints");
  await fs.mkdir(project);
  await fs.mkdir(checkpoints);
  const asset = `${"a".repeat(64)}.jpg`;
  const relative = `src/forms/assets/${asset}`;
  let live = ["form: { name: Main }"];
  const sources: PictureReferenceSources = {
    openTexts: () => live,
    async *sourceTexts() {
      for (const name of await fs.readdir(project))
        yield await fs.readFile(path.join(project, name), "utf8");
    },
    async *checkpointTexts() {
      for (const name of await fs.readdir(checkpoints))
        yield await fs.readFile(path.join(checkpoints, name), "utf8");
    },
  };
  const checkpoint = (text: string) =>
    JSON.stringify({
      schemaVersion: 1,
      sourceUri: "file:///closed.yaml",
      updatedAt: 1,
      state: {
        schemaVersion: 1,
        baseText: text,
        text: "no current picture",
        revision: 1,
        history: [{ before: text, after: "no current picture" }],
        cursor: 1,
        conflicted: false,
      },
    });
  try {
    await fs.writeFile(path.join(project, "Main.yaml"), "no saved picture");
    assert.equal(await pictureAssetReferenced(relative, sources), false);
    // The saved source is clear but an ordinary dirty text editor references it.
    live = [`picture: { path: ${relative} }`];
    assert.equal(await pictureAssetReferenced(relative, sources), true);
    live = [];
    // Closed Designer has no open source; only its external Undo history refers.
    const saved = path.join(checkpoints, "closed-editor.json");
    await fs.writeFile(saved, checkpoint(`picture: { path: ${relative} }`));
    assert.equal(await pictureAssetReferenced(relative, sources), true);
    await fs.writeFile(saved, checkpoint(relative).replace(/a/g, "\\u0061"));
    assert.equal(
      await pictureAssetReferenced(relative, sources),
      true,
      "JSON escapes cannot hide recovery references",
    );
    await fs.writeFile(saved, checkpoint("no reference"));
    assert.equal(await pictureAssetReferenced(relative, sources), false);
    await fs.writeFile(path.join(checkpoints, "pending.tmp"), '{"state":');
    assert.equal(
      await pictureAssetReferenced(relative, sources),
      true,
      "incomplete checkpoint retains the asset",
    );
    await fs.unlink(path.join(checkpoints, "pending.tmp"));
    await fs.writeFile(saved, "null");
    assert.equal(
      await pictureAssetReferenced(relative, sources),
      true,
      "unknown recovery data retains the asset",
    );
    await fs.unlink(saved);
    const unreadable = {
      ...sources,
      async *checkpointTexts(): AsyncGenerator<string> {
        yield await Promise.reject<string>(new Error("storage unavailable"));
      },
    };
    assert.equal(await pictureAssetReferenced(relative, unreadable), true);
    const duringScan = {
      ...sources,
      async *sourceTexts() {
        live = [relative];
        yield "no saved reference";
      },
    };
    assert.equal(
      await pictureAssetReferenced(relative, duringScan),
      true,
      "recheck live buffers after asynchronous filesystem reads",
    );
    live = [];
    let checkpointRevision = 0;
    const initialRevision = checkpointRevision;
    const checkpointPublishedDuringScan = {
      ...sources,
      checkpointStable: () => checkpointRevision === initialRevision,
      async *checkpointTexts() {
        // The directory snapshot was empty. A different editor now persists
        // and closes before the final live-buffer check.
        checkpointRevision++;
        await fs.writeFile(saved, checkpoint(relative));
        yield checkpoint("no reference in the original directory snapshot");
      },
    };
    assert.equal(
      await pictureAssetReferenced(relative, checkpointPublishedDuringScan),
      true,
      "a newly published closed-editor checkpoint invalidates the scan",
    );
    console.log("UserForm picture reference assertions passed");
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}

void run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
