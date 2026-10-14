import * as assert from "assert";
import { mkdtemp, rm, writeFile } from "fs/promises";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";
import { DocumentEditQueue, type EditClient } from "../../src/userFormEditor/edits";
import type {
  EditResult,
  GeometryOperation,
  SemanticOperation,
  SourceTextEdit,
} from "../../src/userFormEditor/protocol";

const formSource = [
  "schemaVersion: 1",
  "kind: xlflow.userform",
  'form: { name: Main, caption: "😀 Café" }',
  "controls:",
  '  - { id: "😀Box", left: 10, top: 20, width: 30, height: 40 }',
  "",
].join("\r\n");

export async function runUserFormEditAssertions(): Promise<void> {
  await assertGeometryTransactionsUndoAndRedo();
  await assertProductivityTransactionsUndoAndRedo();
  await assertCopyReadSharesEditQueue();
  await assertPropertyTransactionsUndoAndRedo();
  await assertFormResizeTransactionsUndoAndRedo();
  await assertStructuralTransactionsUndoAndRedo();
  await assertPageCollectionUndoAndRedo();
  await assertParentTransactionUndoAndRedo();
  await assertConfirmationRejectsStaleAndCancelledEdits();
  await assertStaleResponsesAreDiscarded();
  await assertConnectionInvalidationIsDiscarded();
  await assertConcurrentPanelsAreSerialized();
  await assertUnicodeAndCRLFRanges();
  await assertInvalidRangesAreRejectedAtomically();
  await assertStructuredRejectionPreservesSource();
}

async function assertFormResizeTransactionsUndoAndRedo(): Promise<void> {
  for (const extension of ["yaml", "json"]) {
    const source =
      extension === "yaml"
        ? "schemaVersion: 1\r\nkind: xlflow.userform\r\nform: {name: Main, width: 500, height: 450}\r\ncontrols: []\r\n"
        : '{"schemaVersion":1,"kind":"xlflow.userform","form":{"name":"Main","width":500,"height":450},"controls":[]}\r\n';
    await withDesignerDocument(
      `form-resize-${extension}`,
      source,
      async (document) => {
        const operations: SemanticOperation[] = [
          { type: "setFormProperty", field: "width", value: 520 },
          { type: "setFormProperty", field: "height", value: 470 },
        ];
        const queue = new DocumentEditQueue();
        const client = fakeClient((doc, version, text, requested) => {
          assert.deepStrictEqual(requested, operations);
          return {
            version,
            edits: [
              {
                range: {
                  start: doc.positionAt(text.indexOf("500")),
                  end: doc.positionAt(text.indexOf("500") + 3),
                },
                newText: "520",
              },
              {
                range: {
                  start: doc.positionAt(text.indexOf("450")),
                  end: doc.positionAt(text.indexOf("450") + 3),
                },
                newText: "470",
              },
            ],
          };
        });
        assert.strictEqual(
          await queue.run(document, document.version, operations, client, () => true),
          undefined,
        );
        const resized = source.replace("500", "520").replace("450", "470");
        assert.strictEqual(document.getText(), resized);
        assertActiveTextEditor(document.uri);
        await vscode.commands.executeCommand("undo");
        assert.strictEqual(document.getText(), source, "one Undo restores both form dimensions");
        await vscode.commands.executeCommand("redo");
        assert.strictEqual(document.getText(), resized);
      },
      extension,
    );
  }
}

async function assertProductivityTransactionsUndoAndRedo(): Promise<void> {
  const source = [
    "schemaVersion: 1",
    "kind: xlflow.userform",
    "form: {name: Main}",
    "controls:",
    "  - {id: a, name: Label1, type: Label, left: 10, top: 20}",
    "  - {id: b, name: Label2, type: Label, left: 100, top: 40}",
    "",
  ].join("\r\n");
  await withDesignerDocument("productivity-undo", source, async (document) => {
    const queue = new DocumentEditQueue();
    const client = fakeClient((doc, version, _text, operations) => {
      assert.strictEqual(operations.length, 1);
      const operation = operations[0];
      assert.strictEqual(operation.type, "moveControls");
      const edits = (operation as Extract<SemanticOperation, { type: "moveControls" }>).moves.map(
        (move) => {
          const line = Array.from({ length: doc.lineCount }, (_, index) => doc.lineAt(index)).find(
            (candidate) => candidate.text.includes(`id: ${move.controlId},`),
          )!;
          return {
            range: line.range,
            newText: line.text
              .replace(/left: \d+/, `left: ${move.left}`)
              .replace(/top: \d+/, `top: ${move.top}`),
          };
        },
      );
      return { version, edits };
    });
    const operation: SemanticOperation = {
      type: "moveControls",
      moves: [
        { controlId: "a", left: 18, top: 28 },
        { controlId: "b", left: 108, top: 48 },
      ],
    };
    assert.strictEqual(
      await queue.run(document, document.version, [operation], client, () => true),
      undefined,
    );
    const moved = document.getText();
    assert.ok(moved.includes("left: 18") && moved.includes("left: 108"));
    assertActiveTextEditor(document.uri);
    await vscode.commands.executeCommand("undo");
    assert.strictEqual(document.getText(), source, "one Undo restores every selected control");
    await vscode.commands.executeCommand("redo");
    assert.strictEqual(document.getText(), moved);
  });
}

async function assertCopyReadSharesEditQueue(): Promise<void> {
  await withDesignerDocument("clipboard-queue", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const version = document.version;
    const started = deferred<void>();
    const reply = deferred<EditResult>();
    const operation: GeometryOperation = {
      type: "moveControl",
      controlId: "😀Box",
      left: 18,
      top: 28,
    };
    const moving = queue.run(
      document,
      version,
      [operation],
      fakeClient(() => {
        started.resolve();
        return reply.promise;
      }),
      () => true,
    );
    await withTimeout(started.promise, "edit did not reach the client");
    let reads = 0;
    const copying = queue.capture(
      document,
      version,
      () => true,
      async (text) => {
        reads++;
        return text;
      },
    );
    assert.strictEqual(reads, 0, "copy must wait for same-URI edits");
    reply.resolve({ version, edits: geometryEdits(document.getText(), [operation]) });
    assert.strictEqual(await moving, undefined);
    assert.strictEqual((await copying).error?.code, "staleDocument");
    assert.strictEqual(reads, 0, "a copy based on the old version never reads source");

    const copyStarted = deferred<void>();
    const copyReply = deferred<string>();
    let current = true;
    const pending = queue.capture(
      document,
      document.version,
      () => current,
      async () => {
        copyStarted.resolve();
        return copyReply.promise;
      },
    );
    await withTimeout(copyStarted.promise, "copy did not reach the client");
    current = false;
    copyReply.resolve("stale snapshot");
    const result = await pending;
    assert.strictEqual(result.error?.code, "staleDocument");
    assert.strictEqual(
      result.value,
      undefined,
      "a retired Webview cannot publish a clipboard snapshot",
    );
  });
}

async function assertPageCollectionUndoAndRedo(): Promise<void> {
  const source = [
    "schemaVersion: 1",
    "kind: xlflow.userform",
    "form: {name: Main}",
    "controls:",
    "  - {id: multi, name: Multi, type: MultiPage, selectedIndex: 1}",
    "  - {id: keep, name: Page1, type: Page, parentId: multi}",
    "  - {id: deleted, name: Page2, type: Page, parentId: multi}",
    "  - {id: child, name: Child, type: Label, parentId: deleted}",
    "",
  ].join("\r\n");
  await withDesignerDocument("page-undo", source, async (document) => {
    const initial = document.getText();
    const queue = new DocumentEditQueue();
    const client = fakeClient((doc, version, text, operations) => {
      assert.deepStrictEqual(operations, [
        { type: "removePage", controlId: "deleted", cascade: true },
      ]);
      const edits: SourceTextEdit[] = [scalarEdit(text, "selectedIndex", 0)];
      for (let line = 0; line < doc.lineCount; line++)
        if (/id: (deleted|child),/.test(doc.lineAt(line).text))
          edits.push({ range: doc.lineAt(line).rangeIncludingLineBreak, newText: "" });
      return { version, edits };
    });
    assert.strictEqual(
      await queue.run(
        document,
        document.version,
        [{ type: "removePage", controlId: "deleted", cascade: true }],
        client,
        () => true,
      ),
      undefined,
    );
    const removed = document.getText();
    assert.ok(removed.includes("selectedIndex: 0") && !removed.includes("id: child"));
    assertActiveTextEditor(document.uri);
    await vscode.commands.executeCommand("undo");
    assert.strictEqual(
      document.getText(),
      initial,
      "one Undo restores Page, descendants and selectedIndex",
    );
    await vscode.commands.executeCommand("redo");
    assert.strictEqual(document.getText(), removed);
  });
}

async function assertStructuralTransactionsUndoAndRedo(): Promise<void> {
  const source = [
    "schemaVersion: 1",
    "kind: xlflow.userform",
    "form: {name: Main, width: 480, height: 360}",
    "controls:",
    "  - {id: frame, name: Frame1, type: Frame}",
    '  - {id: other, name: Label1, type: Label, caption: "😀 keep"}',
    "  - {id: child, parentId: frame, name: TextBox1, type: TextBox}",
    "",
  ].join("\r\n");
  await withDesignerDocument("structural-undo", source, async (document) => {
    const queue = new DocumentEditQueue();
    const initial = document.getText();
    const addition =
      "  - {id: control-new, name: Label2, type: Label, left: 10, top: 20, width: 72, height: 18}\r\n";
    const client = fakeClient((doc, version, _text, operations) => {
      if (operations[0].type === "addControl") {
        const position = doc.positionAt(doc.getText().length);
        return {
          version,
          edits: [{ range: { start: position, end: position }, newText: addition }],
        };
      }
      const edits: SourceTextEdit[] = [];
      for (let i = 0; i < doc.lineCount; i++) {
        const line = doc.lineAt(i);
        if (/id: (frame|child),/.test(line.text))
          edits.push({ range: line.rangeIncludingLineBreak, newText: "" });
      }
      return { version, edits };
    });
    assert.strictEqual(
      await queue.run(
        document,
        document.version,
        [
          {
            type: "addControl",
            control: {
              id: "control-new",
              name: "Label2",
              type: "Label",
              left: 10,
              top: 20,
              width: 72,
              height: 18,
            },
          },
        ],
        client,
        () => true,
      ),
      undefined,
    );
    const added = initial + addition;
    assert.strictEqual(document.getText(), added);
    assert.strictEqual(
      await queue.run(
        document,
        document.version,
        [{ type: "removeControl", controlId: "frame", cascade: true }],
        client,
        () => true,
      ),
      undefined,
    );
    const removed = document.getText();
    assert.ok(!removed.includes("id: frame") && !removed.includes("id: child"));
    assert.ok(removed.includes('caption: "😀 keep"'));
    assertActiveTextEditor(document.uri);
    await vscode.commands.executeCommand("undo");
    assert.strictEqual(
      document.getText(),
      added,
      "one undo restores the entire noncontiguous subtree",
    );
    await vscode.commands.executeCommand("undo");
    assert.strictEqual(document.getText(), initial, "the preceding add is its own transaction");
    await vscode.commands.executeCommand("redo");
    assert.strictEqual(document.getText(), added);
    await vscode.commands.executeCommand("redo");
    assert.strictEqual(document.getText(), removed);
  });
}

async function assertParentTransactionUndoAndRedo(): Promise<void> {
  const source = [
    "schemaVersion: 1",
    "kind: xlflow.userform",
    "form: {name: Main, width: 480, height: 360}",
    "controls:",
    "  - {id: frame, name: Frame1, type: Frame}",
    "  - {id: child, parentId: frame, name: TextBox1, type: TextBox}",
    "",
  ].join("\r\n");
  await withDesignerDocument("parent-undo", source, async (document) => {
    const queue = new DocumentEditQueue();
    const initial = document.getText();
    const oldLine = "  - {id: child, parentId: frame, name: TextBox1, type: TextBox}";
    const newLine = "  - {id: child, name: TextBox1, type: TextBox}";
    const client = fakeClient((doc, version, text, operations) => {
      assert.strictEqual(operations.length, 1);
      assert.strictEqual(operations[0].type, "setParent");
      const startOffset = text.indexOf(oldLine);
      assert.ok(startOffset >= 0);
      return {
        version,
        edits: [
          {
            range: {
              start: doc.positionAt(startOffset),
              end: doc.positionAt(startOffset + oldLine.length),
            },
            newText: newLine,
          },
        ],
      };
    });
    const result = await queue.run(
      document,
      document.version,
      [{ type: "setParent", controlId: "child", parentId: "", left: 24, top: 18 }],
      client,
      () => true,
    );
    assert.strictEqual(result, undefined);
    const reparented = document.getText();
    assert.ok(!reparented.includes("parentId: frame"));
    await vscode.commands.executeCommand("undo");
    assert.strictEqual(document.getText(), initial, "one undo restores the previous parent");
    await vscode.commands.executeCommand("redo");
    assert.strictEqual(document.getText(), reparented);
  });
}

async function assertConfirmationRejectsStaleAndCancelledEdits(): Promise<void> {
  await withDesignerDocument("confirmation-stale", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    let requests = 0;
    const client = fakeClient((_document, version) => {
      requests++;
      return { version, edits: [] };
    });
    const operations: SemanticOperation[] = [
      { type: "removeControl", controlId: "frame", cascade: true },
    ];
    const initial = document.getText();
    const cancelled = await queue.run(
      document,
      document.version,
      operations,
      client,
      () => true,
      async () => ({ code: "editCancelled", message: "Cancelled" }),
    );
    assert.strictEqual(cancelled?.code, "editCancelled");
    assert.strictEqual(requests, 0);
    assert.strictEqual(document.getText(), initial);

    const started = deferred<void>();
    const confirmation = deferred<undefined>();
    const pending = queue.run(
      document,
      document.version,
      operations,
      client,
      () => true,
      () => {
        started.resolve();
        return confirmation.promise;
      },
    );
    await withTimeout(started.promise, "confirmation did not open");
    const edit = new vscode.WorkspaceEdit();
    edit.insert(document.uri, new vscode.Position(0, 0), "# external edit\r\n");
    assert.ok(await vscode.workspace.applyEdit(edit));
    const externallyEdited = document.getText();
    confirmation.resolve(undefined);
    assert.strictEqual((await pending)?.code, "staleDocument");
    assert.strictEqual(requests, 0, "an accepted old confirmation cannot reach the LSP");
    assert.strictEqual(document.getText(), externallyEdited);

    const connectionStarted = deferred<void>();
    const connectionConfirmation = deferred<undefined>();
    let current = true;
    const connectionPending = queue.run(
      document,
      document.version,
      operations,
      client,
      () => current,
      () => {
        connectionStarted.resolve();
        return connectionConfirmation.promise;
      },
    );
    await withTimeout(connectionStarted.promise, "connection confirmation did not open");
    current = false;
    connectionConfirmation.resolve(undefined);
    assert.strictEqual((await connectionPending)?.code, "staleDocument");
    assert.strictEqual(requests, 0);
  });
}

async function assertGeometryTransactionsUndoAndRedo(): Promise<void> {
  await withDesignerDocument("undo-redo", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const initial = document.getText();
    const northWestResize: GeometryOperation[] = [
      { type: "moveControl", controlId: "😀Box", left: 15, top: 25 },
      { type: "resizeControl", controlId: "😀Box", width: 25, height: 35 },
    ];
    assert.strictEqual(
      await queue.run(document, document.version, northWestResize, geometryClient(), () => true),
      undefined,
    );
    const afterNorthWestResize = replaceGeometry(initial, {
      left: 15,
      top: 25,
      width: 25,
      height: 35,
    });
    assert.strictEqual(document.getText(), afterNorthWestResize);

    const secondTransaction: GeometryOperation[] = [
      { type: "moveControl", controlId: "😀Box", left: 17, top: 27 },
    ];
    assert.strictEqual(
      await queue.run(document, document.version, secondTransaction, geometryClient(), () => true),
      undefined,
    );
    const afterSecondTransaction = replaceGeometry(afterNorthWestResize, { left: 17, top: 27 });
    assert.strictEqual(document.getText(), afterSecondTransaction);

    assert.strictEqual(
      await queue.run(
        document,
        document.version,
        secondTransaction,
        fakeClient((_doc, version) => ({ version, edits: [] })),
        () => true,
      ),
      undefined,
      "an empty edit response is a successful no-op",
    );
    assert.strictEqual(document.getText(), afterSecondTransaction);
    assertActiveTextEditor(document.uri);

    await vscode.commands.executeCommand("undo");
    assert.strictEqual(
      document.getText(),
      afterNorthWestResize,
      "one undo reverts only transaction two",
    );
    await vscode.commands.executeCommand("undo");
    assert.strictEqual(
      document.getText(),
      initial,
      "the NW move+resize remains one undo transaction",
    );
    await vscode.commands.executeCommand("redo");
    assert.strictEqual(document.getText(), afterNorthWestResize);
    await vscode.commands.executeCommand("redo");
    assert.strictEqual(document.getText(), afterSecondTransaction);
  });
}

async function assertPropertyTransactionsUndoAndRedo(): Promise<void> {
  await withDesignerDocument("property-undo-redo", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const initial = document.getText();
    const client = fakeClient((_doc, version, text, operations) => {
      const op = operations[0];
      assert.strictEqual(op.type, "setFormProperty");
      const old = '"😀 Café"';
      const start = text.indexOf(old);
      assert.ok(start >= 0);
      const from = document.positionAt(start);
      const to = document.positionAt(start + old.length);
      return {
        version,
        edits: [{ range: { start: from, end: to }, newText: JSON.stringify("編集済み 😀") }],
      };
    });
    const op: SemanticOperation = {
      type: "setFormProperty",
      field: "caption",
      value: "編集済み 😀",
    };
    assert.strictEqual(
      await queue.run(document, document.version, [op], client, () => true),
      undefined,
    );
    const edited = initial.replace('"😀 Café"', '"編集済み 😀"');
    assert.strictEqual(document.getText(), edited);
    await vscode.commands.executeCommand("undo");
    assert.strictEqual(document.getText(), initial, "property edit is one native undo step");
    await vscode.commands.executeCommand("redo");
    assert.strictEqual(document.getText(), edited);
    const error = {
      code: "invalid",
      message: "Invalid property",
      diagnostics: [
        { operationIndex: 0, field: "caption", code: "UFV001", message: "Invalid caption" },
      ],
    };
    assert.deepStrictEqual(
      await queue.run(
        document,
        document.version,
        [op],
        fakeClient((_, version) => ({ version, edits: [], error })),
        () => true,
      ),
      error,
    );
    assert.strictEqual(document.getText(), edited, "field diagnostics do not mutate source");
  });
}

async function assertStaleResponsesAreDiscarded(): Promise<void> {
  await withDesignerDocument("stale-response", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const original = document.getText();
    const version = document.version;
    const response = deferred<EditResult>();
    const started = deferred<void>();
    const client = fakeClient(() => {
      started.resolve();
      return response.promise;
    });
    const pending = queue.run(
      document,
      version,
      [{ type: "moveControl", controlId: "😀Box", left: 11, top: 21 }],
      client,
      () => true,
    );
    await withTimeout(started.promise, "edit request did not start");
    response.resolve({
      version: version + 1,
      edits: geometryEdits(original, [
        { type: "moveControl", controlId: "😀Box", left: 11, top: 21 },
      ]),
    });
    assert.strictEqual((await pending)?.code, "staleDocument");
    assert.strictEqual(document.getText(), original, "a stale response must not edit the source");
  });
}

async function assertConnectionInvalidationIsDiscarded(): Promise<void> {
  await withDesignerDocument("connection-change", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const original = document.getText();
    const version = document.version;
    const response = deferred<EditResult>();
    const started = deferred<void>();
    let connectionIsCurrent = true;
    const client = fakeClient(() => {
      started.resolve();
      return response.promise;
    });
    const pending = queue.run(
      document,
      version,
      [{ type: "moveControl", controlId: "😀Box", left: 11, top: 21 }],
      client,
      () => connectionIsCurrent,
    );
    await withTimeout(started.promise, "edit request did not start");
    connectionIsCurrent = false;
    response.resolve({
      version,
      edits: geometryEdits(original, [
        { type: "moveControl", controlId: "😀Box", left: 11, top: 21 },
      ]),
    });
    assert.strictEqual((await pending)?.code, "staleDocument");
    assert.strictEqual(
      document.getText(),
      original,
      "a retired connection must not edit the source",
    );
  });
}

async function assertConcurrentPanelsAreSerialized(): Promise<void> {
  await withDesignerDocument("concurrent-panels", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const original = document.getText();
    const version = document.version;
    const response = deferred<EditResult>();
    const started = deferred<void>();
    let requests = 0;
    const client = fakeClient((_doc, requestVersion) => {
      requests++;
      if (requests === 1) {
        started.resolve();
        return response.promise;
      }
      return { version: requestVersion, edits: [] };
    });
    const firstOperation: GeometryOperation[] = [
      { type: "moveControl", controlId: "😀Box", left: 11, top: 21 },
    ];
    const first = queue.run(document, version, firstOperation, client, () => true);
    await withTimeout(started.promise, "first panel edit request did not start");
    const second = queue.run(
      document,
      version,
      [{ type: "moveControl", controlId: "😀Box", left: 12, top: 22 }],
      client,
      () => true,
    );
    assert.strictEqual(requests, 1, "the second panel must wait while the URI is being edited");

    response.resolve({ version, edits: geometryEdits(original, firstOperation) });
    assert.strictEqual(await first, undefined);
    assert.strictEqual((await second)?.code, "staleDocument");
    assert.strictEqual(
      requests,
      1,
      "the queued panel must reject after the first edit changes the version",
    );
    assert.strictEqual(document.getText(), replaceGeometry(original, { left: 11, top: 21 }));
  });
}

async function assertUnicodeAndCRLFRanges(): Promise<void> {
  await withDesignerDocument("unicode-crlf", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const original = document.getText();
    const operations: GeometryOperation[] = [
      { type: "moveControl", controlId: "😀Box", left: 11, top: 21 },
    ];
    const error = await queue.run(
      document,
      document.version,
      operations,
      geometryClient(),
      () => true,
    );
    assert.strictEqual(error, undefined);
    assert.strictEqual(document.getText(), replaceGeometry(original, { left: 11, top: 21 }));
    assert.ok(
      document.getText().includes('id: "😀Box"'),
      "UTF-16 positions after a surrogate pair stay aligned",
    );
    assert.ok(
      document.getText().includes('caption: "😀 Café"'),
      "unrelated Unicode text is preserved",
    );
    assert.ok(
      document.getText().includes("\r\n"),
      "the workspace edit preserves CRLF line endings",
    );
    assert.ok(
      !/(^|[^\r])\n/.test(document.getText()),
      "no LF-only line endings should be introduced",
    );
  });
}

async function assertInvalidRangesAreRejectedAtomically(): Promise<void> {
  await withDesignerDocument("invalid-ranges", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const original = document.getText();
    const line = document.lineAt(4);
    const validEdit = scalarEdit(original, "left", 99);
    const outOfBounds: SourceTextEdit = {
      range: {
        start: { line: 4, character: line.text.length + 1 },
        end: { line: 4, character: line.text.length + 1 },
      },
      newText: "!",
    };
    const outOfBoundsError = await queue.run(
      document,
      document.version,
      [{ type: "moveControl", controlId: "😀Box", left: 99, top: 20 }],
      fakeClient((_doc, version) => ({ version, edits: [validEdit, outOfBounds] })),
      () => true,
    );
    assert.strictEqual(outOfBoundsError?.code, "invalidResponse");
    assert.strictEqual(document.getText(), original, "one bad range rejects the entire edit batch");

    const emojiOffset = line.text.indexOf("😀");
    assert.ok(emojiOffset >= 0);
    const splitSurrogate: SourceTextEdit = {
      range: {
        start: { line: 4, character: emojiOffset + 1 },
        end: { line: 4, character: emojiOffset + 1 },
      },
      newText: "!",
    };
    const surrogateError = await queue.run(
      document,
      document.version,
      [{ type: "moveControl", controlId: "😀Box", left: 99, top: 20 }],
      fakeClient((_doc, version) => ({ version, edits: [splitSurrogate] })),
      () => true,
    );
    assert.strictEqual(surrogateError?.code, "invalidResponse");
    assert.strictEqual(
      document.getText(),
      original,
      "a range splitting a surrogate pair is rejected",
    );
  });
}

async function assertStructuredRejectionPreservesSource(): Promise<void> {
  await withDesignerDocument("structured-rejection", formSource, async (document) => {
    const queue = new DocumentEditQueue();
    const original = document.getText();
    const rejection = { code: "editNotSupported", message: "This UserForm cannot be edited." };
    const error = await queue.run(
      document,
      document.version,
      [{ type: "moveControl", controlId: "😀Box", left: 99, top: 99 }],
      fakeClient((_doc, version) => ({
        version,
        error: rejection,
        edits: [scalarEdit(original, "left", 99)],
      })),
      () => true,
    );
    assert.deepStrictEqual(
      error,
      rejection,
      "the server's structured rejection is propagated verbatim",
    );
    assert.strictEqual(
      document.getText(),
      original,
      "rejected edits must preserve the exact document text",
    );
  });
}

async function withDesignerDocument(
  name: string,
  source: string,
  assertions: (document: vscode.TextDocument) => Promise<void>,
  extension = "yaml",
): Promise<void> {
  const directory = await mkdtemp(path.join(os.tmpdir(), `xlflow-userform-edits-${name}-`));
  const file = path.join(directory, `Main.${extension}`);
  const uri = vscode.Uri.file(file);
  try {
    await writeFile(file, source, "utf8");
    const document = await vscode.workspace.openTextDocument(uri);
    await vscode.commands.executeCommand("vscode.openWith", uri, "default");
    assert.ok(
      vscode.window.activeTextEditor?.document.uri.toString() === uri.toString(),
      "the isolated edit queue must use a regular text editor",
    );
    await assertions(document);
  } finally {
    try {
      for (const group of vscode.window.tabGroups.all) {
        for (const tab of group.tabs) {
          const input = tab.input;
          if (
            (input instanceof vscode.TabInputText || input instanceof vscode.TabInputCustom) &&
            input.uri.toString() === uri.toString()
          ) {
            await vscode.window.tabGroups.close(tab, true);
          }
        }
      }
    } finally {
      await rm(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
    }
  }
}

function geometryClient(): EditClient {
  return fakeClient((_, version, text, operations) => ({
    version,
    edits: geometryEdits(
      text,
      operations.filter(
        (op): op is GeometryOperation => op.type === "moveControl" || op.type === "resizeControl",
      ),
    ),
  }));
}

function assertActiveTextEditor(uri: vscode.Uri): void {
  assert.ok(
    vscode.window.activeTextEditor?.document.uri.toString() === uri.toString(),
    "undo and redo must run with the source text editor active",
  );
}

function geometryEdits(text: string, operations: GeometryOperation[]): SourceTextEdit[] {
  const values = new Map<string, number>();
  for (const operation of operations) {
    if (operation.type === "moveControl") {
      values.set("left", operation.left);
      values.set("top", operation.top);
    } else {
      values.set("width", operation.width);
      values.set("height", operation.height);
    }
  }
  return [...values]
    .map(([field, value]) => scalarEdit(text, field, value))
    .sort((a, b) => comparePositions(a.range.start, b.range.start));
}

function scalarEdit(text: string, field: string, value: number): SourceTextEdit {
  const lines = text.split(/\r\n|\n/);
  const expression = new RegExp(`\\b${field}:\\s*(-?\\d+(?:\\.\\d+)?)`);
  for (let line = 0; line < lines.length; line++) {
    const match = expression.exec(lines[line]);
    if (!match || match.index === undefined) continue;
    const start = match.index + match[0].lastIndexOf(match[1]);
    return {
      range: {
        start: { line, character: start },
        end: { line, character: start + match[1].length },
      },
      newText: String(value),
    };
  }
  throw new Error(`Fixture does not contain a ${field} scalar`);
}

function replaceGeometry(
  text: string,
  values: Partial<Record<"left" | "top" | "width" | "height", number>>,
): string {
  let result = text;
  for (const [field, value] of Object.entries(values)) {
    result = result.replace(
      new RegExp(`(\\b${field}:\\s*)-?\\d+(?:\\.\\d+)?`),
      (_match, prefix: string) => `${prefix}${value}`,
    );
  }
  return result;
}

function fakeClient(
  respond: (
    document: vscode.TextDocument,
    version: number,
    text: string,
    operations: SemanticOperation[],
  ) => EditResult | Promise<EditResult>,
): EditClient {
  return {
    requestUserFormEdit: async (document, version, text, operations) =>
      respond(document, version, text, operations),
  };
}

function comparePositions(
  a: { line: number; character: number },
  b: { line: number; character: number },
): number {
  return a.line - b.line || a.character - b.character;
}

function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((complete) => {
    resolve = complete;
  });
  return { promise, resolve };
}

async function withTimeout<T>(promise: Promise<T>, message: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      promise,
      new Promise<T>((_, reject) => {
        timer = setTimeout(() => reject(new Error(message)), 10_000);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}
