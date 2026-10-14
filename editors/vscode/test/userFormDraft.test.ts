import * as assert from "node:assert/strict";
import {
  DraftState,
  applySourceTextEdits,
  computeSourceChanges,
  type SourceTextChange,
} from "../src/userFormEditor/draft";
import type { SourceTextEdit } from "../src/userFormEditor/protocol";

function edit(
  startLine: number,
  startCharacter: number,
  endLine: number,
  endCharacter: number,
  newText: string,
): SourceTextEdit {
  return {
    range: {
      start: { line: startLine, character: startCharacter },
      end: { line: endLine, character: endCharacter },
    },
    newText,
  };
}

function applyChanges(text: string, changes: readonly SourceTextChange[]): string {
  let result = text;
  for (const change of [...changes].sort(
    (left, right) => right.start - left.start || right.end - left.end,
  ))
    result = result.slice(0, change.start) + change.newText + result.slice(change.end);
  return result;
}

function testSourceEditsAndDiff() {
  const source = "first 😀\r\nsecond # retain\r\n";
  const updated = applySourceTextEdits(source, [
    edit(1, 0, 1, 6, "renamed"),
    edit(0, 6, 0, 8, "🧪"),
  ]);
  assert.equal(updated, "first 🧪\r\nrenamed # retain\r\n");
  assert.throws(
    () => applySourceTextEdits(source, [edit(1, 0, 1, 4, "x"), edit(1, 3, 1, 6, "y")]),
    /overlap/,
  );
  assert.throws(
    () => applySourceTextEdits(source, [edit(3, 0, 3, 0, "x")]),
    /invalid UTF-16 position/,
  );

  const base = "# preserve comment\r\nform:\r\n  name: First\r\n  width: 400\r\ncontrols: []\r\n";
  const current = base.replace("First", "Second").replace("400", "480");
  const changes = computeSourceChanges(base, current);
  assert.ok(changes.length >= 2, "separated edits remain separate source changes");
  for (let index = 1; index < changes.length; index++)
    assert.ok(changes[index - 1].start <= changes[index].start, "source changes are ordered");
  assert.equal(applyChanges(base, changes), current);
  assert.equal(computeSourceChanges(base, base).length, 0);
}

function testYamlExternalMergeAndUndo() {
  const base =
    "# preserve comment\r\nform:\r\n  name: Main\r\n  width: 400\r\ncontrols:\r\n  - id: button\r\n    name: Button1\r\n    type: CommandButton\r\n    left: 10\r\n    top: 20\r\n";
  const state = new DraftState(base);
  const frame = state.commit(applySourceTextEdits(base, [edit(3, 9, 3, 12, "480")]), "Resize form");
  assert.ok(frame);
  const source = applySourceTextEdits(base, [edit(9, 9, 9, 11, "31")]);
  assert.equal(state.synchronize(source), "merged");

  const merged = source.replace("width: 400", "width: 480");
  assert.equal(state.text, merged);
  assert.equal(frame.before, source, "the existing callback sees its rebased before snapshot");
  assert.equal(frame.after, merged, "the existing callback sees its rebased after snapshot");
  assert.ok(merged.includes("# preserve comment\r\n"));
  assert.ok(merged.endsWith("top: 31\r\n"));

  state.undo(frame);
  assert.equal(state.text, source, "Undo preserves the external source edit");
  state.redo(frame);
  assert.equal(state.text, merged);
}

function testSeparatedYamlEditsPreserveExternalMiddleCaption() {
  const base = "form:\n  width: 400\n  caption: Main\n  height: 300\n";
  const draftText = base.replace("width: 400", "width: 480").replace("height: 300", "height: 320");
  const state = new DraftState(base);
  const frame = state.commit(draftText)!;
  const external = base.replace("caption: Main", "caption: Updated externally");

  assert.equal(state.synchronize(external), "merged");
  const merged = external.replace("width: 400", "width: 480").replace("height: 300", "height: 320");
  assert.equal(state.text, merged);
  assert.equal(frame.before, external);
  assert.equal(frame.after, merged);
  state.undo(frame);
  assert.equal(
    state.text,
    external,
    "undo keeps the external caption between separated draft edits",
  );
}

function testPrettyJsonDifferentPropertiesMergeAndSameLineScalarConflicts() {
  const base = '{\n  "caption": "abcd",\n  "width": 100\n}\n';
  const state = new DraftState(base);
  const frame = state.commit(base.replace('"abcd"', '"draft caption"'))!;
  const external = base.replace('"width": 100', '"width": 120');
  assert.equal(state.synchronize(external), "merged");
  assert.equal(state.text, external.replace('"abcd"', '"draft caption"'));
  state.undo(frame);
  assert.equal(state.text, external, "Undo preserves a distinct external JSON property edit");

  const scalarBase = 'form: {caption: "abcd"}\n';
  const scalar = new DraftState(scalarBase);
  scalar.commit(scalarBase.replace('"abcd"', '"Abcd"'));
  assert.equal(
    scalar.synchronize(scalarBase.replace('"abcd"', '"abcD"')),
    "conflict",
    "edits to separate UTF-16 positions in one scalar must not synthesize a third value",
  );
  assert.equal(scalar.text, scalarBase.replace('"abcd"', '"Abcd"'));
  const insertions = new DraftState(scalarBase);
  insertions.commit(scalarBase.replace('"abcd"', '"abXcd"'));
  assert.equal(
    insertions.synchronize(scalarBase.replace('"abcd"', '"abcYd"')),
    "conflict",
    "different insertion positions within one scalar still conflict",
  );
}

function testUnsafeHistoryIsTrimmedWithoutConflictingCurrentDraft() {
  const base = "form:\n  caption: Main\n  width: 400\n";
  const state = new DraftState(base);
  const captionFrame = state.commit(base.replace("caption: Main", "caption: Saved"))!;
  state.markSaved();
  const savedBase = state.baseText;
  const widthFrame = state.commit(savedBase.replace("width: 400", "width: 480"))!;
  const external = savedBase.replace("caption: Saved", "caption: External");

  assert.equal(state.synchronize(external), "merged");
  const merged = external.replace("width: 400", "width: 480");
  assert.equal(
    state.conflicted,
    false,
    "a historical caption collision does not conflict the pending width edit",
  );
  assert.equal(state.text, merged);
  assert.deepEqual(
    state.appliedFrames,
    [widthFrame],
    "only the safe contiguous history beside the cursor remains",
  );
  assert.throws(() => state.undo(captionFrame), /current undo operation/);

  state.undo(widthFrame);
  assert.equal(state.text, external, "the safe Undo preserves the external caption edit");
  assert.throws(() => state.redo(captionFrame), /current redo operation/);
}

function testCleanDraftExternalEditTrimsUnsafeRedoWithoutConflict() {
  const base = "form:\n  caption: Main\n  width: 400\n";
  const state = new DraftState(base);
  const frame = state.commit(base.replace("caption: Main", "caption: Designer"))!;
  state.undo(frame);
  assert.equal(state.dirty, false);

  const external = base.replace("caption: Main", "caption: External");
  assert.equal(state.synchronize(external), "merged");
  assert.equal(
    state.conflicted,
    false,
    "an undone draft edit cannot conflict the clean source view",
  );
  assert.equal(state.text, external);
  assert.equal(state.dirty, false);
  assert.equal(state.appliedFrames.length, 0);
  assert.throws(() => state.redo(frame), /current redo operation/);
}

function testJsonSamePropertyAndStructuralConflicts() {
  const base = '{\n  "form": {"name": "Main", "width": 400},\n  "controls": []\n}\n';
  const draft = new DraftState(base);
  const widthFrame = draft.commit(base.replace('"width": 400', '"width": 500'))!;
  const external = base.replace('"width": 400', '"width": 600');
  assert.equal(draft.synchronize(external), "conflict");
  assert.equal(draft.conflicted, true);
  assert.equal(
    draft.text,
    widthFrame.after,
    "conflict markers are never inserted into source text",
  );
  assert.throws(() => draft.undo(widthFrame), /conflict/);
  const restoredConflict = DraftState.restore(draft.serialize());
  assert.equal(restoredConflict.conflicted, true);
  assert.equal(restoredConflict.text, draft.text);
  assert.equal(restoredConflict.serialize().pendingSourceText, external);

  const resolution = draft.resolveConflict("designer");
  assert.equal(resolution.before, widthFrame.after);
  assert.equal(resolution.after, widthFrame.after);
  assert.equal(draft.baseText, external);
  assert.equal(
    draft.dirty,
    true,
    "designer resolution keeps the selected draft dirty against latest source",
  );
  assert.throws(() => draft.undo(widthFrame), /current undo operation/);
  assert.equal(
    draft.text,
    widthFrame.after,
    "native Undo must follow the resolution frame before older draft edits",
  );
  draft.undo(resolution);
  assert.equal(draft.text, widthFrame.after);
  draft.undo(widthFrame);
  assert.equal(
    draft.text,
    external,
    "Undo after resolving preserves the latest external source baseline",
  );
  draft.redo(widthFrame);
  draft.redo(resolution);
  assert.equal(draft.text, widthFrame.after);

  const sourceChoice = new DraftState(base);
  const discarded = sourceChoice.commit(base.replace('"width": 400', '"width": 500'))!;
  assert.equal(sourceChoice.synchronize(external), "conflict");
  const useSource = sourceChoice.resolveConflict("source");
  assert.equal(sourceChoice.text, external);
  assert.throws(() => sourceChoice.undo(discarded), /current undo operation/);
  sourceChoice.undo(useSource);
  assert.equal(
    sourceChoice.text,
    discarded.after,
    "the new resolution frame owns its explicit undo",
  );
  sourceChoice.undo(discarded);
  assert.equal(
    sourceChoice.text,
    external,
    "Undo the earlier draft instead of leaving pending work behind an old native checkpoint",
  );
  sourceChoice.redo(discarded);
  sourceChoice.redo(useSource);
  assert.equal(sourceChoice.text, external);

  const structureBase = '{"controls":[]}\n';
  const structure = new DraftState(structureBase);
  structure.commit('{"controls":[{"id":"draft","type":"Label"}]}\n');
  assert.equal(
    structure.synchronize('{"controls":[{"id":"source","type":"Label"}]}\n'),
    "conflict",
    "different structural insertions at the same source location conflict",
  );
  assert.equal(structure.text, '{"controls":[{"id":"draft","type":"Label"}]}\n');
}

function testSaveCheckpointPreservesUndoAndSourceWatcherNoop() {
  const base =
    '{\n  "form": {\n    "name": "Before",\n    "height": 300\n  },\n  "controls": []\n}\n';
  const state = new DraftState(base);
  const frame = state.commit(base.replace('"name": "Before"', '"name": "After"'))!;
  const savedText = state.text;
  state.markSaved();
  assert.equal(state.baseText, savedText);
  assert.equal(frame.before, base);
  assert.equal(frame.after, savedText);

  const revision = state.revision;
  assert.equal(state.synchronize(savedText), "merged");
  assert.equal(
    state.revision,
    revision,
    "the save-applied source event is a no-op for draft state",
  );
  assert.equal(frame.before, base);
  assert.equal(frame.after, savedText);

  const external = savedText.replace('"height": 300', '"height": 320');
  assert.equal(state.synchronize(external), "merged");
  const rebasedBefore = base.replace('"height": 300', '"height": 320');
  const rebasedAfter = savedText.replace('"height": 300', '"height": 320');
  assert.equal(frame.before, rebasedBefore);
  assert.equal(frame.after, rebasedAfter);
  state.undo(frame);
  assert.equal(state.text, rebasedBefore, "Undo after save keeps the later external source edit");
  state.redo(frame);
  assert.equal(state.text, rebasedAfter);

  const restored = DraftState.restore(state.serialize());
  assert.equal(restored.text, state.text);
  assert.equal(restored.baseText, state.baseText);
  assert.equal(restored.dirty, state.dirty);
  assert.equal(restored.appliedFrames.length, 1, "backup restores the applied native-undo prefix");
  const restoredFrame = restored.appliedFrames[0];
  restored.undo(restoredFrame);
  assert.equal(restored.text, rebasedBefore);
}

testSourceEditsAndDiff();
testYamlExternalMergeAndUndo();
testSeparatedYamlEditsPreserveExternalMiddleCaption();
testPrettyJsonDifferentPropertiesMergeAndSameLineScalarConflicts();
testUnsafeHistoryIsTrimmedWithoutConflictingCurrentDraft();
testCleanDraftExternalEditTrimsUnsafeRedoWithoutConflict();
testJsonSamePropertyAndStructuralConflicts();
testSaveCheckpointPreservesUndoAndSourceWatcherNoop();
process.stdout.write("UserForm draft state and source edit assertions passed\n");
