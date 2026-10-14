import type { SourceTextEdit } from "./protocol";

/** A source change whose offsets are UTF-16 code-unit offsets into the base text. */
export interface SourceTextChange {
  readonly start: number;
  readonly end: number;
  readonly newText: string;
}

/** A stable handle for a semantic draft edit. Its snapshots can be rebased in place. */
export interface DraftFrame {
  readonly before: string;
  readonly after: string;
  readonly label?: string;
}

export interface SerializedDraftState {
  readonly schemaVersion: 1;
  readonly baseText: string;
  readonly text: string;
  readonly revision: number;
  readonly history: readonly {
    readonly before: string;
    readonly after: string;
    readonly label?: string;
  }[];
  readonly cursor: number;
  readonly conflicted: boolean;
  readonly pendingSourceText?: string;
}

interface FrameSnapshot {
  before: string;
  after: string;
}

const frameSnapshots = new WeakMap<DraftFrame, FrameSnapshot>();

function createFrame(before: string, after: string, label?: string): DraftFrame {
  const snapshot: FrameSnapshot = { before, after };
  let frame: DraftFrame;
  frame = {
    get before() {
      return snapshot.before;
    },
    get after() {
      return snapshot.after;
    },
    ...(label === undefined ? {} : { label }),
  };
  frameSnapshots.set(frame, snapshot);
  return Object.freeze(frame);
}

function setFrameSnapshot(frame: DraftFrame, before: string, after: string): void {
  const snapshot = frameSnapshots.get(frame);
  if (!snapshot) throw new Error("The draft frame is no longer valid.");
  snapshot.before = before;
  snapshot.after = after;
}

function assertString(value: unknown, name: string): asserts value is string {
  if (typeof value !== "string") throw new TypeError(`${name} must be a string.`);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

interface SourceLineIndex {
  readonly starts: readonly number[];
  readonly ends: readonly number[];
}

function indexLines(text: string): SourceLineIndex {
  const starts = [0];
  const ends: number[] = [];
  for (let index = 0; index < text.length; index++) {
    const character = text.charAt(index);
    if (character !== "\r" && character !== "\n") continue;
    ends.push(index);
    if (character === "\r" && text.charAt(index + 1) === "\n") index++;
    starts.push(index + 1);
  }
  ends.push(text.length);
  return { starts, ends };
}

function positionOffset(
  position: { readonly line: number; readonly character: number },
  lines: SourceLineIndex,
): number {
  if (
    !Number.isSafeInteger(position.line) ||
    !Number.isSafeInteger(position.character) ||
    position.line < 0 ||
    position.character < 0 ||
    position.line >= lines.starts.length
  )
    throw new RangeError("A source edit contains an invalid UTF-16 position.");
  const start = lines.starts[position.line];
  const end = lines.ends[position.line];
  if (position.character > end - start)
    throw new RangeError("A source edit position extends past the end of its line.");
  return start + position.character;
}

function orderedAndValidatedChanges(
  text: string,
  edits: readonly SourceTextChange[],
): SourceTextChange[] {
  const ordered = edits.map((edit) => {
    if (
      !Number.isSafeInteger(edit.start) ||
      !Number.isSafeInteger(edit.end) ||
      edit.start < 0 ||
      edit.start > edit.end ||
      edit.end > text.length ||
      typeof edit.newText !== "string"
    )
      throw new RangeError("A source edit has invalid UTF-16 offsets or replacement text.");
    return { start: edit.start, end: edit.end, newText: edit.newText };
  });
  ordered.sort((left, right) => left.start - right.start || left.end - right.end);

  for (let index = 1; index < ordered.length; index++) {
    const previous = ordered[index - 1];
    const current = ordered[index];
    const previousInsertion = previous.start === previous.end;
    const currentInsertion = current.start === current.end;
    if (previousInsertion && currentInsertion && previous.start === current.start)
      throw new RangeError("Multiple source insertions at the same position are ambiguous.");
    if (!previousInsertion && !currentInsertion && current.start < previous.end)
      throw new RangeError("Source edits overlap.");
    if (previousInsertion && current.start < previous.start && previous.start < current.end)
      throw new RangeError("A source insertion overlaps a replacement.");
    if (currentInsertion && previous.start < current.start && current.start < previous.end)
      throw new RangeError("A source insertion overlaps a replacement.");
  }
  return ordered;
}

function applyOffsetChanges(text: string, edits: readonly SourceTextChange[]): string {
  const ordered = orderedAndValidatedChanges(text, edits);
  let result = text;
  for (let index = ordered.length - 1; index >= 0; index--) {
    const edit = ordered[index];
    result = result.slice(0, edit.start) + edit.newText + result.slice(edit.end);
  }
  return result;
}

/**
 * Applies LSP TextEdits using VS Code's zero-based UTF-16 line/character
 * positions. All ranges are validated before any change is applied.
 */
export function applySourceTextEdits(text: string, edits: readonly SourceTextEdit[]): string {
  assertString(text, "text");
  if (!Array.isArray(edits)) throw new TypeError("edits must be an array.");
  const lines = indexLines(text);
  const changes = edits.map((edit) => {
    if (
      !isRecord(edit) ||
      !isRecord(edit.range) ||
      !isRecord(edit.range.start) ||
      !isRecord(edit.range.end)
    )
      throw new TypeError("A source edit must contain a start and end position.");
    if (typeof edit.newText !== "string")
      throw new TypeError("A source edit replacement must be a string.");
    const start = positionOffset(edit.range.start as SourceTextEdit["range"]["start"], lines);
    const end = positionOffset(edit.range.end as SourceTextEdit["range"]["end"], lines);
    if (start > end) throw new RangeError("A source edit range is reversed.");
    return { start, end, newText: edit.newText };
  });
  return applyOffsetChanges(text, changes);
}

const MAX_DIFF_INPUT = 1_000_000;
const MAX_DIFF_DISTANCE = 1_024;
const MAX_TRACE_CELLS = 1_100_000;
const MAX_DIFF_WORK = 3_000_000;

type DiffSegment =
  | { readonly kind: "equal"; readonly count: number }
  | { readonly kind: "delete" }
  | { readonly kind: "insert"; readonly value: string };

function myersChanges(
  before: string,
  after: string,
  absoluteStart: number,
): SourceTextChange[] | undefined {
  const n = before.length;
  const m = after.length;
  const max = n + m;
  if (max === 0) return [];

  const offset = max + 1;
  const frontier = new Int32Array(max * 2 + 3);
  frontier.fill(-1);
  frontier[offset + 1] = 0;
  const trace: Int32Array[] = [];
  let traceCells = 0;
  let work = 0;
  let foundDepth = -1;

  for (let depth = 0; depth <= Math.min(max, MAX_DIFF_DISTANCE); depth++) {
    const snapshot = frontier.slice(offset - depth - 1, offset + depth + 2);
    traceCells += snapshot.length;
    if (traceCells > MAX_TRACE_CELLS) return undefined;
    trace.push(snapshot);

    for (let diagonal = -depth; diagonal <= depth; diagonal += 2) {
      if (++work > MAX_DIFF_WORK) return undefined;
      const chooseInsertion =
        diagonal === -depth ||
        (diagonal !== depth && frontier[offset + diagonal - 1] < frontier[offset + diagonal + 1]);
      let x = chooseInsertion
        ? frontier[offset + diagonal + 1]
        : frontier[offset + diagonal - 1] + 1;
      if (x < 0) x = 0;
      let y = x - diagonal;
      while (x < n && y < m && before.charAt(x) === after.charAt(y)) {
        x++;
        y++;
        if (++work > MAX_DIFF_WORK) return undefined;
      }
      frontier[offset + diagonal] = x;
      if (x >= n && y >= m) {
        foundDepth = depth;
        break;
      }
    }
    if (foundDepth >= 0) break;
  }
  if (foundDepth < 0) return undefined;

  let x = n;
  let y = m;
  const reversed: DiffSegment[] = [];
  for (let depth = foundDepth; depth > 0; depth--) {
    const snapshot = trace[depth];
    const diagonal = x - y;
    const left = snapshot[diagonal + depth];
    const right = snapshot[diagonal + depth + 2];
    const previousDiagonal =
      diagonal === -depth || (diagonal !== depth && left < right) ? diagonal + 1 : diagonal - 1;
    const previousX = snapshot[previousDiagonal + depth + 1];
    const previousY = previousX - previousDiagonal;
    if (previousX < 0 || previousY < 0) return undefined;

    const insertion = diagonal === -depth || (diagonal !== depth && left < right);
    const snakeX = previousX + (insertion ? 0 : 1);
    const snakeY = previousY + (insertion ? 1 : 0);
    const equalCount = x - snakeX;
    if (equalCount !== y - snakeY || equalCount < 0) return undefined;
    if (equalCount > 0) reversed.push({ kind: "equal", count: equalCount });
    if (insertion) {
      if (previousY >= m) return undefined;
      reversed.push({ kind: "insert", value: after.charAt(previousY) });
    } else {
      if (previousX >= n) return undefined;
      reversed.push({ kind: "delete" });
    }
    x = previousX;
    y = previousY;
  }
  if (x !== y) return undefined;
  if (x > 0) reversed.push({ kind: "equal", count: x });
  reversed.reverse();

  const changes: SourceTextChange[] = [];
  let sourceOffset = 0;
  let pendingStart: number | undefined;
  let replacement = "";
  const flush = () => {
    if (pendingStart === undefined) return;
    changes.push({
      start: absoluteStart + pendingStart,
      end: absoluteStart + sourceOffset,
      newText: replacement,
    });
    pendingStart = undefined;
    replacement = "";
  };

  for (const segment of reversed) {
    if (segment.kind === "equal") {
      flush();
      sourceOffset += segment.count;
    } else {
      pendingStart ??= sourceOffset;
      if (segment.kind === "delete") sourceOffset++;
      else replacement += segment.value;
    }
  }
  flush();
  return sourceOffset === n ? changes : undefined;
}

/**
 * Returns a bounded, source-preserving edit script from base to current.
 * Edits use zero-based UTF-16 offsets into base and are ordered by start.
 */
export function computeSourceChanges(base: string, current: string): SourceTextChange[] {
  assertString(base, "base");
  assertString(current, "current");
  if (base === current) return [];

  let prefix = 0;
  const sharedLength = Math.min(base.length, current.length);
  while (prefix < sharedLength && base.charAt(prefix) === current.charAt(prefix)) prefix++;
  let suffix = 0;
  while (
    suffix < sharedLength - prefix &&
    base.charAt(base.length - suffix - 1) === current.charAt(current.length - suffix - 1)
  )
    suffix++;

  const beforeCore = base.slice(prefix, base.length - suffix);
  const afterCore = current.slice(prefix, current.length - suffix);
  if (beforeCore === afterCore) return [];
  if (beforeCore.length === 0 || afterCore.length === 0)
    return [{ start: prefix, end: prefix + beforeCore.length, newText: afterCore }];
  if (beforeCore.length + afterCore.length > MAX_DIFF_INPUT)
    return [{ start: prefix, end: prefix + beforeCore.length, newText: afterCore }];

  return (
    myersChanges(beforeCore, afterCore, prefix) ?? [
      { start: prefix, end: prefix + beforeCore.length, newText: afterCore },
    ]
  );
}

function lineAtOffset(lines: SourceLineIndex, offset: number): number {
  let low = 0;
  let high = lines.starts.length;
  while (low + 1 < high) {
    const middle = (low + high) >>> 1;
    if (lines.starts[middle] <= offset) low = middle;
    else high = middle;
  }
  return low;
}

function editsConflict(
  lines: SourceLineIndex,
  left: SourceTextChange,
  right: SourceTextChange,
): boolean {
  const leftInsertion = left.start === left.end;
  const rightInsertion = right.start === right.end;
  if (leftInsertion && rightInsertion && left.start === right.start) return true;
  if (leftInsertion && right.start < left.start && left.start < right.end) return true;
  if (rightInsertion && left.start < right.start && right.start < left.end) return true;
  if (!leftInsertion && !rightInsertion && left.start < right.end && right.start < left.end)
    return true;

  // Text-only merging cannot identify YAML/JSON property boundaries. Treat
  // changes on one physical line as conflicting so separate character hunks
  // cannot synthesize a value neither editor produced.
  const firstLine = (edit: SourceTextChange) => lineAtOffset(lines, edit.start);
  const lastLine = (edit: SourceTextChange) =>
    edit.start === edit.end ? firstLine(edit) : lineAtOffset(lines, edit.end - 1);
  const leftFirst = firstLine(left);
  const leftLast = lastLine(left);
  const rightFirst = firstLine(right);
  const rightLast = lastLine(right);
  return leftFirst <= rightLast && rightFirst <= leftLast;
}

function sameEdit(left: SourceTextChange, right: SourceTextChange): boolean {
  return left.start === right.start && left.end === right.end && left.newText === right.newText;
}

function mergeThreeWay(base: string, draft: string, source: string): string | undefined {
  const draftChanges = computeSourceChanges(base, draft);
  const sourceChanges = computeSourceChanges(base, source);
  const lines = indexLines(base);
  const merged = [...draftChanges];
  for (const sourceChange of sourceChanges) {
    if (draftChanges.some((draftChange) => sameEdit(draftChange, sourceChange))) continue;
    if (draftChanges.some((draftChange) => editsConflict(lines, draftChange, sourceChange)))
      return undefined;
    merged.push(sourceChange);
  }
  return applyOffsetChanges(base, merged);
}

// Used only after an explicit conflict-resolution choice. Preserve the native
// history identities, rebasing its snapshots onto the latest source. A past
// Designer value can be restored in the draft without changing the source.
function rebaseResolvedSnapshot(base: string, snapshot: string, source: string): string {
  const changes = computeSourceChanges(base, snapshot);
  const lines = indexLines(base);
  const external = computeSourceChanges(base, source).filter(
    (change) => !changes.some((draftChange) => editsConflict(lines, draftChange, change)),
  );
  return applyOffsetChanges(base, [...changes, ...external]);
}

export class DraftState {
  private _baseText: string;
  private _text: string;
  private _revision = 0;
  private _conflicted = false;
  private _pendingSourceText: string | undefined;
  private history: DraftFrame[] = [];
  private cursor = 0;

  constructor(baseText: string) {
    assertString(baseText, "baseText");
    this._baseText = baseText;
    this._text = baseText;
  }

  get text(): string {
    return this._text;
  }

  get baseText(): string {
    return this._baseText;
  }

  get revision(): number {
    return this._revision;
  }

  get dirty(): boolean {
    return this._text !== this._baseText;
  }

  get conflicted(): boolean {
    return this._conflicted;
  }

  retainConflict(externalText: string): void {
    assertString(externalText, "externalText");
    if (!this._conflicted || this._pendingSourceText !== externalText) this._revision++;
    this._conflicted = true;
    this._pendingSourceText = externalText;
  }

  /** Applied native-undo frames, in callback order, for backup restoration. */
  get appliedFrames(): readonly DraftFrame[] {
    return this.history.slice(0, this.cursor);
  }

  canUndo(frame: DraftFrame): boolean {
    return this.cursor > 0 && this.history[this.cursor - 1] === frame && frame.after === this._text;
  }

  canRedo(frame: DraftFrame): boolean {
    return (
      this.cursor < this.history.length &&
      this.history[this.cursor] === frame &&
      frame.before === this._text
    );
  }

  commit(nextText: string, label?: string): DraftFrame | undefined {
    assertString(nextText, "nextText");
    if (this._conflicted) throw new Error("Resolve the source conflict before editing the draft.");
    if (nextText === this._text) return undefined;
    this.history.splice(this.cursor);
    const frame = createFrame(this._text, nextText, label);
    this.history.push(frame);
    this.cursor++;
    this._text = nextText;
    this._revision++;
    return frame;
  }

  undo(frame: DraftFrame): void {
    if (this._conflicted) throw new Error("Resolve the source conflict before undoing the draft.");
    if (this.cursor === 0 || this.history[this.cursor - 1] !== frame || frame.after !== this._text)
      throw new Error("The draft frame is no longer the current undo operation.");
    this._text = frame.before;
    this.cursor--;
    this._revision++;
  }

  redo(frame: DraftFrame): void {
    if (this._conflicted) throw new Error("Resolve the source conflict before redoing the draft.");
    if (
      this.cursor >= this.history.length ||
      this.history[this.cursor] !== frame ||
      frame.before !== this._text
    )
      throw new Error("The draft frame is no longer the current redo operation.");
    this._text = frame.after;
    this.cursor++;
    this._revision++;
  }

  synchronize(externalText: string): "merged" | "conflict" {
    assertString(externalText, "externalText");
    const mergedCurrent = mergeThreeWay(this._baseText, this._text, externalText);
    const markConflict = (): "conflict" => {
      const changed = !this._conflicted || this._pendingSourceText !== externalText;
      this._conflicted = true;
      this._pendingSourceText = externalText;
      if (changed) this._revision++;
      return "conflict";
    };
    if (mergedCurrent === undefined) return markConflict();

    const snapshots: {
      frame: DraftFrame;
      before: string | undefined;
      after: string | undefined;
    }[] = [];
    const cache = new Map<string, string | undefined>();
    const rebase = (snapshot: string): string | undefined => {
      if (cache.has(snapshot)) return cache.get(snapshot);
      const merged = mergeThreeWay(this._baseText, snapshot, externalText);
      cache.set(snapshot, merged);
      return merged;
    };

    for (const frame of this.history) {
      const before = rebase(frame.before);
      const after = rebase(frame.after);
      snapshots.push({ frame, before, after });
    }

    // History can include edits already undone or superseded by a save. Those
    // frames do not make the current draft conflicted. Keep only the contiguous
    // safe interval touching the cursor so native callbacks cannot restore a
    // snapshot that omits the external edit.
    let firstRetained = 0;
    for (let index = this.cursor - 1; index >= 0; index--) {
      const snapshot = snapshots[index];
      if (snapshot.before === undefined || snapshot.after === undefined) {
        firstRetained = index + 1;
        break;
      }
    }
    let endRetained = snapshots.length;
    for (let index = this.cursor; index < snapshots.length; index++) {
      const snapshot = snapshots[index];
      if (snapshot.before === undefined || snapshot.after === undefined) {
        endRetained = index;
        break;
      }
    }
    const retained = snapshots.slice(firstRetained, endRetained);
    let changed =
      this._baseText !== externalText ||
      this._text !== mergedCurrent ||
      this._conflicted ||
      this._pendingSourceText !== undefined ||
      retained.length !== this.history.length ||
      firstRetained > 0;
    for (const snapshot of retained) {
      if (
        snapshot.before === undefined ||
        snapshot.after === undefined ||
        snapshot.frame.before !== snapshot.before ||
        snapshot.frame.after !== snapshot.after
      )
        changed = true;
    }
    this._baseText = externalText;
    this._text = mergedCurrent;
    for (const snapshot of retained) {
      // Retained frames are known to have successfully rebased above.
      setFrameSnapshot(snapshot.frame, snapshot.before!, snapshot.after!);
    }
    this.history = retained.map(({ frame }) => frame);
    this.cursor -= firstRetained;
    this._conflicted = false;
    this._pendingSourceText = undefined;
    if (changed) this._revision++;
    return "merged";
  }

  markSaved(): void {
    if (this._conflicted) throw new Error("Resolve the source conflict before saving the draft.");
    if (this._baseText === this._text) return;
    this._baseText = this._text;
    this._revision++;
  }

  resolveConflict(choice: "source" | "designer"): DraftFrame {
    if (!this._conflicted || this._pendingSourceText === undefined)
      throw new Error("There is no source conflict to resolve.");
    const source = this._pendingSourceText;
    const before = rebaseResolvedSnapshot(this._baseText, this._text, source);
    const after = choice === "source" ? source : this._text;
    const frame = createFrame(before, after, "Resolve conflict");
    const retained = this.history.slice(0, this.cursor);
    for (const previous of retained) {
      setFrameSnapshot(
        previous,
        rebaseResolvedSnapshot(this._baseText, previous.before, source),
        rebaseResolvedSnapshot(this._baseText, previous.after, source),
      );
    }
    this._baseText = source;
    this._text = after;
    // Keep the existing native stack aligned: Undo resolution restores the
    // prior draft, and Undo that draft restores the latest source baseline.
    this.history = [...retained, frame];
    this.cursor = retained.length + 1;
    this._conflicted = false;
    this._pendingSourceText = undefined;
    this._revision++;
    return frame;
  }

  revert(text: string): void {
    assertString(text, "text");
    const changed =
      this._text !== text ||
      this._baseText !== text ||
      this.history.length > 0 ||
      this._conflicted ||
      this._pendingSourceText !== undefined;
    this._baseText = text;
    this._text = text;
    this.history = [];
    this.cursor = 0;
    this._conflicted = false;
    this._pendingSourceText = undefined;
    if (changed) this._revision++;
  }

  serialize(): SerializedDraftState {
    return {
      schemaVersion: 1,
      baseText: this._baseText,
      text: this._text,
      revision: this._revision,
      history: this.history.map((frame) => ({
        before: frame.before,
        after: frame.after,
        ...(frame.label === undefined ? {} : { label: frame.label }),
      })),
      cursor: this.cursor,
      conflicted: this._conflicted,
      ...(this._pendingSourceText === undefined
        ? {}
        : { pendingSourceText: this._pendingSourceText }),
    };
  }

  static restore(value: unknown): DraftState {
    if (
      !isRecord(value) ||
      value.schemaVersion !== 1 ||
      typeof value.baseText !== "string" ||
      typeof value.text !== "string" ||
      !Number.isSafeInteger(value.revision) ||
      (value.revision as number) < 0 ||
      !Array.isArray(value.history) ||
      !Number.isSafeInteger(value.cursor) ||
      (value.cursor as number) < 0 ||
      (value.cursor as number) > value.history.length ||
      typeof value.conflicted !== "boolean" ||
      (value.pendingSourceText !== undefined && typeof value.pendingSourceText !== "string") ||
      (value.conflicted && value.pendingSourceText === undefined) ||
      (!value.conflicted && value.pendingSourceText !== undefined)
    )
      throw new TypeError("Invalid serialized UserForm draft state.");

    const history = value.history.map((entry) => {
      if (
        !isRecord(entry) ||
        typeof entry.before !== "string" ||
        typeof entry.after !== "string" ||
        (entry.label !== undefined && typeof entry.label !== "string")
      )
        throw new TypeError("Invalid serialized UserForm draft frame.");
      return {
        before: entry.before,
        after: entry.after,
        ...(entry.label === undefined ? {} : { label: entry.label }),
      };
    });
    for (let index = 1; index < history.length; index++) {
      if (history[index - 1].after !== history[index].before)
        throw new TypeError("Serialized UserForm draft history is discontinuous.");
    }
    const cursor = value.cursor as number;
    if (
      (cursor > 0 && history[cursor - 1].after !== value.text) ||
      (cursor < history.length && history[cursor].before !== value.text)
    )
      throw new TypeError("Serialized UserForm draft text does not match its history cursor.");

    const state = new DraftState(value.baseText);
    state._text = value.text;
    state._revision = value.revision as number;
    state.history = history.map((frame) => createFrame(frame.before, frame.after, frame.label));
    state.cursor = cursor;
    state._conflicted = value.conflicted;
    state._pendingSourceText = value.pendingSourceText as string | undefined;
    return state;
  }
}
