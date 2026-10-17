import * as vscode from "vscode";
import { DraftState, applySourceTextEdits, computeSourceChanges } from "./draft";
import type { DraftFrame } from "./draft";
import type { SourceTextEdit } from "./protocol";

/** Shared across all Designer panels for a URI; the source buffer stays untouched. */
export class UserFormCustomDocument implements vscode.CustomDocument {
  public readonly uri: vscode.Uri;
  public readonly sourceUri: vscode.Uri;
  public readonly document: vscode.TextDocument;
  public state: DraftState;
  public saving = false;
  public reconciling = false;
  public closed = false;
  public saveError: string | undefined;
  private readonly changed = new vscode.EventEmitter<void>();
  public readonly onDidChange = this.changed.event;
  private readonly disposed = new vscode.EventEmitter<void>();
  public readonly onDidDispose = this.disposed.event;
  private readonly sourceSubscription: vscode.Disposable;
  private readonly sourceOpenSubscription: vscode.Disposable;
  private readonly sourceWatcher: vscode.FileSystemWatcher | undefined;
  private appliedSource: string | undefined;
  private observedSource: string;
  private sourceEpoch = 0;
  private sourceChangedWhileSaving = false;

  constructor(
    public sourceDocument: vscode.TextDocument,
    state?: DraftState,
    editorUri: vscode.Uri = sourceDocument.uri,
    private readonly sourceMergeSafe: (
      baseText: string,
      externalText: string,
    ) => Promise<boolean> = async () => true,
    private readonly serializeSource: (action: () => Promise<void>) => Promise<void> = (action) =>
      action(),
  ) {
    this.uri = editorUri;
    this.sourceUri = sourceDocument.uri;
    this.state = state ?? new DraftState(sourceDocument.getText());
    this.observedSource = sourceDocument.getText();
    this.document = virtualTextDocument(
      () => this.sourceDocument,
      () => this.previewText,
      () => this.state.revision + this.sourceEpoch,
      () => this.closed,
    );
    this.sourceSubscription = vscode.workspace.onDidChangeTextDocument((event) => {
      if (event.document.uri.toString() !== this.sourceUri.toString()) return;
      this.sourceDocument = event.document;
      const text = event.document.getText();
      // Only our exact save edit is exempt from external-change reconciliation.
      this.observeSource(text);
    });
    this.sourceOpenSubscription = vscode.workspace.onDidOpenTextDocument((document) => {
      if (document.uri.toString() !== this.sourceUri.toString()) return;
      this.sourceDocument = document;
      this.observeSource(document.getText());
    });
    if (this.sourceUri.scheme === "file") {
      const slash = this.sourceUri.path.lastIndexOf("/");
      this.sourceWatcher = vscode.workspace.createFileSystemWatcher(
        new vscode.RelativePattern(
          this.sourceUri.with({ path: this.sourceUri.path.slice(0, slash) }),
          this.sourceUri.path.slice(slash + 1),
        ),
      );
      const refresh = async () => {
        try {
          const source = await vscode.workspace.openTextDocument(this.sourceUri);
          if (this.closed) return;
          this.sourceDocument = source;
          this.observeSource(source.getText());
        } catch (error) {
          if (!this.closed) {
            this.saveError = String(error);
            this.changed.fire();
          }
        }
      };
      this.sourceWatcher.onDidChange(refresh);
      this.sourceWatcher.onDidCreate(refresh);
    }
  }

  get previewText(): string {
    return this.state.conflicted || this.reconciling
      ? this.sourceDocument.getText()
      : this.state.text;
  }

  private observeSource(text: string): void {
    // TextDocument save participants can change the source buffer while
    // save() is awaiting VS Code. The save transaction rebases from that final
    // saved buffer; queueing reconciliation here would race that rebase.
    if (this.saving) {
      this.observedSource = text;
      this.sourceChangedWhileSaving = true;
      return;
    }
    if (this.appliedSource === text) {
      this.observedSource = text;
      return;
    }
    if (this.closed || this.observedSource === text) return;
    this.observedSource = text;
    const epoch = ++this.sourceEpoch;
    this.reconciling = true;
    this.changed.fire();
    void this.serializeSource(async () => {
      if (this.closed || epoch !== this.sourceEpoch) return;
      try {
        await this.reconcileSource(text);
      } catch (error) {
        if (epoch === this.sourceEpoch) {
          this.state.retainConflict(text);
          this.saveError = String(error);
        }
      } finally {
        if (!this.closed && epoch === this.sourceEpoch) {
          this.reconciling = false;
          this.changed.fire();
        }
      }
    });
  }

  async reconcileSource(text: string): Promise<boolean> {
    if (text === this.state.baseText && !this.state.conflicted) return true;
    const revision = this.state.revision;
    const epoch = this.sourceEpoch;
    const safe = !this.state.dirty || (await this.sourceMergeSafe(this.state.baseText, text));
    if (
      this.closed ||
      epoch !== this.sourceEpoch ||
      this.state.revision !== revision ||
      this.sourceDocument.getText() !== text
    )
      return false;
    if (safe) this.state.synchronize(text);
    else this.state.retainConflict(text);
    return true;
  }

  notify(): void {
    this.changed.fire();
  }

  apply(edits: SourceTextEdit[], label?: string): DraftFrame | undefined {
    if (this.closed || this.saving || this.reconciling || this.state.conflicted)
      throw new Error("The Designer draft is unavailable for editing.");
    // TextDocument applies multiline TextEdits using the source document's EOL.
    // Apply the same rule in the draft, before creating its Undo/Redo frame.
    // Otherwise our own publish is mistaken for an external source change.
    const eol = this.sourceDocument.eol === vscode.EndOfLine.CRLF ? "\r\n" : "\n";
    const normalizedEdits = edits.map((edit) => ({
      ...edit,
      newText: edit.newText.replace(/\r\n|\r|\n/g, eol),
    }));
    const frame = this.state.commit(applySourceTextEdits(this.state.text, normalizedEdits), label);
    if (frame) this.changed.fire();
    return frame;
  }

  undo(frame: DraftFrame): void {
    // VS Code retains callbacks for edits superseded by direct source editing.
    // Those entries are neutral: advancing past them must never restore stale
    // source or make the native Undo command fail.
    if (this.closed || !this.state.canUndo(frame)) return;
    this.state.undo(frame);
    this.changed.fire();
  }
  redo(frame: DraftFrame): void {
    if (this.closed || !this.state.canRedo(frame)) return;
    this.state.redo(frame);
    this.changed.fire();
  }

  async save(validate: () => Promise<void>, cancellation: vscode.CancellationToken): Promise<void> {
    if (this.closed) throw new Error("The Designer is closed.");
    if (this.saving)
      throw new Error(vscode.l10n.t("A UserForm source save is already in progress."));
    this.saving = true;
    this.sourceChangedWhileSaving = false;
    this.saveError = undefined;
    this.changed.fire();
    try {
      this.sourceDocument = await vscode.workspace.openTextDocument(this.sourceUri);
      if (!(await this.reconcileSource(this.sourceDocument.getText())))
        throw new Error(
          vscode.l10n.t("The document changed. Retry the operation on the updated form."),
        );
      if (this.state.conflicted)
        throw new Error(
          vscode.l10n.t("Resolve the YAML/JSON and Designer conflict before saving."),
        );
      const version = this.sourceDocument.version;
      const revision = this.state.revision;
      const text = this.state.text;
      await validate();
      if (cancellation.isCancellationRequested) throw new vscode.CancellationError();
      if (
        this.closed ||
        this.sourceDocument.version !== version ||
        this.state.revision !== revision
      )
        throw new Error(
          vscode.l10n.t("The document changed. Retry the operation on the updated form."),
        );
      const changes = computeSourceChanges(this.sourceDocument.getText(), text);
      if (changes.length) {
        const edit = new vscode.WorkspaceEdit();
        edit.set(
          this.sourceUri,
          changes.map(
            (change) =>
              new vscode.TextEdit(
                new vscode.Range(
                  this.sourceDocument.positionAt(change.start),
                  this.sourceDocument.positionAt(change.end),
                ),
                change.newText,
              ),
          ),
        );
        this.appliedSource = text;
        try {
          if (!(await vscode.workspace.applyEdit(edit)))
            throw new Error(vscode.l10n.t("VS Code could not apply the UserForm edit."));
        } finally {
          this.appliedSource = undefined;
        }
      }
      if (cancellation.isCancellationRequested) throw new vscode.CancellationError();
      if (this.sourceDocument.getText() !== text || this.state.revision !== revision)
        throw new Error(
          vscode.l10n.t("The document changed. Retry the operation on the updated form."),
        );
      if (!(await this.sourceDocument.save()))
        throw new Error(
          vscode.l10n.t(
            "Could not save the UserForm source. The Designer draft has been retained.",
          ),
        );
      if (this.sourceDocument.isDirty || this.state.revision !== revision)
        throw new Error(
          vscode.l10n.t("The document changed. Retry the operation on the updated form."),
        );
      // Publication was verified before invoking save. Save participants may
      // then format the source; the successfully saved buffer is authoritative.
      // Rebase from the published draft, rather than its pre-publication base,
      // so our own edit cannot conflict with its formatted counterpart.
      this.state.markSaved();
      this.state.synchronize(this.sourceDocument.getText());
      this.sourceChangedWhileSaving = false;
    } catch (error) {
      this.saveError = String(error);
      throw error;
    } finally {
      const reconcileFailedSave =
        this.sourceChangedWhileSaving && this.state.baseText !== this.sourceDocument.getText();
      this.sourceChangedWhileSaving = false;
      this.saving = false;
      this.changed.fire();
      if (reconcileFailedSave) {
        // A failed native save may still have changed the in-memory source
        // buffer. Reconcile it after releasing the save transaction.
        this.observedSource = this.state.baseText;
        this.observeSource(this.sourceDocument.getText());
      }
    }
  }

  dispose(): void {
    if (this.closed) return;
    this.closed = true;
    this.sourceSubscription.dispose();
    this.sourceOpenSubscription.dispose();
    this.sourceWatcher?.dispose();
    this.disposed.fire();
    this.changed.dispose();
    this.disposed.dispose();
  }
}

export function virtualTextDocument(
  readSource: () => vscode.TextDocument,
  readText: () => string,
  readVersion: () => number,
  closed: () => boolean,
): vscode.TextDocument {
  const lines = () => {
    const text = readText();
    const starts = [0];
    for (let index = 0; index < text.length; index++)
      if (text[index] === "\n") starts.push(index + 1);
    return { text, starts };
  };
  const offsetAt = (position: vscode.Position): number => {
    const { text, starts } = lines();
    const line = Math.max(0, Math.min(position.line, starts.length - 1));
    const end =
      line + 1 < starts.length
        ? starts[line + 1] - (text[starts[line + 1] - 2] === "\r" ? 2 : 1)
        : text.length;
    return starts[line] + Math.max(0, Math.min(position.character, end - starts[line]));
  };
  const positionAt = (offset: number): vscode.Position => {
    const { text, starts } = lines();
    offset = Math.max(0, Math.min(offset, text.length));
    let line = 0;
    while (line + 1 < starts.length && starts[line + 1] <= offset) line++;
    return new vscode.Position(line, offset - starts[line]);
  };
  // VS Code freezes the public TextDocument API. Proxy an empty facade rather
  // than overriding non-configurable properties on that frozen object.
  return new Proxy({} as vscode.TextDocument, {
    get(_target, key) {
      if (key === "version") return readVersion();
      if (key === "isClosed") return closed();
      if (key === "getText")
        return (range?: vscode.Range) =>
          range ? readText().slice(offsetAt(range.start), offsetAt(range.end)) : readText();
      if (key === "lineCount") return lines().starts.length;
      if (key === "offsetAt") return offsetAt;
      if (key === "positionAt") return positionAt;
      if (key === "lineAt")
        return (value: number | vscode.Position) => {
          const lineNumber = typeof value === "number" ? value : value.line;
          const { text, starts } = lines();
          if (lineNumber < 0 || lineNumber >= starts.length)
            throw new RangeError("Invalid line number");
          const rawEnd = starts[lineNumber + 1] ?? text.length;
          const lineText = text.slice(starts[lineNumber], rawEnd).replace(/\r?\n$/, "");
          const range = new vscode.Range(lineNumber, 0, lineNumber, lineText.length);
          return {
            lineNumber,
            text: lineText,
            range,
            rangeIncludingLineBreak:
              lineNumber + 1 < starts.length
                ? new vscode.Range(lineNumber, 0, lineNumber + 1, 0)
                : range,
            firstNonWhitespaceCharacterIndex:
              lineText.search(/\S/) < 0 ? lineText.length : lineText.search(/\S/),
            isEmptyOrWhitespace: /^\s*$/.test(lineText),
          };
        };
      const source = readSource();
      const value = Reflect.get(source, key, source);
      return typeof value === "function" ? value.bind(source) : value;
    },
  });
}
