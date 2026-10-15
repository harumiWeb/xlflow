import { createHash, randomUUID } from "node:crypto";
import * as vscode from "vscode";
import { DraftState } from "./draft";
import type { SerializedDraftState } from "./draft";

export interface DraftCheckpoint {
  readonly schemaVersion: 1;
  readonly sourceUri: string;
  readonly updatedAt: number;
  readonly state: SerializedDraftState;
}

export function decodeDraftCheckpoint(value: unknown, sourceUri: vscode.Uri): DraftCheckpoint {
  if (
    !value ||
    typeof value !== "object" ||
    (value as DraftCheckpoint).schemaVersion !== 1 ||
    (value as DraftCheckpoint).sourceUri !== sourceUri.toString() ||
    !Number.isSafeInteger((value as DraftCheckpoint).updatedAt) ||
    (value as DraftCheckpoint).updatedAt < 0
  )
    throw new Error(
      vscode.l10n.t("The saved Designer draft is invalid. Its recovery data has been retained."),
    );
  const checkpoint = value as DraftCheckpoint;
  try {
    DraftState.restore(checkpoint.state);
    return checkpoint;
  } catch (cause) {
    throw new Error(
      vscode.l10n.t("The saved Designer draft is invalid. Its recovery data has been retained."),
      { cause },
    );
  }
}

/** Durable Designer checkpoints are independent of publishing source files. */
export class DraftCheckpointStore {
  private clock = 0;
  private mutationRevision = 0;
  private writesPending = 0;

  get revision(): number {
    return this.mutationRevision;
  }

  get writing(): boolean {
    return this.writesPending !== 0;
  }

  constructor(private readonly root: vscode.Uri) {}

  location(sourceUri: vscode.Uri): vscode.Uri {
    const key = createHash("sha256").update(sourceUri.toString()).digest("hex");
    return vscode.Uri.joinPath(this.root, `${key}.json`);
  }

  observe(checkpoint: DraftCheckpoint): void {
    this.clock = Math.max(this.clock, checkpoint.updatedAt);
  }

  async read(sourceUri: vscode.Uri): Promise<DraftCheckpoint | undefined> {
    let bytes: Uint8Array;
    try {
      bytes = await vscode.workspace.fs.readFile(this.location(sourceUri));
    } catch (error) {
      if (error instanceof vscode.FileSystemError && error.code === "FileNotFound")
        return undefined;
      throw error;
    }
    try {
      const checkpoint = decodeDraftCheckpoint(
        JSON.parse(Buffer.from(bytes).toString("utf8")),
        sourceUri,
      );
      this.observe(checkpoint);
      return checkpoint;
    } catch (cause) {
      throw new Error(
        vscode.l10n.t("The saved Designer draft is invalid. Its recovery data has been retained."),
        { cause },
      );
    }
  }

  async write(
    sourceUri: vscode.Uri,
    state: SerializedDraftState,
    token: vscode.CancellationToken,
  ): Promise<DraftCheckpoint> {
    if (token.isCancellationRequested) throw new vscode.CancellationError();
    const checkpoint: DraftCheckpoint = {
      schemaVersion: 1,
      sourceUri: sourceUri.toString(),
      updatedAt: (this.clock = Math.max(Date.now(), this.clock + 1)),
      state,
    };
    DraftState.restore(state);
    const destination = this.location(sourceUri);
    const temporary = vscode.Uri.joinPath(this.root, `${randomUUID()}.tmp`);
    this.mutationRevision++;
    this.writesPending++;
    try {
      await vscode.workspace.fs.createDirectory(this.root);
      await vscode.workspace.fs.writeFile(
        temporary,
        Buffer.from(JSON.stringify(checkpoint), "utf8"),
      );
      if (token.isCancellationRequested) throw new vscode.CancellationError();
      await vscode.workspace.fs.rename(temporary, destination, { overwrite: true });
      return checkpoint;
    } finally {
      try {
        await vscode.workspace.fs.delete(temporary);
      } catch {
        /* Renamed or never created. */
      }
      this.writesPending--;
      this.mutationRevision++;
    }
  }
}
