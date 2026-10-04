import * as path from "path";
import type { PreviewResult, HostMessage } from "./protocol";
import { designerDocument } from "./model";

export function isFormSpecPath(file: string, projectRoot: string, formsRoot: string): boolean {
  const specs = path.resolve(projectRoot, formsRoot.replace(/\\/g, "/"), "specs");
  const normalize = (value: string) => (process.platform === "win32" ? value.toLowerCase() : value);
  return (
    normalize(path.dirname(path.resolve(file))) === normalize(specs) &&
    /\.(yaml|yml|json)$/i.test(file)
  );
}

// Version and connection generations both matter: restarted servers may reply to
// a request whose TextDocument version has not changed.
export class PreviewSynchronizer {
  private generation = 0;
  private disposed = false;
  private lastDocument: HostMessage | undefined;
  private status: HostMessage | undefined;

  invalidate(): void {
    this.generation++;
  }
  dispose(): void {
    this.disposed = true;
    this.invalidate();
  }
  replay(send: (message: HostMessage) => void): void {
    if (this.lastDocument) send(this.lastDocument);
    if (this.status) send(this.status);
  }
  async update(
    version: number,
    request: () => Promise<PreviewResult>,
    currentVersion: () => number,
    send: (message: HostMessage) => void,
  ): Promise<void> {
    const generation = ++this.generation;
    let result: PreviewResult;
    try {
      result = await request();
    } catch (error) {
      result = { version, error: { code: "previewUnavailable", message: String(error) } };
    }
    if (
      this.disposed ||
      generation !== this.generation ||
      currentVersion() !== version ||
      result.version !== version
    )
      return;
    if (result.document) {
      this.lastDocument = {
        type: "document",
        version,
        document: designerDocument(
          result.document,
          (result.warnings ?? []).map((w) => w.message),
        ),
      };
      this.status = undefined;
      send(this.lastDocument);
    } else {
      this.status = {
        type: "invalidDocument",
        version,
        error: result.error ?? {
          code: "invalidResponse",
          message: "The preview response has no document.",
        },
      };
      send(this.status);
    }
  }
}
