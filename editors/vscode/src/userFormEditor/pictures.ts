import type * as vscode from "vscode";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import { createHash, randomUUID } from "node:crypto";

export interface PictureValidationRequest {
  uri: string;
  path?: string;
  data?: string;
}

export interface PictureValidationResult {
  path?: string;
  format?: "bmp" | "jpeg";
  sha256?: string;
  dataURI?: string;
  width?: number;
  height?: number;
  error?: { code: string; message: string };
}

export interface PictureSelection {
  token: string;
  name: string;
  dataURI: string;
  width: number;
  height: number;
}

export interface PictureCommitReceipt {
  picture: { path: string };
  /** Opaque receipt; callers must not construct or forward it from a webview. */
  token: string;
}

export interface PictureImportOptions {
  projectRoot: string;
  formsRoot: string;
  sourceURI: string;
  validate: (request: PictureValidationRequest) => Promise<PictureValidationResult>;
  pick?: () => PromiseLike<vscode.Uri | undefined>;
}

interface StagedPicture {
  selection: PictureSelection;
  bytes: Buffer;
  relativePath?: string;
  extension: "bmp" | "jpg";
  digest: string;
}

interface CreatedAsset {
  absolutePath: string;
  relativePath: string;
  digest: string;
  identity: string;
}

const maxTransportBytes = 16 * 1024 * 1024;

function inside(root: string, candidate: string): boolean {
  const relative = path.relative(root, candidate);
  return (
    relative !== "" &&
    !path.isAbsolute(relative) &&
    relative !== ".." &&
    !relative.startsWith(`..${path.sep}`)
  );
}

function portable(relative: string): string {
  return relative.split(path.sep).join("/");
}

function digest(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

/** One editor's staged selection. Selection/Cancel never creates project files.
 * Bind an instance to the source URI and dispose it when that editor closes.
 * Call commit only after checking the editor's target/version is still current.
 */
export class PictureImportSession {
  private staged?: StagedPicture;
  private readonly created = new Map<string, CreatedAsset>();

  constructor(private readonly options: PictureImportOptions) {}

  async open(): Promise<PictureSelection | undefined> {
    const uri = await (this.options.pick?.() ?? this.pick());
    if (!uri) return undefined;
    if (uri.scheme !== "file") throw new Error("Select a local BMP or JPEG file.");
    const root = await fs.realpath(this.options.projectRoot);
    const selected = path.resolve(uri.fsPath);
    const actual = await fs.realpath(selected);
    const lexicalInside = inside(path.resolve(this.options.projectRoot), selected);
    // A path presented as an existing project asset cannot escape via a link.
    if (lexicalInside && !inside(root, actual))
      throw new Error("Picture links must stay inside the project root.");
    const handle = await fs.open(actual, "r");
    let bytes: Buffer;
    try {
      const stat = await handle.stat();
      if (!stat.isFile() || stat.size <= 0 || stat.size > maxTransportBytes)
        throw new Error("Select a BMP or JPEG file no larger than 16 MiB.");
      // A bounded read also protects against a file growing after stat().
      bytes = Buffer.alloc(stat.size + 1);
      const read = await handle.read(bytes, 0, bytes.length, 0);
      if (read.bytesRead !== stat.size)
        throw new Error("The picture changed while it was being selected. Select it again.");
      bytes = bytes.subarray(0, stat.size);
    } finally {
      await handle.close();
    }
    const relativePath = inside(root, actual) ? portable(path.relative(root, actual)) : undefined;
    const result = await this.options.validate({
      uri: this.options.sourceURI,
      data: bytes.toString("base64"),
    });
    if (result.error) throw new Error(result.error.message);
    const hash = digest(bytes);
    if (
      (result.format !== "bmp" && result.format !== "jpeg") ||
      result.sha256 !== hash ||
      result.dataURI !== `data:image/${result.format};base64,${bytes.toString("base64")}` ||
      !Number.isSafeInteger(result.width) ||
      !Number.isSafeInteger(result.height) ||
      result.width! <= 0 ||
      result.height! <= 0
    )
      throw new Error("Invalid picture validation response.");
    // Validate the path through the same rooted Go loader used by pack/push.
    if (relativePath) {
      const current = await this.options.validate({
        uri: this.options.sourceURI,
        path: relativePath,
      });
      if (current.error) throw new Error(current.error.message);
      if (current.sha256 !== hash)
        throw new Error("The picture changed while it was being selected. Select it again.");
    }
    const selection: PictureSelection = {
      token: randomUUID(),
      name: path.basename(selected),
      dataURI: result.dataURI,
      width: result.width!,
      height: result.height!,
    };
    this.staged = {
      selection,
      bytes,
      relativePath,
      extension: result.format === "jpeg" ? "jpg" : "bmp",
      digest: hash,
    };
    return selection;
  }

  private async pick(): Promise<vscode.Uri | undefined> {
    const vscode = await import("vscode");
    return (
      await vscode.window.showOpenDialog({
        canSelectFiles: true,
        canSelectFolders: false,
        canSelectMany: false,
        title: vscode.l10n.t("Select Image picture"),
        filters: { "BMP / JPEG": ["bmp", "jpg", "jpeg"] },
      })
    )?.[0];
  }

  cancel(): void {
    this.staged = undefined;
  }

  async commit(token: string): Promise<PictureCommitReceipt> {
    const staged = this.staged;
    if (!staged || staged.selection.token !== token)
      throw new Error("The picture selection expired. Select it again.");
    if (staged.relativePath) {
      const result = await this.options.validate({
        uri: this.options.sourceURI,
        path: staged.relativePath,
      });
      if (result.error) throw new Error(result.error.message);
      if (result.sha256 !== staged.digest) throw new Error("The picture changed. Select it again.");
      return { token, picture: { path: staged.relativePath } };
    }
    const root = await fs.realpath(this.options.projectRoot);
    const assets = path.resolve(root, this.options.formsRoot.replace(/\\/g, "/"), "assets");
    if (!inside(root, assets))
      throw new Error("The forms assets directory must stay inside the project root.");
    await this.ensureDirectories(root, assets);
    const target = path.join(assets, `${staged.digest}.${staged.extension}`);
    const relativePath = portable(path.relative(root, target));
    try {
      const file = await fs.open(target, "wx+");
      // Register before writing so a partial write is still eligible for cleanup.
      try {
        const stat = await file.stat({ bigint: true });
        const created = {
          absolutePath: target,
          relativePath,
          digest: staged.digest,
          identity: `${stat.dev}:${stat.ino}`,
        };
        this.created.set(token, created);
        try {
          await file.writeFile(staged.bytes);
        } catch (error) {
          // A failed write may leave a prefix. Record its exact content using
          // the original open handle, never a path which could be replaced.
          const current = await file.stat();
          if (current.size <= staged.bytes.length) {
            const partial = Buffer.alloc(current.size);
            const read = await file.read(partial, 0, partial.length, 0);
            if (
              read.bytesRead === partial.length &&
              staged.bytes.subarray(0, partial.length).equals(partial)
            )
              created.digest = digest(partial);
          }
          throw error;
        }
      } finally {
        await file.close();
      }
    } catch (error) {
      if (!(error instanceof Error) || !("code" in error) || error.code !== "EEXIST") throw error;
      const stat = await fs.lstat(target);
      if (
        !stat.isFile() ||
        stat.isSymbolicLink() ||
        stat.size !== staged.bytes.length ||
        !Buffer.from(await fs.readFile(target)).equals(staged.bytes)
      )
        throw new Error("A different asset already uses this content-addressed filename.");
    }
    const checked = await this.options.validate({
      uri: this.options.sourceURI,
      path: relativePath,
    });
    if (checked.error) throw new Error(checked.error.message);
    if (checked.sha256 !== staged.digest)
      throw new Error("The imported picture changed before confirmation.");
    return { token, picture: { path: relativePath } };
  }

  private async ensureDirectories(root: string, directory: string): Promise<void> {
    let current = root;
    for (const segment of path.relative(root, directory).split(path.sep)) {
      current = path.join(current, segment);
      try {
        await fs.mkdir(current);
      } catch (error) {
        if (!(error instanceof Error) || !("code" in error) || error.code !== "EEXIST") throw error;
      }
      const stat = await fs.lstat(current);
      if (!stat.isDirectory() || stat.isSymbolicLink() || !inside(root, await fs.realpath(current)))
        throw new Error("The assets directory cannot traverse symbolic links or junctions.");
    }
  }

  /** Call after a successful semantic edit. Assets deliberately survive Undo. */
  accept(receipt: PictureCommitReceipt): void {
    this.created.delete(receipt.token);
    if (this.staged?.selection.token === receipt.token) this.staged = undefined;
  }

  /** Call on failed semantic edits. isReferenced must include current unsaved
   * project drafts; when references cannot be established, retain the asset.
   */
  async rollback(
    receipt: Pick<PictureCommitReceipt, "token">,
    isReferenced: (relativePath: string) => Promise<boolean>,
  ): Promise<void> {
    const created = this.created.get(receipt.token);
    if (!created) return;
    if (await isReferenced(created.relativePath)) return;
    const root = await fs.realpath(this.options.projectRoot);
    const actual = await fs.realpath(created.absolutePath);
    if (!inside(root, actual) || actual !== created.absolutePath)
      throw new Error("The created asset path changed; it was retained.");
    const file = await fs.open(actual, "r");
    try {
      const stat = await file.stat({ bigint: true });
      if (`${stat.dev}:${stat.ino}` !== created.identity || stat.size > BigInt(maxTransportBytes))
        throw new Error("The created asset was replaced; it was retained.");
      const bytes = await file.readFile();
      if (digest(bytes) !== created.digest)
        throw new Error("The created asset changed; it was retained.");
    } finally {
      await file.close();
    }
    const current = await fs.lstat(actual, { bigint: true });
    if (current.isSymbolicLink() || `${current.dev}:${current.ino}` !== created.identity)
      throw new Error("The created asset was replaced; it was retained.");
    // File identity checks above yield to other editors. Recheck references
    // after those reads, immediately before issuing deletion.
    if (await isReferenced(created.relativePath)) return;
    await fs.unlink(actual);
    this.created.delete(receipt.token);
  }
}
