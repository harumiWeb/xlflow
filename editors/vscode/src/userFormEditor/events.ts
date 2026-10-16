import * as path from "path";
import * as vscode from "vscode";
import { realpath, stat } from "node:fs/promises";
import type { UserFormCodeSource } from "../sidebar";

export interface UserFormEventDescriptor {
  name: string;
  signature: string;
  isDefault: boolean;
}

export interface UserFormEventsRequest {
  uri: string;
  version: number;
  text: string;
  targetId?: string;
}

export interface UserFormEventsResult {
  version: number;
  targetId?: string;
  objectName: string;
  events: UserFormEventDescriptor[];
  error?: string;
}

export interface UserFormEventEditRequest extends UserFormEventsRequest {
  eventName: string;
  codeURI: string;
  codeVersion: number;
  codeText: string;
  indent: string;
  newline: string;
  codeSource: UserFormCodeSource;
}

export interface UserFormEventEditResult {
  version: number;
  codeVersion: number;
  edits: {
    range: EventRange;
    newText: string;
  }[];
  selection: EventRange;
  created: boolean;
  error?: string;
}

export interface EventPosition {
  line: number;
  character: number;
}

export interface EventRange {
  start: EventPosition;
  end: EventPosition;
}

export interface UserFormCodeTarget {
  uri: vscode.Uri;
  codeSource: UserFormCodeSource;
  exists: boolean;
}

/** Serializes event edits across every Designer panel that targets one code file. */
export class UserFormEventQueue {
  private readonly tails = new Map<string, Promise<unknown>>();

  public run<T>(uri: vscode.Uri, operation: () => Promise<T>): Promise<T> {
    const key = uri.toString();
    const previous = this.tails.get(key) ?? Promise.resolve();
    const current = previous.catch(() => undefined).then(operation);
    this.tails.set(key, current);
    return current.finally(() => {
      if (this.tails.get(key) === current) this.tails.delete(key);
    });
  }
}

/** Resolve the configured code artifact without changing the workspace. */
export async function resolveUserFormCodeTarget(
  workspaceFolder: vscode.Uri,
  formsRoot: string,
  codeSource: UserFormCodeSource,
  formName: string,
): Promise<UserFormCodeTarget> {
  if (!isVbaName(formName)) throw new Error(vscode.l10n.t("The UserForm name is invalid."));

  if (path.isAbsolute(formsRoot))
    throw new Error(
      vscode.l10n.t("The configured UserForm source directory must be inside the workspace."),
    );

  const configuredRoot = path.resolve(
    workspaceFolder.fsPath,
    formsRoot.replace(/[\\/]/g, path.sep),
  );
  const relativeRoot = path.relative(workspaceFolder.fsPath, configuredRoot);
  if (
    relativeRoot === ".." ||
    relativeRoot.startsWith(`..${path.sep}`) ||
    path.isAbsolute(relativeRoot)
  )
    throw new Error(
      vscode.l10n.t("The configured UserForm source directory must be inside the workspace."),
    );
  const workspaceRealPath = await realpath(workspaceFolder.fsPath);
  const formsDirectory = await resolveExistingDirectoryCase(
    vscode.Uri.file(configuredRoot),
    workspaceFolder,
    workspaceRealPath,
  );
  if (!formsDirectory)
    throw new Error(vscode.l10n.t("The configured UserForm source directory is missing."));

  if (codeSource === "sidecar") {
    const codeDirectory = await findChildDirectory(
      formsDirectory,
      "code",
      workspaceFolder,
      workspaceRealPath,
    );
    if (!codeDirectory)
      return {
        uri: vscode.Uri.joinPath(formsDirectory, "code", `${formName}.bas`),
        codeSource,
        exists: false,
      };

    const entries = await readDirectory(codeDirectory);
    const candidates: string[] = [];
    for (const [name] of entries) {
      if (
        !name.toLowerCase().endsWith(".bas") ||
        basenameWithoutExtension(name).toLowerCase() !== formName.toLowerCase()
      )
        continue;
      const candidate = vscode.Uri.joinPath(codeDirectory, name);
      await assertWorkspacePathHasNoLinks(candidate, workspaceFolder, workspaceRealPath);
      if ((await stat(candidate.fsPath)).isFile()) candidates.push(name);
    }
    if (candidates.length > 1)
      throw new Error(
        vscode.l10n.t("Multiple UserForm code sidecars match {formName}.", { formName }),
      );
    if (candidates.length === 1) {
      const name = candidates[0];
      return { uri: vscode.Uri.joinPath(codeDirectory, name), codeSource, exists: true };
    }
    return {
      uri: vscode.Uri.joinPath(codeDirectory, `${formName}.bas`),
      codeSource,
      exists: false,
    };
  }

  const candidates = await findFormFiles(formsDirectory, workspaceFolder, workspaceRealPath);
  const matches = candidates.filter(
    (uri) =>
      basenameWithoutExtension(path.basename(uri.fsPath)).toLowerCase() === formName.toLowerCase(),
  );
  if (matches.length === 0)
    throw new Error(
      vscode.l10n.t("The configured .frm code source for {formName} is missing.", { formName }),
    );
  if (matches.length > 1)
    throw new Error(vscode.l10n.t("Multiple .frm files match UserForm {formName}.", { formName }));

  const uri = matches[0];
  const document = await vscode.workspace.openTextDocument(uri);
  const fileName = basenameWithoutExtension(path.basename(uri.fsPath));
  const attributeName = readVBName(document.getText());
  if (
    fileName.toLowerCase() !== formName.toLowerCase() ||
    !attributeName ||
    attributeName.toLowerCase() !== formName.toLowerCase()
  )
    throw new Error(
      vscode.l10n.t("The .frm filename and VB_Name must match UserForm {formName}.", { formName }),
    );
  return { uri, codeSource, exists: true };
}

export function validEventRange(range: EventRange, text: string): boolean {
  return (
    !!range &&
    !!range.start &&
    !!range.end &&
    validPosition(range.start, text) &&
    validPosition(range.end, text) &&
    (range.start.line < range.end.line ||
      (range.start.line === range.end.line && range.start.character <= range.end.character))
  );
}

function validPosition(position: EventPosition, text: string): boolean {
  if (!Number.isSafeInteger(position.line) || !Number.isSafeInteger(position.character))
    return false;
  if (position.line < 0 || position.character < 0) return false;
  const lines = text.split(/\r\n|\r|\n/);
  return position.line < lines.length && position.character <= lines[position.line].length;
}

async function resolveExistingDirectoryCase(
  uri: vscode.Uri,
  workspaceFolder: vscode.Uri,
  workspaceRealPath: string,
): Promise<vscode.Uri | undefined> {
  const resolvedPath = path.resolve(uri.fsPath);
  const rootPath = path.parse(resolvedPath).root;
  let current = vscode.Uri.file(rootPath);
  const segments = resolvedPath.slice(rootPath.length).split(path.sep).filter(Boolean);
  for (const segment of segments) {
    const matches: [string, vscode.FileType][] = [];
    for (const [name, type] of await readDirectory(current)) {
      if (!sameDirectoryName(name, segment)) continue;
      const child = vscode.Uri.joinPath(current, name);
      if (isWorkspacePathOrDescendant(child.fsPath, workspaceFolder.fsPath))
        await assertWorkspacePathHasNoLinks(child, workspaceFolder, workspaceRealPath);
      if ((await stat(child.fsPath)).isDirectory()) matches.push([name, type]);
    }
    if (matches.length > 1)
      throw new Error(
        vscode.l10n.t("The configured path has ambiguous casing at {segment}.", { segment }),
      );
    if (matches.length === 0) return undefined;
    current = vscode.Uri.joinPath(current, matches[0][0]);
  }
  return current;
}

async function findChildDirectory(
  parent: vscode.Uri,
  name: string,
  workspaceFolder: vscode.Uri,
  workspaceRealPath: string,
): Promise<vscode.Uri | undefined> {
  const matches: string[] = [];
  for (const [entry] of await readDirectory(parent)) {
    if (!sameDirectoryName(entry, name)) continue;
    const child = vscode.Uri.joinPath(parent, entry);
    await assertWorkspacePathHasNoLinks(child, workspaceFolder, workspaceRealPath);
    if ((await stat(child.fsPath)).isDirectory()) matches.push(entry);
  }
  if (matches.length > 1)
    throw new Error(
      vscode.l10n.t("Multiple {name} directories exist under {path}.", {
        name,
        path: parent.fsPath,
      }),
    );
  return matches.length ? vscode.Uri.joinPath(parent, matches[0]) : undefined;
}

async function findFormFiles(
  root: vscode.Uri,
  workspaceFolder: vscode.Uri,
  workspaceRealPath: string,
): Promise<vscode.Uri[]> {
  const result: vscode.Uri[] = [];
  const visit = async (directory: vscode.Uri): Promise<void> => {
    for (const [name, type] of await readDirectory(directory)) {
      const lowerName = name.toLowerCase();
      const child = vscode.Uri.joinPath(directory, name);
      let childStat;
      try {
        childStat = await stat(child.fsPath);
      } catch {
        if (lowerName.endsWith(".frm") || (type & vscode.FileType.Directory) !== 0)
          await assertWorkspacePathHasNoLinks(child, workspaceFolder, workspaceRealPath);
        continue;
      }
      if (childStat.isDirectory()) {
        if (["code", "specs", "assets"].some((reserved) => sameDirectoryName(name, reserved)))
          continue;
        await assertWorkspacePathHasNoLinks(child, workspaceFolder, workspaceRealPath);
        await visit(child);
      } else if (childStat.isFile() && lowerName.endsWith(".frm")) {
        await assertWorkspacePathHasNoLinks(child, workspaceFolder, workspaceRealPath);
        result.push(child);
      }
    }
  };
  await visit(root);
  return result;
}

async function assertWorkspacePathHasNoLinks(
  uri: vscode.Uri,
  workspaceFolder: vscode.Uri,
  workspaceRealPath: string,
): Promise<void> {
  const relative = path.relative(workspaceFolder.fsPath, uri.fsPath);
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative))
    throw new Error(
      vscode.l10n.t("The configured UserForm source directory must be inside the workspace."),
    );
  let actualPath: string;
  try {
    actualPath = await realpath(uri.fsPath);
  } catch {
    throw new Error(
      vscode.l10n.t("The UserForm code path cannot traverse symbolic links or junctions: {path}.", {
        path: uri.fsPath,
      }),
    );
  }
  const expectedPath = path.resolve(workspaceRealPath, relative);
  if (!sameFilesystemPath(actualPath, expectedPath))
    throw new Error(
      vscode.l10n.t("The UserForm code path cannot traverse symbolic links or junctions: {path}.", {
        path: uri.fsPath,
      }),
    );
}

function isWorkspacePathOrDescendant(candidate: string, workspaceRoot: string): boolean {
  const relative = path.relative(workspaceRoot, candidate);
  return (
    relative === "" ||
    (relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative))
  );
}

function sameFilesystemPath(left: string, right: string): boolean {
  const normalize = (value: string) => {
    let resolved = path.resolve(value);
    if (process.platform === "win32") {
      resolved = resolved.replace(/^\\\\\?\\UNC\\/i, "\\\\").replace(/^\\\\\?\\/, "");
      return resolved.toLowerCase();
    }
    return resolved;
  };
  return normalize(left) === normalize(right);
}

async function readDirectory(uri: vscode.Uri): Promise<[string, vscode.FileType][]> {
  try {
    return await vscode.workspace.fs.readDirectory(uri);
  } catch (error) {
    if (error instanceof vscode.FileSystemError && error.code === "FileNotFound") return [];
    throw error;
  }
}

function readVBName(text: string): string | undefined {
  return /^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$/im.exec(text)?.[1];
}

function basenameWithoutExtension(name: string): string {
  return name.replace(/\.[^.]*$/, "");
}

function isVbaName(name: string): boolean {
  const characters = Array.from(name);
  return (
    characters.length > 0 &&
    characters.length <= 255 &&
    /^\p{L}$/u.test(characters[0]) &&
    characters.slice(1).every((character) => /^[\p{L}\p{N}_]$/u.test(character))
  );
}

function sameDirectoryName(actual: string, configured: string): boolean {
  return process.platform === "win32"
    ? actual.toLowerCase() === configured.toLowerCase()
    : actual === configured;
}
