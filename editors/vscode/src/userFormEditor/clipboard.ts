import { randomUUID } from "crypto";
import * as vscode from "vscode";
import type {
  ClipboardSnapshot,
  CopyResult,
  DesignerDocument,
  DesignerError,
  ProductivityOperation,
} from "./protocol";
import { isDesignerFrame, isDesignerPage } from "./geometry";
import { controlMetadata, subtreeIds } from "./toolbox";

export interface ClipboardContent {
  snapshot: ClipboardSnapshot;
  document: DesignerDocument;
  projectKey: string;
  hasPicturePaths: boolean;
}

export function projectKey(document: vscode.TextDocument): string | undefined {
  const uri = vscode.workspace.getWorkspaceFolder(document.uri)?.uri;
  if (!uri) return;
  const key = uri.toString();
  return process.platform === "win32" ? key.toLowerCase() : key;
}

// Source snapshots stay in the extension host; Webviews only see availability.
export class DesignerClipboard implements vscode.Disposable {
  private content: ClipboardContent | undefined;
  private readonly changed = new vscode.EventEmitter<void>();
  readonly onDidChange = this.changed.event;
  private readonly repetitions = new Map<string, number>();
  private revision = 0;
  private copyRequest = 0;

  beginCopy(): number {
    return ++this.copyRequest;
  }

  isCurrentCopy(request: number): boolean {
    return request === this.copyRequest;
  }

  finishCopy(request: number, content: ClipboardContent): boolean {
    if (request !== this.copyRequest) return false;
    this.write(content);
    return true;
  }

  get generation(): number {
    return this.revision;
  }

  get available(): boolean {
    return this.content !== undefined;
  }

  read(): ClipboardContent | undefined {
    return this.content && structuredClone(this.content);
  }

  write(content: ClipboardContent): void {
    this.content = structuredClone(content);
    this.revision++;
    this.repetitions.clear();
    this.changed.fire();
  }

  nextOffset(uri: string, parentId: string): number {
    return 8 * ((this.repetitions.get(JSON.stringify([uri, parentId])) ?? 0) + 1);
  }

  pasted(uri: string, parentId: string, generation = this.revision): void {
    if (generation !== this.revision) return;
    const key = JSON.stringify([uri, parentId]);
    this.repetitions.set(key, (this.repetitions.get(key) ?? 0) + 1);
  }

  dispose(): void {
    this.copyRequest++;
    this.content = undefined;
    this.repetitions.clear();
    this.changed.dispose();
  }
}

export function validCopyResult(result: CopyResult, version: number): boolean {
  const snapshot = result.snapshot;
  return (
    result.version === version &&
    !result.error &&
    !!snapshot &&
    ["yaml", "json"].includes(snapshot.format) &&
    typeof snapshot.text === "string" &&
    Array.isArray(snapshot.controlIds) &&
    snapshot.controlIds.length > 0 &&
    new Set(snapshot.controlIds).size === snapshot.controlIds.length &&
    !!result.document &&
    typeof result.document.form?.name === "string" &&
    Array.isArray(result.document.controls) &&
    result.document.controls.every(
      (c) =>
        typeof c.id === "string" &&
        !!c.id &&
        typeof c.name === "string" &&
        typeof c.type === "string",
    ) &&
    snapshot.controlIds.every((id) => result.document!.controls.some((c) => c.id === id)) &&
    typeof result.hasPicturePaths === "boolean"
  );
}

export function siblingSelection(
  document: DesignerDocument,
  ids: readonly string[],
): DesignerError | undefined {
  const controls = ids.map((id) => document.controls.find((c) => c.id === id));
  if (
    !ids.length ||
    new Set(ids).size !== ids.length ||
    controls.some(
      (c) =>
        !c ||
        !controlMetadata.some((m) => m.type.toLowerCase() === c.type.toLowerCase()) ||
        (c.progId && c.progId.toLowerCase() !== `forms.${c.type.toLowerCase()}.1`),
    ) ||
    controls.some((c) => (c!.parentId ?? "") !== (controls[0]!.parentId ?? ""))
  )
    return {
      code: "invalidSelection",
      message: vscode.l10n.t("Select built-in controls with the same parent."),
    };
}

export function pasteOperation(
  content: ClipboardContent,
  target: DesignerDocument,
  targetProject: string,
  parentId: string,
  offset: number,
  pageEditable: boolean,
  uuid: () => string = randomUUID,
): Extract<ProductivityOperation, { type: "pasteControls" }> | DesignerError {
  if (content.hasPicturePaths && content.projectKey !== targetProject)
    return {
      code: "crossProjectPicture",
      message: vscode.l10n.t(
        "Controls with image paths can only be pasted within the same project.",
      ),
    };
  if (parentId) {
    const parent = target.controls.find((c) => c.id === parentId);
    if (!parent || (!isDesignerFrame(parent) && !(pageEditable && isDesignerPage(target, parent))))
      return {
        code: "invalidParent",
        message: vscode.l10n.t("This control cannot be placed in this container."),
      };
  }
  const names = new Set(
    [target.name, ...target.controls.map((c) => c.name)].map((n) => n.toLowerCase()),
  );
  const ids = new Set(target.controls.map((c) => c.id));
  const allocations: { sourceId: string; id: string; name: string }[] = [];
  for (const control of content.document.controls) {
    // Page is copied only as part of a MultiPage subtree.
    const prefix =
      controlMetadata.find((m) => m.type.toLowerCase() === control.type.toLowerCase())?.type ??
      (control.type.toLowerCase() === "page" ? "Page" : undefined);
    if (!prefix)
      return {
        code: "unsupportedCopy",
        message: vscode.l10n.t("Custom controls cannot be copied."),
      };
    let suffix = 1;
    while (names.has(`${prefix}${suffix}`.toLowerCase())) suffix++;
    const name = `${prefix}${suffix}`;
    let id: string | undefined;
    for (let attempt = 0; attempt < 100; attempt++) {
      const candidate = `control-${uuid()}`;
      if (!ids.has(candidate)) {
        id = candidate;
        break;
      }
    }
    if (!id)
      return {
        code: "controlConflict",
        message: vscode.l10n.t("A unique control ID could not be allocated."),
      };
    names.add(name.toLowerCase());
    ids.add(id);
    allocations.push({ sourceId: control.id, id, name });
  }
  return {
    type: "pasteControls",
    snapshot: content.snapshot,
    allocations,
    parentId,
    offset: { left: offset, top: offset },
  };
}

export async function prepareProductivityDeletion(
  operation: Extract<ProductivityOperation, { type: "removeControls" }>,
  document: DesignerDocument,
  confirm: (name: string, descendants: number) => PromiseLike<boolean>,
): Promise<DesignerError | undefined> {
  const invalid = siblingSelection(document, operation.controlIds);
  if (invalid) return invalid;
  const roots = new Set(operation.controlIds);
  const deleted = new Set<string>();
  for (const id of roots)
    for (const descendant of subtreeIds(document.controls, id)) deleted.add(descendant);
  const count = deleted.size - roots.size;
  if (count && !operation.cascade)
    return {
      code: "cascadeRequired",
      message: vscode.l10n.t("Deleting containers requires cascade."),
    };
  if (
    count &&
    !(await confirm(
      operation.controlIds.map((id) => document.controls.find((c) => c.id === id)!.name).join(", "),
      count,
    ))
  )
    return { code: "editCancelled", message: vscode.l10n.t("Control deletion cancelled.") };
}
