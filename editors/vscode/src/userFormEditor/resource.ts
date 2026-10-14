import * as vscode from "vscode";

export const designerResourceScheme = "xlflow-userform-designer";

/** Native Undo is resource-scoped; keep Designer history off the YAML resource. */
export function designerResourceUri(source: vscode.Uri): vscode.Uri {
  if (source.scheme !== "file") throw new Error("The Designer requires a file-backed FormSpec.");
  return vscode.Uri.from({
    scheme: designerResourceScheme,
    authority: source.authority,
    path: source.path,
    query: `source=${encodeURIComponent(source.toString())}`,
  });
}

export function designerSourceUri(resource: vscode.Uri): vscode.Uri {
  if (resource.scheme !== designerResourceScheme) return resource;
  const query = new URLSearchParams(resource.query);
  const values = query.getAll("source");
  if (values.length !== 1 || [...query.keys()].some((key) => key !== "source"))
    throw new Error("Invalid Designer source URI.");
  const source = vscode.Uri.parse(values[0]);
  if (
    source.scheme !== "file" ||
    source.path !== resource.path ||
    source.authority !== resource.authority
  )
    throw new Error("Invalid Designer source URI.");
  return source;
}
