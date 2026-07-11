import * as vscode from "vscode";

const unsupportedInlineSuppressionRules = new Set([
  "VB008",
  "VB009",
  "VB010",
  "VB011",
  "VB012",
  "VB013",
  "VB014",
  "VB028",
  "VB029",
  "VB031",
  "VB032",
  "VBA104",
  "VBA105",
  "VBA106",
  "VBA211",
]);

export function registerLineSuppressionCodeActions(): vscode.Disposable {
  return vscode.Disposable.from(
    vscode.languages.registerCodeActionsProvider(
      { language: "vba", scheme: "file" },
      new XlflowLineSuppressionCodeActionProvider(),
      {
        providedCodeActionKinds: [vscode.CodeActionKind.QuickFix],
      },
    ),
  );
}

export function registerDocumentationCodeActions(): vscode.Disposable {
  return vscode.Disposable.from(
    vscode.commands.registerCommand(
      "xlflow.generateDocumentationComment",
      async (uri: vscode.Uri, range: vscode.Range, snippet: string) => {
        const editor = vscode.window.activeTextEditor;
        if (editor === undefined || editor.document.uri.toString() !== uri.toString()) {
          return;
        }
        await editor.insertSnippet(new vscode.SnippetString(snippet), range);
      },
    ),
    vscode.languages.registerCodeActionsProvider(
      { language: "vba", scheme: "file" },
      new XlflowDocumentationCodeActionProvider(),
      {
        providedCodeActionKinds: [vscode.CodeActionKind.QuickFix],
      },
    ),
  );
}

class XlflowDocumentationCodeActionProvider implements vscode.CodeActionProvider {
  public provideCodeActions(
    document: vscode.TextDocument,
    range: vscode.Range,
  ): vscode.CodeAction[] {
    const action = documentationCommentAction(document, range.start.line);
    return action === undefined ? [] : [action];
  }
}

export function documentationCommentAction(
  document: vscode.TextDocument,
  line: number,
): vscode.CodeAction | undefined {
  const marker = docCommentMarkerRange(document, line);
  if (marker === undefined) {
    return undefined;
  }
  const procedure = nextProcedureDeclaration(document, line + 1);
  if (procedure === undefined) {
    return undefined;
  }
  const snippet = documentationSnippet(procedure);
  if (snippet === "") {
    return undefined;
  }
  const action = new vscode.CodeAction(
    vscode.l10n.t("Generate documentation comment for {0}", procedure.name),
    vscode.CodeActionKind.QuickFix,
  );
  action.isPreferred = true;
  action.command = {
    command: "xlflow.generateDocumentationComment",
    title: action.title,
    arguments: [document.uri, marker, snippet],
  };
  return action;
}

function docCommentMarkerRange(
  document: vscode.TextDocument,
  line: number,
): vscode.Range | undefined {
  if (line < 0 || line >= document.lineCount) {
    return undefined;
  }
  const text = document.lineAt(line).text;
  if (!/^\s*'''$/.test(text)) {
    return undefined;
  }
  return new vscode.Range(line, 0, line, text.length);
}

interface ProcedureDeclaration {
  name: string;
  kind: "sub" | "function" | "property_get" | "property_let" | "property_set";
  parameters: string[];
}

function nextProcedureDeclaration(
  document: vscode.TextDocument,
  startLine: number,
): ProcedureDeclaration | undefined {
  for (let line = startLine; line < document.lineCount; line++) {
    const logical = logicalDeclarationLine(document, line);
    const trimmed = logical.text.trim();
    if (
      trimmed.length === 0 ||
      /^'\s*@(?:ModuleDescription|Description|VariableDescription)\b/i.test(trimmed)
    ) {
      continue;
    }
    if (trimmed.startsWith("'")) {
      return undefined;
    }
    return parseProcedureDeclaration(logical.text);
  }
  return undefined;
}

function logicalDeclarationLine(
  document: vscode.TextDocument,
  startLine: number,
): { text: string; endLine: number } {
  const parts: string[] = [];
  let line = startLine;
  for (; line < document.lineCount; line++) {
    const text = document.lineAt(line).text;
    const withoutContinuation = text.replace(/\s+_\s*$/, "");
    parts.push(withoutContinuation);
    if (withoutContinuation === text) {
      break;
    }
  }
  return { text: parts.join(" "), endLine: line };
}

export function parseProcedureDeclaration(text: string): ProcedureDeclaration | undefined {
  const property =
    /^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Property)\s+(Get|Let|Set)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(([^)]*)\))?/i.exec(
      text,
    );
  if (property !== null) {
    const accessor = property[2].toLowerCase() as "get" | "let" | "set";
    return {
      name: property[3],
      kind: `property_${accessor}`,
      parameters: parseParameterNames(property[4] ?? ""),
    };
  }
  const procedure =
    /^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(([^)]*)\))?/i.exec(
      text,
    );
  if (procedure === null) {
    return undefined;
  }
  return {
    name: procedure[2],
    kind: procedure[1].toLowerCase() === "sub" ? "sub" : "function",
    parameters: parseParameterNames(procedure[3] ?? ""),
  };
}

function parseParameterNames(parameters: string): string[] {
  return parameters
    .split(",")
    .map((parameter) => {
      const left = parameter.split("=")[0]?.trim() ?? "";
      const withoutType = left.replace(/\s+As\s+.+$/i, "").trim();
      const words = withoutType
        .replace(/\([^)]*\)/g, "")
        .trim()
        .split(/\s+/)
        .filter((word) => !/^(Optional|ByVal|ByRef|ParamArray)$/i.test(word));
      return words.at(-1) ?? "";
    })
    .filter((name) => /^[A-Za-z_][A-Za-z0-9_]*$/.test(name));
}

export function documentationSnippet(procedure: ProcedureDeclaration): string {
  return documentationSnippetText(procedure);
}

function documentationSnippetText(procedure: ProcedureDeclaration): string {
  let index = 1;
  const lines = [
    `''' \${${index++}:${procedure.kind.startsWith("property_") ? "Property description." : "Summary."}}`,
  ];
  const args = procedure.kind === "property_get" ? [] : procedure.parameters;
  if (args.length > 0) {
    lines.push("'''", "''' Args:");
    for (const arg of args) {
      lines.push(
        `'''     ${arg}: \${${index++}:${propertySetterKind(procedure.kind) ? "Value description." : "Parameter description."}}`,
      );
    }
  }
  if (procedure.kind === "function" || procedure.kind === "property_get") {
    lines.push("'''", "''' Returns:");
    lines.push(
      `'''     \${${index++}:${procedure.kind === "property_get" ? "Returned property value description." : "Return value description."}}`,
    );
  }
  return lines.join("\n");
}

function propertySetterKind(kind: ProcedureDeclaration["kind"]): boolean {
  return kind === "property_let" || kind === "property_set";
}

class XlflowLineSuppressionCodeActionProvider implements vscode.CodeActionProvider {
  public provideCodeActions(
    document: vscode.TextDocument,
    _range: vscode.Range,
    context: vscode.CodeActionContext,
  ): vscode.CodeAction[] {
    const actions: vscode.CodeAction[] = [];
    const seen = new Set<string>();
    for (const diagnostic of context.diagnostics) {
      const code = diagnosticRuleCode(diagnostic);
      if (code === undefined) {
        continue;
      }
      const key = diagnosticActionKey(diagnostic, code);
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      const nextLineAction = disableNextLineAction(document, diagnostic, code);
      if (nextLineAction !== undefined) {
        actions.push(nextLineAction);
      }
      const lineAction = disableLineAction(document, diagnostic, code);
      if (lineAction !== undefined) {
        actions.push(lineAction);
      }
    }
    return actions;
  }
}

export function diagnosticRuleCode(diagnostic: vscode.Diagnostic): string | undefined {
  if (diagnostic.source !== "xlflow") {
    return undefined;
  }
  const raw = diagnostic.code;
  const code =
    typeof raw === "string" || typeof raw === "number"
      ? String(raw)
      : raw === undefined
        ? ""
        : String(raw.value);
  const normalized = code.trim().toUpperCase();
  if (!/^(?:VB|VBA)\d{3}$/.test(normalized)) {
    return undefined;
  }
  if (unsupportedInlineSuppressionRules.has(normalized)) {
    return undefined;
  }
  return normalized;
}

export function diagnosticActionKey(diagnostic: vscode.Diagnostic, code: string): string {
  const range = diagnostic.range;
  return [
    code,
    range.start.line,
    range.start.character,
    range.end.line,
    range.end.character,
    diagnostic.message,
  ].join(":");
}

function disableNextLineAction(
  document: vscode.TextDocument,
  diagnostic: vscode.Diagnostic,
  code: string,
): vscode.CodeAction | undefined {
  const line = diagnostic.range.start.line;
  const action = new vscode.CodeAction(
    vscode.l10n.t("Suppress {0} on the next line", code),
    vscode.CodeActionKind.QuickFix,
  );
  action.diagnostics = [diagnostic];
  action.isPreferred = true;
  action.edit = new vscode.WorkspaceEdit();
  const previous = existingDirectiveLine(document, line, "xlflow:disable-next-line");
  if (previous !== undefined) {
    if (directiveLineHasCode(document.lineAt(previous).text, "xlflow:disable-next-line", code)) {
      return undefined;
    }
    action.edit.insert(
      document.uri,
      new vscode.Position(previous, document.lineAt(previous).text.length),
      ` ${code}`,
    );
    return action;
  }
  action.edit.insert(
    document.uri,
    new vscode.Position(line, 0),
    disableNextLineComment(document.lineAt(line).text, code, document.eol),
  );
  return action;
}

function disableLineAction(
  document: vscode.TextDocument,
  diagnostic: vscode.Diagnostic,
  code: string,
): vscode.CodeAction | undefined {
  const line = diagnostic.range.start.line;
  const action = new vscode.CodeAction(
    vscode.l10n.t("Suppress {0} on this line", code),
    vscode.CodeActionKind.QuickFix,
  );
  action.diagnostics = [diagnostic];
  action.edit = new vscode.WorkspaceEdit();
  const text = document.lineAt(line).text;
  const suffix = disableLineSuffix(text, code);
  if (suffix === "") {
    return undefined;
  }
  action.edit.insert(document.uri, new vscode.Position(line, text.length), suffix);
  return action;
}

function existingDirectiveLine(
  document: vscode.TextDocument,
  targetLine: number,
  directive: "xlflow:disable-next-line",
): number | undefined {
  const previousLine = targetLine - 1;
  if (previousLine < 0) {
    return undefined;
  }
  const text = document.lineAt(previousLine).text;
  if (new RegExp(`'\\s*${directive}\\b`, "i").test(text)) {
    return previousLine;
  }
  return undefined;
}

export function disableNextLineComment(
  targetLineText: string,
  code: string,
  eol: vscode.EndOfLine,
): string {
  const indent = targetLineText.match(/^\s*/)?.[0] ?? "";
  const newline = eol === vscode.EndOfLine.CRLF ? "\r\n" : "\n";
  return `${indent}' xlflow:disable-next-line ${code}${newline}`;
}

export function disableLineSuffix(lineText: string, code: string): string {
  const existing = /'\s*xlflow:disable-line\b(.*)$/i.exec(lineText);
  if (existing !== null) {
    if (directiveCodesInclude(existing[1], code)) {
      return "";
    }
    return ` ${code}`;
  }
  const spacer = lineText.trimEnd().length === 0 ? "" : " ";
  return `${spacer}' xlflow:disable-line ${code}`;
}

function directiveLineHasCode(lineText: string, directive: string, code: string): boolean {
  const escaped = directive.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = new RegExp(`'\\s*${escaped}\\b(.*)$`, "i").exec(lineText);
  return match !== null && directiveCodesInclude(match[1], code);
}

function directiveCodesInclude(text: string, code: string): boolean {
  return text
    .trim()
    .split(/\s+/)
    .some((candidate) => candidate.toUpperCase() === code.toUpperCase());
}
