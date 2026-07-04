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
  return vscode.languages.registerCodeActionsProvider(
    { language: "vba", scheme: "file" },
    new XlflowLineSuppressionCodeActionProvider(),
    {
      providedCodeActionKinds: [vscode.CodeActionKind.QuickFix],
    },
  );
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
