import * as vscode from "vscode";
import {
  LanguageClient,
  LanguageClientOptions,
  ServerOptions,
  TransportKind,
} from "vscode-languageclient/node";

let client: LanguageClient | undefined;
let suggestTimer: NodeJS.Timeout | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const outputChannel = vscode.window.createOutputChannel("xlflow LSP Dev Client", { log: true });
  const traceOutputChannel = vscode.window.createOutputChannel("xlflow LSP Dev Client Trace", {
    log: true,
  });
  const workspaceFolder = vscode.workspace.workspaceFolders?.[0];
  const cwd = workspaceFolder?.uri.fsPath;

  const serverOptions: ServerOptions = {
    command: "xlflow",
    args: ["lsp", "--stdio", "--log-file", ".xlflow/lsp.log"],
    transport: TransportKind.stdio,
    options: cwd === undefined ? undefined : { cwd },
  };

  const clientOptions: LanguageClientOptions = {
    documentSelector: [
      { scheme: "file", language: "vba" },
      { scheme: "file", pattern: "**/*.{bas,cls,frm}" },
    ],
    synchronize: {
      fileEvents: vscode.workspace.createFileSystemWatcher("**/*.{bas,cls,frm}"),
    },
    outputChannel,
    traceOutputChannel,
  };

  client = new LanguageClient(
    "xlflow-vscode-lsp-dev",
    "xlflow LSP Dev Client",
    serverOptions,
    clientOptions,
  );

  context.subscriptions.push(outputChannel, traceOutputChannel, {
    dispose: () => {
      clearPendingSuggest();
      void client?.stop();
      client = undefined;
    },
  });
  context.subscriptions.push(
    vscode.workspace.onDidChangeTextDocument((event) => {
      scheduleDeclarationSuggest(event.document);
    }),
  );

  try {
    await client.start();
    outputChannel.info(`Started xlflow lsp --stdio${cwd === undefined ? "" : ` in ${cwd}`}`);
  } catch (error) {
    outputChannel.error(`Failed to start xlflow lsp --stdio: ${String(error)}`);
    throw error;
  }
}

export async function deactivate(): Promise<void> {
  clearPendingSuggest();
  const runningClient = client;
  client = undefined;
  await runningClient?.stop();
}

function scheduleDeclarationSuggest(document: vscode.TextDocument): void {
  const editor = vscode.window.activeTextEditor;
  if (editor === undefined || editor.document !== document || document.languageId !== "vba") {
    return;
  }
  const position = editor.selection.active;
  const linePrefix = document.lineAt(position.line).text.slice(0, position.character);
  if (!isDeclarationPrefix(linePrefix)) {
    return;
  }

  clearPendingSuggest();
  suggestTimer = setTimeout(() => {
    suggestTimer = undefined;
    void vscode.commands.executeCommand("editor.action.triggerSuggest");
  }, 75);
}

function clearPendingSuggest(): void {
  if (suggestTimer !== undefined) {
    clearTimeout(suggestTimer);
    suggestTimer = undefined;
  }
}

function isDeclarationPrefix(linePrefix: string): boolean {
  const typed = linePrefix.trimStart();
  if (typed.length === 0 || /[."'():=]/.test(typed)) {
    return false;
  }
  return /^(o|op|opt|opti|optio|option|option\s+\w*|p|pu|pub|publ|publi|public|public\s+\w*|pr|pri|priv|priva|privat|private|private\s+\w*|f|fr|fri|frie|frien|friend|friend\s+\w*|s|su|sub|fu|fun|func|funct|functi|functio|function|d|di|dim|c|co|con|cons|const|t|ty|typ|type|e|en|enu|enum|declare|declare\s+\w*)$/i.test(
    typed,
  );
}
