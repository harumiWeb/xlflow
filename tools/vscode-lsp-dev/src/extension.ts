import * as vscode from "vscode";
import * as crypto from "crypto";
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
  const logFile = await resolveServerLogFile(context, workspaceFolder);
  const serverArgs = ["lsp", "--stdio", "--log-file", logFile];

  const serverOptions: ServerOptions = {
    command: "xlflow",
    args: serverArgs,
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
    outputChannel.info(
      `Started xlflow lsp --stdio${cwd === undefined ? "" : ` in ${cwd}`} with log file ${logFile}`,
    );
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

async function resolveServerLogFile(
  context: vscode.ExtensionContext,
  workspaceFolder: vscode.WorkspaceFolder | undefined,
): Promise<string> {
  if (workspaceFolder !== undefined && (await workspaceHasXlflowToml(workspaceFolder))) {
    return ".xlflow/lsp.log";
  }

  await vscode.workspace.fs.createDirectory(context.logUri);
  const workspaceKey = workspaceFolder?.uri.toString() ?? "no-workspace";
  const workspaceHash = crypto.createHash("sha256").update(workspaceKey).digest("hex").slice(0, 12);
  return vscode.Uri.joinPath(context.logUri, `xlflow-lsp-${workspaceHash}.log`).fsPath;
}

async function workspaceHasXlflowToml(workspaceFolder: vscode.WorkspaceFolder): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(vscode.Uri.joinPath(workspaceFolder.uri, "xlflow.toml"));
    return true;
  } catch {
    return false;
  }
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
