import * as vscode from "vscode";
import {
  LanguageClient,
  LanguageClientOptions,
  ServerOptions,
  TransportKind,
} from "vscode-languageclient/node";

let client: LanguageClient | undefined;

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
      void client?.stop();
      client = undefined;
    },
  });

  try {
    await client.start();
    outputChannel.info(`Started xlflow lsp --stdio${cwd === undefined ? "" : ` in ${cwd}`}`);
  } catch (error) {
    outputChannel.error(`Failed to start xlflow lsp --stdio: ${String(error)}`);
    throw error;
  }
}

export async function deactivate(): Promise<void> {
  const runningClient = client;
  client = undefined;
  await runningClient?.stop();
}
