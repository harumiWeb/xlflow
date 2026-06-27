import * as vscode from "vscode";

export interface XlflowChannels {
  output: vscode.LogOutputChannel;
  trace: vscode.LogOutputChannel;
}

export function createChannels(): XlflowChannels {
  return {
    output: vscode.window.createOutputChannel("xlflow", { log: true }),
    trace: vscode.window.createOutputChannel(vscode.l10n.t("xlflow Language Server Trace"), {
      log: true,
    }),
  };
}

export function appendProcessOutput(
  channel: vscode.OutputChannel,
  stream: string,
  data: Buffer,
): void {
  const text = data.toString();
  if (text.length === 0) {
    return;
  }
  channel.append(`[${stream}] `);
  channel.append(text);
  if (!text.endsWith("\n")) {
    channel.appendLine("");
  }
}
