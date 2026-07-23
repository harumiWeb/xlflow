# Configure VS Code

Install the extension, open the folder containing `xlflow.toml`, and configure `xlflow.path` only when the executable cannot be found on `PATH`.

Useful settings include `xlflow.lsp.enabled`, `xlflow.lsp.logFile`, `xlflow.lsp.trace.server`, `xlflow.lsp.performanceLogging`, `xlflow.codeLens.enabled`, `xlflow.run.saveBeforeRun`, and `xlflow.testing.autoDiscover`. Restart the language server after changing the executable path or when an old process is still attached.

The [VS Code section](../vscode/) describes the user-visible behavior and troubleshooting steps.
