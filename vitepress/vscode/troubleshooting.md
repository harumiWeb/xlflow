# VS Code troubleshooting

Start by opening the folder containing `xlflow.toml`, then use **View → Output → xlflow** to retain the extension log. Each problem below gives the quickest check before you reinstall anything.

## The extension cannot find xlflow

**Symptoms:** the activity bar is absent or commands report that the CLI is unavailable.  
**Likely cause:** PATH is stale or `xlflow.path` points to the wrong executable.  
**Diagnose:** run `xlflow version` in the same environment and `xlflow: Check Environment`.  
**Fix:** set an absolute `xlflow.path`, reload VS Code, and ensure the Windows bridge is installed.  
**Verify:** `xlflow: Retry CLI Detection` succeeds.

## The language server does not start

**Symptoms:** no completion, diagnostics, or navigation; the output channel mentions startup.  
**Diagnose:** enable `xlflow.lsp.trace.server`, inspect the Output channel, and run `xlflow lsp` manually.  
**Fix:** confirm the opened folder contains `xlflow.toml` and that the CLI version supports LSP. Update the CLI and extension together when their versions differ.  
**Verify:** reopen a `.bas` file and confirm that its Problems entries have source `xlflow`.

## Completion or diagnostics are missing

**Symptoms:** the extension loads, but no VBA suggestions or Problems appear.  
**Diagnose:** check `xlflow.lsp.enabled`, file language mode, and the server log.  
**Fix:** restart the language server after changing the path or type database. Unsaved editor content is authoritative until the document closes; CLI lint reads files from disk.  
**Verify:** type a known VBA keyword in a `.bas` file or open a known diagnostic line.

## Sidebar actions fail

**Symptoms:** Pull, Push, Run, or session actions show an error.  
**Diagnose and fix:** run `xlflow status --json` and `xlflow doctor --json` from the project directory. Resolve `vbide_access_denied`, workbook path, session, or recovery errors before retrying a mutating action.  
**Verify:** run the matching CLI command first; when it works, retry the sidebar action.
