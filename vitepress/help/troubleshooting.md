# Troubleshooting

Start with `xlflow version --verbose --json`, then `xlflow doctor --json` and `xlflow status --json`. Keep the JSON output and exit code when reporting a failure.

## Installation

### `xlflow` is not found

**Symptoms:** PowerShell, VS Code, or WSL reports that `xlflow`/`xlflow.exe` is not a command.  
**Likely cause:** the installer directory is not on PATH, or only the Linux frontend/Go binary was installed.  
**Diagnose:** run `Get-Command xlflow`, `where.exe xlflow`, or `command -v xlflow`; then run `xlflow version`.  
**Fix:** reopen the shell, add the install directory to PATH, or set the VS Code `xlflow.path`. Install the Windows release archive when Excel operations need the `.NET` bridge.  
**Verify:** `xlflow version --json` succeeds.

### Excel is not detected

Run `xlflow doctor --json`. Install or repair desktop Excel, run the command on Windows, and ensure the process is not blocked by policy. Source-only `lint`, `analyze`, and `fmt` can still run without Excel.

### VBIDE access is denied

Enable Excel Trust Center → Macro Settings → **Trust access to the VBA project object model**, restart Excel, and rerun `doctor`. The related code is `vbide_access_denied`.

## Project setup

### `xlflow.toml` is not found or the workbook path is invalid

Run commands from the directory containing `xlflow.toml`, or use `xlflow new`/`xlflow init` first. Check the configured `[workbook].path` and that the extension is `.xlsm`, `.xlam`, or `.xlsb`.

### Source and workbook are out of sync

**Symptoms:** edits disappear, `status` reports a newer workbook/source, or push would overwrite changes.  
**Diagnose:** run `xlflow status --json` and inspect `source_of_truth`, `src_newer_than_workbook`, and session dirty metadata.  
**Fix:** workbook authoritative → `xlflow pull`; source authoritative → `xlflow push`; unknown → create a backup and inspect before overwriting.  
**Verify:** rerun `status` and confirm the intended side is current.

## Execution

### Macro is not found

Run `xlflow macros --json` and use the returned `qualified_name`. A private or argument-bearing procedure may not be runnable through the selected path; use a public no-argument entry point or a test wrapper. Related code: `macro_not_found`.

### Source changes are not reflected

Confirm the edited file is under a configured source root, run `lint`, then `push`. If using a session, include `--session`; if using `--no-save`, inspect the live workbook with `--session` and save explicitly.

### Compile error or `0x800A03EC`

Run `xlflow lint --json`, `xlflow analyze --json`, and `xlflow run --diagnostic`. Fix the reported source location or preflight finding before opening Excel. The diagnostic runner is preferred for agents; do not leave a compile dialog unattended.

### A dialog blocks execution or the macro times out

Use `--diagnostic`/headless mode and xlflow dialog mappings for supported UI. Native `MsgBox`/`InputBox` calls should use the XlflowUI wrappers documented in the UI reference. For `macro_timeout`, inspect `status`, check for a UserForm/file picker, and stop or recover the session only after confirming whether Excel is still executing.

## Sessions

### A session is already running or remains dirty

Run `xlflow session status --json`. Reuse the matching session with `--session`, save with `xlflow save --session`, or intentionally discard a managed recovery session with `xlflow session stop --discard`. Never discard an external user-owned workbook automatically.

### Excel remains after stopping

Managed sessions should close their owned Excel process. External sessions only detach. Use `xlflow process list --json`, then `xlflow process cleanup <pid> --json` only when ownership and safety are clear.

### Recovery is required

`workbook_recovery_required` is not lock contention and cannot be fixed with `--wait`. Follow [recovery](../commands/recovery): inspect status, stop/discard or close Excel without saving, clear the verified marker, and confirm status before resuming.

## WSL

### Windows xlflow cannot be found

Install the Windows CLI and bridge, verify Windows PATH interoperability, or set `XLFLOW_WINDOWS_EXE` to the Windows path. Re-run `xlflow doctor --json` from WSL.

### Project path cannot be translated

Move the project under `/mnt/c`, `/mnt/d`, or another Windows-mounted drive. `/home/...` paths are supported for source-only commands but not Excel delegation. Related codes: `windows_xlflow_not_found`, `wsl_project_path_unsupported`, `wsl_path_translation_failed`, and `windows_xlflow_execution_failed`.

## Related references

- [Error codes](../reference/error-codes)
- [Installation](../installation)
- [Sessions and recovery](../guides/sessions)
- [VS Code troubleshooting](../vscode/troubleshooting)
