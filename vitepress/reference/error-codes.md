# Error Codes

Stable error codes appear in `error.code` and command-specific findings.

Common examples:

- `workbook_busy`: another xlflow process holds the target workbook's
  coordination lock. The command fails immediately by default; JSON details
  include the workbook, attempted operation, and whether retrying is supported.
  Current owner metadata is included when available but is not authoritative.
- `workbook_busy_timeout`: an explicit `--wait` did not acquire every required
  workbook lock within the 30-second default or supplied `--wait-timeout`.
- `workbook_busy_cancelled`: Ctrl+C or caller cancellation stopped an explicit
  workbook wait before the command body started.
- `workbook_recovery_required`: the normal OS lock was acquired, but a previous
  operation left the workbook in an uncertain Excel/VBA state. The command body
  does not start, the failure is not retryable, and `--wait` will not recover
  it. Follow `error.details.recovery_actions`.
- `workbook_recovery_verification_failed`: normal `recovery clear` could not
  prove that the recorded Excel PID no longer exists, or the marker did not
  contain verifiable process information. The marker remains.
- `workbook_recovery_publication_failed`: xlflow detected an uncertain
  termination but could not atomically publish its recovery marker. Stop or
  close Excel manually before attempting more workbook work.
- `coordination_recovery_check_failed`: xlflow could not safely read recovery
  state after acquiring the workbook lock. Unsafe operations fail closed.
- `coordination_wait_args_invalid`: `--wait-timeout` was used without `--wait`,
  or the timeout was not positive.
- `coordination_wait_unsupported`: the selected command is not a retryable,
  non-parallel-safe workbook operation.
- `coordination_status_unavailable` (warning): `session status` could not probe
  workbook coordination state. Existing session status fields and exit behavior
  remain available, while the top-level `coordination` field is omitted.
- `vbide_access_denied`
- `macro_failed`
- `macro_not_found`
- `runner_not_invocable`
- `macro_timeout`
- `vba_compile_failed`
- `gui_boundary_detected`
- `source_preflight_failed`
- `output_file_exists`
- `unsupported_image_format`
- `form_not_found`
- `FRM202`
- `spec_validation_failed`
- `fmt_failed`
- `fmt_check_failed`
- `fmt_skipped_unsupported_extension`
- `fmt_args_invalid`
- `fmt_stdin_read_failed`
- `fmt_stdin_write_failed`
- `windows_xlflow_not_found`
- `wsl_project_path_unsupported`
- `wsl_path_translation_failed`
- `windows_xlflow_execution_failed`

Lint codes include `VB001` through `VB015`, `VB018` through `VB023`, `VB026` through `VB029`, `VB031`, `VB032`, and `VB044`. Analyzer codes include `VBA101` through `VBA106` and runtime-risk findings `VBA201` through `VBA212`.

## Recovery map

| Code or symptom                                                | Likely cause                                           | Recovery                                                                                        |
| -------------------------------------------------------------- | ------------------------------------------------------ | ----------------------------------------------------------------------------------------------- |
| `vbide_access_denied`                                          | Excel Trust Center blocks VBIDE automation.            | Enable Trust access, restart Excel, rerun `doctor`.                                             |
| `macro_not_found`                                              | The target name is not a runnable qualified procedure. | Run `macros --json` and use `qualified_name`.                                                   |
| `source_preflight_failed`                                      | Source would trigger a VBE compile/dialog failure.     | Fix the reported source issue, then rerun `lint` and `push`.                                    |
| `vba_compile_failed` / `0x800A03EC`                            | Excel rejected imported VBA or a workbook operation.   | Use `run --diagnostic`, inspect the source location, and retry only after preflight passes.     |
| `macro_timeout`                                                | A macro loop or modal UI did not complete.             | Inspect dialogs and session state; use interactive mode only with a human.                      |
| `workbook_recovery_required`                                   | Excel/VBA termination was not proven safe.             | Follow [recovery](../commands/recovery); `--wait` cannot clear quarantine.                      |
| `windows_xlflow_not_found`                                     | WSL cannot locate the Windows frontend.                | Install it or set `XLFLOW_WINDOWS_EXE`; see [WSL troubleshooting](../help/troubleshooting#wsl). |
| `wsl_project_path_unsupported` / `wsl_path_translation_failed` | The project is outside a Windows-mounted path.         | Move it under `/mnt/c`, `/mnt/d`, or another mounted drive.                                     |

The [symptom-oriented troubleshooting guide](../help/troubleshooting) adds diagnosis and verification steps for each category.
