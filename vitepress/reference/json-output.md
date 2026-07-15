# JSON Output

All commands accept the global `--json` flag.

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

Failures set `status` to `failed` and populate `error`:

```json
{
  "status": "failed",
  "command": "run",
  "error": {
    "code": "macro_failed",
    "message": "Main Err 5: inputPath is required",
    "phase": "invoke_macro"
  },
  "logs": []
}
```

When another xlflow process is using the same workbook, workbook-bound commands
fail immediately with `error.code: "workbook_busy"`. `error.details` includes
`workbook`, `operation`, `resource_scope`, and `retryable`. It may also include
diagnostic `owner` metadata such as `pid`, stable command name, operation kind,
and start time. Owner metadata can be missing after a crash or during cleanup;
the error code, not the presence of `owner`, is the stable contention signal.

Retryable workbook commands may add global `--wait`. Waiting defaults to 30
seconds and can be changed with `--wait-timeout`. Timeout and cancellation use
`workbook_busy_timeout` and `workbook_busy_cancelled`; their details include
`wait_timeout`. JSON mode never writes wait progress to stderr or mixes it into
stdout, so the result remains one valid envelope.

Unknown commands are also structured when `--json` appears before the invalid command:

```json
{
  "status": "failed",
  "command": "xlflow",
  "error": {
    "code": "unknown_command",
    "message": "unknown command \"pussh\"",
    "suggestions": ["push"]
  },
  "logs": []
}
```

Command-specific fields are top-level fields such as `issues`, `analysis`, `macro`, `macros`, `tests`, `diff`, `inspect`, `ui`, `debug`, `backups`, `backup_prune`, `rollback`, `target`, `session`, `coordination`, `warnings`, `hints`, `output`, `forms`, `edit`, `runner`, and `version`. `output` carries `fmt` result summaries, `export-image` output paths, and `form` command artifacts.

`session status` adds top-level `coordination` without changing `session`.
Successful observation returns `{"busy":false}`, `{"busy":true}`, or a busy
object with `resource_scope`, `operation_kind`, `command`, `pid`, and
`started_at`. The value is observational command-start state. Probe failure
omits the field and adds warning `coordination_status_unavailable`.

`backup_prune` is returned by `xlflow backup prune` and may also appear on successful `push` or `rollback` when automatic retention deleted entries, skipped invalid or legacy entries, or encountered a pruning failure. Automatic results include `"automatic": true`; pruning failures are warnings and do not change the successful workbook operation status.

`xlflow run --json` uses a compact failure payload by default. This keeps the fields that are usually relevant for fixing user VBA code and hides xlflow-internal diagnostic detail unless `--verbose` is specified.

Default failure example:

```json
{
  "status": "failed",
  "command": "run",
  "error": {
    "code": "macro_failed",
    "message": "実行時エラー '9':\n\nインデックスが有効範囲にありません。",
    "number": 9,
    "phase": "invoke_macro"
  },
  "macro": {
    "name": "Main.Run",
    "duration_ms": 1115
  },
  "location": {
    "source_path": "src/modules/QRCode.bas",
    "component": "QRCode",
    "component_type": "module",
    "procedure": "AddErrorCorrection",
    "line": 449,
    "end_line": 449,
    "text": "        dividend(i + j + 1) = dividend(i + j + 1) Xor genCoef",
    "confidence": "high",
    "method": "vbe.selection"
  },
  "session": {
    "active": true,
    "mode": "explicit",
    "dirty": true,
    "save_required": true,
    "source_of_truth": "live_workbook",
    "workbook_path": "C:\\dev\\test\\QRCode\\build\\Book.xlsm"
  },
  "target": {
    "kind": "live_session",
    "path": "C:\\dev\\test\\QRCode\\build\\Book.xlsm"
  },
  "warnings": [
    {
      "code": "save_required",
      "message": "The live workbook is newer than disk. Run xlflow save --session to persist workbook changes."
    }
  ],
  "suggestion": "Inspect src/modules/QRCode.bas:449 in AddErrorCorrection. Add targeted XlflowDebug.Log calls around the failing block and rerun."
}
```

Use `xlflow run --json --verbose` when you need xlflow-internal diagnostics such as full `run_diagnostic`, workbook/bridge/runtime metadata, dialog snapshots, or location-capture attempt details for bug reports or dialog-watcher debugging.

`run --ui-stream` and `test --ui-stream` may add a top-level `ui` object when `XlflowUI` dialogs are resolved. The stable field is `ui.events`, where each event may contain keys such as `kind`, `dialog_id`, `response_source`, `resolved_result`, `resolved_value`, `redacted`, and `error`.

Example:

```json
{
  "status": "ok",
  "command": "run",
  "ui": {
    "events": [
      {
        "kind": "msgbox",
        "dialog_id": "confirm-save",
        "response_source": "default",
        "resolved_result": "yes",
        "redacted": false
      },
      {
        "kind": "inputbox",
        "dialog_id": "customer-name",
        "response_source": "default",
        "resolved_value": "[redacted]",
        "redacted": true
      }
    ]
  }
}
```

When `--ui-stream` is enabled, xlflow also writes realtime `XlflowUI` summaries to stderr. Those streamed lines are not part of stdout JSON.

`fmt --json` returns `output` with `changed`, `unchanged`, `skipped`, and `total` summary fields. `fmt --stdin --json` returns the same envelope shape instead of formatted text; the formatted source body is not included in the JSON output.

Example:

```json
{
  "status": "ok",
  "command": "fmt",
  "output": {
    "changed": 2,
    "unchanged": 5,
    "skipped": 1,
    "total": 8
  },
  "logs": []
}
```

Source: `docs/specs/cli-contract.md`.
