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

Command-specific fields are top-level fields such as `issues`, `analysis`, `macro`, `macros`, `tests`, `diff`, `inspect`, `trace`, `ui`, `backups`, `rollback`, `target`, `session`, `warnings`, `hints`, `output`, `forms`, `edit`, `runner`, and `version`. `output` carries `fmt` result summaries, `export-image` output paths, and `form` command artifacts.

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
